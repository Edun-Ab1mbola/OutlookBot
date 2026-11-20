"""
Hotmail / Outlook Telegram Account & OTP Manager Bot (IMAP & OAuth2 Version)
============================================================================
Requirements:
    pip install python-telegram-bot python-dotenv requests

Setup (.env):
    BOT_TOKEN=your_telegram_bot_token
    ADMIN_ID=your_admin_telegram_id_or_ids_comma_separated
"""

import os
import re
import asyncio
import imaplib
import email
import sqlite3
import logging
from functools import wraps
import urllib.request
import urllib.parse
import http.cookiejar
import requests
import msal
from datetime import datetime, timedelta
from dataclasses import dataclass
from dotenv import load_dotenv
from telegram import Update

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException as SeleniumTimeoutException, WebDriverException
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        WEBDRIVER_MANAGER_AVAILABLE = True
    except ImportError:
        WEBDRIVER_MANAGER_AVAILABLE = False
    SELENIUM_AVAILABLE = True
except ImportError:
    WEBDRIVER_MANAGER_AVAILABLE = False
    SELENIUM_AVAILABLE = False

try:
    from playwright.async_api import async_playwright, TimeoutError as AsyncPlaywrightTimeoutError
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters
)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
ADMIN_IDS_RAW = os.getenv("ADMIN_ID", "")
PROXY_URL = (os.getenv("PROXY_URL") or "").strip()
BROWSER_PROXY_URL = (os.getenv("BROWSER_PROXY_URL") or PROXY_URL).strip()

# Support multiple comma-separated admin IDs
ADMIN_IDS = {int(aid.strip()) for aid in ADMIN_IDS_RAW.split(",") if aid.strip().isdigit()}

if not BOT_TOKEN:
    raise ValueError("BOT_TOKEN is missing in .env")

DB_FILE = "hotmail_bot_accounts.db"
GRAPH_INBOX_URL = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
DEFAULT_AUTHORITY = "https://login.microsoftonline.com/consumers"
DEFAULT_SCOPES = ["https://graph.microsoft.com/Mail.Read"]
OUTLOOK_SIGNIN_URL = "https://login.live.com/"
BROWSER_SIGNIN_URLS = [
    "https://login.live.com/",
    "https://outlook.live.com/owa/",
    "https://account.microsoft.com/",
]


def _escape_md(text: str) -> str:
    """Escapes basic Markdown characters for Telegram legacy Markdown messages."""
    for ch in ("\\", "_", "*", "`", "["):
        text = text.replace(ch, f"\\{ch}")
    return text


def _requests_proxies() -> dict[str, str] | None:
    """Returns a requests-compatible proxy mapping when PROXY_URL is configured."""
    if not PROXY_URL:
        return None
    return {"http": PROXY_URL, "https": PROXY_URL}


def _playwright_proxy_config() -> dict[str, str] | None:
    """Returns a Playwright proxy configuration dictionary when a proxy URL is set."""
    if not BROWSER_PROXY_URL:
        return None
    return {"server": BROWSER_PROXY_URL}


def _create_selenium_driver(user_agent: str, headless: bool = True):
    """Creates and configures a Selenium Chrome WebDriver instance."""
    options = ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--window-size=1280,800")
    options.add_argument(f"--user-agent={user_agent}")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    if BROWSER_PROXY_URL:
        options.add_argument(f"--proxy-server={BROWSER_PROXY_URL}")

    if WEBDRIVER_MANAGER_AVAILABLE:
        try:
            service = ChromeService(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
        except Exception as wdm_err:
            logger.warning(f"webdriver-manager driver install failed, falling back to default Chrome driver: {wdm_err}")
            driver = webdriver.Chrome(options=options)
    else:
        driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(35)
    driver.implicitly_wait(2)
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"},
        )
    except Exception:
        pass
    return driver


def _selenium_find_visible(driver, selectors: list[tuple[By, str]], timeout_sec: int = 12):
    """Waits up to timeout_sec for the first visible element matching any (By, selector) pair."""
    deadline = datetime.utcnow() + timedelta(seconds=timeout_sec)
    while datetime.utcnow() < deadline:
        for by_type, selector in selectors:
            try:
                elements = driver.find_elements(by_type, selector)
                for el in elements:
                    if el.is_displayed() and el.is_enabled():
                        return el
            except Exception:
                continue
        import time
        time.sleep(0.4)
    return None


def _fetch_otp_via_selenium_sync(email_addr: str, password_str: str, user_agent: str, headless: bool = True) -> str:
    """Synchronous Selenium WebDriver flow for signing in and checking for an OTP code."""
    import time

    driver = None
    try:
        driver = _create_selenium_driver(user_agent=user_agent, headless=headless)

        last_nav_error = None
        for candidate_url in BROWSER_SIGNIN_URLS:
            try:
                driver.get(candidate_url)
                time.sleep(1.5)
                body_el = driver.find_elements(By.TAG_NAME, "body")
                current_txt = (body_el[0].text if body_el else "").lower()
                if "error" not in (driver.current_url or "").lower() and any(
                    k in current_txt for k in ["sign in", "email", "phone", "skype", "password", "next"]
                ):
                    break
            except Exception as nav_exc:
                last_nav_error = nav_exc
                continue
        else:
            if last_nav_error and not driver.current_url:
                raise last_nav_error

        # Click account tile or "Use another account" if shown
        try:
            acc_tiles = driver.find_elements(By.XPATH, f"//*[self::div or self::button][contains(text(), '{email_addr}')]")
            clicked_existing = False
            for tile in acc_tiles:
                if tile.is_displayed():
                    tile.click()
                    time.sleep(1.2)
                    clicked_existing = True
                    break
            if not clicked_existing:
                other_tiles = driver.find_elements(
                    By.XPATH,
                    "//*[self::div or self::button][contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'use another account')]"
                )
                for tile in other_tiles:
                    if tile.is_displayed():
                        tile.click()
                        time.sleep(1.2)
                        break
        except Exception:
            pass

        email_input = _selenium_find_visible(
            driver,
            [
                (By.CSS_SELECTOR, "input[type='email']"),
                (By.NAME, "loginfmt"),
                (By.NAME, "identifier"),
            ],
            timeout_sec=12,
        )
        if email_input is not None:
            email_input.clear()
            email_input.send_keys(email_addr)
            submit_btn = _selenium_find_visible(
                driver,
                [
                    (By.CSS_SELECTOR, "input[type='submit']"),
                    (By.CSS_SELECTOR, "button[type='submit']"),
                    (By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'next')]"),
                ],
                timeout_sec=6,
            )
            if submit_btn is not None:
                submit_btn.click()
            else:
                email_input.send_keys(Keys.RETURN)
            time.sleep(1.5)
        else:
            logger.info(f"Selenium: email input not immediately visible for {email_addr}; checking next step.")

        # Switch to password option if Microsoft offers alternative buttons
        for alt_text in ["use your password", "enter my password", "password"]:
            try:
                alt_btns = driver.find_elements(
                    By.XPATH,
                    f"//*[self::button or self::a or self::span][contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{alt_text}')]"
                )
                clicked_alt = False
                for btn in alt_btns:
                    if btn.is_displayed():
                        btn.click()
                        time.sleep(1.2)
                        clicked_alt = True
                        break
                if clicked_alt:
                    break
            except Exception:
                continue

        password_input = _selenium_find_visible(
            driver,
            [
                (By.CSS_SELECTOR, "input[type='password']"),
                (By.NAME, "passwd"),
                (By.XPATH, "//input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'passwd')]"),
                (By.XPATH, "//input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'password')]"),
            ],
            timeout_sec=18,
        )

        if password_input is not None:
            password_input.clear()
            password_input.send_keys(password_str.strip())
            pwd_submit = _selenium_find_visible(
                driver,
                [
                    (By.CSS_SELECTOR, "input[type='submit']"),
                    (By.CSS_SELECTOR, "button[type='submit']"),
                    (By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sign in')]"),
                    (By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'next')]"),
                ],
                timeout_sec=6,
            )
            if pwd_submit is not None:
                pwd_submit.click()
            else:
                password_input.send_keys(Keys.RETURN)
        else:
            body_el = driver.find_elements(By.TAG_NAME, "body")
            current_body = body_el[0].text if body_el else ""
            lower_current = current_body.lower()
            current_url = driver.current_url

            if "password sign-in isn't available" in lower_current or "try another method" in lower_current:
                try:
                    try_other_btns = driver.find_elements(
                        By.XPATH,
                        "//*[self::button or self::a or self::div][contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'try another method')]"
                    )
                    for btn in try_other_btns:
                        if btn.is_displayed():
                            btn.click()
                            time.sleep(1.5)
                            body_after = driver.find_elements(By.TAG_NAME, "body")
                            if body_after:
                                current_body = body_after[0].text
                            break
                except Exception:
                    pass
                alt_snippet = " ".join(current_body.split())[:240]
                return f"⚠️ Password sign-in is disabled for `{email_addr}`. Alternative screen: {alt_snippet or 'none'}"

            if any(k in lower_current for k in ["verify your identity", "help us protect", "enter the code", "approve sign in"]):
                return "⚠️ Browser sign-in triggered a security/2FA verification challenge before password entry."
            if any(k in lower_current for k in ["couldn't find an account", "doesn't exist", "enter a valid email", "account doesn't exist"]):
                return f"❌ Microsoft reported that `{email_addr}` was not found or invalid."
            if any(k in lower_current for k in ["other ways to sign in", "sign in another way", "passkey", "windows hello"]):
                return "⚠️ Microsoft prompted for passkey / alternative sign-in instead of a password field."

            snippet = " ".join(current_body.split())[:220]
            return f"⏱️ Password prompt did not appear in time (URL: `{current_url}`). Page text: {snippet or 'empty'}"

        # Handle secondary prompts ("Stay signed in?", "Skip for now", etc.)
        prompt_xpaths = [
            "//input[@id='idBtn_Back']",
            "//button[@id='idBtn_Back']",
            "//button[normalize-space()='No']",
            "//input[@type='submit' and @value='Yes']",
            "//*[self::button or self::a][contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'skip for now')]",
        ]
        for xp in prompt_xpaths:
            try:
                btns = driver.find_elements(By.XPATH, xp)
                clicked_prompt = False
                for b in btns:
                    if b.is_displayed():
                        b.click()
                        time.sleep(1.5)
                        clicked_prompt = True
                        break
                if clicked_prompt:
                    break
            except Exception:
                continue

        body_text = ""
        for _ in range(10):
            time.sleep(2.0)
            body_el = driver.find_elements(By.TAG_NAME, "body")
            body_text = body_el[0].text if body_el else ""
            lower_body = body_text.lower()
            if any(phrase in lower_body for phrase in [
                "verify your identity",
                "help us protect your account",
                "enter the code sent to",
                "approve sign in request"
            ]):
                return "⚠️ Browser sign-in triggered a security/2FA verification challenge."
            if "incorrect account or password" in lower_body:
                return f"❌ Microsoft returned 'Incorrect account or password' for `{email_addr}`."
            if "focused" in lower_body or "other" in lower_body or "inbox" in lower_body or re.search(r'\b\d{4,8}\b', body_text):
                break

        otp_match = re.search(r'\b\d{4,8}\b', body_text)
        if otp_match:
            return f"🌐 Found OTP Code (Selenium): `{otp_match.group(0)}`"

        cleaned = " ".join(body_text.split())
        if not cleaned:
            return "⏱️ Browser sign-in timed out (page remained blank while loading)."
        return f"🌐 Browser loaded page snippet:\n{cleaned[:300]}..."

    except SeleniumTimeoutException as e:
        return f"⏱️ Selenium browser sign-in timed out ({str(e) or 'login prompt or inbox did not load in time'})."
    except WebDriverException as e:
        logger.error(f"Selenium WebDriver Error for {email_addr}: {str(e)}")
        return f"❌ Selenium Error: {str(e)}"
    except Exception as e:
        logger.error(f"Browser OTP Fetch Error for {email_addr}: {str(e)}")
        return f"❌ Browser Error: {str(e)}"
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


@dataclass
class AccountCredential:
    email: str
    password: str
    refresh_token: str
    client_id: str
    raw_line: str = ""


class BatchOutlookInboxChecker:
    def __init__(
        self,
        authority: str = DEFAULT_AUTHORITY,
        scopes: list[str] | None = None,
        top: int = 5,
        timeout: int = 20,
    ) -> None:
        self.authority = authority
        self.scopes = scopes or DEFAULT_SCOPES
        self.top = top
        self.timeout = timeout
        self._msal_apps: dict[str, msal.PublicClientApplication] = {}

    def _get_msal_app(self, client_id: str) -> msal.PublicClientApplication:
        if client_id not in self._msal_apps:
            self._msal_apps[client_id] = msal.PublicClientApplication(
                client_id=client_id,
                authority=self.authority,
            )
        return self._msal_apps[client_id]

    def check_account(self, account: AccountCredential) -> dict:
        """
        Attempts to refresh the token and read recent inbox messages for one account.
        """
        app = self._get_msal_app(account.client_id)
        token_result = app.acquire_token_by_refresh_token(
            refresh_token=account.refresh_token,
            scopes=self.scopes,
        )

        access_token = token_result.get("access_token")
        if not access_token:
            return {
                "email": account.email,
                "ok": False,
                "stage": "token_refresh",
                "error": token_result.get("error", "unknown_error"),
                "description": token_result.get("error_description", "No access_token returned"),
                "new_refresh_token": None,
                "messages": [],
            }

        new_refresh_token = token_result.get("refresh_token", account.refresh_token)

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        }
        params = {
            "$top": self.top,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,bodyPreview,isRead",
        }

        try:
            resp = requests.get(
                GRAPH_INBOX_URL,
                headers=headers,
                params=params,
                proxies=_requests_proxies(),
                timeout=self.timeout,
            )
            resp.raise_for_status()
        except requests.HTTPError as exc:
            body_preview = exc.response.text[:300] if exc.response is not None else ""
            status_code = exc.response.status_code if exc.response is not None else "HTTP"
            return {
                "email": account.email,
                "ok": False,
                "stage": "graph_inbox",
                "error": f"HTTP_{status_code}",
                "description": body_preview or str(exc),
                "new_refresh_token": new_refresh_token,
                "messages": [],
            }
        except requests.RequestException as exc:
            return {
                "email": account.email,
                "ok": False,
                "stage": "network",
                "error": type(exc).__name__,
                "description": str(exc),
                "new_refresh_token": new_refresh_token,
                "messages": [],
            }

        raw_messages = resp.json().get("value", [])
        messages = [
            {
                "subject": m.get("subject"),
                "from": m.get("from", {}).get("emailAddress", {}).get("address"),
                "received_at": m.get("receivedDateTime"),
                "preview": m.get("bodyPreview"),
            }
            for m in raw_messages
        ]

        return {
            "email": account.email,
            "ok": True,
            "stage": "done",
            "error": None,
            "description": f"Fetched {len(messages)} message(s)",
            "new_refresh_token": new_refresh_token,
            "messages": messages,
        }

    def run_batch(self, accounts: list[AccountCredential]) -> list[dict]:
        return [self.check_account(acc) for acc in accounts]


inbox_checker = BatchOutlookInboxChecker(
    authority=DEFAULT_AUTHORITY,
    scopes=DEFAULT_SCOPES,
    top=5,
)


def fetch_otp_via_graph_checker(email_addr: str, password_str: str, refresh_token: str, client_id: str) -> str:
    """Uses BatchOutlookInboxChecker to read recent Graph messages and extract an OTP code."""
    cred = AccountCredential(
        email=email_addr,
        password=password_str or "",
        refresh_token=refresh_token,
        client_id=client_id,
        raw_line=f"{email_addr}|{password_str or ''}|{refresh_token}|{client_id}",
    )
    res = inbox_checker.check_account(cred)

    new_rt = res.get("new_refresh_token")
    if new_rt and new_rt != refresh_token:
        db.update_refresh_token(email_addr, new_rt)

    if not res.get("ok"):
        return f"OAUTH_FAILED [{res.get('stage')}/{res.get('error')}]: {res.get('description')}"

    messages = res.get("messages", [])
    if not messages:
        return "📭 Inbox is empty."

    for msg in messages:
        combined = f"{msg.get('subject') or ''} {msg.get('preview') or ''}"
        otp_match = re.search(r'\b\d{4,8}\b', combined)
        if otp_match:
            return f"🔑 Found OTP Code: `{otp_match.group(0)}`"

    latest = messages[0]
    snippet = (latest.get("preview") or "").strip()
    return f"📩 Latest Email ({latest.get('subject') or 'No Subject'}):\n{snippet[:300]}..."


class DatabaseManager:
    def __init__(self, db_path: str = DB_FILE):
        self.db_path = db_path
        self._init_db()

    def _get_connection(self):
        conn = sqlite3.connect(self.db_path, timeout=15)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS accounts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT UNIQUE NOT NULL,
                    password TEXT NOT NULL,
                    refresh_token TEXT DEFAULT NULL,
                    client_id TEXT DEFAULT NULL,
                    status TEXT DEFAULT 'available',
                    assigned_to INTEGER DEFAULT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            cursor.execute("PRAGMA table_info(accounts)")
            columns = [column[1] for column in cursor.fetchall()]

            if "refresh_token" not in columns:
                cursor.execute("ALTER TABLE accounts ADD COLUMN refresh_token TEXT DEFAULT NULL")
            if "client_id" not in columns:
                cursor.execute("ALTER TABLE accounts ADD COLUMN client_id TEXT DEFAULT NULL")

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS usage_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER NOT NULL,
                    email TEXT NOT NULL,
                    action TEXT NOT NULL,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS user_quotas (
                    user_id INTEGER PRIMARY KEY,
                    username TEXT,
                    quota INTEGER DEFAULT 0,
                    last_active TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            cursor.execute("PRAGMA table_info(user_quotas)")
            quota_cols = [col[1] for col in cursor.fetchall()]
            if "last_active" not in quota_cols:
                cursor.execute("ALTER TABLE user_quotas ADD COLUMN last_active TIMESTAMP DEFAULT NULL")
                cursor.execute("UPDATE user_quotas SET last_active = CURRENT_TIMESTAMP WHERE last_active IS NULL")

            conn.commit()

    def register_user(self, user_id: int, username: str | None):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO user_quotas (user_id, username, quota, last_active)
                VALUES (?, ?, 0, CURRENT_TIMESTAMP)
                ON CONFLICT(user_id) DO UPDATE SET 
                    username=EXCLUDED.username,
                    last_active=CURRENT_TIMESTAMP
            """, (user_id, f"@{username}" if username else f"ID:{user_id}"))
            conn.commit()

    def update_user_activity(self, user_id: int, username: str | None):
        """Updates user username and last active timestamp whenever they interact."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO user_quotas (user_id, username, quota, last_active)
                VALUES (?, ?, 0, CURRENT_TIMESTAMP)
                ON CONFLICT(user_id) DO UPDATE SET 
                    username=EXCLUDED.username,
                    last_active=CURRENT_TIMESTAMP
            """, (user_id, f"@{username}" if username else f"ID:{user_id}"))
            conn.commit()

    def add_quota_by_identifier(self, identifier: str, amount: int) -> tuple[bool, str]:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            target_id = None
            clean_ident = identifier.strip()

            if clean_ident.isdigit():
                target_id = int(clean_ident)
            else:
                uname = clean_ident if clean_ident.startswith("@") else f"@{clean_ident}"
                cursor.execute("SELECT user_id FROM user_quotas WHERE LOWER(username) = LOWER(?)", (uname,))
                row = cursor.fetchone()
                if row:
                    target_id = row[0]

            if target_id:
                cursor.execute("""
                    INSERT INTO user_quotas (user_id, username, quota, last_active)
                    VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(user_id) DO UPDATE SET quota = quota + EXCLUDED.quota
                """, (target_id, f"ID:{target_id}" if clean_ident.isdigit() else clean_ident, amount))
                conn.commit()
                return True, f"Successfully granted `{amount}` account quota(s) to ID `{target_id}`."

            return False, f"User `{identifier}` has not interacted with the bot yet. Please assign using their numerical Telegram User ID (e.g., `/assign 12345678 {amount}`) to pre-grant quota."

    def get_user_remaining_quota(self, user_id: int) -> int:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT quota FROM user_quotas WHERE user_id = ?", (user_id,))
            row = cursor.fetchone()
            return row[0] if row else 0

    def add_account(self, email_addr: str, password_str: str, refresh_token: str | None = None, client_id: str | None = None) -> bool:
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    INSERT INTO accounts (email, password, refresh_token, client_id, status, assigned_to)
                    VALUES (?, ?, ?, ?, 'available', NULL)
                    """,
                    (
                        email_addr.strip().lower(),
                        password_str.strip(),
                        refresh_token.strip() if refresh_token else None,
                        client_id.strip() if client_id else None
                    )
                )
                conn.commit()
                return True
        except sqlite3.IntegrityError:
            return False

    def get_assigned_account(self, user_id: int):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT id, email, password, refresh_token, client_id
                FROM accounts
                WHERE assigned_to = ? AND status = 'assigned'
                LIMIT 1
                """,
                (user_id,)
            )
            row = cursor.fetchone()
            return tuple(row) if row else None

    def assign_next_available(self, user_id: int):
        """
        Atomically assigns one unique, never-assigned, non-used account to user_id.
        """
        conn = sqlite3.connect(self.db_path, timeout=15)
        try:
            conn.execute("BEGIN IMMEDIATE")
            cursor = conn.cursor()

            cursor.execute(
                "SELECT id FROM accounts WHERE assigned_to = ? AND status = 'assigned' LIMIT 1",
                (user_id,)
            )
            if cursor.fetchone():
                conn.rollback()
                return "ALREADY_ACTIVE"

            cursor.execute("SELECT quota FROM user_quotas WHERE user_id = ?", (user_id,))
            q_row = cursor.fetchone()
            if not q_row or q_row[0] <= 0:
                conn.rollback()
                return "NO_QUOTA"

            cursor.execute("""
                SELECT id, email, password, refresh_token, client_id
                FROM accounts
                WHERE status = 'available'
                  AND assigned_to IS NULL
                  AND email NOT IN (
                      SELECT email FROM usage_logs WHERE action IN ('assigned', 'used')
                  )
                ORDER BY id ASC
                LIMIT 1
            """)
            acc = cursor.fetchone()
            if not acc:
                conn.rollback()
                return "NO_ACCOUNTS"

            acc_id, email_addr, password_str, r_token, c_id = acc

            cursor.execute("""
                UPDATE accounts
                SET status = 'assigned', assigned_to = ?
                WHERE id = ? AND status = 'available' AND assigned_to IS NULL
            """, (user_id, acc_id))
            if cursor.rowcount != 1:
                conn.rollback()
                return "NO_ACCOUNTS"

            cursor.execute("UPDATE user_quotas SET quota = quota - 1 WHERE user_id = ? AND quota > 0", (user_id,))
            if cursor.rowcount != 1:
                conn.rollback()
                return "NO_QUOTA"

            cursor.execute(
                "INSERT INTO usage_logs (user_id, email, action) VALUES (?, ?, 'assigned')",
                (user_id, email_addr)
            )
            conn.commit()
            return email_addr, password_str
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def update_refresh_token(self, email_addr: str, new_refresh_token: str):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE accounts SET refresh_token = ? WHERE LOWER(email) = LOWER(?)", (new_refresh_token, email_addr))
            conn.commit()

    def mark_as_used(self, user_id: int) -> str | None:
        conn = sqlite3.connect(self.db_path, timeout=15)
        try:
            conn.execute("BEGIN IMMEDIATE")
            cursor = conn.cursor()
            cursor.execute(
                "SELECT id, email FROM accounts WHERE assigned_to = ? AND status = 'assigned' LIMIT 1",
                (user_id,)
            )
            row = cursor.fetchone()
            if not row:
                conn.rollback()
                return None

            acc_id, email_addr = row
            cursor.execute("UPDATE accounts SET status = 'used' WHERE id = ? AND assigned_to = ?", (acc_id, user_id))
            cursor.execute(
                "INSERT INTO usage_logs (user_id, email, action) VALUES (?, ?, 'used')",
                (user_id, email_addr)
            )
            conn.commit()
            return email_addr
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def reset_accounts_db(self) -> int:
        """Deletes all uploaded accounts from the accounts table and returns how many were removed."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM accounts")
            count = cursor.fetchone()[0]
            cursor.execute("DELETE FROM accounts")
            cursor.execute("DELETE FROM sqlite_sequence WHERE name='accounts'")
            conn.commit()
            return count

    def get_all_users_info(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT 
                    uq.user_id,
                    uq.username,
                    uq.quota,
                    (SELECT COUNT(*) FROM accounts WHERE assigned_to = uq.user_id AND status = 'assigned') as active_count,
                    (SELECT COUNT(*) FROM accounts WHERE assigned_to = uq.user_id AND status = 'used') as used_count,
                    uq.last_active
                FROM user_quotas uq
                ORDER BY uq.last_active DESC
            """)
            return [tuple(r) for r in cursor.fetchall()]

    def get_stats(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT status, COUNT(*) FROM accounts GROUP BY status")
            status_rows = cursor.fetchall()
            status_counts = {row[0]: row[1] for row in status_rows}

            cursor.execute("""
                SELECT 
                    assigned_to,
                    SUM(CASE WHEN status = 'used' THEN 1 ELSE 0 END),
                    SUM(CASE WHEN status = 'assigned' THEN 1 ELSE 0 END)
                FROM accounts 
                WHERE assigned_to IS NOT NULL 
                GROUP BY assigned_to
            """)
            user_stats = [tuple(r) for r in cursor.fetchall()]
            return status_counts, user_stats


db = DatabaseManager()


def track_activity(func):
    """Decorator that updates user activity before running a command or message handler."""
    @wraps(func)
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE, *args, **kwargs):
        user = update.effective_user if update else None
        if user:
            try:
                db.update_user_activity(user.id, user.username)
            except Exception as e:
                logger.error(f"Failed to update user activity for {user.id}: {e}")
        return await func(update, context, *args, **kwargs)
    return wrapper


@track_activity
async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db.register_user(user.id, user.username)
    rem_quota = db.get_user_remaining_quota(user.id)

    welcome_text = (
        f"👋 Welcome, **{_escape_md(user.first_name)}**!\n\n"
        f"📊 Your Remaining Account Quota: `{rem_quota}`\n\n"
        f"**User Commands:**\n"
        f"• `/getacc` — Get an assigned account\n"
        f"• `/getcode` — Fetch latest OTP code from inbox\n"
        f"• `/done` — Mark your current account as used\n"
    )

    if user.id in ADMIN_IDS:
        welcome_text += (
            f"\n**Admin Commands:**\n"
            f"• `/addacc` — Add accounts or upload `.txt` file\n"
            f"• `/assign <user_or_id> <amount>` — Grant account quota\n"
            f"• `/users` — View registered users & activity\n"
            f"• `/stock` — Check unassigned available accounts\n"
            f"• `/stats` — View full system account statistics\n"
            f"• `/reset confirm` — Clear all uploaded accounts from DB\n"
        )

    await update.message.reply_text(welcome_text, parse_mode="Markdown")


def fetch_otp_via_oauth2(email_addr: str, refresh_token: str, client_id: str) -> str:
    authority = "https://login.microsoftonline.com/consumers"
    scopes = ["https://graph.microsoft.com/Mail.Read"]

    try:
        app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority
        )

        token_result = app.acquire_token_by_refresh_token(
            refresh_token=refresh_token,
            scopes=scopes
        )

        access_token = token_result.get("access_token")
        if not access_token:
            error_code = token_result.get("error", "unknown_error")
            error_desc = token_result.get("error_description", "No access token returned")
            return f"❌ MSAL OAuth Error [{error_code}]: {error_desc}"

        new_refresh_token = token_result.get("refresh_token")
        if new_refresh_token and new_refresh_token != refresh_token:
            db.update_refresh_token(email_addr, new_refresh_token)

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json"
        }
        params = {
            "$top": 5,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,body,bodyPreview"
        }

        graph_url = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
        mail_res = requests.get(
            graph_url,
            headers=headers,
            params=params,
            proxies=_requests_proxies(),
            timeout=15,
        )
        mail_res.raise_for_status()
        mail_json = mail_res.json()

        messages = mail_json.get("value", [])
        if not messages:
            return "📭 Inbox is empty."

        latest_msg = messages[0]
        body_content = latest_msg.get("body", {}).get("content", "") or latest_msg.get("bodyPreview", "")

        otp_match = re.search(r'\b\d{4,8}\b', body_content)
        if otp_match:
            return f"🔑 Found OTP Code: `{otp_match.group(0)}`"

        clean_body = re.sub('<[^<]+?>', '', body_content)
        return f"📩 Latest Email Content Snippet:\n{clean_body[:300]}..."

    except Exception as e:
        logger.error(f"OAuth Fetch Error for {email_addr}: {str(e)}")
        return f"OAUTH_FAILED: {str(e)}"


def generate_xoauth2_string(username: str, access_token: str) -> str:
    """Generates the SASL XOAUTH2 string required by Microsoft IMAP."""
    return f"user={username}\x01auth=Bearer {access_token}\x01\x01"


def get_access_token_from_refresh(email_addr: str, refresh_token: str, client_id: str) -> str | None:
    """Exchanges a refresh token for an access token via MSAL."""
    authority = "https://login.microsoftonline.com/consumers"
    scopes = []

    try:
        app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority
        )

        token_result = app.acquire_token_by_refresh_token(
            refresh_token=refresh_token,
            scopes=scopes
        )

        access_token = token_result.get("access_token")
        if access_token:
            new_refresh_token = token_result.get("refresh_token")
            if new_refresh_token and new_refresh_token != refresh_token:
                db.update_refresh_token(email_addr, new_refresh_token)
            return access_token

        logger.error(f"Failed to acquire access token: {token_result.get('error_description')}")
        return None
    except Exception as e:
        logger.error(f"MSAL Token acquisition exception: {str(e)}")
        return None


def fetch_otp_via_imap(email_addr: str, password_str: str, refresh_token: str | None = None, client_id: str | None = None) -> str:
    imap_server = "outlook.office365.com"
    imap_port = 993

    try:
        mail = imaplib.IMAP4_SSL(imap_server, imap_port)

        if refresh_token and client_id:
            access_token = get_access_token_from_refresh(email_addr, refresh_token, client_id)
            if access_token:
                xoauth2_str = generate_xoauth2_string(email_addr, access_token)
                mail.authenticate('XOAUTH2', lambda x: xoauth2_str)
            else:
                return fetch_otp_via_oauth2(email_addr, refresh_token, client_id)
        else:
            mail.login(email_addr, password_str)

        mail.select("inbox")

        status, messages = mail.search(None, "ALL")
        if status != "OK":
            mail.logout()
            return "❌ Failed to search inbox."

        message_ids = messages[0].split()
        if not message_ids:
            mail.logout()
            return "📭 Inbox is empty."

        latest_id = message_ids[-1]
        _, msg_data = mail.fetch(latest_id, "(RFC822)")

        email_body = ""
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            payload = part.get_payload(decode=True)
                            if payload:
                                email_body = payload.decode("utf-8", errors="ignore")
                                break
                else:
                    payload = msg.get_payload(decode=True)
                    if payload:
                        email_body = payload.decode("utf-8", errors="ignore")

        mail.logout()

        otp_match = re.search(r'\b\d{4,8}\b', email_body)
        if otp_match:
            return f"🔑 Found OTP Code: `{otp_match.group(0)}`"

        return f"📩 Latest email content snippet:\n{email_body[:300]}..."

    except imaplib.IMAP4.error as e:
        return f"❌ IMAP Authentication Failed: {str(e)}"
    except Exception as e:
        return f"❌ Error: {str(e)}"


async def _find_password_input(page, total_timeout_sec: int = 18):
    """Polls across frames/selectors for a visible password field."""
    selectors = [
        "input[type='password']",
        "input[name='passwd']",
        "input[id*='passwd' i]",
        "input[id*='password' i]",
        "input[placeholder*='password' i]",
    ]
    deadline = datetime.utcnow() + timedelta(seconds=total_timeout_sec)
    while datetime.utcnow() < deadline:
        for frame in page.frames:
            for sel in selectors:
                try:
                    loc = frame.locator(sel).first
                    if await loc.count() > 0 and await loc.is_visible():
                        return loc
                except Exception:
                    continue
        await page.wait_for_timeout(500)
    return None


async def fetch_otp_via_browser(email_addr: str, password_str: str, headless: bool = True) -> str:
    """Signs in via Outlook Web using Selenium WebDriver."""
    user_agent = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )

    if not SELENIUM_AVAILABLE:
        return "❌ Selenium is not installed (`pip install selenium webdriver-manager`)."

    return await asyncio.to_thread(
        _fetch_otp_via_selenium_sync,
        email_addr,
        password_str,
        user_agent,
        headless,
    )

    if SELENIUM_AVAILABLE:
        return await asyncio.to_thread(
            _fetch_otp_via_selenium_sync,
            email_addr,
            password_str,
            user_agent,
            headless,
        )

    if PLAYWRIGHT_AVAILABLE:
        try:
            async with async_playwright() as p:
                launch_kwargs = {
                    "headless": headless,
                    "args": ["--disable-blink-features=AutomationControlled"],
                }
                proxy_cfg = _playwright_proxy_config()
                if proxy_cfg:
                    launch_kwargs["proxy"] = proxy_cfg

                browser = await p.chromium.launch(**launch_kwargs)
                context = await browser.new_context(
                    user_agent=user_agent,
                    viewport={"width": 1280, "height": 800}
                )
                page = await context.new_page()
                page.set_default_timeout(35000)

                last_nav_error = None
                for candidate_url in BROWSER_SIGNIN_URLS:
                    try:
                        await page.goto(candidate_url, wait_until="domcontentloaded")
                        await page.wait_for_timeout(1500)
                        current_txt = (await page.inner_text("body")).lower()
                        if "error" not in page.url.lower() and any(
                            k in current_txt for k in ["sign in", "email", "phone", "skype", "password", "next"]
                        ):
                            break
                    except Exception as nav_exc:
                        last_nav_error = nav_exc
                        continue
                else:
                    if last_nav_error and not page.url:
                        raise last_nav_error

                try:
                    existing_acc = page.locator(f"div[role='button']:has-text('{email_addr}'), button:has-text('{email_addr}')").first
                    if await existing_acc.is_visible(timeout=2000):
                        await existing_acc.click()
                        await page.wait_for_timeout(1200)
                    else:
                        other_acc = page.locator("div[role='button']:has-text('Use another account'), button:has-text('Use another account'), div:has-text('Use another account')").first
                        if await other_acc.is_visible(timeout=2000):
                            await other_acc.click()
                            await page.wait_for_timeout(1200)
                except Exception:
                    pass

                email_input = page.locator("input[type='email'], input[name='loginfmt'], input[name='identifier']").first
                try:
                    await email_input.wait_for(state="visible", timeout=12000)
                    await email_input.fill(email_addr)
                    submit_btn = page.locator("input[type='submit'], button[type='submit'], button:has-text('Next')").first
                    await submit_btn.click()
                    await page.wait_for_timeout(1500)
                except AsyncPlaywrightTimeoutError:
                    logger.info(f"Email input not immediately visible for {email_addr}; checking password/account screen.")

                for alt_text in ["Use your password", "Enter my password", "Password"]:
                    try:
                        pwd_alt = page.locator(f"button:has-text('{alt_text}'), a:has-text('{alt_text}'), span:has-text('{alt_text}')").first
                        if await pwd_alt.is_visible(timeout=2000):
                            await pwd_alt.click()
                            await page.wait_for_timeout(1200)
                            break
                    except Exception:
                        continue

                password_input = await _find_password_input(page, total_timeout_sec=18)
                if password_input is not None:
                    await password_input.fill(password_str.strip())
                    pwd_submit = page.locator("input[type='submit'], button[type='submit'], button:has-text('Sign in'), button:has-text('Next')").first
                    await pwd_submit.click()
                else:
                    current_body = ""
                    try:
                        current_body = await page.inner_text("body")
                    except Exception:
                        pass
                    lower_current = current_body.lower()

                    if "password sign-in isn't available" in lower_current or "try another method" in lower_current:
                        try:
                            try_other = page.locator("button:has-text('Try another method'), a:has-text('Try another method'), div[role='button']:has-text('Try another method')").first
                            if await try_other.is_visible(timeout=2500):
                                await try_other.click()
                                await page.wait_for_timeout(1500)
                                current_body = await page.inner_text("body")
                        except Exception:
                            pass
                        await browser.close()
                        alt_snippet = " ".join(current_body.split())[:240]
                        return f"⚠️ Password sign-in is disabled for `{email_addr}`. Alternative screen: {alt_snippet or 'none'}"

                    await browser.close()

                    if any(k in lower_current for k in ["verify your identity", "help us protect", "enter the code", "approve sign in"]):
                        return "⚠️ Browser sign-in triggered a security/2FA verification challenge before password entry."
                    if any(k in lower_current for k in ["couldn't find an account", "doesn't exist", "enter a valid email", "account doesn't exist"]):
                        return f"❌ Microsoft reported that `{email_addr}` was not found or invalid."
                    if any(k in lower_current for k in ["other ways to sign in", "sign in another way", "passkey", "windows hello"]):
                        return "⚠️ Microsoft prompted for passkey / alternative sign-in instead of a password field."

                    snippet = " ".join(current_body.split())[:220]
                    return f"⏱️ Password prompt did not appear in time (URL: `{page.url}`). Page text: {snippet or 'empty'}"

                for prompt_selector in [
                    "input#idBtn_Back",
                    "button#idBtn_Back",
                    "button:has-text('No')",
                    "input[type='submit'][value='Yes']",
                    "button:has-text('Skip for now')",
                    "a:has-text('Skip for now')"
                ]:
                    try:
                        prompt_btn = page.locator(prompt_selector).first
                        if await prompt_btn.is_visible(timeout=2500):
                            await prompt_btn.click()
                            await page.wait_for_timeout(1500)
                            break
                    except Exception:
                        continue

                body_text = ""
                for _ in range(10):
                    await page.wait_for_timeout(2000)
                    body_text = await page.inner_text("body")
                    lower_body = body_text.lower()
                    if any(phrase in lower_body for phrase in [
                        "verify your identity",
                        "help us protect your account",
                        "enter the code sent to",
                        "approve sign in request"
                    ]):
                        await browser.close()
                        return "⚠️ Browser sign-in triggered a security/2FA verification challenge."
                    if "incorrect account or password" in lower_body:
                        await browser.close()
                        return f"❌ Microsoft returned 'Incorrect account or password' for `{email_addr}`."
                    if "focused" in lower_body or "other" in lower_body or "inbox" in lower_body or re.search(r'\b\d{4,8}\b', body_text):
                        break

                otp_match = re.search(r'\b\d{4,8}\b', body_text)
                await browser.close()

                if otp_match:
                    return f"🌐 Found OTP Code (Browser): `{otp_match.group(0)}`"

                cleaned = " ".join(body_text.split())
                if not cleaned:
                    return "⏱️ Browser sign-in timed out (page remained blank while loading)."
                return f"🌐 Browser loaded page snippet:\n{cleaned[:300]}..."

        except AsyncPlaywrightTimeoutError as e:
            return f"⏱️ Browser sign-in timed out ({str(e) or 'login prompt or inbox did not load in time'})."
        except Exception as e:
            logger.error(f"Browser OTP Fetch Error for {email_addr}: {str(e)}")
            return f"❌ Browser Error: {str(e)}"

    try:
        cookie_jar = http.cookiejar.CookieJar()
        handlers = [urllib.request.HTTPCookieProcessor(cookie_jar)]
        if PROXY_URL:
            handlers.append(urllib.request.ProxyHandler({"http": PROXY_URL, "https": PROXY_URL}))
        opener = urllib.request.build_opener(*handlers)
        opener.addheaders = [
            ("User-Agent", user_agent),
            ("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"),
        ]

        login_page_req = opener.open(OUTLOOK_SIGNIN_URL, timeout=15)
        html_content = login_page_req.read().decode("utf-8", errors="ignore")

        action_match = re.search(r'<form[^>]*action=["\']([^"\']+)["\']', html_content, re.IGNORECASE)
        if action_match:
            post_url = action_match.group(1)
            if post_url.startswith("/"):
                post_url = f"https://login.live.com{post_url}"
            payload = urllib.parse.urlencode({
                "loginfmt": email_addr,
                "passwd": password_str,
            }).encode("utf-8")
            post_req = urllib.request.Request(post_url, data=payload, method="POST")
            post_req.add_header("Content-Type", "application/x-www-form-urlencoded")
            response = opener.open(post_req, timeout=15)
            html_content = response.read().decode("utf-8", errors="ignore")

        clean_text = re.sub(r'<[^>]+>', ' ', html_content)
        clean_text = " ".join(clean_text.split())

        otp_match = re.search(r'\b\d{4,8}\b', clean_text)
        if otp_match:
            return f"🌐 Found OTP Code (Fallback HTTP): `{otp_match.group(0)}`"

        return (
            "⚠️ Playwright is not installed (`pip install playwright && playwright install chromium`), "
            f"and fallback HTTP flow found no OTP code. Snippet: {clean_text[:200]}..."
        )
    except Exception as e:
        logger.error(f"Fallback Browser Flow Error for {email_addr}: {str(e)}")
        return f"❌ Browser Fallback Error: {str(e)}"


@track_activity
async def getacc_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db.register_user(user.id, user.username)

    existing = db.get_assigned_account(user.id)
    if existing:
        _, em, pw, _, _ = existing
        await update.message.reply_text(f"⚠️ You already have an active account:\n📧 `{em}`\n🔑 `{pw}`", parse_mode="Markdown")
        return

    assigned = db.assign_next_available(user.id)
    if assigned == "ALREADY_ACTIVE":
        existing = db.get_assigned_account(user.id)
        if existing:
            _, em, pw, _, _ = existing
            await update.message.reply_text(f"⚠️ You already have an active account:\n📧 `{em}`\n🔑 `{pw}`", parse_mode="Markdown")
        return
    elif assigned == "NO_QUOTA":
        await update.message.reply_text("❌ You have 0 assigned account credits. Ask an admin to grant you account permissions using `/assign`.")
        return
    elif assigned == "NO_ACCOUNTS" or not assigned:
        await update.message.reply_text("❌ No available accounts in stock right now.")
        return

    em, pw = assigned
    rem_quota = db.get_user_remaining_quota(user.id)
    await update.message.reply_text(f"✅ Assigned Account:\n📧 `{em}`\n🔑 `{pw}`\n\n📊 Remaining Quota: `{rem_quota}` account(s).", parse_mode="Markdown")


@track_activity
async def assign_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    if len(context.args) < 2:
        await update.message.reply_text("Usage: `/assign <@username or UserID> <amount>`\nExample: `/assign @JohnDoe 5`", parse_mode="Markdown")
        return

    target = context.args[0]
    try:
        amount = int(context.args[1])
    except ValueError:
        await update.message.reply_text("❌ Amount must be a valid number.")
        return

    success, msg = db.add_quota_by_identifier(target, amount)
    await update.message.reply_text(f"{'✅' if success else '❌'} {msg}", parse_mode="Markdown")


@track_activity
async def users_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    try:
        users_data = db.get_all_users_info()
        if not users_data:
            await update.message.reply_text("👥 No registered users yet.")
            return

        total_users = len(users_data)
        online_count = 0
        fifteen_mins_ago = datetime.utcnow() - timedelta(minutes=15)

        user_lines = []
        for u_id, uname, quota, active, used, last_active_str in users_data:
            is_online = False
            if last_active_str:
                try:
                    clean_ts = str(last_active_str).split('.')[0].replace('T', ' ')
                    last_dt = datetime.strptime(clean_ts, "%Y-%m-%d %H:%M:%S")
                    if last_dt >= fifteen_mins_ago:
                        is_online = True
                except Exception:
                    pass

            status_indicator = "🟢 Online" if is_online else "⚪ Offline"
            if is_online:
                online_count += 1

            safe_uname = _escape_md(str(uname or f"ID:{u_id}"))
            safe_last = _escape_md(str(last_active_str or "N/A"))

            user_lines.append(
                f"{status_indicator} *{safe_uname}* (`{u_id}`)\n"
                f"   ├ Quota Balance: `{quota or 0}`\n"
                f"   ├ Active Acc: `{active or 0}`\n"
                f"   ├ Used Accs: `{used or 0}`\n"
                f"   └ Last Active: `{safe_last}`"
            )

        header = (
            f"👥 *User Activity & Quota Report*\n"
            f"📊 Total Registered: `{total_users}` | 🟢 Active Now (15m): `{online_count}`\n\n"
        )

        chunk = header
        for block in user_lines:
            entry = block + "\n\n"
            if len(chunk) + len(entry) > 3900:
                await update.message.reply_text(chunk, parse_mode="Markdown")
                chunk = entry
            else:
                chunk += entry

        if chunk:
            await update.message.reply_text(chunk, parse_mode="Markdown")
    except Exception as e:
        logger.error(f"Error in /users handler: {e}", exc_info=True)
        await update.message.reply_text(f"❌ Failed to load users report: {e}")


@track_activity
async def getcode_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    assigned = db.get_assigned_account(user_id)
    if not assigned:
        await update.message.reply_text("❌ You do not have an active account. Use `/getacc` first.")
        return

    _, em, pw, r_token, c_id = assigned

    status_msg = await update.message.reply_text("🔄 Checking Hotmail inbox...")

    has_oauth = bool(r_token and r_token.strip() and c_id and c_id.strip())

    if has_oauth:
        result = fetch_otp_via_graph_checker(em, pw, r_token.strip(), c_id.strip())
        if "OAUTH_FAILED" in result:
            imap_res = fetch_otp_via_imap(em, pw, r_token.strip(), c_id.strip())
            if not imap_res.startswith("❌") and "OAUTH_FAILED" not in imap_res:
                result = imap_res
            else:
                result = f"{result}\n{imap_res}"
    else:
        result = fetch_otp_via_imap(em, pw, None, None)

    if result.startswith("❌") or "OAUTH_FAILED" in result or result.startswith("📩"):
        browser_result = await fetch_otp_via_browser(em, pw, headless=True)
        if browser_result.startswith("🌐 Found OTP Code"):
            result = browser_result
        elif result.startswith("❌") or "OAUTH_FAILED" in result:
            result = f"{result}\n\n{browser_result}"

    if status_msg:
        await status_msg.edit_text(result, parse_mode="Markdown")
    else:
        await update.message.reply_text(result, parse_mode="Markdown")


@track_activity
async def done_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    used_email = db.mark_as_used(user_id)
    if not used_email:
        await update.message.reply_text("❌ You don't have any active account to mark as done.")
        return

    await update.message.reply_text(f"✅ Account `{used_email}` marked as used and locked.", parse_mode="Markdown")


@track_activity
async def addacc_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    def parse_raw_order_content(text: str) -> list[tuple[str, str, str | None, str | None]]:
        """Parses multi-line raw marketplace order text and extracts account credentials."""
        accounts = []
        clean_text = text.replace('\xa0', ' ').replace('\u200b', '')
        blocks = re.split(r'(?:Item\s+\d+:|Item\s+\d+)', clean_text, flags=re.IGNORECASE)

        for block in blocks:
            if not block.strip():
                continue

            email_match = re.search(r'Email:\s*(\S+)', block, re.IGNORECASE)
            pass_match = re.search(r'Password:\s*(\S+)', block, re.IGNORECASE)
            cred1_match = re.search(r'Additional credential 1:\s*(\S+)', block, re.IGNORECASE)
            cred2_match = re.search(r'Additional credential 2:\s*(\S+)', block, re.IGNORECASE)

            if email_match and pass_match:
                em = email_match.group(1).strip()
                pw = pass_match.group(1).strip()
                r_token = cred1_match.group(1).strip() if cred1_match else None
                c_id = cred2_match.group(1).strip() if cred2_match else None
                accounts.append((em, pw, r_token, c_id))

        return accounts

    def parse_and_add_text(raw_text: str) -> int:
        """Parses raw text in pipe-separated formats or multi-line order block formats."""
        count = 0
        clean_text = raw_text.replace('\xa0', ' ').replace('\u200b', '')

        order_accounts = parse_raw_order_content(clean_text)
        if order_accounts:
            for em, pw, r_token, c_id in order_accounts:
                if db.add_account(em, pw, refresh_token=r_token, client_id=c_id):
                    count += 1
            return count

        for line in clean_text.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue

            if "|" in line:
                parts = [p.strip() for p in line.split("|")]
                if len(parts) >= 4:
                    if db.add_account(parts[0], parts[1], refresh_token=parts[2], client_id=parts[3]):
                        count += 1
                elif len(parts) == 2:
                    if db.add_account(parts[0], parts[1]):
                        count += 1

        return count

    if update.message and update.message.document:
        doc = update.message.document
        if not doc.file_name.endswith(".txt"):
            await update.message.reply_text("❌ Please upload a valid `.txt` file.")
            return

        file = await context.bot.get_file(doc.file_id)
        file_bytes = await file.download_as_bytearray()
        content = file_bytes.decode("utf-8", errors="ignore")

        added_count = parse_and_add_text(content)
        await update.message.reply_text(f"✅ Successfully bulk imported **{added_count}** accounts from `.txt` file!", parse_mode="Markdown")
        return

    if not context.args:
        await update.message.reply_text("Usage:\nUpload a `.txt` file or send:\n`/addacc email|password` or\n`/addacc email|pass|refresh_token|client_id`", parse_mode="Markdown")
        return

    text_input = update.message.text.partition(' ')[2]
    added_count = parse_and_add_text(text_input)
    await update.message.reply_text(f"✅ Successfully added **{added_count}** account(s).", parse_mode="Markdown")


@track_activity
async def stats_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    status_counts, user_stats = db.get_stats()

    available = status_counts.get("available", 0)
    assigned = status_counts.get("assigned", 0)
    used = status_counts.get("used", 0)

    report = (
        f"📊 **System Account Stats**\n\n"
        f"🟢 Available: `{available}`\n"
        f"🟡 Assigned (Active): `{assigned}`\n"
        f"🔴 Used / Completed: `{used}`\n\n"
        f"👥 **User Breakdown:**\n"
    )

    if not user_stats:
        report += "_No user usage records found yet._"
    else:
        for u_id, used_cnt, active_cnt in user_stats:
            report += f"• User ID `{u_id}` -> Used: `{used_cnt or 0}` | Active: `{active_cnt or 0}`\n"

    await update.message.reply_text(report, parse_mode="Markdown")


@track_activity
async def stock_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    status_counts, _ = db.get_stats()
    available = status_counts.get("available", 0)

    await update.message.reply_text(
        f"📦 **Account Stock Report**\n\n"
        f"🟢 **Available Unassigned Accounts:** `{available}`",
        parse_mode="Markdown"
    )


@track_activity
async def reset_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    if not context.args or context.args[0].lower() != "confirm":
        await update.message.reply_text(
            "⚠️ **Warning:** This will delete **all** uploaded accounts from the database.\n"
            "Type `/reset confirm` to proceed.",
            parse_mode="Markdown"
        )
        return

    deleted_count = db.reset_accounts_db()
    await update.message.reply_text(
        f"🗑️ Cleared **{deleted_count}** account(s) from the accounts database.",
        parse_mode="Markdown"
    )


def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("getacc", getacc_handler))
    app.add_handler(CommandHandler("getcode", getcode_handler))
    app.add_handler(CommandHandler("done", done_handler))
    app.add_handler(CommandHandler("assign", assign_handler))
    app.add_handler(CommandHandler("users", users_handler))
    app.add_handler(CommandHandler("stock", stock_handler))
    app.add_handler(CommandHandler("reset", reset_handler))
    app.add_handler(CommandHandler("addacc", addacc_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, addacc_handler))
    app.add_handler(CommandHandler("stats", stats_handler))

    logger.info("Bot is starting polling...")
    app.run_polling()


if __name__ == "__main__":
    main()