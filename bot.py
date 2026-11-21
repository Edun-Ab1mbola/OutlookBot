import os
import re
import time
import logging
import sqlite3
import requests
from typing import Optional, Tuple, List, Dict

from dotenv import load_dotenv

import email
import logging
import sqlite3
import requests
import io
from email.header import decode_header
from typing import Optional, Tuple, List, Dict

from dotenv import load_dotenv
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters
)

# ---------------------------------------------------------------------------
def parse_credential_line(raw_text: str) -> List[Tuple[str, str, Optional[str], Optional[str]]]:
    """
    Extracts email|password[|refresh_token|client_id] tuples from packed, 
    unspaced, or multi-line text blocks using regex patterns.
    """
    found_accounts = []
    lines = raw_text.splitlines()
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Split by pipe '|' or colon ':'
        parts = [p.strip() for p in re.split(r'[|:]', line) if p.strip()]
        
        # Validate that the first field looks like an email address
        if len(parts) >= 2 and "@" in parts[0]:
            email_val = parts[0]
            pass_val = parts[1]
            ref_token = parts[2] if len(parts) >= 3 else None
            c_id = parts[3] if len(parts) >= 4 else None
            
            found_accounts.append((email_val, pass_val, ref_token, c_id))
            
    return found_accounts

# ---------------------------------------------------------------------------
load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
API_ID = os.getenv("API_ID", "")
API_HASH = os.getenv("API_HASH", "")
PROXY_URL = os.getenv("PROXY_URL", "")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "")

admin_ids_raw = os.getenv("ADMIN_IDS", "")
ADMIN_IDS = [int(x.strip()) for x in admin_ids_raw.split(",") if x.strip().isdigit()]

# Logging Setup
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Standard Microsoft Endpoints
MS_GRAPH_TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
MS_GRAPH_MESSAGES_URL = "https://graph.microsoft.com/v1.0/me/messages"

# ---------------------------------------------------------------------------
# Helper Utility Functions
# ---------------------------------------------------------------------------
def _escape_md(text: str) -> str:
    """Escapes Markdown special characters for safe Telegram delivery."""
    if not text:
        return ""
    escape_chars = r"_*[]()~`>#+-=|{}.!"
    return re.sub(f"([{re.escape(escape_chars)}])", r"\\\1", text)

def extract_otp(text: str) -> Optional[str]:
    """Finds 4-8 digit numeric codes commonly present in verification messages."""
    patterns = [
        r"(?:code|verification|otp|pin|is)\b[:\s\-]*([0-9]{4,8})\b",
        r"\b([0-9]{4,8})\b"
    ]
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            return matches[0]
    return None

# ---------------------------------------------------------------------------
# Account Checking Services (Microsoft Graph OAuth 2.0 ONLY)
# ---------------------------------------------------------------------------
def fetch_otp_via_graph_checker(email_str: str, password: str, refresh_token: str, client_id: str) -> str:
    """Uses standard Microsoft OAuth 2.0 refresh token grant to access Graph API."""
    if not refresh_token or not client_id:
        return (
            "❌ *OAuth 2.0 Error:* Missing `refresh_token` or `client_id` for this account.\n"
            "⚠️ Microsoft has permanently disabled basic password/IMAP auth. "
            "Please re-add this account with its valid OAuth tokens."
        )

    proxies = {"http": PROXY_URL, "https": PROXY_URL} if PROXY_URL else None

    token_payload = {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": "https://graph.microsoft.com/Mail.Read offline_access"
    }
    if MS_CLIENT_SECRET:
        token_payload["client_secret"] = MS_CLIENT_SECRET

    try:
        response = requests.post(
            MS_GRAPH_TOKEN_URL,
            data=token_payload,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            proxies=proxies,
            timeout=20
        )
        try:
            token_data = response.json()
        except ValueError:
            return f"❌ *OAuth 2.0 Token Failed (HTTP {response.status_code}):*\n`{_escape_md(response.text[:300])}`"

        if response.status_code != 200 or "access_token" not in token_data:
            err = token_data.get("error_description") or token_data.get("error") or "Unknown OAuth error"
            logger.error(f"OAuth error for {email_str} (HTTP {response.status_code}): {err}")
            return f"❌ *OAuth 2.0 Token Failed (HTTP {response.status_code}):*\n_{_escape_md(err)}_"

        access_token = token_data["access_token"]
        graph_headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json"
        }
        # Explicitly order newest messages first by receivedDateTime
        params_with_order = {
            "$top": 8,
            "$orderby": "receivedDateTime DESC",
            "$select": "id,subject,bodyPreview,receivedDateTime,from,isRead"
        }
        params_fallback = {
            "$top": 8,
            "$select": "id,subject,bodyPreview,receivedDateTime,from,isRead"
        }

        # Check Inbox first, then Junk/Spam and all messages if needed
        folders_to_check = [
            ("Inbox", "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"),
            ("Junk Email", "https://graph.microsoft.com/v1.0/me/mailFolders/junkemail/messages"),
            ("All Messages", MS_GRAPH_MESSAGES_URL),
        ]

        max_attempts = 4
        retry_delay_sec = 4
        collected_messages = []
        graph_errors = []

        for attempt in range(1, max_attempts + 1):
            collected_messages = []
            seen_ids = set()
            graph_errors = []

            for folder_label, endpoint_url in folders_to_check:
                try:
                    mail_res = requests.get(
                        endpoint_url,
                        headers=graph_headers,
                        params=params_with_order,
                        proxies=proxies,
                        timeout=20
                    )
                    # Retry without $orderby if Graph returns 400 Bad Request on sorting
                    if mail_res.status_code == 400:
                        mail_res = requests.get(
                            endpoint_url,
                            headers=graph_headers,
                            params=params_fallback,
                            proxies=proxies,
                            timeout=20
                        )

                    if mail_res.status_code == 200:
                        msgs = mail_res.json().get("value", [])
                        for m in msgs:
                            m_id = m.get("id")
                            if m_id and m_id not in seen_ids:
                                seen_ids.add(m_id)
                                m["_folder"] = folder_label
                                collected_messages.append(m)
                    else:
                        try:
                            err_json = mail_res.json()
                            err_msg = err_json.get("error", {}).get("message", mail_res.text[:180])
                        except ValueError:
                            err_msg = mail_res.text[:180]
                        graph_errors.append(f"{folder_label} (HTTP {mail_res.status_code}): {err_msg}")
                except requests.RequestException as folder_err:
                    logger.warning(f"Error querying {folder_label} for {email_str}: {folder_err}")
                    graph_errors.append(f"{folder_label}: {folder_err}")

            # Sort newest first by receivedDateTime
            collected_messages.sort(key=lambda x: x.get("receivedDateTime", ""), reverse=True)

            # Stop early if any of the top recent messages contains an extractable OTP code
            has_otp = any(
                extract_otp(f"{m.get('subject', '')} {m.get('bodyPreview', '')}") is not None
                for m in collected_messages[:4]
            )
            if has_otp or attempt == max_attempts:
                break

            time.sleep(retry_delay_sec)

        if not collected_messages:
            if graph_errors:
                errs_formatted = "\n".join(f"• `{_escape_md(e)}`" for e in graph_errors[:3])
                return (
                    "❌ *Graph API Mail Fetch Issue:*\n"
                    f"{errs_formatted}\n\n"
                    "💡 *Common fixes:* Ensure the `refresh_token` is not expired/revoked, matches the `client_id`, "
                    "and has `Mail.Read` + `offline_access` permissions."
                )
            return (
                "📩 No messages found across Inbox, Junk Email, or All Messages.\n"
                "💡 *Tip:* If you know emails exist, verify that the OAuth token has `Mail.Read` scope "
                "and check whether the messages were just delivered."
            )

        out = []
        for msg in collected_messages[:6]:
            folder_tag = _escape_md(msg.get("_folder", "Mail"))
            subj = _escape_md(msg.get("subject") or "No Subject")
            received_raw = msg.get("receivedDateTime") or ""
            received_str = _escape_md(received_raw.replace("T", " ").replace("Z", " UTC")[:19]) if received_raw else "Unknown Time"
            sender_obj = msg.get("from", {}).get("emailAddress", {})
            sender_addr = _escape_md(sender_obj.get("address") or sender_obj.get("name") or "Unknown Sender")
            preview_raw = msg.get("bodyPreview") or ""
            preview = _escape_md(preview_raw[:180])
            code = extract_otp(f"{msg.get('subject', '')} {preview_raw}")
            code_str = f"\n🔑 *Extracted OTP:* `{code}`" if code else ""
            out.append(
                f"• *[{folder_tag}]* *{subj}*\n"
                f"  🕒 `{received_str}` | From: `{sender_addr}`{code_str}\n"
                f"  _{preview}_"
            )

        return "📧 *Recent Messages (via Microsoft Graph OAuth 2.0):*\n\n" + "\n\n".join(out)
    except requests.RequestException as req_err:
        logger.error(f"Network error during Graph OAuth check for {email_str}: {req_err}")
        return f"❌ *Network Error:* `{_escape_md(str(req_err))}`"


class DatabaseManager:
    def __init__(self, db_path: str = "accounts.db"):
        self.db_path = db_path
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self.db_path)

    def _init_db(self):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    user_id INTEGER PRIMARY KEY,
                    username TEXT,
                    quota INTEGER DEFAULT 0
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS accounts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT UNIQUE,
                    password TEXT,
                    refresh_token TEXT,
                    client_id TEXT,
                    status TEXT DEFAULT 'available',
                    assigned_to INTEGER
                )
            """)
            conn.commit()

    def register_user(self, user_id: int, username: str):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO users (user_id, username, quota)
                VALUES (?, ?, 0)
                ON CONFLICT(user_id) DO UPDATE SET username = excluded.username
                """,
                (user_id, username)
            )
            conn.commit()

    def _resolve_user_id(self, cur: sqlite3.Cursor, identifier: str) -> Optional[int]:
        clean = identifier.strip().lstrip("@")
        if clean.isdigit():
            uid = int(clean)
            cur.execute("SELECT user_id FROM users WHERE user_id = ?", (uid,))
            row = cur.fetchone()
            return row[0] if row else uid

        cur.execute("SELECT user_id FROM users WHERE LOWER(username) = LOWER(?)", (clean,))
        row = cur.fetchone()
        return row[0] if row else None

    def add_quota_by_identifier(self, identifier: str, amount: int) -> Tuple[bool, str]:
        """Batch-assigns up to `amount` available accounts directly to a user by ID or username."""
        if amount <= 0:
            return False, "Amount must be greater than 0."

        with self._connect() as conn:
            cur = conn.cursor()
            target_uid = self._resolve_user_id(cur, identifier)
            if not target_uid:
                return False, f"User `{_escape_md(identifier)}` not found in registered users. Ask them to run /start first."

            cur.execute(
                """
                SELECT id, email, password FROM accounts
                WHERE status = 'available'
                ORDER BY id ASC
                LIMIT ?
                """,
                (amount,)
            )
            rows = cur.fetchall()
            if not rows:
                return False, "No available accounts in inventory to assign."

            acc_ids = [r[0] for r in rows]
            placeholders = ",".join("?" for _ in acc_ids)
            cur.execute(
                f"""
                UPDATE accounts
                SET status = 'assigned', assigned_to = ?
                WHERE id IN ({placeholders})
                """,
                [target_uid, *acc_ids]
            )
            conn.commit()

            assigned_list = "\n".join(f"• `{em}` | `{pw}`" for _, em, pw in rows)
            return True, (
                f"Assigned *{len(rows)}* account(s) to user `{target_uid}`:\n\n{assigned_list}"
            )

    def add_account(
        self,
        email_addr: str,
        password: str,
        refresh_token: Optional[str] = None,
        client_id: Optional[str] = None
    ) -> bool:
        try:
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(
                    """
                    INSERT OR IGNORE INTO accounts (email, password, refresh_token, client_id, status)
                    VALUES (?, ?, ?, ?, 'available')
                    """,
                    (email_addr.strip(), password.strip(), refresh_token, client_id)
                )
                conn.commit()
                return cur.rowcount > 0
        except sqlite3.Error as e:
            logger.error(f"Database error adding account {email_addr}: {e}")
            return False

    def get_assigned_accounts_for_user(self, user_id: int) -> List[Tuple[int, str, str, Optional[str], Optional[str]]]:
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT id, email, password, refresh_token, client_id
                FROM accounts
                WHERE assigned_to = ? AND status = 'assigned'
                ORDER BY id ASC
                """,
                (user_id,)
            )
            return cur.fetchall()

    def mark_as_used(self, user_id: int) -> Optional[str]:
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT id, email FROM accounts
                WHERE assigned_to = ? AND status = 'assigned'
                ORDER BY id ASC
                LIMIT 1
                """,
                (user_id,)
            )
            row = cur.fetchone()
            if not row:
                return None

            acc_id, acc_email = row
            cur.execute(
                "UPDATE accounts SET status = 'used' WHERE id = ?",
                (acc_id,)
            )
            conn.commit()
            return acc_email

    def get_stats(self) -> Tuple[Dict[str, int], List[Tuple[int, str, int, int]]]:
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT status, COUNT(*) FROM accounts GROUP BY status")
            status_counts = {row[0]: row[1] for row in cur.fetchall()}

            cur.execute("""
                SELECT
                    u.user_id,
                    u.username,
                    SUM(CASE WHEN a.status = 'assigned' THEN 1 ELSE 0 END) AS assigned_cnt,
                    SUM(CASE WHEN a.status = 'used' THEN 1 ELSE 0 END) AS used_cnt
                FROM users u
                LEFT JOIN accounts a ON a.assigned_to = u.user_id
                GROUP BY u.user_id, u.username
                ORDER BY u.user_id ASC
            """)
            user_stats = [
                (r[0], r[1] or "", r[2] or 0, r[3] or 0)
                for r in cur.fetchall()
            ]
            return status_counts, user_stats


db = DatabaseManager()

# ---------------------------------------------------------------------------
# Telegram Bot Command Handlers
# ---------------------------------------------------------------------------
def track_activity(func):
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE, *args, **kwargs):
        if update.effective_user:
            db.register_user(update.effective_user.id, update.effective_user.username or "Unknown")
        return await func(update, context, *args, **kwargs)
    return wrapper

@track_activity
async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    u_id = update.effective_user.id
    assigned_accs = db.get_assigned_accounts_for_user(u_id)
    is_admin = u_id in ADMIN_IDS

    msg_lines = [
        "👋 *Welcome to Outlook Account Manager Bot!*\n",
        f"👤 *Your User ID:* `{u_id}`",
        f"📦 *Your Assigned Accounts:* `{len(assigned_accs)}` active account(s)\n",
        "📌 *User Commands:*",
        "/start - Show menu and assigned count",
        "/getacc - Claim an available account assigned to you",
        "/myaccs - List all accounts assigned to you",
        "/getcode - Retrieve current mailbox code/OTP (e.g. `/getcode` or `/getcode user@hotmail.com`)",
        "/done - Mark account complete (e.g. `/done` or `/done user@hotmail.com`)"
    ]

    if is_admin:
        msg_lines.extend([
            "\n👑 *Admin Commands:*",
            "/assign `<user_id|username> <amount>` - Directly batch-assign accounts to a user",
            "/addacc `<credentials>` - Add account(s) manually or upload `.txt` file",
            "/checkraw `<credentials>` - Test an account directly via OAuth/IMAP",
            "/inventory - View overall account inventory breakdown",
            "/stats - View all registered users, ID, assigned, and used counts"
        ])

    await update.message.reply_text("\n".join(msg_lines), parse_mode="Markdown")

@track_activity
async def myaccs_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Lists all accounts currently assigned to the user."""
    u_id = update.effective_user.id
    accs = db.get_assigned_accounts_for_user(u_id)
    
    if not accs:
        await update.message.reply_text(
            "❌ *No Accounts Assigned!*\n"
            "You currently have no active assigned accounts. Please contact the administrator.",
            parse_mode="Markdown"
        )
        return

    out = [f"📋 *Your Assigned Accounts ({len(accs)} total):*\n"]
    for idx, (acc_id, em, pw, r_tok, c_id) in enumerate(accs, 1):
        out.append(f"{idx}. 📧 `{em}` | 🔑 `{pw}`")

    out.append("\n💡 *To fetch OTP:* Run `/getcode <email>`\n💡 *To mark complete:* Run `/done <email>`")
    await update.message.reply_text("\n".join(out), parse_mode="Markdown")


@track_activity
async def getacc_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Allows a user to pull/view their assigned account(s) or claim one if available."""
    u_id = update.effective_user.id
    assigned_accs = db.get_assigned_accounts_for_user(u_id)

    if not assigned_accs:
        await update.message.reply_text(
            "❌ *No Assigned Accounts Available*\n"
            "You do not have any accounts assigned to you right now. Please contact the administrator.",
            parse_mode="Markdown"
        )
        return

    acc_id, em, pw, r_tok, c_id = assigned_accs[0]
    await update.message.reply_text(
        f"✅ *Assigned Account ({len(assigned_accs)} active total):*\n\n"
        f"📧 *Email:* `{em}`\n"
        f"🔑 *Password:* `{pw}`\n\n"
        "Use `/getcode` to fetch OTP codes or `/done` when finished.",
        parse_mode="Markdown"
    )


@track_activity
async def getcode_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    u_id = update.effective_user.id
    assigned_accs = db.get_assigned_accounts_for_user(u_id)
    
    if not assigned_accs:
        await update.message.reply_text("❌ You do not have any active assigned accounts. Contact admin.")
        return

    target_acc = None
    if context.args:
        query = context.args[0].strip().lower()
        for acc in assigned_accs:
            if query == acc[1].lower() or query == str(acc[0]):
                target_acc = acc
                break
        if not target_acc:
            await update.message.reply_text("❌ Specified email/ID is not found in your assigned active accounts.")
            return
    else:
        # Default to the first assigned account
        target_acc = assigned_accs[0]

    acc_id, em, pw, r_tok, c_id = target_acc
    status_msg = await update.message.reply_text(f"🔄 Checking mailbox for `{em}` via Microsoft Graph OAuth 2.0...", parse_mode="Markdown")

    result = fetch_otp_via_graph_checker(em, pw, r_tok, c_id)
    await status_msg.edit_text(f"📧 *Account:* `{em}`\n\n{result}", parse_mode="Markdown")

@track_activity
async def checkraw_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Allows admins to test account strings directly via OAuth 2.0 without saving to DB."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    if not context.args:
        await update.message.reply_text("Usage:\n`/checkraw email|password|refresh_token|client_id`", parse_mode="Markdown")
        return

    raw_input = update.message.text.partition(' ')[2].strip()
    parts = [p.strip() for p in raw_input.split('|')]

    if len(parts) < 2:
        await update.message.reply_text("❌ Format requires at least `email|password` (plus `refresh_token|client_id` for OAuth).")
        return

    em = parts[0]
    pw = parts[1]
    r_token = parts[2] if len(parts) >= 3 else None
    c_id = parts[3] if len(parts) >= 4 else None

    status_msg = await update.message.reply_text(f"🔄 Testing account `{em}` via Microsoft Graph OAuth 2.0...", parse_mode="Markdown")
    result = fetch_otp_via_graph_checker(em, pw, r_token, c_id)
    await status_msg.edit_text(result, parse_mode="Markdown")

@track_activity
async def checkraw_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Allows admins to test account strings directly without saving to DB."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    if not context.args:
        await update.message.reply_text("Usage:\n`/checkraw email|password|refresh_token|client_id`", parse_mode="Markdown")
        return

    raw_input = update.message.text.partition(' ')[2].strip()
    parts = [p.strip() for p in raw_input.split('|')]

    if len(parts) < 2:
        await update.message.reply_text("❌ Format requires at least `email|password`.")
        return

    em = parts[0]
    pw = parts[1]
    r_token = parts[2] if len(parts) >= 3 else None
    c_id = parts[3] if len(parts) >= 4 else None

    status_msg = await update.message.reply_text(f"🔄 Testing account `{em}`...")

    if r_token and c_id:
        result = fetch_otp_via_graph_checker(em, pw, r_token, c_id)
        if "OAUTH_FAILED" in result:
            imap_res = fetch_otp_via_imap(em, pw, r_token, c_id)
            result = f"{result}\n\n*IMAP Fallback Result:*\n{imap_res}"
    else:
        result = fetch_otp_via_imap(em, pw, None, None)

    await status_msg.edit_text(result, parse_mode="Markdown")

@track_activity
async def done_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    u_id = update.effective_user.id
    email_used = db.mark_as_used(u_id)
    if email_used:
        await update.message.reply_text(
            f"✅ *Account Completed!*\n\n"
            f"Marked account `{email_used}` as used and closed session.\n"
            f"You can now use `/getacc` to request your next account.",
            parse_mode="Markdown"
        )
    else:
        await update.message.reply_text("❌ You do not have an active assigned account to mark as completed.")

@track_activity
async def assign_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Allows admins to assign quota to a user by User ID or username."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    if len(context.args) < 2:
        await update.message.reply_text(
            "ℹ️ *Usage:*\n`/assign <user_id_or_username> <amount>`\n\n*Examples:*\n`/assign 12345678 10`\n`/assign @john_doe 5`",
            parse_mode="Markdown"
        )
        return

    target = context.args[0]
    try:
        amount = int(context.args[1])
    except ValueError:
        await update.message.reply_text("❌ Amount must be a valid integer.")
        return

    success, msg = db.add_quota_by_identifier(target, amount)
    prefix = "✅ " if success else "❌ "
    await update.message.reply_text(f"{prefix}{msg}", parse_mode="Markdown")

@track_activity
async def addacc_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles both inline text, /addacc command, and automatically processed .txt file documents."""
    if update.effective_user.id not in ADMIN_IDS:
        return

    raw_text = ""

    # Check if user uploaded a .txt file document
    if update.message.document:
        doc = update.message.document
        if doc.file_name.endswith(".txt") or doc.mime_type == "text/plain":
            file = await context.bot.get_file(doc.file_id)
            file_bytes = await file.download_as_bytearray()
            raw_text = file_bytes.decode('utf-8', errors='ignore')
        else:
            await update.message.reply_text("❌ Please upload a valid `.txt` file.")
            return
    else:
        # Read from message text after /addacc
        raw_text = update.message.text.partition(' ')[2].strip()

    if not raw_text:
        await update.message.reply_text(
            "ℹ️ *Usage:*\n"
            "1. Simply upload a `.txt` file into the chat.\n"
            "2. Or type `/addacc email|password|refresh_token|client_id`",
            parse_mode="Markdown"
        )
        return

    accounts = parse_credential_line(raw_text)
    
    if not accounts:
        await update.message.reply_text("❌ Could not parse any valid account formats from the input/file.")
        return

    added_count = 0
    for em, pw, r_tok, c_id in accounts:
        if db.add_account(em, pw, r_tok, c_id):
            added_count += 1

    await update.message.reply_text(
        f"✅ Parsed *{len(accounts)}* line(s) and added *{added_count}* new account(s) to inventory.",
        parse_mode="Markdown"
    )

@track_activity
async def inventory_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Displays total inventory metrics."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    status_counts, _ = db.get_stats()
    avail = status_counts.get("available", 0)
    assigned = status_counts.get("assigned", 0)
    used = status_counts.get("used", 0)
    total = avail + assigned + used

    msg = (
        "📦 *Account Inventory Summary*\n\n"
        f"🟢 *Available:* `{avail}`\n"
        f"🟡 *Assigned:* `{assigned}`\n"
        f"🔴 *Used:* `{used}`\n"
        f"📊 *Total:* `{total}`"
    )
    await update.message.reply_text(msg, parse_mode="Markdown")

@track_activity
async def stats_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Displays user breakdown (Username, ID, Assigned, Used)."""
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Unauthorized.")
        return

    _, user_stats = db.get_stats()

    msg_lines = ["👥 *User Statistics Breakdown*\n"]

    if not user_stats:
        msg_lines.append("_No registered users found._")
    else:
        for u_id, uname, assigned_cnt, used_cnt in user_stats:
            user_display = f"@{_escape_md(uname)}" if uname else "No Username"
            msg_lines.append(
                f"👤 *User:* {user_display}\n"
                f"🆔 *ID:* `{u_id}`\n"
                f"🟡 *Active Assigned Accs:* `{assigned_cnt}`\n"
                f"🔴 *Used Accs:* `{used_cnt}`\n"
                "─────────────────"
            )

    await update.message.reply_text("\n".join(msg_lines), parse_mode="Markdown")

# ---------------------------------------------------------------------------
# Application Initialization
# ---------------------------------------------------------------------------
def main():
    if not BOT_TOKEN:
        logger.error("BOT_TOKEN is missing in .env file! Exiting...")
        return

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("help", start_handler))
    app.add_handler(CommandHandler("getacc", getacc_handler))
    app.add_handler(CommandHandler("myaccs", myaccs_handler))
    app.add_handler(CommandHandler("getcode", getcode_handler))
    app.add_handler(CommandHandler("checkraw", checkraw_handler))
    app.add_handler(CommandHandler("done", done_handler))
    
    # Admin Handlers
    app.add_handler(CommandHandler("assign", assign_handler))
    app.add_handler(CommandHandler("addacc", addacc_handler))
    app.add_handler(CommandHandler("inventory", inventory_handler))
    app.add_handler(CommandHandler("stats", stats_handler))

    # Automatic .txt File Processing Handler (any .txt file document sent to the bot by admin)
    app.add_handler(MessageHandler(filters.Document.MimeType("text/plain") | filters.Document.FileExtension("txt"), addacc_handler))

    logger.info("Bot started successfully. Polling...")
    app.run_polling()

if __name__ == "__main__":
    main()