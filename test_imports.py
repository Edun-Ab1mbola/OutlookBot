import sys, os

print(f"Running Python version: {sys.version}")
print(f"CWD is: {os.getcwd()}")

try: 
    import pillow; print("✓ Pillow loaded successfully.") 
except ImportError as e: 
    print(f"✗ Pillow missing or wrong path: {e}")

try: 
    import selenium; print("✓ Selenium loaded.") 
except ImportError as e: 
    print(f"✗ Selenium missing. Run 'pip install selenium'.")

from pathlib import Path
DATA_DIR = Path("data") / "users_db.json"
if DATA_DIR.exists():
    with open(DATA_DIR) as f: data = json.load(f); print(f"DB file exists and contains keys: {list(data.keys())[:3]}...")
else:
    from pathlib import Path; DataDir.mkdir(parents=True, exist_ok=True); (Path("data")).joinpath("users_db.json").touch() # Create sample DB if empty

print("\n✅ All imports successful! Ready for bot.py.")
