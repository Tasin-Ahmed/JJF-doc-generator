"""
config.py — Single source of truth for all settings and paths.

To switch from local to cloud:
  - Update ASSETS_DIR, INPUT_DIR, OUTPUT_DIR here (or via .env)
  - Update storage/file_storage.py for cloud I/O

To add Telegram or Web secrets:
  - Add them to .env and load them here.
"""

import os
from dotenv import load_dotenv

load_dotenv()

# ── Directory Paths ───────────────────────────────────────────────────────────
ASSETS_DIR = os.getenv("ASSETS_DIR", "assets/")
INPUT_DIR  = os.getenv("INPUT_DIR",  "input/")
OUTPUT_DIR = os.getenv("OUTPUT_DIR", "output/")

# ── Asset File Paths ──────────────────────────────────────────────────────────
LOGO_PATH   = os.path.join(ASSETS_DIR, "JJF letter_head.png")
SLOGAN_PATH = os.path.join(ASSETS_DIR, "JJF header bottom line.png")

# ── Output File Names ─────────────────────────────────────────────────────────
JSON_OUTPUT_FILENAME            = "invoice_data.json"
COMMERCIAL_INVOICE_FILENAME     = "Commercial_Invoice.docx"
PACKING_LIST_FILENAME           = "Packing_List.docx"

# ── Telegram (uncomment when ready) ──────────────────────────────────────────
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")

# ── Railway / Webhook Deployment ─────────────────────────────────────────────
# PORT is automatically injected by Railway at runtime. Default 8000 for local dev.
PORT = int(os.getenv("PORT", 8000))
# WEBHOOK_URL is your Railway app's public URL (e.g. https://your-app.up.railway.app).
# Set this in Railway dashboard after your first deploy. Leave empty for local dev.
WEBHOOK_URL = os.getenv("WEBHOOK_URL", "")

# ── Web (uncomment when ready) ────────────────────────────────────────────────
# WEB_HOST = os.getenv("WEB_HOST", "0.0.0.0")
# WEB_PORT = int(os.getenv("WEB_PORT", 8000))
