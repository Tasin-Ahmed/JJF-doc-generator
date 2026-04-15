"""
telegram_bot.py — Telegram interface for the Invoice Document Generator.

=====================================================================
  MODE: WEBHOOK (for Railway / cloud deployment)
=====================================================================

How it works:
  - Instead of polling Telegram servers 24/7 (which eats Railway hours),
    we run a lightweight FastAPI web server.
  - Telegram sends updates to our webhook URL via HTTP POST.
  - The server only wakes up when a message arrives → minimal compute usage.

To deploy on Railway:
  1. pip install -r requirements.txt
  2. Set these env vars in Railway dashboard:
       TELEGRAM_BOT_TOKEN = your bot token from @BotFather
       WEBHOOK_URL        = https://<your-app>.up.railway.app
  3. Railway will auto-detect the Procfile and start the server.

To run locally (for testing):
  python interfaces/telegram_bot.py

This file contains ZERO business logic.
All processing is delegated to app.services.invoice_service.process_invoice().
"""

# ==========================================================================
#  ORIGINAL POLLING-BASED CODE (COMMENTED OUT)
#  Kept for reference. This runs 24/7 and is NOT suitable for Railway.
# ==========================================================================
#
# import os
# import sys
#
# # Ensure the project root is on sys.path so `app` package is importable
# sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
#
# import logging
# from telegram import Update
# from telegram.ext import (
#     Application, CommandHandler, MessageHandler, filters, ContextTypes
# )
#
# from app.config import TELEGRAM_BOT_TOKEN, OUTPUT_DIR, ASSETS_DIR
# from app.services.invoice_service import process_invoice
#
# logging.basicConfig(level=logging.INFO)
# logger = logging.getLogger(__name__)
#
#
# async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
#     await update.message.reply_text(
#         "👋 Send me a commercial invoice Excel file (.xls or .xlsx)\n"
#         "I'll generate the Commercial Invoice and Packing List documents for you."
#     )
#
#
# async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
#     doc = update.message.document
#     if not doc.file_name.lower().endswith(('.xls', '.xlsx')):
#         await update.message.reply_text("❌ Please send an Excel file (.xls or .xlsx)")
#         return
#
#     await update.message.reply_text("⏳ Processing your invoice...")
#
#     # Download the file
#     file = await context.bot.get_file(doc.file_id)
#     input_path = os.path.join("input", doc.file_name)
#     await file.download_to_drive(input_path)
#
#     try:
#         result = process_invoice(
#             excel_file_path=input_path,
#             assets_dir=ASSETS_DIR,
#             output_dir=OUTPUT_DIR,
#         )
#
#         # Send back the two .docx files
#         await update.message.reply_document(open(result["invoice_docx_path"], 'rb'),
#                                             caption="✅ Commercial Invoice")
#         await update.message.reply_document(open(result["packing_list_docx_path"], 'rb'),
#                                             caption="✅ Packing List")
#
#         if result["validation"]["errors"]:
#             errors = "\n".join(f"• {e}" for e in result["validation"]["errors"])
#             await update.message.reply_text(f"⚠️ Validation issues:\n{errors}")
#
#     except Exception as e:
#         logger.error(f"Error processing invoice: {e}")
#         await update.message.reply_text(f"❌ Error: {e}")
#
#
# def main():
#     app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
#     app.add_handler(CommandHandler("start", start))
#     app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
#     logger.info("Bot started...")
#     app.run_polling()
#
#
# if __name__ == "__main__":
#     main()


# ==========================================================================
#  NEW WEBHOOK-BASED CODE (for Railway deployment)
# ==========================================================================

# ── Standard library imports ──────────────────────────────────────────────────
import os           # For file path operations (joining, checking existence)
import sys          # For modifying the Python module search path
import logging      # For structured logging output

# ── Add project root to sys.path ─────────────────────────────────────────────
# When this file runs from interfaces/, Python can't find the `app` package
# unless we explicitly add the project root (one level up) to the search path.
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# ── Third-party imports — Telegram ───────────────────────────────────────────
from telegram import Update, Bot                        # Core Telegram types
from telegram.ext import (                              # Bot framework components
    Application,                                        # Main bot application
    CommandHandler,                                     # Handles /start, /help etc.
    MessageHandler,                                     # Handles file uploads
    filters,                                            # Filters for message types
    ContextTypes,                                       # Type hints for handlers
)

# ── Third-party imports — FastAPI ────────────────────────────────────────────
from fastapi import FastAPI, Request                    # Web framework for webhook endpoint
from fastapi.responses import JSONResponse              # JSON response helper
from contextlib import asynccontextmanager              # For lifespan event management

# ── Project imports ──────────────────────────────────────────────────────────
from app.config import (
    TELEGRAM_BOT_TOKEN,     # Bot token from @BotFather (set in .env or Railway)
    OUTPUT_DIR,             # Directory where generated documents are saved
    ASSETS_DIR,             # Directory containing logo and slogan images
    PORT,                   # Server port (auto-injected by Railway, default 8000)
    WEBHOOK_URL,            # Public URL of your Railway app (set after first deploy)
)
from app.services.invoice_service import process_invoice  # Core business logic

# ── Configure logging ────────────────────────────────────────────────────────
# INFO level shows startup messages, incoming requests, and processing results.
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


# ==========================================================================
#  TELEGRAM BOT HANDLERS
#  These are the same handlers as the polling version — they define
#  what happens when a user sends /start or uploads a document.
# ==========================================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Handle /start command.
    Sends a welcome message explaining what the bot does.
    """
    await update.message.reply_text(
        "👋 Send me a commercial invoice Excel file (.xls or .xlsx)\n"
        "I'll generate the Commercial Invoice and Packing List documents for you."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Handle incoming document (file upload).
    Validates the file type, processes it through the invoice pipeline,
    and sends back the generated .docx files.
    """
    # Get the uploaded document metadata from the Telegram message
    doc = update.message.document

    # Validate file type — only accept Excel files
    if not doc.file_name.lower().endswith(('.xls', '.xlsx')):
        await update.message.reply_text("❌ Please send an Excel file (.xls or .xlsx)")
        return  # Stop processing if wrong file type

    # Acknowledge receipt — user knows the bot is working
    await update.message.reply_text("⏳ Processing your invoice...")

    # Download the file from Telegram servers to local disk
    file = await context.bot.get_file(doc.file_id)  # Get file object with download URL
    input_path = os.path.join("input", doc.file_name)  # Save to input/ directory
    await file.download_to_drive(input_path)  # Actually download the file

    try:
        # Run the full invoice processing pipeline:
        # Excel → JSON → Commercial Invoice .docx → Packing List .docx
        result = process_invoice(
            excel_file_path=input_path,     # Path to the downloaded Excel file
            assets_dir=ASSETS_DIR,          # Logo and slogan images
            output_dir=OUTPUT_DIR,          # Where to save generated documents
        )

        # Send the Commercial Invoice .docx back to the user
        await update.message.reply_document(
            open(result["invoice_docx_path"], 'rb'),    # Open file in binary read mode
            caption="✅ Commercial Invoice"              # Caption shown under the file
        )
        # Send the Packing List .docx back to the user
        await update.message.reply_document(
            open(result["packing_list_docx_path"], 'rb'),
            caption="✅ Packing List"
        )

        # If there were validation warnings, notify the user
        if result["validation"]["errors"]:
            errors = "\n".join(f"• {e}" for e in result["validation"]["errors"])
            await update.message.reply_text(f"⚠️ Validation issues:\n{errors}")

    except Exception as e:
        # Log the full error for debugging, send a clean message to the user
        logger.error(f"Error processing invoice: {e}")
        await update.message.reply_text(f"❌ Error: {e}")


# ==========================================================================
#  TELEGRAM APPLICATION (python-telegram-bot framework)
#  Build the bot application and register handlers.
#  This is shared between webhook and local-test modes.
# ==========================================================================

# Create the Telegram bot application using the builder pattern.
# .token() sets the bot's authentication token from environment config.
# .build() finalizes the application object.
ptb_app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()

# Register the /start command handler — triggers when user sends "/start"
ptb_app.add_handler(CommandHandler("start", start))

# Register the document handler — triggers when user uploads any file
ptb_app.add_handler(MessageHandler(filters.Document.ALL, handle_document))


# ==========================================================================
#  FASTAPI WEB SERVER (Webhook endpoint)
#  This is the web server that Railway runs. Telegram sends HTTP POST
#  requests to /webhook whenever a user messages the bot.
# ==========================================================================

@asynccontextmanager
async def lifespan(app: FastAPI):
    """
    FastAPI lifespan event handler.
    Runs setup code on startup and cleanup code on shutdown.

    On startup:
      - Initialize the Telegram bot application
      - Register the webhook URL with Telegram (so Telegram knows where to send updates)

    On shutdown:
      - Cleanly stop the Telegram bot application
    """
    # ── STARTUP ──────────────────────────────────────────────────────────────
    # Initialize the python-telegram-bot application (opens HTTP session, etc.)
    await ptb_app.initialize()
    # Start the application (begins processing updates from the queue)
    await ptb_app.start()

    # Register webhook with Telegram servers.
    # This tells Telegram: "Send all updates for this bot to this URL"
    if WEBHOOK_URL:
        # Construct the full webhook endpoint URL
        webhook_endpoint = f"{WEBHOOK_URL}/webhook"
        # Call Telegram's setWebhook API — all future updates go to this URL
        await ptb_app.bot.set_webhook(url=webhook_endpoint)
        logger.info(f"✅ Webhook registered: {webhook_endpoint}")
    else:
        # If no WEBHOOK_URL is set, warn — webhook won't work without it.
        # This is expected during local development.
        logger.warning(
            "⚠️  WEBHOOK_URL is not set. "
            "Webhook will NOT be registered with Telegram. "
            "Set WEBHOOK_URL in .env or Railway dashboard."
        )

    # yield passes control to the running application
    yield

    # ── SHUTDOWN ─────────────────────────────────────────────────────────────
    # Cleanly stop the bot (closes HTTP sessions, cancels pending tasks)
    await ptb_app.stop()
    await ptb_app.shutdown()


# Create the FastAPI app with lifespan management and metadata
fastapi_app = FastAPI(
    title="Invoice Generator Telegram Bot",     # Shown in auto-generated API docs
    description="Webhook-based Telegram bot for generating commercial invoices",
    lifespan=lifespan,                          # Use our startup/shutdown handler
)


@fastapi_app.get("/")
async def health_check():
    """
    Health check endpoint.
    Railway pings this to verify the service is alive.
    Returns a simple JSON response with status info.
    """
    return JSONResponse({
        "status": "ok",                                  # Service is running
        "service": "Invoice Generator Telegram Bot",     # Service name
        "mode": "webhook",                               # Confirms we're in webhook mode
    })


@fastapi_app.post("/webhook")
async def telegram_webhook(request: Request):
    """
    Webhook endpoint — receives Telegram updates via HTTP POST.

    Flow:
      1. Telegram sends a JSON payload containing the user's message
      2. We parse it into a Telegram Update object
      3. We feed it into the python-telegram-bot framework for processing
      4. The appropriate handler (start/handle_document) processes it
      5. We return 200 OK to Telegram so it knows we received the update

    This endpoint is called by Telegram's servers, NOT by your users.
    """
    # Parse the raw JSON body from Telegram's POST request
    json_data = await request.json()

    # Convert the raw JSON dict into a python-telegram-bot Update object
    # Update.de_json() deserializes the Telegram API response format
    update = Update.de_json(data=json_data, bot=ptb_app.bot)

    # Feed the update into the bot framework's processing pipeline.
    # This triggers the matching handler (CommandHandler or MessageHandler)
    # and runs the corresponding async function (start or handle_document).
    await ptb_app.process_update(update)

    # Return 200 OK — Telegram expects this to confirm receipt.
    # If we don't respond with 200, Telegram will retry the request.
    return JSONResponse({"status": "ok"})


# ==========================================================================
#  LOCAL DEVELOPMENT ENTRY POINT
#  When you run `python interfaces/telegram_bot.py` locally,
#  it starts the FastAPI server with uvicorn for testing.
# ==========================================================================

if __name__ == "__main__":
    import uvicorn  # ASGI server — runs FastAPI apps

    logger.info(f"🚀 Starting webhook server on port {PORT}...")
    logger.info(f"📡 Webhook URL: {WEBHOOK_URL or 'NOT SET (local dev mode)'}")

    # Start the uvicorn ASGI server:
    #   - "interfaces.telegram_bot:fastapi_app" tells uvicorn which app to serve
    #   - host="0.0.0.0" binds to all network interfaces (required for Railway)
    #   - port=PORT uses Railway's injected port or defaults to 8000
    #   - reload=True enables auto-restart on code changes (useful for local dev)
    uvicorn.run(
        "interfaces.telegram_bot:fastapi_app",  # Import path to the FastAPI app
        host="0.0.0.0",                         # Listen on all interfaces
        port=PORT,                              # Railway injects PORT, default 8000
        reload=True,                            # Auto-reload on file changes (dev only)
    )
