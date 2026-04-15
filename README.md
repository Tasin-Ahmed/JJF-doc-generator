# Commercial Invoice Generator

Extracts data from Excel invoice files and generates:
- `Commercial_Invoice.docx`
- `Packing_List.docx`

---

## Project Structure

```
project/
├── main.py                        ← local CLI entry point
├── app/
│   ├── config.py                  ← all settings (edit paths/secrets here)
│   ├── services/
│   │   └── invoice_service.py     ← core pipeline (extract → validate → generate)
│   ├── extractor/
│   │   └── invoice_extractor.py   ← reads Excel, returns structured data
│   ├── generators/
│   │   ├── base.py                ← shared XML helpers, fonts, page setup
│   │   ├── commercial_invoice.py  ← builds Commercial Invoice .docx
│   │   └── packing_list.py        ← builds Packing List .docx
│   └── utils/
│       └── helpers.py             ← validate, summarize (no file I/O)
├── interfaces/
│   ├── telegram_bot.py            ← plug in later (instructions inside)
│   └── web_app.py                 ← plug in later (instructions inside)
├── storage/
│   └── file_storage.py            ← all file I/O (swap to S3/GCS here later)
├── assets/                        ← put PNG logo files here
├── input/                         ← put Excel invoice files here
└── output/                        ← generated files appear here
```

---

## Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Copy and configure environment
cp .env.example .env
# Edit .env if you need to change folder paths

# 3. Add your logo files to assets/
#    JJF letter_head.png
#    JJF header bottom line.png
#    (Documents still generate without them — text fallback is used)
```

---

## Local Usage

```bash
# Drop your Excel file into input/ then run:
python main.py invoice.xlsx

# Output files appear in output/
#   output/invoice_data.json
#   output/Commercial_Invoice.docx
#   output/Packing_List.docx
```

---

## Adding Telegram (when ready)

1. `pip install python-telegram-bot`
2. Add `TELEGRAM_BOT_TOKEN=your_token` to `.env`
3. Uncomment `TELEGRAM_BOT_TOKEN` in `app/config.py`
4. Uncomment all code in `interfaces/telegram_bot.py`
5. Run: `python interfaces/telegram_bot.py`

---

## Adding Web Interface (when ready)

1. `pip install fastapi uvicorn python-multipart`
2. Uncomment all code in `interfaces/web_app.py`
3. Run: `uvicorn interfaces.web_app:app --host 0.0.0.0 --port 8000`

---

## Switching to Cloud Storage (when ready)

Only one file needs to change: `storage/file_storage.py`
Replace the local file read/write functions with S3/GCS calls.
Nothing else in the codebase changes.
