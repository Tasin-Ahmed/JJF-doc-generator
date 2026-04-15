"""
main.py — Local CLI entry point for testing.

Usage:
    python main.py input/your_invoice.xlsx

This is ONLY used for local testing.
When you move to Telegram or Web, use interfaces/telegram_bot.py or interfaces/web_app.py instead.
They all call the same app.services.invoice_service.process_invoice() under the hood.
"""

import sys
import os

# Allow imports from project root
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.config import ASSETS_DIR, INPUT_DIR, OUTPUT_DIR
from app.services.invoice_service import process_invoice


def main():
    if len(sys.argv) < 2:
        print("Usage: python main.py <excel_filename>")
        print("Example: python main.py invoice.xlsx")
        print(f"\nDrop your Excel file into the '{INPUT_DIR}' folder first.")
        sys.exit(1)

    filename       = sys.argv[1]
    excel_file_path = filename if os.path.isabs(filename) else os.path.join(INPUT_DIR, filename)

    if not os.path.exists(excel_file_path):
        print(f"❌ File not found: {excel_file_path}")
        print(f"   Make sure your file is in the '{INPUT_DIR}' folder.")
        sys.exit(1)

    result = process_invoice(
        excel_file_path=excel_file_path,
        assets_dir=ASSETS_DIR,
        output_dir=OUTPUT_DIR,
    )

    if result["validation"]["errors"]:
        print("\n⚠️  There were validation errors. Check the output files carefully.")
        sys.exit(0)

    print("\n✅ Processing complete. Check the 'output/' folder.")


if __name__ == "__main__":
    main()
