"""
invoice_service.py — Core business logic. Interface-agnostic.

This is the single function that Telegram, Web, or CLI will call.
It does not know HOW the file arrived or WHERE the result will be sent.
It just processes an invoice and returns paths to the output files.
"""

from typing import Any, Dict

from app.extractor.invoice_extractor import CommercialInvoiceExtractor
from app.generators.commercial_invoice import create_commercial_invoice
from app.generators.packing_list import create_packing_list
from app.utils.helpers import (
    validate_extracted_data,
    print_validation_report,
    generate_field_summary,
)
from storage.file_storage import save_json
from app.config import (
    ASSETS_DIR,
    OUTPUT_DIR,
    JSON_OUTPUT_FILENAME,
)


def process_invoice(excel_file_path: str,
                    assets_dir: str = ASSETS_DIR,
                    output_dir: str = OUTPUT_DIR) -> Dict[str, Any]:
    """
    Full pipeline: Excel → JSON → Commercial Invoice .docx → Packing List .docx

    Args:
        excel_file_path : absolute or relative path to the input .xls/.xlsx file
        assets_dir      : folder containing logo/slogan PNGs
        output_dir      : folder where all output files will be saved

    Returns:
        {
            "invoice_data"         : dict,          # raw extracted data
            "json_path"            : str,           # path to saved JSON
            "invoice_docx_path"    : str,           # path to Commercial Invoice .docx
            "packing_list_docx_path": str,          # path to Packing List .docx
            "validation"           : dict,          # errors & warnings
            "summary"              : str,           # human-readable field summary
        }
    """
    # Commented out — exposes file paths in production logs
    # print(f"\n{'='*60}")
    # print(f"Processing: {excel_file_path}")
    # print(f"{'='*60}")

    # ── Step 1: Extract ───────────────────────────────────────────────────────
    extractor = CommercialInvoiceExtractor(excel_file_path)
    extractor.load_file(sheet_name='INVOICE')
    invoice_data = extractor.extract_all_data(include_packing_list=True)

    # ── Step 2: Validate ──────────────────────────────────────────────────────
    validation = validate_extracted_data(invoice_data)
    # Commented out — exposes extracted invoice data in production logs
    # print_validation_report(validation)
    # print(generate_field_summary(invoice_data))

    # ── Step 3: Save JSON ─────────────────────────────────────────────────────
    json_path = save_json(invoice_data, output_dir, JSON_OUTPUT_FILENAME)

    # ── Step 4: Generate Documents ────────────────────────────────────────────
    invoice_docx_path      = create_commercial_invoice(invoice_data, assets_dir, output_dir)
    packing_list_docx_path = create_packing_list(invoice_data, assets_dir, output_dir)

    # Commented out — exposes output file paths in production logs
    # print(f"\n✓ All done.")
    # print(f"  JSON:          {json_path}")
    # print(f"  Invoice:       {invoice_docx_path}")
    # print(f"  Packing List:  {packing_list_docx_path}")

    return {
        "invoice_data":              invoice_data,
        "json_path":                 json_path,
        "invoice_docx_path":         invoice_docx_path,
        "packing_list_docx_path":    packing_list_docx_path,
        "validation":                validation,
        "summary":                   generate_field_summary(invoice_data),
    }
