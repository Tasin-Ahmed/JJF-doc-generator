"""
helpers.py — Utility functions: validation, JSON export, field summary.

These are pure functions — no file I/O paths, no interface dependencies.
File saving is handled by storage/file_storage.py.
"""

import json
from typing import Any, Dict


def validate_extracted_data(invoice_data: Dict[str, Any]) -> Dict[str, list]:
    """Validate extracted invoice data. Returns dict with 'errors' and 'warnings'."""
    issues = {'errors': [], 'warnings': []}

    if not invoice_data['invoice_details']['invoice_no']:
        issues['errors'].append("Invoice number is missing")
    if not invoice_data['invoice_details']['invoice_date']:
        issues['errors'].append("Invoice date is missing")
    if not invoice_data['shipper']['name']:
        issues['errors'].append("Shipper name is missing")
    if not invoice_data['goods']['description']:
        issues['errors'].append("Goods description is missing")
    if invoice_data['goods']['total_amount_usd'] <= 0:
        issues['errors'].append("Total amount is zero or missing")

    if not invoice_data['notify_party']['name']:
        issues['warnings'].append("Notify party name is missing")
    if not invoice_data['containers']:
        issues['warnings'].append("No container information found")
    if not invoice_data['amount_in_words']:
        issues['warnings'].append("Amount in words is missing")
    if not invoice_data['port_info']['port_of_loading']:
        issues['warnings'].append("Port of loading is missing")
    if not invoice_data['port_info']['port_of_discharge']:
        issues['warnings'].append("Port of discharge is missing")

    return issues


def print_validation_report(issues: Dict[str, list]):
    """Print a formatted validation report to stdout."""
    print("\n" + "=" * 60)
    print("VALIDATION REPORT")
    print("=" * 60)

    if not issues['errors'] and not issues['warnings']:
        print("\n✓ All validations passed!")
    else:
        if issues['errors']:
            print(f"\n❌ ERRORS ({len(issues['errors'])})")
            for e in issues['errors']:
                print(f"   • {e}")
        if issues['warnings']:
            print(f"\n⚠️  WARNINGS ({len(issues['warnings'])})")
            for w in issues['warnings']:
                print(f"   • {w}")

    print("=" * 60)


def generate_field_summary(invoice_data: Dict[str, Any]) -> str:
    """Generate a human-readable completeness summary of all extracted fields."""
    total_fields  = 0
    filled_fields = 0

    def count(data):
        nonlocal total_fields, filled_fields
        if isinstance(data, dict):
            for v in data.values():
                if isinstance(v, (dict, list)):
                    count(v)
                else:
                    total_fields += 1
                    if v and v != "" and v != 0:
                        filled_fields += 1
        elif isinstance(data, list):
            for item in data:
                count(item)

    count(invoice_data)
    pct = (filled_fields / total_fields * 100) if total_fields else 0

    lines = [
        "=" * 70,
        "INVOICE DATA SUMMARY",
        "=" * 70,
        f"\nData Completeness: {filled_fields}/{total_fields} fields filled ({pct:.1f}%)",
        "",
        "KEY INFORMATION:",
        f"  Invoice No:   {invoice_data['invoice_details']['invoice_no']}",
        f"  Invoice Date: {invoice_data['invoice_details']['invoice_date']}",
        f"  Shipper:      {invoice_data['shipper']['name']}",
        f"  Consignee:    {invoice_data['notify_party']['name']}",
        f"  Total Amount: ${invoice_data['goods']['total_amount_usd']:,.2f}",
        f"  Quantity:     {invoice_data['goods']['quantity_mt']} MT",
        f"  Containers:   {len(invoice_data['containers'])}",
        "=" * 70,
    ]
    return "\n".join(lines)


def invoice_data_to_json_string(invoice_data: Dict[str, Any]) -> str:
    """Serialize invoice data to a JSON string (used by file_storage)."""
    return json.dumps(invoice_data, indent=2, ensure_ascii=False)
