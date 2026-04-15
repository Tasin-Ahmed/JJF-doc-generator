"""
file_storage.py — All file I/O in one place.

Today:  reads/writes to local disk.
Later:  swap this file to upload/download from S3, GCS, or Azure Blob.
        The rest of the codebase never needs to change.
"""

import os
import json
from typing import Any, Dict


def ensure_dir(directory: str):
    """Create directory if it doesn't exist."""
    os.makedirs(directory, exist_ok=True)


def save_json(data: Dict[str, Any], output_dir: str, filename: str) -> str:
    """Save invoice data as a JSON file. Returns the full output path."""
    ensure_dir(output_dir)
    path = os.path.join(output_dir, filename)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"✓ Saved JSON: {path}")
    return path


def load_json(output_dir: str, filename: str) -> Dict[str, Any]:
    """Load invoice data from a JSON file. Returns parsed dict."""
    path = os.path.join(output_dir, filename)
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    print(f"✓ Loaded JSON: {path}")
    return data


def save_docx(doc, output_dir: str, filename: str) -> str:
    """Save a python-docx Document object to disk. Returns the full output path."""
    ensure_dir(output_dir)
    path = os.path.join(output_dir, filename)
    doc.save(path)
    print(f"✓ Saved DOCX: {path}")
    return path


def get_input_file_path(input_dir: str, filename: str) -> str:
    """Resolve full path of an input file and verify it exists."""
    path = os.path.join(input_dir, filename)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Input file not found: {path}")
    return path
