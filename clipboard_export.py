"""Shared clipboard / Excel export helpers for UI copy and bulk-paste automation."""
from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Optional

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None

DEFAULT_CLIPBOARD_COL_DEFS = [
    ("date", "Date"),
    ("value_date", "Value date"),
    ("voucher", "Voucher"),
    ("company", "Company"),
    ("account", "Account"),
    ("account_name", "Account name"),
    ("payee_name", "Payee name"),
    ("invoice", "Invoice"),
    ("description", "Description"),
    ("debit", "Debit"),
    ("credit", "Credit"),
    ("currency", "Currency"),
    ("sales_order_id", "Sales order id"),
    ("bank_account", "Bank account"),
    ("offset_account_type", "Offset account type"),
    ("offset_account", "Offset account"),
    ("method_of_payment", "Method of payment"),
    ("payment_status", "Payment status"),
    ("demand_number", "Demand number"),
    ("reference_date", "Reference date"),
    ("payment_reference", "Payment reference"),
    ("use_deposit_slip", "Use a deposit slip"),
    ("crm_transaction_type", "CRM transaction type"),
    ("original_payment_voucher", "Original Payment Voucher"),
    ("reversal_payment_voucher", "Reversal Payment Voucher"),
]

TRANSACTION_FIELD_KEYS = [key for key, _label in DEFAULT_CLIPBOARD_COL_DEFS]
TRANSACTION_FIELD_LABELS = {key: label for key, label in DEFAULT_CLIPBOARD_COL_DEFS}

EXTRA_FIELD_LABELS = {
    "transaction_uuid": "Transaction UUID (API)",
    "account_date": "Account date (API)",
    "account_number": "Account number (API)",
    "transaction_amount": "Transaction amount (API)",
    "transaction_description": "Transaction description (API)",
    "mode_of_transaction": "Mode of transaction (API)",
    "receipt_number": "Receipt number (API)",
    "txn_source": "Txn source (API)",
    "created_at": "Created at (API)",
    "updated_at": "Updated at (API)",
    "created_by": "Created by (API)",
}
PRESET_FIELD_LABELS = {**TRANSACTION_FIELD_LABELS, **EXTRA_FIELD_LABELS}
PRESET_FIELD_KEYS = list(dict.fromkeys(TRANSACTION_FIELD_KEYS + list(EXTRA_FIELD_LABELS)))


def normalize_col_def(item) -> dict:
    if isinstance(item, dict):
        key = str(item.get("key", "")).strip()
        label = str(item.get("label", TRANSACTION_FIELD_LABELS.get(key, key))).strip() or key
        return {
            "key": key,
            "label": label,
            "default_value": str(item.get("default_value", "") or ""),
        }
    if isinstance(item, (list, tuple)) and len(item) >= 2:
        key = str(item[0]).strip()
        label = str(item[1]).strip() or key
        default_value = str(item[2]).strip() if len(item) > 2 else ""
        return {"key": key, "label": label, "default_value": default_value}
    raise ValueError("Invalid column definition")


def normalize_col_defs(col_defs: list) -> list:
    normalized = []
    for item in col_defs:
        col = normalize_col_def(item)
        if col.get("key"):
            normalized.append(col)
    return normalized


def slug_custom_field_key(header: str, used_keys: set) -> str:
    base = re.sub(r"[^a-z0-9]+", "_", header.strip().lower()).strip("_") or "custom_field"
    if not base.startswith("custom_"):
        base = f"custom_{base}"
    candidate = base
    suffix = 2
    while candidate in used_keys:
        candidate = f"{base}_{suffix}"
        suffix += 1
    return candidate


def resolve_clipboard_cell(txn: dict, col: dict) -> str:
    key = col.get("key", "")
    value = txn.get(key, "")
    if value is None or str(value).strip() == "":
        return str(col.get("default_value", "") or "")
    return str(value)


def clipboard_columns_path() -> Path:
    env_path = os.environ.get("SOBHA_CLIPBOARD_COLUMNS_PATH")
    if env_path:
        return Path(env_path).expanduser()
    return Path.home() / ".config" / "sobha-reconciliation" / "clipboard_columns.json"


def load_clipboard_col_defs() -> list:
    defaults = [normalize_col_def(item) for item in DEFAULT_CLIPBOARD_COL_DEFS]
    path = clipboard_columns_path()
    if not path.exists():
        return defaults
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list) or not data:
            return defaults
        loaded = normalize_col_defs(data)
        return loaded or defaults
    except Exception:
        return defaults


def save_clipboard_col_defs(col_defs: list) -> None:
    path = clipboard_columns_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = [normalize_col_def(item) for item in col_defs]
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)


def excel_escape_cell(value) -> str:
    text = str(value if value is not None else "")
    if any(ch in text for ch in ("\t", "\n", "\r", '"')):
        return '"' + text.replace('"', '""') + '"'
    return text


def _build_data_rows(transactions: list, col_defs: Optional[list] = None) -> tuple[list, list]:
    columns = normalize_col_defs(col_defs or load_clipboard_col_defs())
    headers = [col["label"] for col in columns]
    data_rows = [
        [resolve_clipboard_cell(txn, col) for col in columns]
        for txn in transactions
    ]
    return headers, data_rows


def build_excel_clipboard_text(transactions: list, col_defs: Optional[list] = None) -> str:
    """Build tab-separated rows (Excel-ready) with header row."""
    headers, data_rows = _build_data_rows(transactions, col_defs)

    if Workbook is not None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Transactions"
        worksheet.append(headers)
        for row in data_rows:
            worksheet.append(row)
        rows = list(worksheet.iter_rows(values_only=True))
    else:
        rows = [tuple(headers)] + [tuple(row) for row in data_rows]

    lines = [
        "\t".join(excel_escape_cell(cell) for cell in row)
        for row in rows
    ]
    return "\r\n".join(lines)


def col_defs_for_d365_paste(col_defs: Optional[list] = None) -> list:
    """Return exactly the 25 D365 journal columns in grid order (no custom extras)."""
    source = {col["key"]: col for col in normalize_col_defs(col_defs or load_clipboard_col_defs())}
    return [
        source.get(
            key,
            {"key": key, "label": label, "default_value": ""},
        )
        for key, label in DEFAULT_CLIPBOARD_COL_DEFS
    ]


def build_paste_clipboard_text(transactions: list, col_defs: Optional[list] = None) -> str:
    """Build tab-separated data rows only (no header) for D365 grid paste."""
    d365_cols = col_defs_for_d365_paste(col_defs)
    _headers, data_rows = _build_data_rows(transactions, d365_cols)
    lines = [
        "\t".join(excel_escape_cell(cell) for cell in row)
        for row in data_rows
    ]
    return "\r\n".join(lines)
