"""Build D365 Data management import workbooks for Customer payment journal lines."""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Optional

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None

_CONFIG_PATH = Path(__file__).resolve().parent / "config" / "dmf_customer_payment_journal_line_columns.json"

DMF_IMPORT_COLUMNS = [
    "JournalBatchNumber",
    "LineNumber",
    "dataAreaId",
    "AccountType",
    "AccountDisplayValue",
    "TransactionDate",
    "TransactionText",
    "DebitAmount",
    "CreditAmount",
    "CurrencyCode",
    "OffsetAccountType",
    "OffsetAccountDisplayValue",
    "PaymentMethodName",
    "PaymentReference",
]


def load_dmf_column_config() -> dict:
    if _CONFIG_PATH.exists():
        with open(_CONFIG_PATH, encoding="utf-8") as handle:
            return json.load(handle)
    return {"columns": [{"field": name} for name in DMF_IMPORT_COLUMNS]}


def _sanitize_amount(value) -> Optional[float]:
    if value is None:
        return None
    raw = str(value).replace("₹", "").replace(",", "").replace(" ", "").strip()
    if not raw:
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def _parse_date(value: str) -> Optional[datetime]:
    if not value or not str(value).strip():
        return None
    text = str(value).strip()
    formats = (
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d-%b-%Y",
        "%d-%B-%Y",
    )
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def find_duplicate_payment_references(records: list) -> list[str]:
    seen: dict[str, int] = {}
    duplicates: list[str] = []
    for record in records:
        ref = str(record.get("payment_reference", "")).strip()
        if not ref:
            continue
        seen[ref] = seen.get(ref, 0) + 1
        if seen[ref] == 2:
            duplicates.append(ref)
    return duplicates


def map_record_to_dmf_row(
    record: dict,
    *,
    journal_batch_number: str,
    company: str,
    line_number: int,
) -> dict:
    tx_date = (
        _parse_date(str(record.get("date", "")).strip())
        or _parse_date(str(record.get("value_date", "")).strip())
        or _parse_date(str(record.get("reference_date", "")).strip())
    )
    credit = _sanitize_amount(record.get("credit", ""))
    debit = _sanitize_amount(record.get("debit", ""))
    currency = str(record.get("currency", "")).strip() or "INR"
    offset_type = str(record.get("offset_account_type", "")).strip() or "Bank"
    return {
        "JournalBatchNumber": journal_batch_number,
        "LineNumber": line_number,
        "dataAreaId": str(record.get("company", "")).strip() or company,
        "AccountType": "Cust",
        "AccountDisplayValue": str(record.get("account", "")).strip(),
        "TransactionDate": tx_date,
        "TransactionText": str(record.get("description", "") or record.get("payment_reference", "")).strip(),
        "DebitAmount": debit if debit is not None else 0.0,
        "CreditAmount": credit if credit is not None else 0.0,
        "CurrencyCode": currency,
        "OffsetAccountType": offset_type,
        "OffsetAccountDisplayValue": str(record.get("offset_account", "")).strip(),
        "PaymentMethodName": str(record.get("method_of_payment", "")).strip(),
        "PaymentReference": str(record.get("payment_reference", "")).strip(),
    }


def build_import_workbook(
    records: list,
    journal_batch_number: str,
    company: str,
    *,
    output_path: Optional[Path] = None,
) -> Path:
    if Workbook is None:
        raise RuntimeError("openpyxl is required for DMF import workbooks.")
    if not journal_batch_number.strip():
        raise ValueError("journal_batch_number is required for DMF import.")
    if not company.strip():
        raise ValueError("company (dataAreaId) is required for DMF import.")

    duplicates = find_duplicate_payment_references(records)
    if duplicates:
        joined = ", ".join(duplicates[:5])
        suffix = "..." if len(duplicates) > 5 else ""
        print(f"Warning: Duplicate payment references in batch: {joined}{suffix}")

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(DMF_IMPORT_COLUMNS)

    for index, record in enumerate(records, start=1):
        row = map_record_to_dmf_row(
            record,
            journal_batch_number=journal_batch_number.strip(),
            company=company.strip(),
            line_number=index,
        )
        sheet.append([row.get(column) for column in DMF_IMPORT_COLUMNS])

    if output_path is None:
        safe_batch = re.sub(r"[^A-Za-z0-9_-]+", "_", journal_batch_number.strip())
        handle = tempfile.NamedTemporaryFile(
            prefix=f"sobha_dmf_{safe_batch}_",
            suffix=".xlsx",
            delete=False,
        )
        handle.close()
        output_path = Path(handle.name)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    print(f"DMF import workbook written: {output_path} ({len(records)} rows).")
    return output_path
