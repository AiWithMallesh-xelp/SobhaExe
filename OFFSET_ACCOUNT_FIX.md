# Offset Account Field Fix

## Problem
The automation was not filling the **Offset account** field in D365 journal entries, even though the data was present in the records.

## Root Cause
The `_process_sub_batch()` function in `automation.py` was missing the logic to:
1. Extract the offset account value from records
2. Locate the offset account input field
3. Fill the offset account field during automation

## Changes Made

### 1. Extract Offset Account Value
Added extraction of `offset_account` from record data:
```python
offset_acc = str(record.get("offset_account", "")).strip()
```

### 2. Locate Offset Account Field
Added locator for the offset account input field:
```python
offset_account_field = page.locator("input[id^='LedgerJournalTrans_OffsetAccount_'][id$='_input']")
if idx > 0:
    offset_account_field = offset_account_field.first
```

### 3. Fill Offset Account Field
Added filling logic after the credit field:
```python
# Fill Offset account field
if offset_acc:
    try:
        offset_account_field.wait_for(state="visible", timeout=5000)
        offset_account_field.click()
        offset_account_field.press("Control+A")
        offset_account_field.press("Backspace")
        offset_account_field.press_sequentially(offset_acc, delay=20)
        print(f"Filled offset account: {offset_acc}")
    except Exception as err:
        print(f"Warning: Could not fill offset account '{offset_acc}': {err}")
```

### 4. Updated Field Clearing Functions
Updated both fast clear and fallback clear functions to include offset account:

**Fallback Clear:**
```python
offset_account_clear = page.locator("input[id^='LedgerJournalTrans_OffsetAccount_'][id$='_input']")
# ... handle first/subsequent rows
_wipe(offset_account_clear)
```

**Fast Clear (JavaScript):**
```javascript
clearValue("input[id^='LedgerJournalTrans_OffsetAccount_'][id$='_input']")
```

**Verification Function:**
```javascript
read("input[id^='LedgerJournalTrans_OffsetAccount_'][id$='_input']")
```

### 5. Manual Wipe Before Fill
Added offset account clearing in the manual wipe logic:
```python
try:
    offset_account_field.click()
    offset_account_field.press("Control+A")
    offset_account_field.press("Backspace")
except Exception:
    pass
```

## Field Selector Details
The offset account field uses the following HTML attributes:
- **ID Pattern:** `LedgerJournalTrans_OffsetAccount_{index}_segmentedEntryLookup_input`
- **Role:** `combobox`
- **Aria-label:** `BankAccount` (or similar, varies by account type)

The selector `input[id^='LedgerJournalTrans_OffsetAccount_'][id$='_input']` matches this pattern.

## Testing Recommendations
1. Test with records that have offset account values
2. Test with records that have empty offset account values
3. Test with multiple sub-batches to ensure field clearing works correctly
4. Verify the field is filled before the "Save" button is clicked

## Data Flow
```
API Response → offset_account field
    ↓
sales_receipt_generation.py (UI displays it)
    ↓
automation.py (now fills it in D365)
```

## Files Modified
- `automation.py` - Added offset account handling in `_process_sub_batch()` function
