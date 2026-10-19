# Sales & Tax ID Lookup & Data Cleaning Solutions

This document provides modern Excel, legacy Excel, and Python automated solutions for reconciling sales records (`orders_raw.csv`) with tax reference data (`tax_ids_reference.csv`).

---

## 1. Modern Excel Formula (for cell `orders!E2`)

```excel
=XLOOKUP(TRIM(B2), tax_ids_reference!$A:$A, tax_ids_reference!$C:$C, "Pending Verification")
```

### Key Elements:
* **`TRIM(B2)`**: Removes irregular leading, trailing, and duplicate spaces from the Customer ID in Column B.
* **`tax_ids_reference!$A:$A`**: Lookup array containing reference Customer IDs (Column A).
* **`tax_ids_reference!$C:$C`**: Return array containing Tax IDs (Column C).
* **`"Pending Verification"`**: Built-in 4th argument (`[if_not_found]`) of `XLOOKUP` that automatically replaces `#N/A` for unregistered IDs (e.g., `CUST999`).

### Table / Structured Reference Equivalent:
If formatted as Excel Tables:
```excel
=XLOOKUP(TRIM([@[Customer ID]]), TaxReferenceTable[Customer ID], TaxReferenceTable[Tax ID], "Pending Verification")
```

---

## 2. Legacy Excel Formula (Pre-Excel 2021 / Pre-Office 365)

```excel
=IFERROR(INDEX(tax_ids_reference!$C:$C, MATCH(TRIM(B2), tax_ids_reference!$A:$A, 0)), "Pending Verification")
```

### Execution Flow:
1. **`TRIM(B2)`**: Cleans the Customer ID lookup value.
2. **`MATCH(TRIM(B2), tax_ids_reference!$A:$A, 0)`**: Locates the exact row index in Column A.
3. **`INDEX(tax_ids_reference!$C:$C, ...)`**: Retrieves the corresponding Tax ID from Column C.
4. **`IFERROR(..., "Pending Verification")`**: Traps `#N/A` errors for unregistered customers and returns the fallback status.

---

## 3. Concise 5-Line Python pandas Pipeline

Implemented in [`clean_and_merge.py`](clean_and_merge.py):

```python
import pandas as pd

orders, tax = pd.read_csv('orders_raw.csv'), pd.read_csv('tax_ids_reference.csv')
orders['Customer ID'] = orders['Customer ID'].astype(str).str.strip()
merged = orders.merge(tax[['Customer ID', 'Tax ID']], on='Customer ID', how='left').fillna({'Tax ID': 'Pending Verification'})
merged.to_csv('orders_enriched.csv', index=False)
```

---

## Enterprise Edge Case: Non-Breaking Spaces (`CHAR(160)`)

When data originates from web scrapers or ERP exports, spaces may be non-breaking (`ASCII 160` / `&nbsp;`), which standard `TRIM()` does not remove.
* **Excel Fix**: `TRIM(SUBSTITUTE(B2, CHAR(160), " "))`
* **Python Fix**: `orders['Customer ID'] = orders['Customer ID'].astype(str).str.replace('\xa0', ' ').str.strip()`
