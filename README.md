# itc-reco-tool

## Testing notes

### Regression scenario: BR0164900501-style invoice
Use a pair of rows where PR and 2B have the same normalized invoice number and equal taxable + tax components (IGST/CGST/SGST, with optional CESS differences within tolerance).

Expected result:
- Remark = `Matched`
- MatchMode can be `Invoice+Amount` or `InvoiceOnly` fallback, but must still be `Matched` when taxable/tax are within configured tolerances.
- `TaxDiff` should be `0.00` (or within tolerance), and export should include `ComputedTotalTax` and `ComputedCESS` for audit.

### Additional checks
- Leave CESS unmapped on one or both sides: reconciliation still runs and treats CESS as `0`.
- Provide only Taxable Value + IGST/CGST/SGST (without Total Tax): total tax is computed correctly.
- Provide Taxable Value + Total Tax (without IGST/CGST/SGST): reconciliation still runs.
- Check `Party Summary` tab: supports search and mismatch-only filter; export includes `Party_Summary` sheet.
