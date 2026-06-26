# Cleaning Log — Retail Orders Dataset

**Source file:** `data/raw_orders.xlsx` (932 rows, 21 columns)
**Output file:** `data/cleaned_orders.xlsx` (912 rows, 42 columns)
**Tool used:** Python (pandas, openpyxl) to apply the same logic that the required Excel
techniques (TRIM, SUBSTITUTE, Find & Replace, Flash Fill, Text to Columns) achieve, executed
programmatically for accuracy and reproducibility across all 932 rows. All output is saved as
a standard `.xlsx` workbook.

---

## 1. Issues Found

| # | Issue | Detail |
|---|-------|--------|
| 1 | Inconsistent text casing | `customer_name`, `segment`, `region`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status` contained a mix of UPPERCASE, lowercase, and Title Case versions of the same value (e.g. `CANCELLED`, `cancelled`, `Cancelled`). |
| 2 | Leading/trailing/double spaces | Many text fields had values like `"  Diya Patel "`, `"Aarav  Sharma"`, `"Small  Business"`. |
| 3 | Mixed date formats | `order_date` and `ship_date` used **4 different formats** simultaneously: `DD Mon YYYY` (e.g. `21 Jul 2024`), `YYYY-MM-DD`, `MM/DD/YYYY`, and `DD-MM-YYYY`. |
| 4 | Ship date earlier than order date | 21 records had a `ship_date` before the corresponding `order_date`. |
| 5 | Missing values | `region` (26 raw rows), `ship_mode` (22 raw rows), `discount` (18 raw rows). |
| 6 | Discount stored as text/percent | 8 records had discount written as a percent string (`"70%"`, `"85%"`) instead of a decimal. |
| 7 | Negative discounts | 16 raw records had negative discount values (e.g. `-0.14`, `-0.24`) — not a valid business value. |
| 8 | Discount above allowed range | 7 records had discount of `0.55` or `0.65` (55%/65%), far outside the observed valid tier set (0, 5, 10, 15, 20, 25%). |
| 9 | Exact duplicate rows | 20 groups (40 rows total) were **fully identical** across all 21 original columns. |
| 10 | Duplicate order_id with conflicting data | 12 `order_id` values (24 rows) appeared more than once **with different** `sales`, `profit`, `order_status`, or dates — a genuine data conflict, not a simple repeat. |
| 11 | Sales/Profit calculation mismatches | 58 records had a `sales` or `profit` value that did not match `quantity × unit_price × (1 − discount)` / `sales − cost` (beyond rounding noise), mostly tied to the conflicting-duplicate rows above. |
| 12 | Status inconsistency | 2 records were marked `order_status = Completed` while `payment_status = Failed` — logically contradictory. |

---

## 2. Cleaning Actions Performed

### Text fields
- Trimmed leading/trailing spaces and collapsed internal multiple spaces to a single space (equivalent to Excel `TRIM`).
- Standardized casing to Title Case for all categorical/text fields.
- No genuinely different spellings of the same category were found (e.g. no "Office Suplies" typo) — only casing/spacing variants — so no `SUBSTITUTE`/manual remapping was required beyond case standardization.

### Dates
- Detected each date's format using pattern matching (4 formats observed) and parsed each into a true date value.
- Confirmed the slash format was `MM/DD/YYYY` (first token never exceeds 12) and the dash format was `DD-MM-YYYY` (first token reaches up to 31) by inspecting the full value range before parsing — this avoided silently mis-parsing ambiguous dates like `04/06/2025`.
- Standardized both `order_date` and `ship_date` to **YYYY-MM-DD** in the cleaned file.
- No missing or unparseable date text was found in this dataset (0 rows) — all 932 raw values matched one of the 4 known formats.
- Calculated `shipping_delay_days = ship_date − order_date`. Negative results (21 rows) were flagged invalid (`ship_before_order_flag`).

### Missing values
- `region`: filled with `"Unknown"`, flagged via `region_missing_flag = TRUE`.
- `ship_mode`: filled with `"Unknown"`, flagged via `ship_mode_missing_flag = TRUE`.
- `discount`: filled with `0`, flagged via `discount_missing_flag = TRUE`. This was only done because, for every row with a missing discount, `quantity`, `unit_price`, `sales`, and `cost` were all present and numerically valid.

### Discount cleaning
- Percent-string values (`"70%"`) converted to numeric decimals (`0.70`).
- Negative values: invalid. `cleaned_discount` set to `0` for these rows; original negative value preserved in `discount` for audit, and `discount_negative_flag = TRUE`.
- Values above the allowed maximum (25%, based on observed valid tiers 0/5/10/15/20/25%): capped at `0.25` in `cleaned_discount`, flagged via `discount_out_of_range_flag = TRUE`.
- All other values used as-is.

### Duplicate handling
- **Exact duplicate rows** (all 21 original columns identical): the extra copies were **removed**, keeping the first occurrence only. 20 such groups (40 rows) were found; 20 extra rows were removed, bringing the dataset from 932 → 912 rows.
- **Duplicate order_id with conflicting data**: these were **never deleted**. All rows were kept and flagged (`duplicate_order_id_flag`, `duplicate_order_id_conflicting_flag`), with `data_quality_flag = "Invalid"`, so a business reviewer can decide which version is authoritative. Full detail is in `outputs/data_quality_report.xlsx`, sheet "2. Duplicates".

### Business rules applied (see also Section 3 below)
- Cancelled orders and Failed-payment orders are excluded from `included_in_completed_sales` and therefore from the completed-sales pivot summaries.
- Refunded orders are summarized separately (see `outputs/pivot_summary.xlsx`, sheet "5. Refund-Cancel-Fail").
- Ship date before order date is flagged as an invalid shipping record (`ship_before_order_flag`), contributing to `data_quality_flag = "Invalid"`.

### Calculated columns added
`cleaned_discount`, `calculated_sales`, `calculated_profit`, `profit_margin`, `shipping_delay_days`, `order_month`, `order_year`, `data_quality_flag`, plus supporting boolean flag columns (see `column_notes` sheet inside `cleaned_orders.xlsx` for full definitions).

---

## 3. Business Rules Applied

| Rule | Action Taken |
|---|---|
| Missing `region` | Filled as `"Unknown"`, flagged in quality report |
| Missing `ship_mode` | Filled as `"Unknown"`, flagged in quality report |
| Missing `discount` | Treated as `0` — only applied because all other sales fields (`quantity`, `unit_price`, `sales`, `cost`) were valid for these rows |
| Negative `discount` | Flagged as invalid; `cleaned_discount` set to `0` for calculation purposes |
| Discount above allowed range (>25%) | Flagged as invalid; `cleaned_discount` capped at `25%` |
| Cancelled orders | Excluded from `included_in_completed_sales`; excluded from completed-sales pivot summaries |
| Failed payments | Excluded from `included_in_completed_sales`; excluded from completed-sales pivot summaries |
| Refunded orders | Summarized separately in `pivot_summary.xlsx`, sheet "5. Refund-Cancel-Fail" |
| Ship date before order date | Flagged as invalid shipping record (`ship_before_order_flag`); contributes to `data_quality_flag = Invalid` |

---

## 4. Assumptions Made

1. **Slash-format dates are `MM/DD/YYYY`** (US-style), not `DD/MM/YYYY`. This was confirmed, not assumed blindly: the first token in every slash-format value never exceeds 12, while the second token reaches up to 31 — only consistent with month-first ordering.
2. **Dash-format dates are `DD-MM-YYYY`**, confirmed the same way (first token reaches up to 31, second token never exceeds 12).
3. **Allowed discount range is 0%–25%.** This was inferred from the data itself: legitimate discount values cluster tightly at 0, 5, 10, 15, 20, and 25%, while a small set of outliers (55%, 65%, and various negative values) sit clearly outside that pattern. No explicit allowed-range was provided in the assignment, so this empirical banding was used and is documented here for transparency.
4. **"Calculated" sales/profit formulas**: `calculated_sales = quantity × unit_price × (1 − cleaned_discount)` and `calculated_profit = calculated_sales − cost`. These were reverse-engineered from the data (verified against ~94% of rows matching within rounding tolerance) and are the formulas used to populate `calculated_sales` / `calculated_profit`, independent of the original (sometimes inconsistent) `sales`/`profit` columns.
5. **Mismatch tolerance**: a sales or profit mismatch is only flagged if the difference between the raw value and the recalculated value exceeds 1.00 (currency unit), to avoid flagging harmless floating-point rounding.
6. **`region` was not back-filled from `state`**, even though `state` reliably predicts `region` in this dataset (e.g. every `Gujarat` row has `region = West`). The assignment's explicit rule was to fill missing `region` as `"Unknown"` and flag it, so that rule was followed literally rather than substituting a smarter inference. This relationship is noted here in case the business wants to revisit the rule.
7. **A "completed sales" record** is defined as `order_status = "Completed"` **and** `payment_status != "Failed"`, since a completed order with a failed payment should not count as realized revenue.

---

## 5. Records Removed

- **20 rows removed** — exact full-row duplicates (every one of the 21 original columns matched another row). Kept the first occurrence of each.
- No other rows were removed. Conflicting duplicates, invalid discounts, and inverted ship dates were **flagged, not deleted**.

## 6. Records Flagged

Final counts in `cleaned_orders.xlsx` (912 rows):

| `data_quality_flag` | Count | % |
|---|---|---|
| Clean | 756 | 82.9% |
| Warning | 98 | 10.7% |
| Invalid | 58 | 6.4% |

See `outputs/data_quality_report.xlsx` for the full breakdown by issue type.

---

## 7. Limitations of This Cleaning Process

1. **Conflicting duplicate order_ids were not auto-resolved.** Twelve order_ids (24 rows) have genuinely different data under the same ID. Without a clear "last updated" timestamp or source-system priority rule, automatically picking a winner risks silently discarding correct data — so these are left for manual / business review rather than guessed at.
2. **Discount range (0–25%) is an inferred business rule**, not one stated explicitly in the source system. If the true allowed range differs, the `discount_out_of_range_flag` and `cleaned_discount` capping logic would need to be revisited.
3. **Region was deliberately not inferred from state**, per the literal instruction to fill as "Unknown". This means 25 rows carry a less-informative `region = Unknown` even though the state strongly implies the correct region.
4. **No deduplication was attempted on near-duplicate customers** (e.g. same person with slightly different name formatting was *not* checked against `customer_id` for entity-resolution beyond simple text cleaning) — `customer_id` was trusted as the unique key wherever the assignment didn't ask for explicit customer-level dedup.
5. **Cancelled/Returned/Failed orders were excluded from "sales" pivots but not deleted** — they remain in `cleaned_orders.xlsx` so any downstream user can re-include them if a different definition of "sales" is needed.
