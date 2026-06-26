# Part 1 — Business Data Cleaning, Validation & Excel Reporting

## Problem Summary

As a business analyst for a retail company, I was given order-level sales data exported from
multiple internal systems (`data/raw_orders.xlsx`). The raw export contained inconsistent text
formatting, mixed date formats, duplicate records, missing values, invalid discounts,
sales/profit calculation mismatches, and inconsistent order/payment status combinations. The
goal was to produce a clean, validated, analysis-ready dataset and a set of summary reports
suitable for business review.

## Dataset Description

- **File:** `data/raw_orders.xlsx`
- **Size:** 932 order-level records, 21 columns
- **Columns:** `order_id`, `order_date`, `ship_date`, `customer_id`, `customer_name`, `segment`,
  `region`, `state`, `city`, `category`, `sub_category`, `product_name`, `ship_mode`, `quantity`,
  `unit_price`, `discount`, `sales`, `cost`, `profit`, `payment_status`, `order_status`
- **Time range:** January 2024 – October 2025
- **Geography:** 5 regions (North, South, East, West, plus rows with missing region) across 19
  Indian states and 20 cities

## Tools Used

- **Python** (pandas, openpyxl) — used to apply the cleaning logic at scale and reliably (the
  same logic that Excel's TRIM, SUBSTITUTE, Find & Replace, Flash Fill, and Text to Columns
  achieve manually), and to build formatted `.xlsx` output files.
- **Microsoft Excel–compatible output** — all deliverables (`cleaned_orders.xlsx`,
  `data_quality_report.xlsx`, `pivot_summary.xlsx`) are standard `.xlsx` workbooks with native
  Excel Tables, AutoFilter, number formatting, and conditional color-coding, fully editable in
  Excel.
- **LibreOffice** — used headlessly to verify the workbooks open with zero formula errors.

## Cleaning Steps Performed

1. **Text fields** — trimmed leading/trailing/extra internal spaces and standardized casing to
   Title Case across `customer_name`, `segment`, `region`, `state`, `city`, `category`,
   `sub_category`, `ship_mode`, `payment_status`, and `order_status`.
2. **Dates** — detected 4 mixed formats (`DD Mon YYYY`, `YYYY-MM-DD`, `MM/DD/YYYY`,
   `DD-MM-YYYY`) present in `order_date`/`ship_date`, parsed each correctly based on the
   numeric range of each token, and standardized both fields to `YYYY-MM-DD`. Added
   `shipping_delay_days` and flagged the 21 records where `ship_date` fell before `order_date`.
3. **Duplicates** — removed 20 fully-identical duplicate rows; flagged (but kept) 24 rows
   belonging to 12 `order_id` values that had conflicting data under the same ID.
4. **Missing values** — filled missing `region`/`ship_mode` as `"Unknown"` and missing
   `discount` as `0`, all flagged for traceability.
5. **Discounts** — converted percent-string discounts (e.g. `"70%"`) to numeric, flagged
   negative discounts and discounts above the observed valid range (25%), and produced a
   `cleaned_discount` column safe to use in calculations.
6. **Calculated columns** — added `calculated_sales`, `calculated_profit`, `profit_margin`,
   `order_month`, `order_year`, and an overall `data_quality_flag` (Clean / Warning / Invalid).

Full detail — including every issue count and the exact logic used — is in
[`outputs/cleaning_log.md`](outputs/cleaning_log.md).

## Business Rules Applied

| Rule | Action |
|---|---|
| Missing region / ship_mode | Filled `"Unknown"`, flagged |
| Missing discount | Treated as 0 (only when other sales fields valid), flagged |
| Negative discount | Flagged invalid, excluded from calculation (set to 0) |
| Discount above allowed range (>25%) | Flagged invalid, capped at 25% |
| Cancelled orders / Failed payments | Excluded from completed-sales summaries |
| Refunded orders | Summarized separately |
| Ship date before order date | Flagged as invalid shipping record |

## Summary of Data Quality Issues Found

| Issue | Count (of 932 raw rows) |
|---|---|
| Exact duplicate rows | 40 rows (20 groups) |
| Duplicate order_id with conflicting data | 24 rows (12 order_ids) |
| Missing region | 26 |
| Missing ship_mode | 22 |
| Missing discount | 18 |
| Negative discount | 16 |
| Discount above allowed range | 7 |
| Ship date before order date | 21 |
| Sales/profit calculation mismatches | 58 |
| Order status / payment status inconsistency | 2 |

Final result in `cleaned_orders.xlsx` (912 rows after removing exact duplicates):

| data_quality_flag | Count | % |
|---|---|---|
| Clean | 756 | 82.9% |
| Warning | 98 | 10.7% |
| Invalid | 58 | 6.4% |

Full breakdown by category is in `outputs/data_quality_report.xlsx`.

## Summary of Final Pivot Reports

`outputs/pivot_summary.xlsx` contains 6 summaries (all completed-sales pivots are scoped to
`order_status = Completed` and `payment_status != Failed`):

1. **Sales and profit by region** — sorted by total sales, descending
2. **Sales and profit by category and sub-category** — sorted by category, then sales
3. **Order count by ship mode** — sorted by order count, descending, with avg. shipping delay
4. **Profit margin by customer segment** — sorted by profit margin, descending
5. **Refunded / cancelled / failed orders by region** — broken down by issue type
6. **Monthly sales trend** — Jan 2024 – Oct 2025, sorted chronologically

## Key Business Insights

- **South** is the top-performing region by completed sales (₹16.1L), closely followed by
  **West** (₹15.0L) and **East** (₹13.8L).
- **Technology** is the highest-revenue category (₹21.9L in completed sales), ahead of
  **Furniture** (₹19.3L) and **Office Supplies** (₹18.5L).
- **Home Office** customers generate the best profit margin (29.9%), narrowly ahead of
  **Corporate** (29.4%); **Small Business** has the thinnest margin (28.3%).
- Total completed sales across the cleaned dataset are **≈₹59.7L**, with **≈₹17.3L** in
  completed profit.
- **310 of 912 orders (34%)** are Cancelled, Returned, or had a Failed payment — a meaningful
  share of order volume that does not convert to realized revenue and may warrant a closer
  operational review (highest cancellation volume is concentrated in the **North** region).

## Assumptions and Limitations

- Slash-format dates were determined to be `MM/DD/YYYY` and dash-format dates `DD-MM-YYYY`,
  based on the numeric ranges actually present in the data (not guessed).
- The valid discount range (0–25%) was inferred from the clustering of legitimate values in the
  data, since no explicit range was provided.
- `region` was filled as `"Unknown"` for missing values per the assignment's explicit rule, even
  though `state` could reliably predict the correct region in nearly all cases.
- Conflicting duplicate `order_id` records (12 IDs / 24 rows) were **flagged, not auto-resolved**
  — there was no reliable way to determine which version of a conflicting record was correct.
- Full limitations list: see the final section of [`outputs/cleaning_log.md`](outputs/cleaning_log.md).

## Screenshots

| File | Shows |
|---|---|
| `screenshots/raw_data_preview.png` | Raw dataset before cleaning — mixed casing, blank cells, mixed date formats |
| `screenshots/cleaned_data_preview.png` | Cleaned dataset with calculated columns and color-coded data quality flags |
| `screenshots/pivot_summary_1.png` | Sales and profit by region |
| `screenshots/pivot_summary_2.png` | Monthly sales trend |

## Repository Structure

```
part1_data_cleaning/
├── data/
│   ├── raw_orders.xlsx          (unchanged original)
│   └── cleaned_orders.xlsx      (cleaned + calculated columns)
├── outputs/
│   ├── data_quality_report.xlsx
│   ├── pivot_summary.xlsx
│   └── cleaning_log.md
├── screenshots/
│   ├── raw_data_preview.png
│   ├── cleaned_data_preview.png
│   ├── pivot_summary_1.png
│   └── pivot_summary_2.png
└── README.md
```
