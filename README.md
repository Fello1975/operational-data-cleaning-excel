# operational-data-cleaning-excel

Excel VBA project to normalize hierarchical operational Excel reports into analysis-ready tables.

## Problem
Operational inbound reports are often delivered in a hierarchical or merged layout:
- Shipment-level fields appear only once
- SKU-level rows underneath contain blank cells
- The format is readable for humans but unusable for analytics

This structure makes it impossible to:
- Build PivotTables
- Filter or sort reliably
- Load the data into Power Query or SQL

## Solution
A VBA macro was developed to normalize the report by:
- Iterating through all columns
- Detecting blank cells caused by hierarchical layouts
- Filling values downward from the last non-empty parent row
- Producing a flat, row-level dataset ready for analysis

## Before / After

### Before (Unmodified Operational Report)
![Unmodified Report](screenshots/Unmodified_Repository.png)

### After (Normalized, Analysis-Ready Table)
![Modified Report](screenshots/Modified_Repository.png)

## Tools Used
- Microsoft Excel
- VBA (Visual Basic for Applications)

## How to Use
1. Open the unmodified report (`01_raw_report/Unfixed_Inbound_mock_data.xlsx`)
2. Press `ALT + F11` to open the VBA Editor
3. Import the macro from `02_cleaning_vba/FillBlanksAcrossColumns.bas`
4. Run the macro
5. A clean, analysis-ready table is generated

## Outcome
- The normalized dataset can be used for PivotTables, dashboards, and SQL exports
- Manual data preparation time is eliminated
- The process can be reused for similar operational reports

## Notes
All data included in this repository is simulated and anonymized.
No proprietary or confidential information is used.
