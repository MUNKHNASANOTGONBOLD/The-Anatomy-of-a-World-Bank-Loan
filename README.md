# World Bank IDA Portfolio: Credits, Grants & Guarantees

Python cleaning pipeline + Power BI-ready star schema + dashboard prototype, built on real, public World Bank data.

## Dataset

**Source:** [World Bank Finances One Open Data: IDA Statement of Credits, Grants and Guarantees](https://financesone.worldbank.org/ida-statement-of-credits-grants-and-guarantees-latest-available-snapshot/DS00001), snapshot as of 6/30/2026.

**Scale:** 11,404 IDA credits/grants across 131 countries, 30 original columns.

## Pipeline

1. **Python cleaning** (`clean_ida_data.py`): standardized column names to snake_case, parsed all date fields, trimmed/title-cased text fields, backfilled 4 missing project names, set missing service charge rates (grants carry none) to 0, derived `disbursement_ratio`, `repayment_ratio`, `is_active`, `board_approval_fy`. Zero duplicate credit numbers, zero negative principal amounts found.
2. **Star schema export**: `fact_credits` (11,404 rows) + `dim_region`, `dim_country`, `dim_status`, `dim_project` lookup tables, all as CSVs.
3. **Power BI data model** (`build_powerbi_model.py`): packages the star schema into a single Excel workbook with a Summary tab and a DAX_Measures reference tab (10 measures ready to paste into Power BI).
4. **Dashboard prototype**: two versions:
   - `IDA_Portfolio_Dashboard.html`: single-page Power BI-style preview (KPIs, trend, status, region/country, top projects)
   - `IDA_Portfolio_Dashboard_v2.html`: multi-page explorer with a sidebar, live region/status filters, and a capital-flow (Principal → Disbursed → Repaid) visualization

## Repo structure

```
data/
  ida_raw.csv                    Original uploaded snapshot
  ida_clean.csv                  Cleaned flat file (34 columns)
  fact_credits.csv               Fact table (star schema)
  dim_region.csv / dim_country.csv / dim_status.csv / dim_project.csv
IDA_PowerBI_Data_Model.xlsx      Excel workbook: star schema + Summary + DAX measures
clean_ida_data.py                Cleaning + star-schema export script
build_powerbi_model.py           Excel workbook builder
IDA_Portfolio_Dashboard.html     Single-page dashboard prototype
IDA_Portfolio_Dashboard_v2.html  Multi-page explorer with sidebar filters
```

## Key numbers

- Total original principal: ~$632B
- Total disbursed: ~$479B (76% disbursement ratio)
- Total repaid to IDA: ~$158B
- 7 regions, 131 countries, 10 credit statuses

## Tools

Python (pandas, NumPy) · Excel · Power BI (DAX) · Chart.js (prototype)
