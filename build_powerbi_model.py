import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

NAVY = "1F3864"
header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor=NAVY)
title_font = Font(name="Arial", size=14, bold=True, color=NAVY)
sub_font = Font(name="Arial", size=10, italic=True, color="595959")
normal_font = Font(name="Arial", size=10)
thin = Side(style="thin", color="B7B7B7")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

def write_df(ws, df, start_row=1):
    for j, col in enumerate(df.columns, start=1):
        ws.cell(row=start_row, column=j, value=col)
    style_header(ws, start_row, len(df.columns))
    for i, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for j, val in enumerate(row, start=1):
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = normal_font
    for j, col in enumerate(df.columns, start=1):
        maxlen = max(len(str(col)), df[col].map(lambda x: len(str(x))).max() if len(df) else 0)
        ws.column_dimensions[get_column_letter(j)].width = min(max(maxlen + 2, 10), 42)

wb = Workbook()

# ---------- ReadMe ----------
ws = wb.active
ws.title = "ReadMe"
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 110
ws["B2"] = "World Bank IDA — Credits, Grants & Guarantees Data Model"
ws["B2"].font = title_font
ws["B3"] = "Star-schema data model for Power BI, cleaned from the World Bank's public IDA Statement of Credits, Grants and Guarantees (Finances One Open Data)."
ws["B3"].font = sub_font
ws["B3"].alignment = Alignment(wrap_text=True)
ws["B5"] = "Source"
ws["B5"].font = Font(bold=True)
ws["B6"] = "financesone.worldbank.org/ida-statement-of-credits-grants-and-guarantees-latest-available-snapshot/DS00001 — real, public World Bank data (snapshot as of 6/30/2026)."
ws["B6"].alignment = Alignment(wrap_text=True)
ws["B8"] = "Tabs"
ws["B8"].font = Font(bold=True)
notes = [
    "fact_credits — one row per IDA credit/grant (11,404 rows): amounts, dates, ratios, status/region/country keys.",
    "dim_region, dim_country, dim_status, dim_project — lookup dimensions for the star schema.",
    "Summary — headline portfolio KPIs.",
    "DAX_Measures — measure definitions to paste into Power BI after loading the tables.",
]
for i, n in enumerate(notes, start=9):
    ws.cell(row=i, column=2, value="• " + n).font = normal_font
    ws.cell(row=i, column=2).alignment = Alignment(wrap_text=True)
ws["B14"] = "Cleaning notes: column names standardized to snake_case; dates parsed; text fields trimmed/title-cased; missing project names (4 rows) backfilled with '(Unnamed Project)'; missing service charge rates (grants, which carry none) set to 0; derived disbursement_ratio, repayment_ratio, is_active, board_approval_fy."
ws["B14"].font = sub_font
ws["B14"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[14].height = 45

# ---------- Load data ----------
fact = pd.read_csv("data/fact_credits.csv")
dim_region = pd.read_csv("data/dim_region.csv")
dim_country = pd.read_csv("data/dim_country.csv")
dim_status = pd.read_csv("data/dim_status.csv")
dim_project = pd.read_csv("data/dim_project.csv")

ws = wb.create_sheet("fact_credits")
write_df(ws, fact)

ws = wb.create_sheet("dim_region")
write_df(ws, dim_region)

ws = wb.create_sheet("dim_country")
write_df(ws, dim_country)

ws = wb.create_sheet("dim_status")
write_df(ws, dim_status)

ws = wb.create_sheet("dim_project")
write_df(ws, dim_project)

# ---------- Summary ----------
ws = wb.create_sheet("Summary")
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 20
ws["B2"] = "Portfolio Summary"
ws["B2"].font = title_font

total_principal = fact["original_principal_usd"].sum()
total_disbursed = fact["disbursed_usd"].sum()
total_repaid = fact["repaid_to_ida_usd"].sum()
total_undisbursed = fact["undisbursed_usd"].sum()
active_count = int(fact["is_active"].sum())

stats = [
    ("Total Credits/Grants", f"{len(fact):,}"),
    ("Countries Covered", f"{dim_country.shape[0]}"),
    ("Total Original Principal (US$)", f"${total_principal:,.0f}"),
    ("Total Disbursed (US$)", f"${total_disbursed:,.0f}"),
    ("Total Repaid to IDA (US$)", f"${total_repaid:,.0f}"),
    ("Total Undisbursed (US$)", f"${total_undisbursed:,.0f}"),
    ("Active Credits (Repaying/Effective/Disbursing)", f"{active_count:,}"),
    ("Overall Disbursement Ratio", f"{total_disbursed/total_principal:.1%}"),
]
r = 4
for label, val in stats:
    ws.cell(row=r, column=2, value=label).font = normal_font
    ws.cell(row=r, column=3, value=val).font = Font(bold=True, size=10)
    r += 1

r += 2
ws.cell(row=r, column=2, value="By Region").font = Font(bold=True, color=NAVY)
r += 1
by_region = (
    fact.merge(dim_region, on="region_id")
    .groupby("region")
    .agg(total_principal=("original_principal_usd", "sum"),
         total_disbursed=("disbursed_usd", "sum"),
         count=("credit_number", "count"))
    .reset_index()
    .sort_values("total_principal", ascending=False)
)
write_df(ws, by_region, start_row=r)

# ---------- DAX Measures ----------
ws = wb.create_sheet("DAX_Measures")
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 95
ws["B2"] = "DAX Measures — paste into Power BI after loading the star schema"
ws["B2"].font = title_font
ws.cell(row=4, column=2, value="Measure Name")
ws.cell(row=4, column=3, value="DAX Formula")
style_header(ws, 4, 2)

measures = [
    ("Total Original Principal", "Total Original Principal = SUM(fact_credits[original_principal_usd])"),
    ("Total Disbursed", "Total Disbursed = SUM(fact_credits[disbursed_usd])"),
    ("Total Repaid", "Total Repaid = SUM(fact_credits[repaid_to_ida_usd])"),
    ("Total Undisbursed", "Total Undisbursed = SUM(fact_credits[undisbursed_usd])"),
    ("Disbursement Ratio", "Disbursement Ratio = DIVIDE([Total Disbursed], [Total Original Principal], 0)"),
    ("Repayment Ratio", "Repayment Ratio = DIVIDE([Total Repaid], [Total Disbursed], 0)"),
    ("Active Credit Count", "Active Credit Count = CALCULATE(COUNTROWS(fact_credits), fact_credits[is_active] = TRUE)"),
    ("Avg Credit Size", "Avg Credit Size = DIVIDE([Total Original Principal], COUNTROWS(fact_credits), 0)"),
    ("YoY Principal Growth %",
     "YoY Principal Growth % = \nVAR CurrentFY = [Total Original Principal]\nVAR PriorFY = CALCULATE([Total Original Principal], fact_credits[board_approval_fy] = MAX(fact_credits[board_approval_fy]) - 1)\nRETURN DIVIDE(CurrentFY - PriorFY, PriorFY, 0)"),
    ("Rank Region by Principal", "Rank Region by Principal = RANKX(ALL(dim_region[region]), [Total Original Principal])"),
]
for name, formula in measures:
    r = ws.max_row + 1
    ws.cell(row=r, column=2, value=name).font = Font(bold=True, size=10)
    ws.cell(row=r, column=2).alignment = Alignment(vertical="top", wrap_text=True)
    c = ws.cell(row=r, column=3, value=formula)
    c.font = Font(name="Consolas", size=9)
    c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.row_dimensions[r].height = 32

order = ["ReadMe", "fact_credits", "dim_region", "dim_country", "dim_status", "dim_project", "Summary", "DAX_Measures"]
wb._sheets = [wb[n] for n in order]

wb.save("data/IDA_PowerBI_Data_Model.xlsx")
print("Saved.")
