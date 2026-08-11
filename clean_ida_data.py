"""
Clean the World Bank IDA Statement of Credits, Grants and Guarantees dataset
and export a Power BI-ready star schema (fact + dimension tables).

Source: World Bank Finances One Open Data
https://financesone.worldbank.org/ida-statement-of-credits-grants-and-guarantees-latest-available-snapshot/DS00001
"""
import pandas as pd
import numpy as np

RAW_PATH = "data/ida_raw.csv"

df = pd.read_csv(RAW_PATH)
print(f"Loaded {len(df):,} rows, {df.shape[1]} columns.")

# ---------------------------------------------------------------
# 1. Standardize column names (snake_case, no special chars)
# ---------------------------------------------------------------
rename_map = {
    "End of Period": "end_of_period",
    "Credit Number": "credit_number",
    "Region": "region",
    "Country / Economy Code": "country_code",
    "Country / Economy": "country",
    "Borrower": "borrower",
    "Credit Status": "credit_status",
    "Service Charge Rate": "service_charge_rate",
    "Currency of Commitment": "currency",
    "Project ID": "project_id",
    "Project Name": "project_name",
    "Original Principal Amount (US$)": "original_principal_usd",
    "Cancelled Amount (US$)": "cancelled_usd",
    "Undisbursed Amount (US$)": "undisbursed_usd",
    "Disbursed Amount (US$)": "disbursed_usd",
    "Repaid to IDA (US$)": "repaid_to_ida_usd",
    "Due to IDA (US$)": "due_to_ida_usd",
    "Exchange Adjustment (US$)": "exchange_adjustment_usd",
    "Borrowers Obligation (US$)": "borrowers_obligation_usd",
    "Sold 3rd Party (US$)": "sold_3rd_party_usd",
    "Repaid 3rd Party (US$)": "repaid_3rd_party_usd",
    "Due 3rd Party (US$)": "due_3rd_party_usd",
    "Credits Held (US$)": "credits_held_usd",
    "First Repayment Date": "first_repayment_date",
    "Last Repayment Date": "last_repayment_date",
    "Agreement Signing Date": "agreement_signing_date",
    "Board Approval Date": "board_approval_date",
    "Effective Date (Most Recent)": "effective_date",
    "Closed Date (Most Recent)": "closed_date",
    "Last Disbursement Date": "last_disbursement_date",
}
df = df.rename(columns=rename_map)

# ---------------------------------------------------------------
# 2. Fix data types
# ---------------------------------------------------------------
date_cols = [
    "end_of_period", "first_repayment_date", "last_repayment_date",
    "agreement_signing_date", "board_approval_date", "effective_date",
    "closed_date", "last_disbursement_date",
]
for c in date_cols:
    df[c] = pd.to_datetime(df[c], errors="coerce")

# Standardize text fields: trim + consistent casing for categoricals
for c in ["region", "country", "borrower", "credit_status", "currency", "project_name"]:
    df[c] = df[c].astype(str).str.strip()

df["region"] = df["region"].str.title()
df["credit_status"] = df["credit_status"].str.title()
df["project_name"] = df["project_name"].replace("nan", np.nan)

# ---------------------------------------------------------------
# 3. Handle missing values (documented, not silently dropped)
# ---------------------------------------------------------------
# service_charge_rate missing -> genuinely not applicable for some instrument
# types (e.g. grants carry no service charge); flag rather than impute.
df["service_charge_rate"] = df["service_charge_rate"].fillna(0.0)
missing_project_name = df["project_name"].isna().sum()
df["project_name"] = df["project_name"].fillna("(Unnamed Project)")

# ---------------------------------------------------------------
# 4. Derived / enrichment columns
# ---------------------------------------------------------------
df["board_approval_fy"] = df["board_approval_date"].dt.year
df["disbursement_ratio"] = np.where(
    df["original_principal_usd"] > 0,
    df["disbursed_usd"] / df["original_principal_usd"],
    np.nan,
)
df["repayment_ratio"] = np.where(
    df["disbursed_usd"] > 0,
    df["repaid_to_ida_usd"] / df["disbursed_usd"],
    np.nan,
)
df["is_active"] = df["credit_status"].isin(["Repaying", "Effective", "Disbursing"])

# ---------------------------------------------------------------
# 5. Data quality checks
# ---------------------------------------------------------------
dupes = df.duplicated(subset=["credit_number"]).sum()
negative_amounts = (df["original_principal_usd"] < 0).sum()
print(f"Duplicate credit numbers: {dupes}")
print(f"Negative principal amounts: {negative_amounts}")
print(f"Project names backfilled: {missing_project_name}")

# ---------------------------------------------------------------
# 6. Save cleaned flat file
# ---------------------------------------------------------------
df.to_csv("data/ida_clean.csv", index=False)
print(f"\nSaved cleaned dataset: {df.shape[0]:,} rows, {df.shape[1]} columns.")

# ---------------------------------------------------------------
# 7. Build star schema for Power BI
# ---------------------------------------------------------------
# DimRegion
dim_region = (
    df[["region"]].drop_duplicates().reset_index(drop=True)
)
dim_region.insert(0, "region_id", dim_region.index + 1)

# DimCountry
dim_country = (
    df[["country_code", "country", "region"]].drop_duplicates().reset_index(drop=True)
)
dim_country.insert(0, "country_id", dim_country.index + 1)

# DimStatus
dim_status = df[["credit_status"]].drop_duplicates().reset_index(drop=True)
dim_status.insert(0, "status_id", dim_status.index + 1)

# DimProject
dim_project = (
    df[["project_id", "project_name"]].drop_duplicates(subset=["project_id"]).reset_index(drop=True)
)

# FactCredits — join back to dimension keys
fact = df.merge(dim_country[["country_code", "country_id"]], on="country_code", how="left")
fact = fact.merge(dim_status, on="credit_status", how="left")
fact = fact.merge(dim_region, on="region", how="left")

fact_cols = [
    "credit_number", "project_id", "country_id", "region_id", "status_id",
    "currency", "service_charge_rate",
    "original_principal_usd", "cancelled_usd", "undisbursed_usd", "disbursed_usd",
    "repaid_to_ida_usd", "due_to_ida_usd", "borrowers_obligation_usd", "credits_held_usd",
    "disbursement_ratio", "repayment_ratio", "is_active", "board_approval_fy",
    "agreement_signing_date", "board_approval_date", "effective_date", "closed_date",
]
fact_credits = fact[fact_cols]

dim_region.to_csv("data/dim_region.csv", index=False)
dim_country.drop(columns=["region"]).to_csv("data/dim_country.csv", index=False)
dim_status.to_csv("data/dim_status.csv", index=False)
dim_project.to_csv("data/dim_project.csv", index=False)
fact_credits.to_csv("data/fact_credits.csv", index=False)

print("\nStar schema exported:")
print(f"  fact_credits.csv   {fact_credits.shape}")
print(f"  dim_region.csv     {dim_region.shape}")
print(f"  dim_country.csv    {dim_country.shape}")
print(f"  dim_status.csv     {dim_status.shape}")
print(f"  dim_project.csv    {dim_project.shape}")
