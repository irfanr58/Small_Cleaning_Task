import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
import os
from typing import cast
import xlsxwriter as pw
from xlsxwriter.chart import Chart      # ← add this
import polars as pl
import re


path = r"C:\Users\irfan\OneDrive\Desktop\Python Projects\Small_Cleaning Task\Dirty_csv_Dataset.csv"

# Load the CSV directly into a DataFrame
df = pd.read_csv(path)

# convert pandas df to polars, clean, then convert back
pl_df = pl.from_pandas(df)

cleaned = (
    pl_df
    .rename(
        {
            col: re.sub(r'\s+', '_', col.strip().replace('-', '').lower())
            for col in pl_df.columns
        }
    )  # rename columns: strip whitespace, remove hyphens, lowercase and convert spaces to underscores
    .with_columns([
        pl.col("first_name")
          .str.to_lowercase()
          .str.to_titlecase()
          .alias("first_name"),
        pl.col("last_name")
          .str.to_lowercase()
          .str.to_titlecase()
          .alias("last_name")
    ])  # normalize first and last names to Title Case
    .filter(
        (pl.col("first_name") != "") &
        (pl.col("last_name")  != "")
    )  # drop rows with empty first or last name
    .with_columns(
        pl.col("email")
          .str.split(r";")
          .alias("email")
    )  # split multiple emails into a list
    .explode("email")  # expand each email entry into its own row
    .filter(
        pl.col("email")
          .str.contains(r"\.com")
    )  # keep only email addresses ending with .com

    .with_columns(
        pl.col("signup_date")
          .str.strptime(pl.Date, fmt="%m/%d/%Y")  # parse your m/d/Y strings into Dates
          .dt.strftime("%F")                      # format back to "YYYY-MM-DD"
          .alias("signup_date")
    )

)

df = cleaned.to_pandas()


# define output folder and file
final_dir = r"C:\Users\irfan\OneDrive\Desktop\Python Projects\Small_Cleaning Task"
os.makedirs(final_dir, exist_ok=True)
final_path = os.path.join(final_dir, "Cleaned_3.xlsx")


with pd.ExcelWriter(final_path, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Sheet1", index=False) 
    worksheet = writer.sheets["Sheet1"]


    for idx, col in enumerate(df.columns):
        # compute the max cell length (as string)
        cell_max = df[col].astype(str).map(len).max()
        # header length
        hdr_len = len(col)
        # add +2 padding, but ensure header alone gets at least hdr_len+2
        width = max(cell_max, hdr_len) + 3
        worksheet.set_column(idx, idx, width)


                             
os.startfile(final_path)