import io
from datetime import datetime

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Account Comparison", layout="centered")

st.title("Account Comparison")
st.write("Upload the CCD CSV and the Guidepoint invoice XLSX, then run the comparison.")

ccd_file = st.file_uploader("CCD Add/Replace (CSV)", type=["csv"], accept_multiple_files=False)
invoice_file = st.file_uploader("Guidepoint Invoice (XLSX)", type=["xlsx"], accept_multiple_files=False)

run = st.button("Run comparison", type="primary", disabled=not (ccd_file and invoice_file))


def _run_comparison(df_csv: pd.DataFrame, df_xlsx: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df_csv = df_csv.copy()
    df_xlsx = df_xlsx.copy()

    df_csv.rename(columns={"Account Name": "Account Name", "CCD Quantity": "CCD Quantity"}, inplace=True)
    df_xlsx.columns = df_xlsx.columns.astype(str).str.replace("\n", " ", regex=False).str.strip()

    if {"Ship To Dealer", "Shipped"}.issubset(set(df_xlsx.columns)):
        df_xlsx.rename(columns={"Ship To Dealer": "Account Name", "Shipped": "Total Quantity"}, inplace=True)
    elif {"Ship To", "New Unit"}.issubset(set(df_xlsx.columns)):
        df_xlsx.rename(columns={"Ship To": "Account Name"}, inplace=True)
        df_xlsx["Total Quantity"] = pd.to_numeric(df_xlsx["New Unit"], errors="coerce").fillna(0)
        df_xlsx = (
            df_xlsx.groupby("Account Name", as_index=False)["Total Quantity"]
            .sum()
            .sort_values("Account Name")
        )
    elif {"Ship to Customer Name", "New Unit"}.issubset(set(df_xlsx.columns)):
        df_xlsx.rename(columns={"Ship to Customer Name": "Account Name"}, inplace=True)
        df_xlsx["Total Quantity"] = pd.to_numeric(df_xlsx["New Unit"], errors="coerce").fillna(0)
        df_xlsx = (
            df_xlsx.groupby("Account Name", as_index=False)["Total Quantity"]
            .sum()
            .sort_values("Account Name")
        )
    else:
        raise KeyError(
            "Unrecognized Guidepoint invoice format. "
            "Expected either columns: Ship To Dealer + Shipped (legacy) "
            "or Ship To + New Unit (new Summary format). "
            "Found columns: "
            + ", ".join(map(str, df_xlsx.columns))
        )

    required_csv_cols = {"Account Name", "CCD Quantity"}
    required_xlsx_cols = {"Account Name", "Total Quantity"}

    missing_csv = sorted(required_csv_cols - set(df_csv.columns))
    if missing_csv:
        raise KeyError(
            "CSV is missing required column(s): "
            + ", ".join(missing_csv)
            + "\nFound columns: "
            + ", ".join(map(str, df_csv.columns))
        )

    missing_xlsx = sorted(required_xlsx_cols - set(df_xlsx.columns))
    if missing_xlsx:
        raise KeyError(
            "XLSX is missing required column(s): "
            + ", ".join(missing_xlsx)
            + "\nFound columns: "
            + ", ".join(map(str, df_xlsx.columns))
        )

    merged = pd.merge(df_csv, df_xlsx, on="Account Name", how="inner")
    merged["Difference"] = merged["CCD Quantity"] - merged["Total Quantity"]
    sheet1 = merged[["Account Name", "CCD Quantity", "Total Quantity", "Difference"]]

    not_in_csv = df_xlsx[~df_xlsx["Account Name"].isin(df_csv["Account Name"])]
    sheet2 = not_in_csv[["Account Name", "Total Quantity"]].copy()
    sheet2.rename(columns={"Total Quantity": "TOTAL CCD Value"}, inplace=True)

    return sheet1, sheet2


if run:
    try:
        df_csv = pd.read_csv(ccd_file)
        df_xlsx = pd.read_excel(invoice_file)

        sheet1, sheet2 = _run_comparison(df_csv, df_xlsx)

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            sheet1.to_excel(writer, sheet_name="Matched Accounts", index=False)
            sheet2.to_excel(writer, sheet_name="XLSX Only Accounts", index=False)

        out.seek(0)

        output_filename = f"ACCOUNT_COMPARISON_{datetime.now().strftime('%b_%d_%Y')}.xlsx"

        st.success("Comparison complete.")
        st.download_button(
            label="Download output Excel",
            data=out,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Preview: Matched Accounts", expanded=False):
            st.dataframe(sheet1, use_container_width=True)

        with st.expander("Preview: XLSX Only Accounts", expanded=False):
            st.dataframe(sheet2, use_container_width=True)

    except KeyError as e:
        st.error(str(e))
    except Exception as e:
        st.error(str(e))
