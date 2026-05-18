import io
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

# Make `src/account_comparison.py` importable when this lives at project root
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent / "SCRIPT (DO NOT TOUCH)"))

from account_comparison import run_comparison, write_excel  # noqa: E402


st.set_page_config(page_title="Guidepoint Invoice Reconciliation", layout="centered")

st.title("Guidepoint Invoice Reconciliation")
st.write(
    "Upload the SFDC Customer Orders report (with Shipping Info) and the Guidepoint "
    "hardware invoice. The script verifies every line on the invoice against SFDC orders."
)

with st.expander("Expected inputs", expanded=False):
    st.markdown(
        """
        **SFDC report** (`.xlsx` or `.csv`) — required columns:
        Order Name, DW Location Name, Created Date, Expected Date of Arrival,
        Quantity Shipped, Number of Connected Car Devices, Number of CCD Replacement,
        Shipping Record Number.

        Report type: *Customer Orders with Shipping Details*.
        Filter: `Expected Date of Arrival` in invoice month + 1 month prior.
        Rows with blank shipping info are dropped automatically.

        **Guidepoint invoice** (`.xlsx`) — must contain a `Data` sheet with columns:
        Type, Reference Nbr., Date, Dealer Ship To, Amount, New Units, Refurb Units,
        Shipping Costs, Customs Fees.

        Harnesses excluded. Credit memos handled via signed quantities.
        """
    )

sfdc_file = st.file_uploader(
    "SFDC report (Customer Orders w/ Shipping Info)",
    type=["xlsx", "csv"],
)
gp_file = st.file_uploader(
    "Guidepoint hardware invoice",
    type=["xlsx"],
)

run = st.button("Run reconciliation", type="primary", disabled=not (sfdc_file and gp_file))


def _stash_upload(upload, suffix: str) -> Path:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(upload.getbuffer())
    tmp.close()
    return Path(tmp.name)


if run:
    try:
        sfdc_path = _stash_upload(sfdc_file, suffix=Path(sfdc_file.name).suffix)
        gp_path = _stash_upload(gp_file, suffix=Path(gp_file.name).suffix)

        with st.spinner("Reconciling…"):
            sheets, window_start = run_comparison(sfdc_path, gp_path)

        invoice_month_ts = window_start + pd.offsets.MonthBegin(1)
        invoice_month = invoice_month_ts.strftime("%b").upper()
        invoice_year = invoice_month_ts.strftime("%Y")

        st.success(
            f"Reconciled invoice month: {invoice_month} {invoice_year} "
            f"(SFDC window: {window_start.date()} → {(invoice_month_ts + pd.offsets.MonthEnd(0)).date()})"
        )

        # Build the Excel output
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp_path = Path(tmp.name)
        write_excel(sheets, tmp_path)
        with open(tmp_path, "rb") as f:
            buf = io.BytesIO(f.read())
        buf.seek(0)

        st.download_button(
            label=f"📥 Download ACCOUNT_COMPARISON_{invoice_month}_{invoice_year}.xlsx",
            data=buf,
            file_name=f"ACCOUNT_COMPARISON_{invoice_month}_{invoice_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Summary metrics
        invoice_detail = sheets["Invoice Line Detail"]
        rollup = sheets["Ship-To Rollup"]
        unmatched = sheets["Unmatched Invoice Lines"]

        matched_lines = len(invoice_detail) - len(unmatched)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Invoice lines", len(invoice_detail))
        col2.metric("Matched", matched_lines)
        col3.metric("Unmatched", len(unmatched))
        col4.metric("Ship-tos", len(rollup))

        # Previews
        with st.expander("Invoice Line Detail", expanded=True):
            st.dataframe(invoice_detail, use_container_width=True)
        with st.expander("Ship-To Rollup"):
            st.dataframe(rollup, use_container_width=True)
        with st.expander("Unmatched Invoice Lines (action items)"):
            st.dataframe(unmatched, use_container_width=True)

    except KeyError as e:
        st.error(f"Column mismatch:\n{e}")
    except Exception as e:
        st.error(f"{type(e).__name__}: {e}")
        st.exception(e)

