"""
Guidepoint Invoice ↔ SFDC Customer Order Reconciliation
========================================================
Verifies every line on the Guidepoint hardware invoice against SFDC Customer
Orders (joined to Shipping Information). The invoice is the source of truth —
for each Guidepoint line, we attach the matching SFDC order(s) so you can
confirm what Guidepoint billed for actually corresponds to a real order.

Inputs (one file each in the matching INPUTS folder):
    INPUTS/SFDC_ORDERS/        SFDC report export (.xlsx or .csv)
    INPUTS/GUIDEPOINT_INVOICE/ Guidepoint hardware invoice (.xlsx)

Output (in OUTPUTS/):
    ACCOUNT_COMPARISON_<MONTH>.xlsx with 4 sheets:
        1. Invoice Line Detail    — GP lines with qty >= 0, + matched SFDC order(s)
        2. Returns                — GP lines with qty < 0 (credit memos / returns)
        3. Ship-To Rollup         — GP lines summed by Dealer Ship To (net of returns)
        4. Unmatched Invoice Lines — GP invoice lines with no SFDC match (action items)

Join key: normalized Location Name ↔ Dealer Ship To.
Fuzzy fallback: rapidfuzz at 90% confidence.
"""

import re
import sys
from pathlib import Path

import pandas as pd

try:
    from rapidfuzz import fuzz, process
    HAVE_RAPIDFUZZ = True
except ImportError:
    HAVE_RAPIDFUZZ = False


PROJECT_ROOT = Path(__file__).resolve().parent.parent
FUZZY_THRESHOLD = 90


# ============================================================
# File discovery
# ============================================================

def _single_file_in_dir(folder: Path) -> Path:
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"Missing folder: {folder}")
    files = [p for p in folder.iterdir() if p.is_file() and not p.name.startswith(".")]
    if len(files) != 1:
        names = "\n".join(f"- {p.name}" for p in sorted(files, key=lambda p: p.name)) or "(no files found)"
        raise ValueError(f"Expected exactly 1 file in {folder}\nFound {len(files)}:\n{names}")
    return files[0]


def _read_sfdc_export(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    df = pd.read_excel(path)
    # If columns are unnamed, SFDC added a title row — re-read with header on row 1
    if any("Unnamed" in str(c) for c in df.columns):
        df = pd.read_excel(path, header=1)
    return df


# ============================================================
# Normalization & fuzzy matching
# ============================================================

def normalize_name(s) -> str:
    if pd.isna(s):
        return ""
    s = str(s).lower()
    s = re.sub(r"\b(inc|llc|ltd|co|corp|company|corporation)\b\.?", "", s)
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def fuzzy_match_one(name: str, candidates: list, threshold: int = FUZZY_THRESHOLD):
    if not HAVE_RAPIDFUZZ or not name or not candidates:
        return None, 0
    result = process.extractOne(name, candidates, scorer=fuzz.token_sort_ratio)
    if result is None:
        return None, 0
    best, score, _ = result
    if score >= threshold:
        return best, score
    return None, score


# ============================================================
# SFDC side
# ============================================================

SFDC_REQUIRED_COLS = {
    "Order Name",
    "Dealerware Location: Dealerware Location Name",
    "Account Name",
    "Created Date",
    "Expected Date of Arrival",
    "Quantity Shipped",
    "Shipping Record Number",
}

# The location column name varies by report type
SFDC_LOCATION_COL = "Dealerware Location: Dealerware Location Name"


def load_sfdc(path: Path, window_start: pd.Timestamp, window_end: pd.Timestamp) -> pd.DataFrame:
    df = _read_sfdc_export(path)
    df.columns = df.columns.astype(str).str.replace("\n", " ", regex=False).str.strip()

    missing = sorted(SFDC_REQUIRED_COLS - set(df.columns))
    if missing:
        raise KeyError(
            "SFDC report is missing required column(s): "
            + ", ".join(missing)
            + "\nFound: "
            + ", ".join(map(str, df.columns))
        )

    # Normalize the location column to a standard name used throughout
    df.rename(columns={SFDC_LOCATION_COL: "DW Location Name"}, inplace=True)

    df["Expected Date of Arrival"] = pd.to_datetime(df["Expected Date of Arrival"], errors="coerce")
    df["Created Date"] = pd.to_datetime(df["Created Date"], errors="coerce")
    df["Quantity Shipped"] = pd.to_numeric(df["Quantity Shipped"], errors="coerce")

    before = len(df)
    df = df[df["Expected Date of Arrival"].notna()]
    df = df[df["DW Location Name"].notna()]
    df = df[df["Quantity Shipped"].notna()]
    dropped = before - len(df)
    if dropped:
        print(f"SFDC: dropped {dropped} rows with missing shipping info / location / quantity")

    mask = (df["Expected Date of Arrival"] >= window_start) & (df["Expected Date of Arrival"] <= window_end)
    df = df[mask].copy()
    print(f"SFDC: {len(df)} orders within window {window_start.date()} → {window_end.date()}")

    df["Location Key"] = df["DW Location Name"].apply(normalize_name)
    return df


# ============================================================
# Guidepoint side
# ============================================================

GP_REQUIRED_COLS = {
    "Type",
    "Reference Nbr.",
    "Date",
    "Dealer Ship To",
    "Amount",
    "New Units",
    "Refurb Units",
    "Shipping Costs",
    "Customs Fees",
}


def load_guidepoint_lines(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Data")
    df.columns = df.columns.astype(str).str.replace("\n", " ", regex=False).str.strip()

    missing = sorted(GP_REQUIRED_COLS - set(df.columns))
    if missing:
        raise KeyError(
            "Guidepoint Data sheet is missing required column(s): "
            + ", ".join(missing)
            + "\nFound: "
            + ", ".join(map(str, df.columns))
        )

    df = df[df["Type"].isin(["Invoice", "Credit Memo"])].copy()
    df = df[df["Dealer Ship To"].notna()].copy()

    for col in ["Amount", "New Units", "Refurb Units", "Shipping Costs", "Customs Fees"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Net Units"] = df["New Units"] + df["Refurb Units"]
    df["Net Device Amount"] = df["Amount"] - df["Shipping Costs"] - df["Customs Fees"]
    df["Location Key"] = df["Dealer Ship To"].apply(normalize_name)

    df = df.reset_index(drop=True)
    df["GP Line ID"] = df.index + 1
    return df


# ============================================================
# Match table: GP Dealer Ship To → SFDC DW Location Name
# ============================================================

def build_match_table(gp: pd.DataFrame, sfdc: pd.DataFrame) -> pd.DataFrame:
    gp_locs = gp[["Dealer Ship To", "Location Key"]].drop_duplicates().copy()
    sfdc_locs = sfdc[["DW Location Name", "Location Key"]].drop_duplicates().copy()

    # Pass 1: exact normalized match
    merged = gp_locs.merge(sfdc_locs, on="Location Key", how="left")
    merged["Match Type"] = merged["DW Location Name"].apply(
        lambda x: "exact" if pd.notna(x) else None
    )
    merged["Fuzzy Score"] = None

    # Pass 2: fuzzy on remaining
    if HAVE_RAPIDFUZZ:
        sfdc_name_list = sfdc_locs["DW Location Name"].dropna().tolist()
        for idx in merged[merged["DW Location Name"].isna()].index:
            gp_name = merged.loc[idx, "Dealer Ship To"]
            best, score = fuzzy_match_one(gp_name, sfdc_name_list, FUZZY_THRESHOLD)
            if best is not None:
                merged.loc[idx, "DW Location Name"] = best
                merged.loc[idx, "Match Type"] = "fuzzy"
                merged.loc[idx, "Fuzzy Score"] = score
            else:
                near, near_score = fuzzy_match_one(gp_name, sfdc_name_list, threshold=0)
                if near is not None:
                    merged.loc[idx, "Fuzzy Score"] = near_score

    return merged[["Dealer Ship To", "DW Location Name", "Match Type", "Fuzzy Score"]]


# ============================================================
# Output builders
# ============================================================

def build_invoice_line_detail(gp: pd.DataFrame, sfdc: pd.DataFrame, match_table: pd.DataFrame) -> pd.DataFrame:
    sfdc_per_loc = sfdc.groupby("DW Location Name", as_index=False).agg(
        **{
            "SFDC Qty Shipped (Total)": ("Quantity Shipped", "sum"),
            "SFDC Order Count": ("Order Name", "size"),
            "SFDC Order Names": ("Order Name", lambda s: ", ".join(sorted(set(s)))),
            "SFDC Earliest Created": ("Created Date", "min"),
            "SFDC Latest Expected Arrival": ("Expected Date of Arrival", "max"),
        }
    )

    gp_lines = gp.merge(match_table, on="Dealer Ship To", how="left")
    detail = gp_lines.merge(sfdc_per_loc, on="DW Location Name", how="left")

    out = detail[[
        "GP Line ID",
        "Type",
        "Reference Nbr.",
        "Date",
        "Dealer Ship To",
        "Net Units",
        "New Units",
        "Refurb Units",
        "Amount",
        "Net Device Amount",
        "Shipping Costs",
        "DW Location Name",
        "Match Type",
        "Fuzzy Score",
        "SFDC Qty Shipped (Total)",
        "SFDC Order Count",
        "SFDC Order Names",
        "SFDC Earliest Created",
        "SFDC Latest Expected Arrival",
    ]].copy()

    out.rename(columns={
        "Type": "GP Type",
        "Reference Nbr.": "GP Reference Nbr.",
        "Date": "GP Invoice Date",
        "Net Units": "GP Net Units",
        "New Units": "GP New Units",
        "Refurb Units": "GP Refurb Units",
        "Amount": "GP Total Amount",
        "Net Device Amount": "GP Net Device Amount",
        "Shipping Costs": "GP Shipping Costs",
    }, inplace=True)

    out.sort_values(["Dealer Ship To", "GP Invoice Date"], inplace=True)
    return out


def build_shipto_rollup(gp: pd.DataFrame, sfdc: pd.DataFrame, match_table: pd.DataFrame) -> pd.DataFrame:
    gp_agg = gp.groupby("Dealer Ship To", as_index=False).agg(
        **{
            "GP Net Units": ("Net Units", "sum"),
            "GP Total Amount": ("Amount", "sum"),
            "GP Net Device Amount": ("Net Device Amount", "sum"),
            "GP Shipping Costs": ("Shipping Costs", "sum"),
            "GP Invoice Line Count": ("Net Units", "size"),
        }
    )

    gp_agg = gp_agg.merge(match_table, on="Dealer Ship To", how="left")

    sfdc_per_loc = sfdc.groupby("DW Location Name", as_index=False).agg(
        **{
            "SFDC Qty Shipped": ("Quantity Shipped", "sum"),
            "SFDC Order Count": ("Order Name", "size"),
            "SFDC Order Names": ("Order Name", lambda s: ", ".join(sorted(set(s)))),
            "SFDC Earliest Created": ("Created Date", "min"),
            "SFDC Latest Expected Arrival": ("Expected Date of Arrival", "max"),
        }
    )

    out = gp_agg.merge(sfdc_per_loc, on="DW Location Name", how="left")
    out["Quantity Difference (SFDC - GP)"] = out["SFDC Qty Shipped"].fillna(0) - out["GP Net Units"]

    out = out[[
        "Dealer Ship To",
        "DW Location Name",
        "Match Type",
        "Fuzzy Score",
        "GP Net Units",
        "SFDC Qty Shipped",
        "Quantity Difference (SFDC - GP)",
        "GP Total Amount",
        "GP Net Device Amount",
        "GP Shipping Costs",
        "GP Invoice Line Count",
        "SFDC Order Count",
        "SFDC Order Names",
        "SFDC Earliest Created",
        "SFDC Latest Expected Arrival",
    ]].copy()

    # Sort: unmatched first, then by abs diff descending
    out["_p1"] = out["Match Type"].apply(lambda x: 0 if pd.isna(x) else 1)
    out["_p2"] = out["Quantity Difference (SFDC - GP)"].abs()
    out.sort_values(["_p1", "_p2"], ascending=[True, False], inplace=True)
    out.drop(columns=["_p1", "_p2"], inplace=True)
    return out


# ============================================================
# Window inference
# ============================================================

def infer_invoice_window(gp_path: Path) -> tuple:
    df = pd.read_excel(gp_path, sheet_name="Data")
    df = df[df["Type"].isin(["Invoice", "Credit Memo"])]
    dates = pd.to_datetime(df["Date"], errors="coerce").dropna()
    if dates.empty:
        today = pd.Timestamp.today()
        end = today.replace(day=1) - pd.Timedelta(days=1)
        invoice_month_start = end.replace(day=1)
    else:
        ym = dates.dt.to_period("M").mode()[0]
        invoice_month_start = ym.to_timestamp()
    invoice_month_end = invoice_month_start + pd.offsets.MonthEnd(0) + pd.offsets.MonthEnd(1)
    window_start = invoice_month_start - pd.offsets.MonthBegin(1)
    return window_start, invoice_month_end


# ============================================================
# Public API
# ============================================================

def run_comparison(sfdc_path: Path, gp_path: Path):
    window_start, window_end = infer_invoice_window(gp_path)
    print(f"Invoice window: {window_start.date()} → {window_end.date()}")

    sfdc = load_sfdc(sfdc_path, window_start, window_end)
    gp = load_guidepoint_lines(gp_path)
    print(f"SFDC: {len(sfdc)} orders across {sfdc['DW Location Name'].nunique()} locations")
    print(f"GP:   {len(gp)} invoice lines across {gp['Dealer Ship To'].nunique()} ship-tos")

    match_table = build_match_table(gp, sfdc)
    matched_ct = match_table["DW Location Name"].notna().sum()
    print(f"GP ship-tos matched to SFDC location: {matched_ct} / {len(match_table)}")

    all_lines = build_invoice_line_detail(gp, sfdc, match_table)
    shipto_rollup = build_shipto_rollup(gp, sfdc, match_table)

    # Split by sign: any negative net units = return, regardless of Type label.
    # (Type='Credit Memo' usually has negative units, but split on the actual quantity
    # to be safe — a misclassified row shouldn't slip through.)
    returns = all_lines[all_lines["GP Net Units"] < 0].copy()
    invoice_detail = all_lines[all_lines["GP Net Units"] >= 0].copy()

    # Unmatched = invoice lines (non-returns) with no SFDC match. Real action items.
    unmatched = invoice_detail[invoice_detail["DW Location Name"].isna()].copy()
    cols = [
        "GP Line ID", "GP Type", "GP Reference Nbr.", "GP Invoice Date",
        "Dealer Ship To", "GP Net Units", "GP New Units", "GP Refurb Units",
        "GP Total Amount", "GP Net Device Amount", "GP Shipping Costs", "Fuzzy Score",
    ]
    unmatched = unmatched[[c for c in cols if c in unmatched.columns]].copy()

    print(f"Invoice lines (qty >= 0): {len(invoice_detail)}")
    print(f"Returns (qty < 0):        {len(returns)}")
    print(f"Ship-to rollup rows:      {len(shipto_rollup)}")
    print(f"Unmatched invoice lines:  {len(unmatched)}")

    if not HAVE_RAPIDFUZZ:
        print("WARNING: rapidfuzz not installed — fuzzy matching disabled. `pip install rapidfuzz`")

    return {
        "Invoice Line Detail": invoice_detail,
        "Returns": returns,
        "Ship-To Rollup": shipto_rollup,
        "Unmatched Invoice Lines": unmatched,
    }, window_start


# ============================================================
# Excel writer
# ============================================================

def write_excel(sheets: dict, out_path: Path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            sheet_name = name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                if len(df) == 0:
                    width = len(str(col)) + 2
                else:
                    max_data_len = df[col].astype(str).str.len().max()
                    if pd.isna(max_data_len):
                        max_data_len = 0
                    width = min(max(len(str(col)), int(max_data_len)) + 2, 50)
                ws.set_column(i, i, width)


# ============================================================
# CLI entry point
# ============================================================

def main():
    try:
        sfdc_folder = PROJECT_ROOT / "INPUTS" / "SFDC_ORDERS"
        if not sfdc_folder.exists():
            for legacy in ("CCD_ADD", "CCD_ADD_REPLACE"):
                alt = PROJECT_ROOT / "INPUTS" / legacy
                if alt.exists():
                    sfdc_folder = alt
                    break

        sfdc_path = _single_file_in_dir(sfdc_folder)
        gp_path = _single_file_in_dir(PROJECT_ROOT / "INPUTS" / "GUIDEPOINT_INVOICE")

        sheets, window_start = run_comparison(sfdc_path, gp_path)

        invoice_month = (window_start + pd.offsets.MonthBegin(1)).strftime("%b").upper()
        out_path = PROJECT_ROOT / "OUTPUTS" / f"ACCOUNT_COMPARISON_{invoice_month}.xlsx"
        write_excel(sheets, out_path)
        print(f"\n✓ Output: {out_path}")
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        raise


if __name__ == "__main__":
    main()

