"""
IPHS – Loading List vs Stock Checker
Streamlit app to compare a vessel loading list against IPHS stock (Empties + Live).
"""

import io
from typing import Optional, Tuple, List

import pandas as pd
import streamlit as st


# -----------------------
# Helper / utility funcs
# -----------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize column names:
      - lowercase
      - strip spaces
      - replace spaces with underscores
      - remove dots
    Returns a new dataframe with replaced columns (data unchanged).
    """
    df = df.copy()
    cols = [str(c).lower().strip().replace(" ", "_").replace(".", "") for c in df.columns]
    df.columns = cols
    return df


def detect_container_column(
    df: pd.DataFrame, keywords: List[str]
) -> Optional[str]:
    """
    Detect the first column name that contains any of the keywords.
    Returns the column name (already normalized) or None.
    """
    for c in df.columns:
        for kw in keywords:
            if kw in c:
                return c
    return None


def prepare_loading_list(df: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    """
    Read/prepare the loading list dataframe:
      - assume df is already read from the first sheet
      - normalize columns
      - detect first column that contains 'container'
      - create _cntr column (string, upper, strip)
    Returns prepared df and the container column name detected.
    Raises ValueError if container column not found.
    """
    df = normalize_columns(df)
    cntr_col = detect_container_column(df, ["container"])
    if cntr_col is None:
        raise ValueError(
            "Could not detect a container column in the loading list. "
            "Search looked for a column containing the word 'container'."
        )
    df["_cntr"] = df[cntr_col].astype(str).str.upper().str.strip()
    return df, cntr_col


def prepare_stock(empties_df: pd.DataFrame, live_df: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    """
    Prepare stock dataframe by:
      - normalizing columns on both empties & live
      - concatenating them (empties first, then live)
      - detecting container column as first column containing 'container', 'unit' or 'cntr'
      - creating _cntr (string, upper, strip)
    Returns concatenated stock df and detected container column name.
    Raises ValueError if expected sheets missing or container column not found.
    """
    empties_df = normalize_columns(empties_df) if not empties_df.empty else pd.DataFrame()
    live_df = normalize_columns(live_df) if not live_df.empty else pd.DataFrame()

    stock_df = pd.concat(
        [
            empties_df.assign(_source="Empties") if not empties_df.empty else pd.DataFrame(),
            live_df.assign(_source="Live") if not live_df.empty else pd.DataFrame(),
        ],
        ignore_index=True,
        sort=False,
    )

    if stock_df.empty:
        raise ValueError("Both Empties and Live sheets are empty or missing in the stock file.")

    cntr_col = detect_container_column(stock_df, ["container", "unit", "cntr"])
    if cntr_col is None:
        raise ValueError(
            "Could not detect a container/unit/cntr column in stock sheets. "
            "Search looked for columns containing 'container', 'unit' or 'cntr'."
        )

    stock_df["_cntr"] = stock_df[cntr_col].astype(str).str.upper().str.strip()
    return stock_df, cntr_col


def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """
    Convert a dataframe to an in-memory Excel file (bytes).
    """
    towrite = io.BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    towrite.seek(0)
    return towrite.read()


# -----------------------
# Streamlit UI
# -----------------------
st.set_page_config(page_title="IPHS – Loading List vs Stock Checker", layout="wide")
st.title("IPHS – Loading List vs Stock Checker")

st.markdown(
    """
Upload two Excel files:
- A **loading list** (first sheet will be used automatically),
- A **stock file** containing sheets named **Empties** and **Live**.
"""
)

col1, col2 = st.columns(2)
with col1:
    load_file = st.file_uploader("Upload Loading List (.xlsx) - first sheet will be used", type=["xlsx"], key="load")
with col2:
    stock_file = st.file_uploader("Upload Stock file (.xlsx) - must contain sheets named 'Empties' and 'Live'", type=["xlsx"], key="stock")

# Require both uploads
if not load_file or not stock_file:
    st.info("Please upload both files (Loading List and Stock file) to proceed.")
    st.stop()

# Main processing with error handling
try:
    # Read loading list: first sheet by default
    load_xl = pd.ExcelFile(load_file)
    first_sheet_name = load_xl.sheet_names[0]
    df_load_raw = load_xl.parse(first_sheet_name)

    # Read stock: require sheets named exactly "Empties" and "Live"
    stock_xl = pd.ExcelFile(stock_file)
    snames = stock_xl.sheet_names

    # Attempt to find exact sheet names "Empties" and "Live"
    if "Empties" not in snames or "Live" not in snames:
        # Show clear error per requirement
        missing = []
        if "Empties" not in snames:
            missing.append("Empties")
        if "Live" not in snames:
            missing.append("Live")
        raise ValueError(
            f"Stock file is missing required sheet(s): {', '.join(missing)}. "
            "Please ensure the stock workbook contains sheets named exactly 'Empties' and 'Live'."
        )

    df_empties_raw = stock_xl.parse("Empties")
    df_live_raw = stock_xl.parse("Live")

    # Prepare dataframes
    df_load, load_container_col = prepare_loading_list(df_load_raw)
    df_stock, stock_container_col = prepare_stock(df_empties_raw, df_live_raw)

    # Matching logic: outer merge with indicator so we can split matched/unmatched
    merged_all = pd.merge(
        df_load,
        df_stock,
        on="_cntr",
        how="outer",
        indicator=True,
        suffixes=("_load", "_stock"),
    )

    # Matched: present in both
    matched_df = merged_all[merged_all["_merge"] == "both"].copy().reset_index(drop=True)
    # Unmatched in Stock: present in loading list but NOT in stock (left_only)
    unmatched_in_stock_df = merged_all[merged_all["_merge"] == "left_only"].copy().reset_index(drop=True)
    # Unused stock: present in stock but not in loading list (right_only)
    unused_stock_df = merged_all[merged_all["_merge"] == "right_only"].copy().reset_index(drop=True)

    # Summary numbers
    total_in_loading_list = df_load["_cntr"].nunique()
    total_in_stock = df_stock["_cntr"].nunique()
    total_matched = matched_df["_cntr"].nunique()

    # UI: Summary
    st.subheader("Summary")
    st.metric("Total containers in loading list", f"{total_in_loading_list:,}")
    st.metric("Total containers in stock (Empties + Live)", f"{total_in_stock:,}")
    st.metric("Total matched containers", f"{total_matched:,}")

    st.markdown("---")

    # Matched Containers table
    st.subheader("Matched Containers")
    st.write("Containers that appear in BOTH the Loading List and IPHS Stock (Empties + Live).")
    st.dataframe(matched_df, height=400)

    # Download for matched
    buf_matched = df_to_excel_bytes(matched_df, sheet_name="Matched")
    st.download_button(
        label="Download Matched Containers (.xlsx)",
        data=buf_matched,
        file_name="matched_containers.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("---")

    # Unmatched in Stock (appear in loading list but not in stock)
    st.subheader("Unmatched in Stock (In Loading List but NOT in Stock)")
    st.write("These containers were in the loading list but do not appear in Empties or Live.")
    st.dataframe(unmatched_in_stock_df, height=300)
    buf_unmatched = df_to_excel_bytes(unmatched_in_stock_df, sheet_name="Unmatched_in_Stock")
    st.download_button(
        label="Download Unmatched in Stock (.xlsx)",
        data=buf_unmatched,
        file_name="unmatched_in_stock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("---")

    # Unused stock (appear in stock but not on loading list)
    st.subheader("Unused Stock (Not on Loading List)")
    st.write("These containers are in Empties or Live but are NOT present in the loading list.")
    st.dataframe(unused_stock_df, height=300)
    buf_unused = df_to_excel_bytes(unused_stock_df, sheet_name="Unused_Stock")
    st.download_button(
        label="Download Unused Stock (.xlsx)",
        data=buf_unused,
        file_name="unused_stock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.error(f"Error: {e}")
    st.stop()

