import streamlit as st
import pandas as pd
from io import BytesIO

# ======================
# PAGE CONFIG
# ======================
st.set_page_config(
    page_title="StockSync",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================
# CUSTOM CSS (IMPORTANT)
# ======================
st.markdown("""
<style>

/* Main background */
[data-testid="stAppViewContainer"] {
    background: #f6f9fc;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0b1f3a, #102a4c);
    color: white;
}

/* Sidebar text */
[data-testid="stSidebar"] * {
    color: white;
}

/* Cards */
.card {
    background: white;
    padding: 18px;
    border-radius: 16px;
    box-shadow: 0px 6px 18px rgba(0,0,0,0.08);
    margin-bottom: 15px;
}

/* Metrics */
.metric-box {
    background: white;
    padding: 15px;
    border-radius: 14px;
    text-align: center;
    box-shadow: 0px 4px 12px rgba(0,0,0,0.06);
}

/* Title */
.title {
    font-size: 34px;
    font-weight: 700;
    color: #0b1f3a;
}

/* Subtitle */
.subtitle {
    font-size: 15px;
    color: #6b7c93;
}

/* Button override */
.stDownloadButton button {
    background-color: #16a34a;
    color: white;
    border-radius: 10px;
    padding: 10px 18px;
    border: none;
}

</style>
""", unsafe_allow_html=True)

# ======================
# SIDEBAR
# ======================
with st.sidebar:
    st.title("📦 StockSync")
    st.write("Smart Inventory & Order System")

    menu = st.radio(
        "Navigation",
        ["Home", "Upload & Process", "Processed Data", "About"]
    )

# ======================
# HEADER
# ======================
st.markdown('<div class="title">Smart Inventory & Order Reconciliation System</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Automate stock tracking between reserved quantities and incoming orders in real time.</div>', unsafe_allow_html=True)

st.divider()

# ======================
# PROCESSING FUNCTION
# ======================
def process_data(df):

    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["XEE Reserved Qty"] = pd.to_numeric(df["XEE Reserved Qty"], errors="coerce").fillna(0)
    df["Reserved Quantity"] = pd.to_numeric(df["Reserved Quantity"], errors="coerce").fillna(0)

    df["AV"] = df.apply(
        lambda r: r["Reserved Quantity"] if r["XEE Reserved Qty"] == 0
        else r["XEE Reserved Qty"] - r["Reserved Quantity"],
        axis=1
    )

    df["No Order"] = df.apply(
        lambda r: r["Quantity"] - r["AV"]
        if r["XEE Reserved Qty"] == 0
        else r["Quantity"] - r["AV"] - r["XEE Reserved Qty"],
        axis=1
    )

    return df[["No.", "Description", "Quantity", "AV", "No Order"]]


def save_excel(original_sheets, processed_df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in original_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

        processed_df.to_excel(writer, sheet_name="Processed Data", index=False)

    output.seek(0)
    return output

# ======================
# UPLOAD SECTION (CARD STYLE)
# ======================
st.markdown('<div class="card">', unsafe_allow_html=True)

uploaded_file = st.file_uploader("📤 Upload Excel File (.xlsx)", type=["xlsx"])

st.markdown("</div>", unsafe_allow_html=True)

# ======================
# MAIN LOGIC
# ======================
if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    sheets = {s: pd.read_excel(xls, sheet_name=s) for s in xls.sheet_names}
    df = sheets[xls.sheet_names[0]]

    processed = process_data(df)

    # ======================
    # METRICS ROW
    # ======================
    col1, col2, col3, col4 = st.columns(4)

    col1.markdown(f'<div class="metric-box"><h3>{len(df)}</h3><p>Rows Loaded</p></div>', unsafe_allow_html=True)
    col2.markdown(f'<div class="metric-box"><h3>{len(processed)}</h3><p>Rows Processed</p></div>', unsafe_allow_html=True)
    col3.markdown(f'<div class="metric-box"><h3>5</h3><p>Columns Generated</p></div>', unsafe_allow_html=True)
    col4.markdown(f'<div class="metric-box"><h3>OK</h3><p>Status</p></div>', unsafe_allow_html=True)

    st.divider()

    # ======================
    # PREVIEW TABLE
    # ======================
    st.subheader("📊 Processed Data Preview")
    st.dataframe(processed, use_container_width=True)

    # ======================
    # DOWNLOAD
    # ======================
    excel_file = save_excel(sheets, processed)

    st.download_button(
        "⬇️ Download Processed Excel",
        data=excel_file,
        file_name="stocksync_processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )