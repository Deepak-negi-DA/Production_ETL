import streamlit as st
import pandas as pd
from io import BytesIO


# -------- Consolidation Function --------
def consolidate_excel_sheets(file):
    # Read the entire Excel file
    xls = pd.ExcelFile(file)
    master_df = pd.DataFrame()

    # Loop through sheets
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Skip completely empty sheets
        if df.empty:
            continue

        # Align columns (handle sheets with different structures)
        if master_df.empty:
            master_df = df.copy()
        else:
            master_df = pd.concat([master_df, df], ignore_index=True)

    return master_df


# -------- Streamlit App --------
st.title("📊 Excel Sheet Consolidator")
st.write("Upload a multi-sheet Excel file, and this tool will merge all sheets into a single 'Master' sheet.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    # Run consolidation
    master_df = consolidate_excel_sheets(uploaded_file)

    st.success(f"✅ Consolidation complete! Total rows: {len(master_df)}")

    # Show preview
    st.subheader("Preview of Master Sheet")
    st.dataframe(master_df.head(100))  # show only first 100 rows for performance

    # Prepare Excel for download
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        master_df.to_excel(writer, index=False, sheet_name="Master")

    st.download_button(
        label="📥 Download Master Sheet",
        data=output.getvalue(),
        file_name="master_consolidated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

