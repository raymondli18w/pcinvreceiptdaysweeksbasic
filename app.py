import streamlit as st
import pandas as pd
from datetime import date

# Page config
st.set_page_config(page_title="Receipt Aging Adder", layout="wide")
st.title("📦 Add Days & Weeks Since Receipt")
st.caption("Upload a Piece Inventory .xlsx file — new aging columns will be added on the right.")

# User selects "today"
today = st.date_input("📅 Set 'As Of' Date (default = today)", value=date.today())

# File uploader
uploaded_file = st.file_uploader("📤 Upload Piece Inventory Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Load data
        df = pd.read_excel(uploaded_file, header=0)
        original_cols = df.columns.tolist()
        df.columns = df.columns.astype(str).str.strip()

        # Check for Receipt Date
        if 'Receipt Date' not in df.columns:
            st.error(f"'Receipt Date' column not found. Available columns: {', '.join(df.columns)}")
            st.stop()

        # Parse Receipt Date as MM/DD/YYYY (Jan 9, not Sep 1)
        df['Receipt Date'] = pd.to_datetime(
            df['Receipt Date'],
            errors='coerce',
            dayfirst=False,  # Critical: MM/DD/YYYY
            infer_datetime_format=True
        ).dt.date

        # Drop rows with invalid dates (optional — or keep as NaT)
        # df = df.dropna(subset=['Receipt Date']).copy()

        # Calculate aging
        def calc_days(receipt_date):
            if pd.isnull(receipt_date):
                return None
            return (today - receipt_date).days

        df['Days Since Receipt Date'] = df['Receipt Date'].apply(calc_days)
        df['Weeks Since Receipt Date'] = (df['Days Since Receipt Date'] / 7).round(2)

        # Weeks from start of month (only if receipt is from prior month)
        start_of_month = today.replace(day=1)
        def weeks_from_start_month(receipt_date):
            if pd.isnull(receipt_date):
                return None
            if receipt_date.year == today.year and receipt_date.month == today.month:
                return None
            return round((today - start_of_month).days / 7, 2)

        df['Weeks from Start of Month to Today'] = df['Receipt Date'].apply(weeks_from_start_month)

        # Reorder: original columns first, then new ones at the end
        new_cols = [
            'Days Since Receipt Date',
            'Weeks Since Receipt Date',
            'Weeks from Start of Month to Today'
        ]
        reordered_cols = original_cols + [col for col in new_cols if col not in original_cols]
        df_output = df[reordered_cols]

        # Preview
        st.success("✅ Processed successfully!")
        st.subheader("Preview (first 20 rows):")
        st.dataframe(df_output.head(20))

        # Download button
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_output.to_excel(writer, index=False, sheet_name='Aging Report')
        output.seek(0)

        st.download_button(
            label="📥 Download Updated Excel File",
            data=output,
            file_name="piece_inventory_with_aging.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.stop()
