import sys
import streamlit as st
from streamlit.web import cli as stcli

if __name__ == '__main__':
    # Check if we are already running inside an active Streamlit server instance
    try:
        from streamlit.runtime import exists as st_exists
        is_running_in_streamlit = st_exists()
    except ImportError:
        is_running_in_streamlit = False


# =========================================================
# THE REST OF YOUR CODE CONTINUES EXACTLY THE SAME...
# =========================================================
st.set_page_config(page_title="Sales Dashboard", layout="wide")
# ... (rest of your imports and functions remain untouched)

# ===========================
# IMPORT OTHER LIBRARIES
# ===========================

import pandas as pd
import os
import matplotlib.pyplot as plt
import io
import zipfile

# ===========================
# STREAMLIT UI SETUP
# ===========================

st.set_page_config(page_title="Sales Dashboard", layout="wide")
st.title("Welcome To Sales Dashboard")
st.write("Note - Only upload item report item wise .......... ! ")

FILE_INPUT_PATH = st.file_uploader("Upload your file here", type=["csv", "xlsx"])


def main_work(df):

    # Find header row (where "Date" exists)
    keywords = ["Date", "Invoice No.", "Item Name"]
    
    try:
        header_row = df[df.apply(lambda row: all(k in row.values for k in keywords), axis=1)].index[0]
    except IndexError:
        st.error("Could not find the standard headers ('Date', 'Invoice No.', 'Item Name') in the uploaded file.")
        return

    # Read again with correct header
    if FILE_INPUT_PATH.name.endswith("csv"):
        FILE_INPUT_PATH.seek(0)
        df = pd.read_csv(FILE_INPUT_PATH, header=header_row)
    else:
        FILE_INPUT_PATH.seek(0)
        df = pd.read_excel(FILE_INPUT_PATH, header=header_row)

    # Remove unwanted rows (like "Total")
    df = df[df["Date"] != "Total"]

    # ===========================
    # COMMON SETTINGS
    # ===========================
    HSN_LENGTH = 4
    DEFAULT_HSN = "2106"


    if "HSN" in df.columns:
        filtered_HSN = df["HSN"].fillna(DEFAULT_HSN).map(lambda x: str(x)[:HSN_LENGTH])
    else:
        filtered_HSN = pd.Series([DEFAULT_HSN] * len(df))

    # ===========================
    # TAX FILTERS
    # ===========================
    df_40 = df[(df.get('CGST@20 Amount', 0) > 0) | (df.get('SGST@20 Amount', 0) > 0)]
    df_5 = df[(df.get('CGST@2.5 Amount', 0) > 0) | (df.get('SGST@2.5 Amount', 0) > 0)]
    df_non_tax = df[df.get('Non Taxable', 0) > 0]

    # ===========================
    # TITLE OF WEB
    # ===========================
    df["Date"] = pd.to_datetime(df["Date"])
    df["Month-Year"] = df["Date"].dt.month_name() + " - " + df["Date"].dt.year.astype(str)
    st.markdown(f"## This Is Your {df['Month-Year'].iloc[0]} Report...... ")

    # ===========================
    # FUNCTION: HSN SUMMARY
    # ===========================
    def generate_hsn_summary(data, group_key, agg_dict, title):
        try:
            if not data.empty:
                # Align group_key to the filtered data's index length
                current_group_key = group_key.loc[data.index]
                result = data.groupby(current_group_key).agg({k: v for k, v in agg_dict.items() if k in data.columns}).round(0)
        
                st.write("---")
                st.subheader(title)
                st.dataframe(result, use_container_width=True)
                return result
            else:
                return None
        except Exception as e:
            st.warning(f"Could not generate {title}: {str(e)}")
            return None

    # ===========================
    # HSN SUMMARIES
    # ===========================
    HSN_SHOW_20 = generate_hsn_summary(
        df_40, filtered_HSN,
        {"Qty.": "sum", "Sub Total": "sum", "CGST@20 Amount": "sum", "SGST@20 Amount": "sum"},
        "40% TAX REPORT WITH HSN WISE"
    )

    HSN_SHOW_5 = generate_hsn_summary(
        df_5, filtered_HSN,
        {"Qty.": "sum", "Sub Total": "sum", "CGST@2.5 Amount": "sum", "SGST@2.5 Amount": "sum"},
        "5% TAX REPORT WITH HSN WISE"
    )

    HSN_SHOW_non_tax = generate_hsn_summary(
        df_non_tax, filtered_HSN,
        {"Qty.": "sum", "Sub Total": "sum", "Non Taxable": "sum"},
        "NON TAXABLE REPORT WITH HSN WISE"
    )

    # ===========================
    # B2B & B2C SPLIT
    # ===========================
    COMMON_COLUMNS = [
        'Date', 'HSN', 'Invoice No.', 'GST', 'Qty.', 'Sub Total', 'Discount', 
        'Final Total', 'CGST@2.5 Amount', 'SGST@2.5 Amount', 'CGST@20 Amount', 
        'SGST@20 Amount', 'Non Taxable'
    ]
    
    # Keep only available common columns
    available_cols = [col for col in COMMON_COLUMNS if col in df.columns]
    df_common = df[available_cols].copy()

    if "GST" in df_common.columns:
        df_b2b = df_common[df_common["GST"].notna()].copy()
        df_b2c = df_common[df_common["GST"].isna()].copy()
    else:
        df_b2b = pd.DataFrame(columns=df_common.columns)
        df_b2c = df_common.copy()

    def add_hsn_group(data):
        if 'HSN' in data.columns:
            data['HSN_Group'] = data['HSN'].fillna(DEFAULT_HSN).map(lambda x: str(x)[:HSN_LENGTH])
        else:
            data['HSN_Group'] = DEFAULT_HSN
        return data

    # ===========================
    # FUNCTION: B2B SUMMARY
    # ===========================
    def generate_b2b_summary(data, agg_dict, title):
        try:
            if not data.empty:
                data = add_hsn_group(data)
                groupby_cols = [col for col in ["GST", 'Date', 'Invoice No.', 'HSN_Group'] if col in data.columns]
                actual_agg = {k: v for k, v in agg_dict.items() if k in data.columns}
                result = data.groupby(groupby_cols).agg(actual_agg).round(0)
          
                st.write("---")
                st.subheader(title)
                st.dataframe(result, use_container_width=True)
                return result
            return None
        except Exception as e:
            st.warning(f"Could not generate {title}: {str(e)}")
            return None

    # ===========================
    # FUNCTION: B2C SUMMARY
    # ===========================
    def generate_b2c_summary(data, agg_dict, title):
        try:
            if not data.empty:
                data = add_hsn_group(data)
                actual_agg = {k: v for k, v in agg_dict.items() if k in data.columns}
                result = data.groupby(['HSN_Group']).agg(actual_agg).round(0)
               
                st.write("---")
                st.subheader(title)
                st.dataframe(result, use_container_width=True)
                return result
            return None
        except Exception as e:
            st.warning(f"Could not generate {title}: {str(e)}")
            return None

    # ===========================
    # B2B & B2C FILTERS
    # ===========================
    df_b2b_5 = df_b2b[(df_b2b.get('CGST@2.5 Amount', 0) > 0) | (df_b2b.get('SGST@2.5 Amount', 0) > 0)] if not df_b2b.empty else df_b2b
    df_b2b_40 = df_b2b[(df_b2b.get('CGST@20 Amount', 0) > 0) | (df_b2b.get('SGST@20 Amount', 0) > 0)] if not df_b2b.empty else df_b2b
    df_b2b_NT = df_b2b[df_b2b.get('Non Taxable', 0) > 0] if not df_b2b.empty else df_b2b

    df_b2c_5 = df_b2c[(df_b2c.get('CGST@2.5 Amount', 0) > 0) | (df_b2c.get('SGST@2.5 Amount', 0) > 0)] if not df_b2c.empty else df_b2c
    df_b2c_40 = df_b2c[(df_b2c.get('CGST@20 Amount', 0) > 0) | (df_b2c.get('SGST@20 Amount', 0) > 0)] if not df_b2c.empty else df_b2c
    df_b2c_NT = df_b2c[df_b2c.get('Non Taxable', 0) > 0] if not df_b2c.empty else df_b2c

    # ===========================
    # GENERATE B2B/B2C REPORTS
    # ===========================
    b2b_c_dict = {'Qty.': "sum", 'Sub Total': "sum", 'Discount': "sum", 'Final Total': "sum", 'CGST@2.5 Amount': "sum", 'SGST@2.5 Amount': "sum"}
    b2b_h_dict = {'Qty.': "sum", 'Sub Total': "sum", 'Discount': "sum", 'Final Total': "sum", 'CGST@20 Amount': "sum", 'SGST@20 Amount': "sum"}
    b2b_n_dict = {'Qty.': "sum", 'Sub Total': "sum", 'Discount': "sum", 'Final Total': "sum", "Non Taxable": "sum"}

    HSN_SHOW_b2b_5 = generate_b2b_summary(df_b2b_5, b2b_c_dict, "B2B SALES REPORT - 5% TAX")
    HSN_SHOW_b2b_40 = generate_b2b_summary(df_b2b_40, b2b_h_dict, "B2B SALES REPORT - 40% TAX")
    HSN_SHOW_b2b_NT = generate_b2b_summary(df_b2b_NT, b2b_n_dict, "B2B SALES REPORT - NON TAXABLE")

    HSN_SHOW_b2c_5 = generate_b2c_summary(df_b2c_5, b2b_c_dict, "B2C SALES REPORT - 5% TAX")
    HSN_SHOW_b2c_40 = generate_b2c_summary(df_b2c_40, b2b_h_dict, "B2C SALES REPORT - 40% TAX")
    HSN_SHOW_b2c_NT = generate_b2c_summary(df_b2c_NT, b2b_n_dict, "B2C SALES REPORT - NON TAXABLE")

    # ===========================
    # VISUAL CHARTS
    # ===========================
    st.write("---")
    st.header(" Analytical Graphs")
    
    col1, col2, col3 = st.columns(3)

    # Chart 1: Popular Items
    with col1:
        if "Item Name" in df.columns and "Final Total" in df.columns:
            Item_sold = df.groupby("Item Name").agg({"Final Total": "sum"})
            max_item_sold = Item_sold.sort_values(by="Final Total", ascending=False).head(5)
            
            fig1, ax1 = plt.subplots()
            ax1.bar(x=max_item_sold.index, height=max_item_sold["Final Total"], color='skyblue')
            ax1.set_title("Top 5 Items Sold")
            ax1.tick_params(axis='x', rotation=45)
            st.pyplot(fig1)
        else:
            st.warning("Missing 'Item Name' or 'Final Total' for charting.")
            fig1 = None

    # Chart 2: Daily Sales
    with col2:
        day_cols = [c for c in ["Qty.", "Sub Total", "Final Total"] if c in df.columns]
        if "Date" in df.columns and day_cols:
            All_day = df.groupby("Date").agg({c: "sum" for c in day_cols}).round(0)
            show_top_five = All_day.sort_values(by="Final Total", ascending=False).head(5)
            
            fig2, ax2 = plt.subplots()
            ax2.bar(x=show_top_five.index.astype(str), height=show_top_five["Final Total"], color='lightgreen')
            ax2.set_title("Best Performing Days")
            ax2.tick_params(axis='x', rotation=45)
            st.pyplot(fig2)
        else:
            fig2 = None

    # Chart 3: Payment Type Breakdown
    with col3:
        if "Payment Type" in df.columns and "Final Total" in df.columns:
            payment_method = df.groupby("Payment Type").agg({"Final Total": "sum"})
            
            fig3, ax3 = plt.subplots()
            ax3.pie(payment_method["Final Total"], labels=payment_method.index, autopct="%1.1f%%", startangle=90)
            ax3.set_title("Payment Distribution")
            st.pyplot(fig3)
        else:
            fig3 = None

    # ===========================
    # CUSTOMER SUMMARY (B2B)
    # ===========================
    b2b_customer_summary = None
    if "GST" in df.columns and "Name" in df.columns:
        st.write("---")
        st.header(" Customer-wise B2B Summary")
        
        cust_agg = {}
        for k, v in {'Invoice No.': 'nunique', 'Qty.': 'sum', 'Sub Total': 'sum', 'Discount': 'sum', 'Tax': 'sum', 'Final Total': 'sum'}.items():
            if k in df.columns:
                cust_agg[k] = v

        b2b_customer_summary = df.groupby(['GST', 'Name']).agg(cust_agg).round(2)
        st.dataframe(b2b_customer_summary, use_container_width=True)

    # ===========================
    # SYSTEM DOWNLOAD IMPLEMENTATIONS
    # ===========================
    st.write("---")
    st.header(" Export Generated Outputs")

    def df_to_excel_bytes(dataframe):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            dataframe.to_excel(writer, index=True)
        output.seek(0)
        return output.getvalue()

    # Buffer 1: Data Reports
    excel_zip_buffer = io.BytesIO()
    with zipfile.ZipFile(excel_zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        if HSN_SHOW_b2b_40 is not None: zip_file.writestr("B2B_40.xlsx", df_to_excel_bytes(HSN_SHOW_b2b_40))
        if HSN_SHOW_b2b_5 is not None: zip_file.writestr("B2B_5.xlsx", df_to_excel_bytes(HSN_SHOW_b2b_5))
        if HSN_SHOW_b2b_NT is not None: zip_file.writestr("B2B_NT.xlsx", df_to_excel_bytes(HSN_SHOW_b2b_NT))
        if HSN_SHOW_b2c_40 is not None: zip_file.writestr("B2C_40.xlsx", df_to_excel_bytes(HSN_SHOW_b2c_40))
        if HSN_SHOW_b2c_5 is not None: zip_file.writestr("B2C_5.xlsx", df_to_excel_bytes(HSN_SHOW_b2c_5))
        if HSN_SHOW_b2c_NT is not None: zip_file.writestr("B2C_NT.xlsx", df_to_excel_bytes(HSN_SHOW_b2c_NT))
        if b2b_customer_summary is not None: zip_file.writestr("B2B_customer_summary.xlsx", df_to_excel_bytes(b2b_customer_summary))
    
    excel_zip_buffer.seek(0)

    st.download_button(
        label=" Download Data Sheets (ZIP)",
        data=excel_zip_buffer,
        file_name="sales_data_sheets.zip",
        mime="application/zip",
        key="data_zip"
    )

    # Buffer 2: Graphical Output Charts
    def fig_to_bytes(fig_obj):
        if fig_obj is None: return b""
        buf = io.BytesIO()
        fig_obj.savefig(buf, format="png", bbox_inches='tight')
        buf.seek(0)
        return buf.getvalue()

    chart_zip_buffer = io.BytesIO()
    with zipfile.ZipFile(chart_zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        if fig1: zip_file.writestr("Top_Items.png", fig_to_bytes(fig1))
        if fig2: zip_file.writestr("Best_Days.png", fig_to_bytes(fig2))
        if fig3: zip_file.writestr("Payment_Mode.png", fig_to_bytes(fig3))
        
    chart_zip_buffer.seek(0)

    st.download_button(
        label=" Download Dashboard Charts (ZIP)",
        data=chart_zip_buffer,
        file_name="dashboard_charts.zip",
        mime="application/zip",
        key="charts_zip"
    )
    st.success(" Dashboard fully structured successfully!")


def one_for_all():
    try:
        if FILE_INPUT_PATH is not None:
            if FILE_INPUT_PATH.name.endswith("csv"):
                df = pd.read_csv(FILE_INPUT_PATH, header=None)
                main_work(df)
            elif FILE_INPUT_PATH.name.endswith('.xlsx'):
                df = pd.read_excel(FILE_INPUT_PATH, header=None)
                main_work(df)
            else:
                st.error("Unsupported file format. Please upload CSV or XLSX file.")
        else:
            st.info("Please upload an item report file to begin parsing...")
    except Exception as e:
        st.error(f"Critical execution failure: {str(e)}")


one_for_all()
