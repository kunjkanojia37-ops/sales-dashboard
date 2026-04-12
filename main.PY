import pandas as pd
import os
import streamlit as st
import matplotlib.pyplot as plt
import io


# ===========================
# FILE LOAD
# ===========================
# base_dir = os.path.dirname(__file__)
 #os.path.join(base_dir,"raw data","Item_Report_With_CustomerOrder_Details_2026_04_02_13_25_26.xlsx")
st.title("Welcome To Sales Dashboard")


FILE_INPUT_PATH = st.file_uploader("Upload your file here",type=["csv","xlsx"])

st.write("Note - Only upload item report item wise .......... ! ")



def main_work(df):
#  Find header row (where "Date" exists)

 keywords = ["Date", "Invoice No.", "Item Name"]
 header_row = df[df.apply(lambda row: all(k in row.values for k in keywords), axis=1)].index[0]

# Read again with correct header
 df = pd.read_excel(FILE_INPUT_PATH, header=header_row)

#  Remove unwanted rows (like "Total")
 df = df[df["Date"] != "Total"]


 print(df.info())
 print(df.head(6))

# ===========================
# COMMON SETTINGS
# ===========================

 HSN_LENGTH = 4
 DEFAULT_HSN = "2106"

 filtered_HSN = df["HSN"].fillna(DEFAULT_HSN).map(lambda x: str(x)[:HSN_LENGTH])

# ===========================
# TAX FILTERS
# ===========================
 df_40 = df[(df['CGST@20 Amount'] > 0) | (df['SGST@20 Amount'] > 0)]
 df_5 = df[(df['CGST@2.5 Amount'] > 0) | (df['SGST@2.5 Amount'] > 0)]
 df_non_tax = df[df['Non Taxable'] > 0]

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
            result = data.groupby(group_key).agg(agg_dict).round(0)
    
            print("\n", "=" * 50)
            print(title)
            print(result)
            print("=" * 50)
            
            st.write("\n","=" * 50)
            st.write(title)
            st.write("\n","=" * 50)
            st.dataframe(result)
   
            return result
        else:
            print("Data of excel sheet is empty please provide data.....!")
            return None

    except ValueError:
        print("Invalid Input Provided in Data...........!")
        return None


# ===========================
# HSN SUMMARIES
# ===========================
 HSN_SHOW_20 = generate_hsn_summary(
    df_40,
    filtered_HSN,
    {
        "Qty.": "sum",
        "Sub Total": "sum",
        "CGST@20 Amount": "sum",
        "SGST@20 Amount": "sum"
    },
    "40% TAX REPORT WITH HSN WISE"
 )

 HSN_SHOW_5 = generate_hsn_summary(
    df_5,
    filtered_HSN,
    {
        "Qty.": "sum",
        "Sub Total": "sum",
        "CGST@2.5 Amount": "sum",
        "SGST@2.5 Amount": "sum"
    },
    "5% TAX REPORT WITH HSN WISE"
 )

 HSN_SHOW_non_tax = generate_hsn_summary(
    df_non_tax,
    filtered_HSN,
    {
        "Qty.": "sum",
        "Sub Total": "sum",
        "Non Taxable": "sum"
    },
    "NON TAXABLE REPORT WITH HSN WISE"
 )

# ===========================
# COMMON COLUMNS
# ===========================
 COMMON_COLUMNS = [
    'Date', 'HSN', 'Invoice No.', 'GST', 'Qty.',
    'Sub Total', 'Discount', 'Final Total',
    'CGST@2.5 Amount', 'SGST@2.5 Amount',
    'CGST@20 Amount', 'SGST@20 Amount', 'Non Taxable'
 ]

# ===========================
# B2B & B2C SPLIT
# ===========================
 df_common = df[COMMON_COLUMNS]

 df_b2b = df_common[df_common["GST"].notna()].copy()
 df_b2c = df_common[df_common["GST"].isna()].copy()

# ===========================
# FUNCTION: ADD HSN GROUP
# ===========================
 def add_hsn_group(data):
    data['HSN_Group'] = data['HSN'].fillna(DEFAULT_HSN).map(lambda x: str(x)[:HSN_LENGTH])
    return data


# ===========================
# FUNCTION: B2B SUMMARY
# ===========================
 def generate_b2b_summary(data, tax_type, agg_dict, title):
    try:
        if not data.empty:
            data = add_hsn_group(data)

            result = data.groupby([ "GST", 'Date', 'Invoice No.','HSN_Group']).agg(agg_dict).round(0)
      
            print("\n", "=" * 50)
            print(title)
            print(result)
            print("=" * 50)
            
            st.write("\n","=" * 50)
            st.write(title)
            st.write("\n","=" * 50)
            st.dataframe(result)
      

            return result
        else:
            print("Data of excel sheet is empty please provide data.....!")
            return None

    except ValueError:
        print("Invalid Input Provided in Data...........!")
        return None


# ===========================
# FUNCTION: B2C SUMMARY
# ===========================
 def generate_b2c_summary(data, agg_dict, title):
    try:
        if not data.empty:
            data = add_hsn_group(data)

            result = data.groupby(['HSN_Group']).agg(agg_dict).round(0)
           
            print("\n", "=" * 50)
            print(title)
            print(result)
            print("=" * 50)

            st.write("\n","=" * 50)
            st.write(title)
            st.write("\n","=" * 50)
            st.dataframe(result)

            return result
        else:
            print("Data of excel sheet is empty please provide data.....!")
            return None

    except ValueError:
        print("Invalid Input Provided in Data...........!")
        return None


# ===========================
# B2B FILTERS
# ===========================
 df_b2b_5 = df_b2b[(df_b2b['CGST@2.5 Amount'] > 0) | (df_b2b['SGST@2.5 Amount'] > 0)]
 df_b2b_40 = df_b2b[(df_b2b['CGST@20 Amount'] > 0) | (df_b2b['SGST@20 Amount'] > 0)]
 df_b2b_NT = df_b2b[df_b2b['Non Taxable'] > 0]

# ===========================
# B2C FILTERS
# ===========================
 df_b2c_5 = df_b2c[(df_b2c['CGST@2.5 Amount'] > 0) | (df_b2c['SGST@2.5 Amount'] > 0)]
 df_b2c_40 = df_b2c[(df_b2c['CGST@20 Amount'] > 0) | (df_b2c['SGST@20 Amount'] > 0)]
 df_b2c_NT = df_b2c[df_b2c['Non Taxable'] > 0]

# ===========================
# GENERATE REPORTS
# ===========================
 HSN_SHOW_b2b_5 = generate_b2b_summary(df_b2b_5,
    "5%",
    {
        'Qty.': "sum", 'Sub Total': "sum", 'Discount': "sum",
        'Final Total': "sum", 'CGST@2.5 Amount': "sum", 'SGST@2.5 Amount': "sum"
    },
    "B2B SALES REPORT - 5% TAX"
 )

 HSN_SHOW_b2b_40 = generate_b2b_summary(df_b2b_40,
    "40%",
    {
        'Qty.': "sum", 'Sub Total': "sum", 'Discount': "sum",
        'Final Total': "sum", 'CGST@20 Amount': "sum", 'SGST@20 Amount': "sum"
    },
    "B2B SALES REPORT - 40% TAX"
 )

 HSN_SHOW_b2b_NT = generate_b2b_summary(df_b2b_NT,
    "NON",
    {
        'Qty.': "sum", 'Sub Total': "sum", 'Discount': "sum",
        'Final Total': "sum", "Non Taxable": "sum"
    },
    "B2B SALES REPORT - NON TAXABLE"
 )

 HSN_SHOW_b2c_5 = generate_b2c_summary(df_b2c_5,
    {
        'Qty.': "sum", 'Sub Total': "sum", 'Discount': "sum",
        'Final Total': "sum", 'CGST@2.5 Amount': "sum", 'SGST@2.5 Amount': "sum"
    },
    "B2C SALES REPORT - 5% TAX"
 )

 HSN_SHOW_b2c_40 = generate_b2c_summary(df_b2c_40,
    {
        'Qty.': "sum", 'Sub Total': "sum", 'Discount': "sum",
        'Final Total': "sum", 'CGST@20 Amount': "sum", 'SGST@20 Amount': "sum"
    },
    "B2C SALES REPORT - 40% TAX"
 )

 HSN_SHOW_b2c_NT = generate_b2c_summary(df_b2c_NT,
    {
        'Qty.': "sum", 'Sub Total': "sum", 'Discount': "sum",
        'Final Total': "sum", "Non Taxable": "sum"
    },
    "B2C SALES REPORT - NON TAXABLE"
 )

# ===========================
# EXPORT FUNCTION (UNCHANGED LOGIC)
# ===========================
 def export_report():
    try:
        base_path = r"D:\python\project\BRIJWASI SWEET AND RESTORANT\output"
        os.makedirs(base_path, exist_ok=True)

        # HSN Summary
        with pd.ExcelWriter(os.path.join(base_path, "Hsn_summary.xlsx"), engine="openpyxl") as writer:
            if HSN_SHOW_20 is not None:
                HSN_SHOW_20.to_excel(writer, sheet_name='HSN_40%')
            if HSN_SHOW_5 is not None:
                HSN_SHOW_5.to_excel(writer, sheet_name='HSN_5%')
            if HSN_SHOW_non_tax is not None:
                HSN_SHOW_non_tax.to_excel(writer, sheet_name='HSN_NT')

        # B2C
        with pd.ExcelWriter(os.path.join(base_path, "B2C.xlsx"), engine="openpyxl") as writer:
            if HSN_SHOW_b2c_5 is not None:
                HSN_SHOW_b2c_5.to_excel(writer, sheet_name='B2C_5%')
            if HSN_SHOW_b2c_40 is not None:
                HSN_SHOW_b2c_40.to_excel(writer, sheet_name='B2C_40%')
            if HSN_SHOW_b2c_NT is not None:
                HSN_SHOW_b2c_NT.to_excel(writer, sheet_name='B2C_NT')

        # B2B
        with pd.ExcelWriter(os.path.join(base_path, "B2B.xlsx"), engine="openpyxl") as writer:
            if HSN_SHOW_b2b_5 is not None:
                HSN_SHOW_b2b_5.to_excel(writer, sheet_name='B2B_5%')
            if HSN_SHOW_b2b_40 is not None:
                HSN_SHOW_b2b_40.to_excel(writer, sheet_name='B2B_40%')
            if HSN_SHOW_b2b_NT is not None:
                HSN_SHOW_b2b_NT.to_excel(writer, sheet_name='B2B_NT')

        print("All reports exported successfully!")

    except ValueError:
        print("Invalid Input Provided in Data...........!")


 export_report()

 print("\nNow check your data at:", os.getcwd())


# ===========================
# PATH CONFIG
# ===========================
# BASE_PATH = os.path.dirname(__file__)
# GRAPH_PATH = os.path.join(BASE_PATH, "graphical presentation")
# OUTPUT_PATH = os.path.join(BASE_PATH, "output")

# os.makedirs(GRAPH_PATH, exist_ok=True)
# os.makedirs(OUTPUT_PATH, exist_ok=True)


# ===========================
# FUNCTION: PRINT FORMAT
# ===========================
 def print_block(title):
    print("\n", "=" * 50)
    print(f"{title}\n")
    print("=" * 50)
    st.write("\n", "=" * 50)
    st.title(f"{title}\n")
    st.write("=" * 50)

# ===========================
# POPULAR ITEMS
# ===========================
 Item_sold = df.groupby("Item Name").agg({
    "Final Total": "sum"
 })

 max_item_sold = Item_sold.sort_values(by="Final Total", ascending=False).head(5)

 print_block("Best Item sold in the month is")

# GRAPH: POPULAR ITEMS
 fig1 , ax1 = plt.subplots()

 ax1.bar(x=max_item_sold.index,
    height=max_item_sold["Final Total"]
 )
 ax1.set_title("Top Items")
 ax1.tick_params(axis='x', rotation=45)
 
 st.pyplot(fig1)

# ===========================
# DAILY SALES SUMMARY
# ===========================
 All_day = df.groupby("Date").agg({
    "Qty.": "sum",
    "Sub Total": "sum",
    "CGST@2.5 Amount": "sum",
    "SGST@2.5 Amount": "sum",
    "Non Taxable": "sum",
    "CGST@20 Amount": "sum",
    "SGST@20 Amount": "sum",
    "Final Total": "sum"
 }).round(0)

 print_block("The sales in a day is")

 best_day_sorted = All_day.sort_values(by="Final Total", ascending=False)
 show_top_five = best_day_sorted.head(5)

# GRAPH: TOP DAYS
 fig2 , ax2 = plt.subplots()
 ax2.bar(
    x=show_top_five.index,
    height=show_top_five["Final Total"]
 )

 ax2.set_title("Best Day in Month")
 ax2.tick_params(axis='x', rotation=45)
 st.pyplot(fig2)

# ===========================
# PAYMENT MODE
# ===========================
 payment_method = df.groupby("Payment Type").agg({
    "Final Total": "sum"
 }).sort_values(by="Payment Type", ascending=False)

 print_block("The mode of payment is")

 payment_mode = df["Payment Type"].sort_values(ascending=False).unique()

 fig3 , ax3 = plt.subplots()
 ax3.pie(
    payment_method["Final Total"],
    labels=payment_mode,
    colors=["orange", "pink", "yellow", "blue"],
    autopct="%1.1f%%"
 )
 ax3.set_title("Mode of payment")
 st.pyplot(fig3)



# ===========================
# CUSTOMER SUMMARY (B2B)
# ===========================
 print("\n" + "=" * 80)
 print("CUSTOMER-WISE B2B SALES SUMMARY")
 print("=" * 80)

 b2b_customer_summary = df.groupby(['GST', 'Name']).agg({
    'Invoice No.': 'nunique',
    'Qty.': 'sum',
    'Sub Total': 'sum',
    'Discount': 'sum',
    'Tax': 'sum',
    'Final Total': 'sum'
 }).round(2)

 b2b_customer_summary.columns = [
    'Number_of_Invoices',
    'Total_Quantity',
    'Total_Taxable_Value',
    'Total_discount',
    'Total_tax',
    'Total_spend'
 ]

 print(b2b_customer_summary.to_string())
 st.dataframe(b2b_customer_summary)

# ===========================
# DOWNLOAD FUNCTIONS
# ===========================

 import zipfile
 import io 

 zip_buffer = io.BytesIO()
 def df_to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=True)
    output.seek(0)
    return output.getvalue()
 zip_buffer.seek(0)
 
 zip_buffer = io.BytesIO()

 with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

    if HSN_SHOW_b2b_40 is not None:
        zip_file.writestr("B2B_40.xlsx", df_to_excel_bytes(HSN_SHOW_b2b_40))

    if HSN_SHOW_b2b_5 is not None:
        zip_file.writestr("B2B_5.xlsx", df_to_excel_bytes(HSN_SHOW_b2b_5))

    if HSN_SHOW_b2b_NT is not None:
        zip_file.writestr("B2B_NT.xlsx", df_to_excel_bytes(HSN_SHOW_b2b_NT))

    if HSN_SHOW_b2c_40 is not None:
        zip_file.writestr("B2C_40.xlsx", df_to_excel_bytes(HSN_SHOW_b2c_40))

    if HSN_SHOW_b2c_5 is not None:
        zip_file.writestr("B2C_5.xlsx", df_to_excel_bytes(HSN_SHOW_b2c_5))

    if HSN_SHOW_b2c_NT is not None:
        zip_file.writestr("B2C_NT.xlsx", df_to_excel_bytes(HSN_SHOW_b2c_NT))

    if b2b_customer_summary is not None:
        zip_file.writestr("B2B_customer_summary.xlsx", df_to_excel_bytes(b2b_customer_summary))

 zip_buffer.seek(0)
 
 st.success("✅ Reports generated successfully!")

 st.download_button(
    label="Download all files (ZIP)",
    data=zip_buffer,
    file_name="all_files.zip",
    mime="application/zip"  
 )

 def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    buf.seek(0)
    return buf.getvalue()
 
 img1 = fig_to_bytes(fig1)
 img2 = fig_to_bytes(fig2)
 img3 = fig_to_bytes(fig3)
 zip_buffer = io.BytesIO()

 with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

   # Charts 

    zip_file.writestr("Top_Items.png", img1)
    zip_file.writestr("Best_Days.png", img2)
    zip_file.writestr("Payment_Mode.png", img3)

 zip_buffer.seek(0)
 
 st.download_button(
    "Download Charts",
    data=zip_buffer,
    file_name="full_dashboard.zip",
    mime="application/zip",
    key="zip_with_charts"
)

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
    st.info("Please upload a file to begin...")

 except Exception as e:
        st.error(f"Error reading file: {str(e)}")

one_for_all()