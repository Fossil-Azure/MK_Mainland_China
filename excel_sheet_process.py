import streamlit as st
import pandas as pd
import io

# Set the page configuration
st.set_page_config(page_title="Data Processing App", layout="wide")

st.title("Data Processing Application")
st.write("Please upload your PIM Master Data and SKU List to proceed.")

# Create two columns for the uploaders
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. PIM Master Data")
    pim_file = st.file_uploader("Upload PIM Master Data (Excel)", type=["xls", "xlsx"], key="pim_upload")

with col2:
    st.subheader("2. SKU List")
    sku_file = st.file_uploader("Upload SKU List (Excel)", type=["xls", "xlsx"], key="sku_upload")

# Check if both files have been uploaded
if pim_file is not None and sku_file is not None:
    try:
        # Read the uploaded Excel files
        pim_df = pd.read_excel(pim_file)
        sku_df = pd.read_excel(sku_file)

        st.success("Files successfully uploaded! Processing data...")
        st.divider()

        # =====================================================================
        # DATA PROCESSING LOGIC
        # =====================================================================

        # Merge the SKU List with the PIM Master Data based on the "SKU" column.
        merged_df = pd.merge(sku_df[['SKU']], pim_df, on="SKU", how="left")

        # --- FIX FOR UPC NUMBERS ---
        # Fill missing values with an empty string, convert to string, and remove trailing '.0'
        if "UPC" in merged_df.columns:
            merged_df["UPC"] = merged_df["UPC"].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)

        num_rows = len(merged_df)

        # ---------------------------------------------------------------------
        # 1. Item Master Sheet
        # ---------------------------------------------------------------------
        df_item_master = pd.DataFrame(index=range(num_rows), columns=[
            "Purpose Code", "Total Number of line items", "Assigned Identification",
            "Style Number", "UPC number", "Color Code", "Size code",
            "Style Description/Platform Name", "Style Color", "Style Size",
            "Product Description Code"
        ])

        # Static Values
        df_item_master["Purpose Code"] = 2
        df_item_master["Total Number of line items"] = 1
        df_item_master["Assigned Identification"] = 1

        # Dynamic Values mapped from merged data
        df_item_master["Style Number"] = merged_df["SKU"]
        df_item_master["UPC number"] = merged_df["UPC"]
        df_item_master["Color Code"] = merged_df["SAP Color"]
        df_item_master["Size code"] = merged_df["Case Size"]
        df_item_master["Style Description/Platform Name"] = merged_df["Platform"]
        df_item_master["Style Color"] = merged_df["Case Color"]
        df_item_master["Style Size"] = merged_df["Case Size"]

        # ---------------------------------------------------------------------
        # 2. Cost- ML Sheet
        # ---------------------------------------------------------------------
        df_cost_ml = pd.DataFrame(index=range(num_rows), columns=[
            "Relation", "AccountCode", "AccountRelation", "ItemCode",
            "UPC", "SKU", "From date", "To date", "Unit", "Amount", "Currency"
        ])

        # Static Values
        df_cost_ml["Relation"] = "000::Price (Purch)"
        df_cost_ml["AccountCode"] = "002::All"
        df_cost_ml["ItemCode"] = "000::Table"
        df_cost_ml["Unit"] = "Pcs"
        df_cost_ml["Currency"] = "CNY"

        # Dynamic Values mapped from merged data
        df_cost_ml["SKU"] = merged_df["SKU"]
        df_cost_ml["UPC"] = merged_df["UPC"]

        # ---------------------------------------------------------------------
        # 3. PO ML-FP Sheet
        # ---------------------------------------------------------------------
        df_po_ml = pd.DataFrame(index=range(num_rows), columns=[
            "vendor code", "UPC", "Style", "Color", "Size"
        ])

        # Static Values
        df_po_ml["vendor code"] = "V21100001"

        # Dynamic Values mapped from merged data
        df_po_ml["Style"] = merged_df["SKU"]
        df_po_ml["Color"] = merged_df["Case Color"]
        df_po_ml["Size"] = merged_df["Case Size"]
        # Assuming you also want the cleaned UPC here based on the column name
        df_po_ml["UPC"] = merged_df["UPC"]

        # ---------------------------------------------------------------------
        # 4. Price-ML Sheet
        # ---------------------------------------------------------------------
        df_price_ml = pd.DataFrame(index=range(num_rows), columns=[
            "Relation", "AccountCode", "AccountRelation", "ItemCode",
            "UPC", "SKU", "From date", "To date", "Unit", "Amount", "Currency"
        ])

        # Static Values
        df_price_ml["Relation"] = "004::Price (sales)"
        df_price_ml["AccountCode"] = "002::All"
        df_price_ml["ItemCode"] = "000::Table"
        df_price_ml["Unit"] = "Pcs"
        df_price_ml["Currency"] = "CNY"

        # Dynamic Values mapped from merged data
        df_price_ml["SKU"] = merged_df["SKU"]
        df_price_ml["UPC"] = merged_df["UPC"]


        # =====================================================================

        # Function to generate an Excel file in memory
        def convert_dfs_to_excel(df_dict):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name, df in df_dict.items():
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
            processed_data = output.getvalue()
            return processed_data


        # Dictionary mapping the sheet names to their respective DataFrames
        sheets_to_export = {
            "Item Master": df_item_master,
            "Cost- ML": df_cost_ml,
            "PO ML-FP": df_po_ml,
            "Price-ML": df_price_ml
        }

        # Generate the final Excel file bytes
        excel_data = convert_dfs_to_excel(sheets_to_export)

        # Display the Download section
        st.subheader("Download Output")
        st.write("Your processed Excel file is ready. Click the button below to download it.")

        st.download_button(
            label="📥 Download Processed Excel File",
            data=excel_data,
            file_name="MK_Mainline_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except KeyError as e:
        st.error(
            f"Mapping Error: Could not find the expected column in your uploaded files. Please ensure the column {e} exists and is spelled exactly as expected.")
    except Exception as e:
        st.error(f"An error occurred: {e}")

elif pim_file is None or sku_file is None:
    st.info("Waiting for both files to be uploaded...")