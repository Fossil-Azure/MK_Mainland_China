import streamlit as st
import pandas as pd
import io
import requests
import concurrent.futures
import re
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, Alignment, Border, Side

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
        pim_df = pd.read_excel(pim_file)
        sku_df = pd.read_excel(sku_file)

        st.divider()

        # =====================================================================
        # DATA PROCESSING LOGIC
        # =====================================================================
        merged_df = pd.merge(sku_df[['SKU']], pim_df, on="SKU", how="left")
        num_rows = len(merged_df)


        # --- SAFE GET HELPER (Fixes the Pandas 'dtype str' bug) ---
        def safe_get(col_name):
            if col_name in merged_df.columns:
                return merged_df[col_name].fillna("").astype(str)
            return pd.Series([""] * num_rows)


        # --- FIX FOR UPC NUMBERS ---
        if "UPC" in merged_df.columns:
            merged_df["UPC"] = safe_get("UPC").str.replace(r'\.0$', '', regex=True)

        columns = [
            "SEASON", "Season\n(Remark Carryover)", "Picture", "Category",
            "Style Color", "Group ", "Subgroup", "UPC", "Style Code",
            "Style Description", "Color Code", "Color Name", "Size",
            "Material Description", "", "USD", "RMB", "HKD", "TWD", "MOP",
            "GC TTL Units", "Mainland TTL Units", "HK", "TW", "MC", "Buffer",
            "Delivery", "RMB-Incl.VAT", "HKD.1", "TWD.1", "MOP.1", "Size.1"
        ]

        # Initialize Dataframe and fill any default NaNs with empty strings
        df_out = pd.DataFrame(index=range(num_rows), columns=columns).fillna("")

        # Static Values
        df_out["SEASON"] = "F26"
        df_out["Season\n(Remark Carryover)"] = "F26"

        # Universal Dynamic Values
        df_out["Style Color"] = safe_get("SKU")
        df_out["Style Code"] = safe_get("SKU")
        df_out["UPC"] = safe_get("UPC")
        df_out["Category"] = safe_get("Product Type")
        df_out["Picture"] = safe_get("Main Image")

        # =====================================================================
        # --- CONDITIONAL MAPPINGS (WATCHES VS JEWELRY) ---
        # =====================================================================
        is_jewelry = safe_get("Product Type").str.upper().str.contains("JEWELRY", na=False)

        # 1. GROUP COLUMN
        df_out["Group "] = safe_get("Platform")
        if "Group" in merged_df.columns:
            df_out.loc[is_jewelry, "Group "] = merged_df.loc[is_jewelry, "Group"].fillna("").astype(str)

        # 2. SUBGROUP COLUMN
        df_out["Subgroup"] = safe_get("Platform")
        if "Primary Color Jewelry" in merged_df.columns:
            df_out.loc[is_jewelry, "Subgroup"] = merged_df.loc[is_jewelry, "Primary Color Jewelry"].fillna("").astype(
                str)

        # 3. STYLE DESCRIPTION COLUMN
        platform_str = safe_get("Platform")
        size_str = safe_get("Case Size").str.replace(r'\.0$', '', regex=True)
        df_out["Style Description"] = (platform_str + " " + size_str).str.strip()

        if "Product Name" in merged_df.columns:
            def clean_jewelry_description(row):
                p_name = str(row.get("Product Name", "")) if pd.notna(row.get("Product Name")) else ""
                brand = str(row.get("Brand", "")) if pd.notna(row.get("Brand")) else ""
                if p_name and brand:
                    cleaned = re.sub(re.escape(brand), '', p_name, flags=re.IGNORECASE)
                    return cleaned.strip()
                return p_name.strip()


            merged_df["Cleaned Product Name"] = merged_df.apply(clean_jewelry_description, axis=1)
            df_out.loc[is_jewelry, "Style Description"] = merged_df.loc[is_jewelry, "Cleaned Product Name"].fillna(
                "").astype(str)
        else:
            df_out.loc[is_jewelry, "Style Description"] = ""

        # 4. COLOR CODE COLUMN
        df_out["Color Code"] = "-"
        if "SAP Color" in merged_df.columns:
            df_out.loc[is_jewelry, "Color Code"] = merged_df.loc[is_jewelry, "SAP Color"].fillna("").astype(str)

        # 5. COLOR NAME COLUMN
        df_out["Color Name"] = safe_get("Ecomm Top Ring Color")
        if "Silhouette Jewelry" in merged_df.columns:
            df_out.loc[is_jewelry, "Color Name"] = merged_df.loc[is_jewelry, "Silhouette Jewelry"].fillna("").astype(
                str)

        # 6. SIZE COLUMN
        df_out["Size"] = safe_get("Case Size")
        df_out["Size.1"] = safe_get("Case Size")
        df_out.loc[is_jewelry, "Size"] = "-"

        # 7. MATERIAL DESCRIPTION COLUMN
        df_out["Material Description"] = safe_get("Ecomm Case Material")
        if "Primary Material Jewelry" in merged_df.columns:
            df_out.loc[is_jewelry, "Material Description"] = merged_df.loc[
                is_jewelry, "Primary Material Jewelry"].fillna("").astype(str)

        # Force the entire Dataframe to pure strings to ensure Openpyxl doesn't crash on hidden floats
        df_out = df_out.fillna("").astype(str)


        # =====================================================================
        # HIGH-SPEED IMAGE DOWNLOADER
        # =====================================================================
        def fetch_image(url):
            try:
                response = requests.get(str(url).strip(), timeout=5)
                if response.status_code == 200:
                    return url, response.content
            except:
                pass
            return url, None


        def generate_assortment_excel(df):
            unique_urls = [url for url in df["Picture"].unique() if url and str(url).strip().startswith("http")]
            image_cache = {}

            if unique_urls:
                st.info("Downloading images...")
                progress_bar = st.progress(0)
                with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
                    futures = {executor.submit(fetch_image, url): url for url in unique_urls}
                    for i, future in enumerate(concurrent.futures.as_completed(futures)):
                        url = futures[future]
                        image_cache[url] = future.result()[1]
                        progress_bar.progress((i + 1) / len(unique_urls))
                progress_bar.empty()

            st.info("Constructing and formatting the Excel file...")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, startrow=5, sheet_name="F26 LICENSEE")
                worksheet = writer.sheets["F26 LICENSEE"]

                # =============================================================
                # STYLES DEFINITION
                # =============================================================
                header_font = Font(bold=True, color="000000")
                header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

                data_center = Alignment(horizontal='center', vertical='center', wrap_text=False)
                data_left = Alignment(horizontal='left', vertical='center', wrap_text=False)

                thin_border = Border(
                    left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                    top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
                )

                worksheet.cell(row=2, column=16, value="PRICE")
                worksheet.cell(row=2, column=21, value="ORDER UNITS")
                worksheet.cell(row=2, column=27, value="DELIVERY")
                worksheet.cell(row=2, column=28, value="COST")
                worksheet.cell(row=3, column=1, value="F26")

                worksheet.cell(row=5, column=21, value=2400)
                worksheet.cell(row=5, column=22, value=1545)
                worksheet.cell(row=5, column=23, value=164)
                worksheet.cell(row=5, column=24, value=549)
                worksheet.cell(row=5, column=25, value=142)

                # Set general column widths
                for col_letter in [chr(i) for i in range(65, 91)] + ['AA', 'AB', 'AC', 'AD', 'AE', 'AF']:
                    worksheet.column_dimensions[col_letter].width = 15

                worksheet.column_dimensions['C'].width = 10.5
                worksheet.column_dimensions['O'].width = 3

                # Format Headers (Row 6)
                for col_num in range(1, 33):
                    cell = worksheet.cell(row=6, column=col_num)
                    cell.font = header_font
                    cell.alignment = header_align
                    cell.border = thin_border

                # Format Top Section Headers (Row 2 & 5)
                for col_num in [16, 21, 27, 28]:
                    cell = worksheet.cell(row=2, column=col_num)
                    cell.font = header_font
                    cell.alignment = header_align
                    cell.border = thin_border

                for col_num in range(21, 26):
                    cell = worksheet.cell(row=5, column=col_num)
                    cell.font = header_font
                    cell.alignment = header_align
                    cell.border = thin_border

                # Fill Data Rows & Apply Formatting
                for idx, url in enumerate(df["Picture"]):
                    row_number = idx + 7

                    worksheet.row_dimensions[row_number].height = 60

                    for col_num in range(1, 33):
                        cell = worksheet.cell(row=row_number, column=col_num)
                        cell.alignment = data_center if col_num not in [10, 14] else data_left
                        cell.border = thin_border

                    cell_ref = f"C{row_number}"
                    worksheet[cell_ref].value = ""

                    img_bytes = image_cache.get(url)
                    if img_bytes:
                        try:
                            image_stream = io.BytesIO(img_bytes)
                            img = OpenpyxlImage(image_stream)

                            img.width = 65
                            img.height = 65

                            worksheet.add_image(img, cell_ref)
                        except Exception:
                            worksheet[cell_ref].value = "Img Format Error"
                    elif pd.notna(url) and str(url).strip().startswith("http"):
                        worksheet[cell_ref].value = "Link Error"

            return output.getvalue()


        excel_data = generate_assortment_excel(df_out)

        st.success("File successfully generated! Formatting and images are applied.")

        st.subheader("Download Output")
        st.download_button(
            label="📥 Download Formatted Excel File",
            data=excel_data,
            file_name="Assortment_Formatted_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except KeyError as e:
        st.error(f"Mapping Error: Could not find column {e}.")
    except Exception as e:
        st.error(f"An error occurred: {e}")

elif pim_file is None or sku_file is None:
    st.info("Waiting for both files to be uploaded...")