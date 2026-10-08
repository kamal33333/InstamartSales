import streamlit as st
import pandas as pd
import io

def process_instamart_data():
    st.set_page_config(page_title="Instamart Data Processor", layout="wide")
    st.title("🛒 Instamart Sales Data Processor")
    st.write("Upload your raw Instamart sales file to process it and calculate total packs and SP totals.")

    # Upload raw sales file
    uploaded_file = st.file_uploader("Upload Sales File (Excel format)", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            with st.spinner("Loading and processing data..."):
                # 1. Load sales data
                sales_df = pd.read_excel(uploaded_file)
                
                # 2. Hardcoded Master Data (Extracted from instamartsku.xlsx)
                # This entirely replaces the need for an external master file
                master_data = [
                    {"Item Code": "135218", "SP": 179.00, "Retail Code": "NONW-004", "No. Of Packs": 1},
                    {"Item Code": "206297", "SP": 135.00, "Retail Code": "NAP-006", "No. Of Packs": 2},
                    {"Item Code": "262674", "SP": 150.00, "Retail Code": "NAP-009", "No. Of Packs": 2},
                    {"Item Code": "367043", "SP": 130.00, "Retail Code": "KT-002", "No. Of Packs": 1},
                    {"Item Code": "432851", "SP": 150.00, "Retail Code": "BT-017", "No. Of Packs": 1},
                    {"Item Code": "509253", "SP": 250.00, "Retail Code": "BT-019", "No. Of Packs": 1},
                    {"Item Code": "713515", "SP": 249.00, "Retail Code": "FT-002", "No. Of Packs": 1},
                    {"Item Code": "758911", "SP": 230.00, "Retail Code": "NAP-024", "No. Of Packs": 2},
                    {"Item Code": "285835", "SP": 325.00, "Retail Code": "K-ONEGSM70", "No. Of Packs": 1},
                    {"Item Code": "427622", "SP": 344.00, "Retail Code": "K-ONEGSM75", "No. Of Packs": 1},
                    {"Item Code": "432124", "SP": 200.00, "Retail Code": "NAP-028", "No. Of Packs": 4},
                    {"Item Code": "443486", "SP": 67.50,  "Retail Code": "NAP-029", "No. Of Packs": 1},
                    {"Item Code": "603435", "SP": 165.00, "Retail Code": "NAP-009", "No. Of Packs": 3},
                    {"Item Code": "781783", "SP": 312.00, "Retail Code": "BT-043", "No. Of Packs": 1},
                    {"Item Code": "937132", "SP": 152.00, "Retail Code": "FT-003", "No. Of Packs": 2}
                ]
                
                master_subset = pd.DataFrame(master_data)

                # 3. Prepare data for merging
                if 'ITEM_CODE' not in sales_df.columns:
                    st.error("Error: The uploaded file does not contain an 'ITEM_CODE' column.")
                    return

                sales_df['ITEM_CODE'] = sales_df['ITEM_CODE'].astype(str)
                master_subset['Item Code'] = master_subset['Item Code'].astype(str)

                # 4. Merge data based on Item Code
                merged_df = sales_df.merge(master_subset, left_on='ITEM_CODE', right_on='Item Code', how='left')

                # 5. Calculations
                if 'UNITS_SOLD' in merged_df.columns:
                    # Calculate 'SP total'
                    merged_df['SP total'] = merged_df['SP'] * merged_df['UNITS_SOLD']
                    # Calculate 'total packs' (Units Sold * Pack Size)
                    merged_df['total packs'] = merged_df['UNITS_SOLD'] * merged_df['No. Of Packs']
                else:
                    st.warning("Warning: 'UNITS_SOLD' column is missing. Skipping SP total and total packs calculations.")

                # 6. Formatting the output
                if 'Unnamed: 0' in merged_df.columns:
                    merged_df.rename(columns={'Unnamed: 0': 'sr'}, inplace=True)

                columns_to_keep = [
                    'sr', 'BRAND', 'ORDERED_DATE', 'CITY', 'AREA_NAME', 'STORE_ID', 
                    'L1_CATEGORY', 'L2_CATEGORY', 'L3_CATEGORY', 'PRODUCT_NAME', 
                    'VARIANT', 'ITEM_CODE', 'COMBO', 'COMBO_ITEM_CODE', 'COMBO_UNITS_SOLD', 
                    'BASE_MRP', 'UNITS_SOLD', 'GMV', 'SP', 'SP total', 'Retail Code', 
                    'No. Of Packs', 'total packs'
                ]

                # Filter to only keep requested columns that exist in the merged data
                final_cols = [col for col in columns_to_keep if col in merged_df.columns]
                output_df = merged_df[final_cols]

                # Show success message and preview
                st.success(f"Success! {len(output_df)} rows processed.")
                st.dataframe(output_df.head(10))

                # 7. Convert DataFrame to in-memory Excel file for downloading
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    output_df.to_excel(writer, index=False)
                
                processed_data = output.getvalue()

                # 8. Download Button
                st.download_button(
                    label="📥 Download Processed File",
                    data=processed_data,
                    file_name="instamartOctober_processed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")

if __name__ == "__main__":
    process_instamart_data()
