import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
import tempfile
import os
from shipping_processor import process_shipping_list

st.set_page_config(
    page_title="Shipping List Processor",
    page_icon="🚢",
    layout="wide"
)

st.title("🚢 Shipping List Processor")
st.markdown("""
This application processes shipping lists to generate export and re-import receipts according to specified business rules.
Upload your Excel files below to get started.
""")

# File upload section
st.header("📁 Upload Files")
col1, col2 = st.columns(2)

with col1:
    shipping_list_file = st.file_uploader("Upload Shipping List Excel File", type=['xlsx'])
    policy_file = st.file_uploader("Upload Policy Excel File", type=['xlsx'])

with col2:
    shipping_rate_file = st.file_uploader("Upload Shipping Rate Excel File", type=['xlsx'])
    exchange_rate_file = st.file_uploader("Upload Exchange Rate Excel File", type=['xlsx'])

# Process files when all are uploaded
if all([shipping_list_file, policy_file, shipping_rate_file, exchange_rate_file]):
    st.success("All files uploaded successfully! Ready to process.")
    
    if st.button("Process Files", type="primary"):
        with st.spinner("Processing files..."):
            try:
                # Create a temporary directory to store uploaded files
                with tempfile.TemporaryDirectory() as temp_dir:
                    # Save uploaded files to temporary directory
                    shipping_list_path = os.path.join(temp_dir, "shipping_list.xlsx")
                    policy_path = os.path.join(temp_dir, "policy.xlsx")
                    shipping_rate_path = os.path.join(temp_dir, "shipping_rate.xlsx")
                    exchange_rate_path = os.path.join(temp_dir, "exchange_rate.xlsx")
                    
                    # Define output paths
                    output_fob_path = os.path.join(temp_dir, "fob_prices.xlsx")
                    output_export_path = os.path.join(temp_dir, "export_receipt.xlsx")
                    output_reimport_path = os.path.join(temp_dir, "reimport_receipt.xlsx")
                    
                    # Save uploaded files
                    with open(shipping_list_path, "wb") as f:
                        f.write(shipping_list_file.getvalue())
                    with open(policy_path, "wb") as f:
                        f.write(policy_file.getvalue())
                    with open(shipping_rate_path, "wb") as f:
                        f.write(shipping_rate_file.getvalue())
                    with open(exchange_rate_path, "wb") as f:
                        f.write(exchange_rate_file.getvalue())
                    
                    # Process the files using the command line processor
                    success = process_shipping_list(
                        shipping_list_path,
                        policy_path,
                        shipping_rate_path,
                        exchange_rate_path,
                        output_fob_path,
                        output_export_path,
                        output_reimport_path
                    )
                    
                    if success:
                        st.success("Files processed successfully!")
                        
                        # Display results
                        st.header("📊 Results")
                        
                        # FOB Prices
                        st.subheader("FOB Prices")
                        fob_df = pd.read_excel(output_fob_path)
                        st.dataframe(fob_df)
                        
                        # Export Receipt
                        st.subheader("Export Receipt")
                        export_df = pd.read_excel(output_export_path)
                        st.dataframe(export_df)
                        
                        # Re-import Receipt
                        st.subheader("Re-import Receipt")
                        reimport_df = pd.read_excel(output_reimport_path)
                        st.dataframe(reimport_df)
                        
                        # Download section
                        st.header("📥 Download Results")
                        
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            with open(output_fob_path, 'rb') as f:
                                st.download_button(
                                    "Download FOB Prices",
                                    f,
                                    file_name="fob_prices.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        
                        with col2:
                            with open(output_export_path, 'rb') as f:
                                st.download_button(
                                    "Download Export Receipt",
                                    f,
                                    file_name="export_receipt.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        
                        with col3:
                            with open(output_reimport_path, 'rb') as f:
                                st.download_button(
                                    "Download Re-import Receipt",
                                    f,
                                    file_name="reimport_receipt.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                    else:
                        st.error("Failed to process files. Please check the console for error messages.")
                        
            except Exception as e:
                st.error(f"An error occurred while processing the files: {str(e)}")
else:
    st.info("Please upload all required files to begin processing.")

# Add sidebar with information
with st.sidebar:
    st.header("ℹ️ About")
    st.markdown("""
    This application helps process shipping lists by:
    - De-duplicating shipping list items
    - Calculating CIF prices
    - Generating export receipts
    - Generating re-import receipts
    
    For more information, please refer to the documentation.
    """) 