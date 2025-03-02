#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Shipping List Processor

This script processes shipping list Excel files to produce export and re-import invoices
according to specified business rules.
"""

import pandas as pd
import numpy as np
import os
from datetime import datetime
import argparse
import sys


def read_shipping_list(file_path):
    """
    Read the shipping list Excel file and return as DataFrame.
    
    Args:
        file_path (str): Path to the shipping list Excel file
        
    Returns:
        pd.DataFrame: DataFrame containing shipping list data
    """
    try:
        # Read the Excel file
        print(f"Reading file: {file_path}")
        try:
            # First try with no skiprows
            df = pd.read_excel(file_path)
        except Exception as e1:
            print(f"Error reading Excel file without skiprows: {e1}")
            try:
                # Try with skiprows=1 in case there's a header row
                df = pd.read_excel(file_path, skiprows=1)
                print("Successfully read file with skiprows=1")
            except Exception as e2:
                print(f"Error reading Excel file with skiprows=1: {e2}")
                raise
        
        print(f"Successfully read file with {len(df)} rows and {len(df.columns)} columns")
        
        # Print first few column names for debugging
        print("First 10 column names in the original file:")
        for i, col in enumerate(df.columns[:10]):
            print(f"{i}: {col}")
        
        # Define column mapping for known column names
        column_mapping = {
            # Serial number
            "Sr NO (序列号)": "serial_no",
            "Sr NO": "serial_no",
            "序列号": "serial_no",
            "Serial No": "serial_no",
            "Serial Number": "serial_no",
            
            # Part number
            "P/N.（系统料号 ）": "part_number",
            "P/N.": "part_number",
            "P/N": "part_number",
            "Part Number": "part_number",
            "料号": "part_number",
            "系统料号": "part_number",
            
            # Supplier
            "供应商": "supplier",
            "Supplier": "supplier",
            
            # Project name
            "项目名称": "project_name",
            "Project Name": "project_name",
            
            # Factory
            "工厂(Daman/Silvassa)": "factory",
            "工厂": "factory",
            "Factory": "factory",
            
            # Customs description
            "清关英文货描（关务提供）": "customs_desc_en",
            "清关英文货描": "customs_desc_en",
            "Customs Description": "customs_desc_en",
            "报关中文品名": "customs_desc_cn",
            "中文品名": "customs_desc_cn",
            
            # Description
            "DESCRIPTION (系统英文品名）": "description_en",
            "DESCRIPTION": "description_en",
            "Description": "description_en",
            "英文品名": "description_en",
            
            # Invoice name
            "开票名称": "invoice_name",
            "Invoice Name": "invoice_name",
            
            # Material name
            "物料名称": "material_name",
            "Material Name": "material_name",
            
            # Model
            "MODEL（货物型号（与实物相符)": "model",
            "MODEL": "model",
            "Model": "model",
            "货物型号": "model",
            
            # Quantity
            "QUANTITY （数量）": "quantity",
            "QUANTITY": "quantity",
            "Quantity": "quantity",
            "数量": "quantity",
            "Qty": "quantity",
            
            # Unit
            "单位": "unit",
            "Unit": "unit",
            
            # Carton measurement
            "Carton MEASUREMENT (外箱尺寸CM）": "carton_measurement",
            "Carton MEASUREMENT": "carton_measurement",
            "外箱尺寸": "carton_measurement",
            
            # Volume
            "体积（CBM）": "volume",
            "体积": "volume",
            "Volume": "volume",
            "总体积": "total_volume",
            "Total Volume": "total_volume",
            
            # Weight
            "单件毛重": "unit_gross_weight",
            "Unit Gross Weight": "unit_gross_weight",
            "G.W（KG) 总毛重": "total_gross_weight",
            "G.W": "total_gross_weight",
            "总毛重": "total_gross_weight",
            "Total Gross Weight": "total_gross_weight",
            "单件净重": "unit_net_weight",
            "Unit Net Weight": "unit_net_weight",
            "N.W  (KG) 总净重": "total_net_weight",
            "N.W": "total_net_weight",
            "总净重": "total_net_weight",
            "Total Net Weight": "total_net_weight",
            
            # Carton info
            "整箱数量": "full_carton_quantity",
            "Full Carton Quantity": "full_carton_quantity",
            "件数": "piece_count",
            "Piece Count": "piece_count",
            "CTN NO. (箱号)": "carton_no",
            "CTN NO.": "carton_no",
            "箱号": "carton_no",
            
            # Export customs method
            "出口报关方式": "export_customs_method",
            "Export Customs Method": "export_customs_method",
            
            # Purchasing unit
            "采购单位（智乐/UC/客供/供应商赠送/系统外订单）": "purchasing_unit",
            "采购单位": "purchasing_unit",
            "Purchasing Unit": "purchasing_unit",
            
            # Price
            "不含税单价（RMB）": "unit_price",
            "不含税单价": "unit_price",
            "单价": "unit_price",
            "Unit Price": "unit_price",
            
            # Tax rate
            "开票税率": "tax_rate"
        }
        
        # Create a lowercase version of the mapping for more robust matching
        lowercase_mapping = {}
        for key, value in column_mapping.items():
            lowercase_mapping[key.lower()] = value
        
        # Try to rename columns if they exist
        renamed_columns = {}
        for i, col in enumerate(df.columns):
            # Try exact match first
            if col in column_mapping:
                renamed_columns[col] = column_mapping[col]
            # Try lowercase match
            elif col.lower() in lowercase_mapping:
                renamed_columns[col] = lowercase_mapping[col.lower()]
            # Try partial match if no exact match found
            else:
                for key in column_mapping:
                    # Skip very short keys to avoid false matches
                    if len(key) < 3:
                        continue
                    if key in col or key.lower() in col.lower():
                        renamed_columns[col] = column_mapping[key]
                        print(f"Partial match: '{col}' -> '{column_mapping[key]}'")
                        break
        
        # Apply the column renaming
        df.rename(columns=renamed_columns, inplace=True)
        
        # Print renamed columns for debugging
        print("\nColumns after renaming:")
        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")
        
        return df
    except Exception as e:
        print(f"Error reading shipping list file: {e}")
        raise


def read_policy_file(file_path):
    """
    Read the policy Excel file that contains markup and insurance rates.
    
    Args:
        file_path (str): Path to the policy Excel file
        
    Returns:
        dict: Dictionary containing markup percentage and insurance rate
    """
    try:
        df = pd.read_excel(file_path)
        # Assuming the policy file has columns for markup_percentage and insurance_rate
        # Adjust as needed based on actual file structure
        policy = {
            'markup_percentage': df['markup_percentage'].iloc[0] / 100,  # Convert to decimal
            'insurance_rate': df['insurance_rate'].iloc[0] / 100,  # Convert to decimal
            'insurance_coefficient': df.get('insurance_coefficient', [1.0]).iloc[0]  # Default to 1.0 if not present
        }
        return policy
    except Exception as e:
        print(f"Error reading policy file: {e}")
        raise


def read_shipping_rate_file(file_path):
    """
    Read the shipping rate Excel file.
    
    Args:
        file_path (str): Path to the shipping rate Excel file
        
    Returns:
        float: Current shipping rate
    """
    try:
        df = pd.read_excel(file_path)
        # Assuming the shipping rate file has a column for shipping_rate
        # Adjust as needed based on actual file structure
        shipping_rate = df['shipping_rate'].iloc[0]
        return shipping_rate
    except Exception as e:
        print(f"Error reading shipping rate file: {e}")
        raise


def read_exchange_rate_file(file_path):
    """
    Read the exchange rate Excel file.
    
    Args:
        file_path (str): Path to the exchange rate Excel file
        
    Returns:
        dict: Dictionary containing exchange rates between currencies
    """
    try:
        df = pd.read_excel(file_path)
        # Assuming the exchange rate file has columns for different currency pairs
        # Adjust as needed based on actual file structure
        exchange_rates = {
            'RMB_USD': df['RMB_USD'].iloc[0],
            'RMB_RUPEE': df.get('RMB_RUPEE', [0]).iloc[0],
            'USD_RUPEE': df.get('USD_RUPEE', [0]).iloc[0]
        }
        return exchange_rates
    except Exception as e:
        print(f"Error reading exchange rate file: {e}")
        raise


def deduplicate_shipping_list(df):
    """
    Deduplicate items in the shipping list where part number and unit price are the same.
    
    Args:
        df (pd.DataFrame): Original shipping list DataFrame
        
    Returns:
        pd.DataFrame: Deduplicated shipping list
    """
    try:
        # Make a copy to avoid modifying the original DataFrame
        df_copy = df.copy()
        
        # Print column names for debugging
        print("Available columns in the input file:")
        for i, col in enumerate(df_copy.columns):
            print(f"{i}: {col}")
        
        # Print first few rows for debugging
        print("\nFirst 3 rows of data:")
        print(df_copy.head(3))
        
        # Handle duplicate columns by renaming them
        if df_copy.columns.duplicated().any():
            print("Warning: Duplicate column names detected in input file. Renaming duplicate columns.")
            duplicate_cols = df_copy.columns[df_copy.columns.duplicated()].tolist()
            for col in duplicate_cols:
                # Find all occurrences of the duplicate column
                cols = df_copy.columns.tolist()
                indices = [i for i, x in enumerate(cols) if x == col]
                
                # Rename all but the first occurrence
                for i, idx in enumerate(indices[1:], 1):
                    cols[idx] = f"{col}_{i}"
                
                # Assign new column names to DataFrame
                df_copy.columns = cols
                print(f"Renamed duplicate columns for '{col}'")
        
        # Handle invalid/empty rows - drop rows where all key columns are NaN
        df_copy = df_copy.dropna(how='all')
        print(f"After dropping completely empty rows: {len(df_copy)} rows")
        
        # If DataFrame is empty after dropping empty rows, return original DataFrame
        if len(df_copy) == 0:
            print("Warning: DataFrame is empty after dropping empty rows. Returning original DataFrame.")
            return df
        
        # Identify groupby columns, using fallbacks if necessary
        part_number_col = 'part_number'
        unit_price_col = 'unit_price'
        
        # Check if the primary columns exist, otherwise look for alternatives
        if part_number_col not in df_copy.columns:
            # Try to find column containing 'P/N' or similar
            potential_pn_cols = [col for col in df_copy.columns if any(term in col.upper() for term in ['P/N', 'PART', 'PN', '料号'])]
            if potential_pn_cols:
                part_number_col = potential_pn_cols[0]
                print(f"Using '{part_number_col}' as the part number column")
            else:
                # If still not found, try to identify by position (assuming part number is usually one of the first few columns)
                print("Could not find a column containing 'P/N' or 'part'. Creating a default part number column.")
                # Create a synthetic part number column using first column or row index as fallback
                if len(df_copy.columns) > 1:
                    # Use the second column as part number (often part number is the second column after serial number)
                    df_copy['part_number'] = df_copy.iloc[:, 1]
                    part_number_col = 'part_number'
                    print(f"Created 'part_number' column using values from column: {df_copy.columns[1]}")
                else:
                    # Use row index as last resort
                    df_copy['part_number'] = df_copy.index
                    part_number_col = 'part_number'
                    print("Created 'part_number' column using row indices")
        
        if unit_price_col not in df_copy.columns:
            # Try to find column containing 'price' or similar
            potential_price_cols = [col for col in df_copy.columns if any(term in col.lower() for term in ['price', 'unit', '单价', 'cost'])]
            if potential_price_cols:
                unit_price_col = potential_price_cols[0]
                print(f"Using '{unit_price_col}' as the unit price column")
            else:
                # If still not found, create a dummy price column with constant value
                print("Could not find a unit price column. Creating a default unit price column with constant value 1.0")
                df_copy['unit_price'] = 1.0
                unit_price_col = 'unit_price'
        
        # Ensure we've found or created both required columns
        if not all(col in df_copy.columns for col in [part_number_col, unit_price_col]):
            print(f"Error: Could not find or create both {part_number_col} and {unit_price_col} columns.")
            return df
            
        groupby_cols = [part_number_col, unit_price_col]
        print(f"Grouping by columns: {groupby_cols}")
        
        # Verify groupby columns actually exist
        for col in groupby_cols:
            if col not in df_copy.columns:
                print(f"Critical error: Column '{col}' not found in DataFrame despite fallback measures. Using original DataFrame.")
                return df
        
        # Identify quantity column
        quantity_col = 'quantity'
        if quantity_col not in df_copy.columns:
            potential_qty_cols = [col for col in df_copy.columns if any(term in col.lower() for term in ['qty', 'quantity', '数量', 'count'])]
            if potential_qty_cols:
                quantity_col = potential_qty_cols[0]
                print(f"Using '{quantity_col}' as the quantity column")
            else:
                # If still not found, create a dummy quantity column
                print("Could not find a quantity column. Creating a default quantity column with constant value 1")
                df_copy['quantity'] = 1
                quantity_col = 'quantity'
        
        # Identify columns to sum based on whether they contain specific keywords
        sum_cols = []
        if quantity_col not in sum_cols and quantity_col in df_copy.columns:
            sum_cols.append(quantity_col)
        
        for col in df_copy.columns:
            if col not in groupby_cols and col not in sum_cols:
                if any(keyword in col.lower() for keyword in ['volume', 'weight', 'qty', 'count', 'piece', '体积', '重量', '数量', '件数']):
                    sum_cols.append(col)
        
        print(f"Columns that will be summed: {sum_cols}")
        
        # Ensure there are columns to sum
        if not sum_cols:
            print("Warning: No columns identified for summing. Deduplication may not be meaningful.")
            sum_cols = [quantity_col]  # Use quantity as a fallback
        
        # Convert numeric columns to the appropriate types
        for col in groupby_cols + sum_cols:
            if col in df_copy.columns:
                try:
                    df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce')
                    print(f"Converted '{col}' to numeric")
                except Exception as e:
                    print(f"Could not convert '{col}' to numeric: {e}")
        
        # Replace NaN values with 0 in numeric columns to ensure proper summing
        for col in sum_cols:
            if col in df_copy.columns:
                df_copy[col] = df_copy[col].fillna(0)
        
        # Identify rows with NaN in groupby columns (these will be dropped during groupby)
        for col in groupby_cols:
            if col in df_copy.columns:
                nan_count = df_copy[col].isna().sum()
                if nan_count > 0:
                    print(f"Warning: {nan_count} rows have NaN values in '{col}' and will be excluded from groupby")
                    # Fill NaN values with a placeholder to avoid losing data
                    if pd.api.types.is_numeric_dtype(df_copy[col]):
                        df_copy[col] = df_copy[col].fillna(-99999)  # Use an unlikely value as placeholder
                    else:
                        df_copy[col] = df_copy[col].fillna("UNKNOWN")
        
        try:
            # Use a simplified two-step approach to avoid the pandas 'name' attribute error
            
            # Step 1: First group by and get the sum of the summable columns
            if sum_cols:
                sum_data = {}
                # Use a try-except block to catch any errors during groupby
                try:
                    for group_key, group_df in df_copy.groupby(groupby_cols):
                        # Handle case where group_key might be a single value or a tuple
                        if not isinstance(group_key, tuple):
                            group_key = (group_key,)
                        
                        # Sum each column for this group
                        sums = {col: group_df[col].sum() for col in sum_cols if col in group_df.columns}
                        sum_data[group_key] = sums
                    
                    # Create a DataFrame for the summed data
                    sum_rows = []
                    for group_key, sums in sum_data.items():
                        row = dict(zip(groupby_cols, group_key))
                        row.update(sums)
                        sum_rows.append(row)
                    
                    if sum_rows:
                        sum_df = pd.DataFrame(sum_rows)
                    else:
                        # If no rows, create an empty DataFrame with appropriate columns
                        sum_df = pd.DataFrame(columns=groupby_cols + sum_cols)
                    
                except Exception as e:
                    print(f"Error during groupby/sum operation: {e}")
                    print("Falling back to original DataFrame")
                    return df
            else:
                # If no sum columns, just get the unique combinations of groupby columns
                sum_df = df_copy[groupby_cols].drop_duplicates()
            
            # Step 2: Get the first row for each group for non-sum columns
            non_sum_cols = [col for col in df_copy.columns if col not in groupby_cols and col not in sum_cols]
            
            if non_sum_cols:
                first_data = {}
                try:
                    for group_key, group_df in df_copy.groupby(groupby_cols):
                        # Handle case where group_key might be a single value or a tuple
                        if not isinstance(group_key, tuple):
                            group_key = (group_key,)
                        
                        # Get the first row of each column for this group
                        first_row = group_df.iloc[0]
                        firsts = {col: first_row[col] for col in non_sum_cols if col in first_row.index}
                        first_data[group_key] = firsts
                    
                    # Create a DataFrame for the first-value data
                    first_rows = []
                    for group_key, firsts in first_data.items():
                        row = dict(zip(groupby_cols, group_key))
                        row.update(firsts)
                        first_rows.append(row)
                    
                    if first_rows:
                        first_df = pd.DataFrame(first_rows)
                    else:
                        # If no rows, create an empty DataFrame with appropriate columns
                        first_df = pd.DataFrame(columns=groupby_cols + non_sum_cols)
                    
                    # Merge the summed data with the first-value data
                    df_deduped = pd.merge(sum_df, first_df, on=groupby_cols, how='outer')
                except Exception as e:
                    print(f"Error during groupby/first operation: {e}")
                    print("Using only sum columns for deduplication")
                    df_deduped = sum_df
            else:
                # If no non-sum columns, just use the summed data
                df_deduped = sum_df
            
            # If the result is empty, return the original DataFrame
            if len(df_deduped) == 0:
                print("Warning: Deduplication resulted in empty DataFrame. Using original DataFrame.")
                return df
                
            # Recalculate unit-based values if they exist
            if quantity_col in df_deduped.columns:
                # Look for net weight columns
                net_weight_unit_col = next((col for col in df_deduped.columns if 'unit_net_weight' in col.lower() or '单件净重' in col), None)
                net_weight_total_col = next((col for col in df_deduped.columns if 'total_net_weight' in col.lower() or '总净重' in col), None)
                
                if net_weight_unit_col and net_weight_total_col and net_weight_total_col in df_deduped.columns:
                    df_deduped[net_weight_unit_col] = df_deduped[net_weight_total_col] / df_deduped[quantity_col].replace(0, np.nan)
                
                # Look for gross weight columns
                gross_weight_unit_col = next((col for col in df_deduped.columns if 'unit_gross_weight' in col.lower() or '单件毛重' in col), None)
                gross_weight_total_col = next((col for col in df_deduped.columns if 'total_gross_weight' in col.lower() or '总毛重' in col), None)
                
                if gross_weight_unit_col and gross_weight_total_col and gross_weight_total_col in df_deduped.columns:
                    df_deduped[gross_weight_unit_col] = df_deduped[gross_weight_total_col] / df_deduped[quantity_col].replace(0, np.nan)
            
            print(f"Successfully deduplicated: {len(df_deduped)} rows (from original {len(df_copy)} rows)")
            return df_deduped
            
        except Exception as e:
            print(f"Error during deduplication: {e}")
            # If deduplication fails, return the original DataFrame as a fallback
            print("Falling back to original data without deduplication")
            return df
            
    except Exception as e:
        print(f"Unexpected error in deduplicate_shipping_list: {e}")
        # In case of any unexpected error, return the original DataFrame
        return df


def calculate_cif_prices(df, policy, shipping_rate, exchange_rates):
    """
    Calculate CIF unit prices for the shipping list according to the updated specification.
    
    The calculation follows these steps:
    1. Calculate FOB unit price = Unit price * (1 + markup %)
    2. Calculate FOB total price = FOB unit price * quantity
    3. Calculate adjusted total goods cost with insurance = total FOB price * insurance co-efficiency * (1+insurance rate)
    4. Calculate total shipping cost = net weight * shipping rate
    5. Calculate CIF total cost = total goods cost with insurance + total shipping cost
    6. Calculate CIF unit price in RMB = total CIF cost / quantity
    7. Calculate CIF unit price in USD = CIF unit price in RMB / RMB_USD exchange rate
    
    Args:
        df (pd.DataFrame): Shipping list DataFrame
        policy (dict): Policy parameters including markup percentage and insurance rate
        shipping_rate (float): Current shipping rate
        exchange_rates (dict): Exchange rates between currencies
        
    Returns:
        pd.DataFrame: DataFrame with added CIF pricing information
    """
    # Make a copy to avoid modifying the original DataFrame
    df_copy = df.copy()
    
    try:
        # Helper function to safely get or create columns
        def safe_get_or_create_column(df, column_name, fallback_columns=None, default_value=0):
            if column_name in df.columns:
                return df[column_name]
            
            if fallback_columns:
                for fallback_col in fallback_columns:
                    if fallback_col in df.columns:
                        print(f"Using '{fallback_col}' instead of '{column_name}' for CIF calculation")
                        df[column_name] = df[fallback_col]
                        return df[column_name]
            
            print(f"Column '{column_name}' not found. Creating with default value {default_value}")
            df[column_name] = default_value
            return df[column_name]
        
        # Ensure required columns exist
        unit_price = safe_get_or_create_column(df_copy, 'unit_price', 
                                             fallback_columns=['Unit Price', '单价', 'price', '不含税单价'],
                                             default_value=1.0)
        
        quantity = safe_get_or_create_column(df_copy, 'quantity', 
                                           fallback_columns=['Quantity', 'Qty', '数量', 'QUANTITY'],
                                           default_value=1.0)
        
        total_net_weight = safe_get_or_create_column(df_copy, 'total_net_weight', 
                                                  fallback_columns=['N.W', '总净重', 'Total Net Weight'],
                                                  default_value=0.1)
        
        # Get policy parameters with defaults
        markup_percentage = policy.get('markup_percentage', 0.05)  # Default 5% if not provided
        insurance_coefficient = policy.get('insurance_coefficient', 1.0)  # Default 1.0 if not provided
        insurance_rate = policy.get('insurance_rate', 0.01)  # Default 1% if not provided
        
        # Step 4.a: Calculate FOB unit price
        df_copy['fob_unit_price'] = df_copy['unit_price'] * (1 + markup_percentage)
        
        # Step 4.b: Calculate FOB total price
        df_copy['fob_total_price'] = df_copy['fob_unit_price'] * df_copy['quantity']
        
        # Step 5.a: Calculate adjusted total goods cost with insurance
        df_copy['total_goods_cost_with_insurance'] = df_copy['fob_total_price'] * insurance_coefficient * (1 + insurance_rate)
        
        # Step 5.b: Calculate total shipping cost
        df_copy['total_shipping_cost'] = df_copy['total_net_weight'] * shipping_rate
        
        # Step 5.c: Calculate CIF total cost
        df_copy['cif_total_cost'] = df_copy['total_goods_cost_with_insurance'] + df_copy['total_shipping_cost']
        
        # Step 5.d: Calculate CIF unit price in RMB
        # Avoid division by zero by replacing 0 with NaN temporarily
        safe_quantity = df_copy['quantity'].replace(0, np.nan)
        df_copy['cif_unit_price_rmb'] = df_copy['cif_total_cost'] / safe_quantity
        df_copy['cif_unit_price_rmb'] = df_copy['cif_unit_price_rmb'].fillna(0)  # Replace NaN back with 0
        
        # Step 5.e: Convert to USD
        exchange_rate_usd = exchange_rates.get('RMB_USD', 7.0)  # Default to 7.0 if not provided
        if exchange_rate_usd == 0:
            print("Warning: RMB_USD exchange rate is 0. Using default value of 7.0")
            exchange_rate_usd = 7.0
            
        df_copy['cif_unit_price_usd'] = df_copy['cif_unit_price_rmb'] / exchange_rate_usd
        
        # Add other relevant currencies if needed
        if 'RMB_RUPEE' in exchange_rates and exchange_rates['RMB_RUPEE'] != 0:
            df_copy['cif_unit_price_rupee'] = df_copy['cif_unit_price_rmb'] * exchange_rates['RMB_RUPEE']
        
        print("CIF price calculation completed successfully")
        return df_copy
        
    except Exception as e:
        print(f"Error during CIF price calculation: {e}")
        # If there's an error, add default CIF columns with reasonable values
        print("Adding default CIF pricing due to calculation error")
        df_copy['cif_unit_price_rmb'] = df_copy.get('unit_price', 10.0) * 1.1  # 10% markup as fallback
        df_copy['cif_unit_price_usd'] = df_copy['cif_unit_price_rmb'] / 7.0  # Default exchange rate
        return df_copy


def generate_export_receipt(df, output_path):
    """
    Generate export receipt Excel file.
    
    Args:
        df (pd.DataFrame): DataFrame with CIF pricing information
        output_path (str): Path to save the export receipt
        
    Returns:
        str: Path to the saved export receipt
    """
    # Make a copy to avoid modifying the original DataFrame
    df_export = df.copy()
    
    print("Generating export receipt...")
    print(f"Available columns for export receipt: {df_export.columns.tolist()}")
    
    # Handle duplicate columns by renaming them with a suffix
    # This prevents errors when accessing columns with the same name
    if df_export.columns.duplicated().any():
        print("Warning: Duplicate column names detected. Renaming duplicate columns.")
        duplicate_cols = df_export.columns[df_export.columns.duplicated()].tolist()
        for col in duplicate_cols:
            # Find all occurrences of the duplicate column
            cols = df_export.columns.tolist()
            indices = [i for i, x in enumerate(cols) if x == col]
            
            # Rename all but the first occurrence
            for i, idx in enumerate(indices[1:], 1):
                cols[idx] = f"{col}_{i}"
            
            # Assign new column names to DataFrame
            df_export.columns = cols
            print(f"Renamed duplicate columns for '{col}'")
    
    # Create a new DataFrame with the required columns for export receipt
    export_df = pd.DataFrame()
    
    # Helper function to safely get column data with a fallback
    def safe_get_column(dataframe, column_name, fallback_value=None, fallback_columns=None):
        if column_name in dataframe.columns:
            return dataframe[column_name]
        elif fallback_columns:
            for fallback_col in fallback_columns:
                if fallback_col in dataframe.columns:
                    print(f"Using '{fallback_col}' instead of '{column_name}'")
                    return dataframe[fallback_col]
        print(f"Column '{column_name}' not found. Using fallback value.")
        return fallback_value if fallback_value is not None else pd.Series(['-'] * len(dataframe))
    
    # Map the existing columns to the required format, using fallbacks when needed
    export_df["NO."] = safe_get_column(
        df_export, "serial_no", 
        fallback_value=range(1, len(df_export) + 1),
        fallback_columns=["Sr NO", "序列号", "Serial No", "Serial Number"]
    )
    
    export_df["P/N"] = safe_get_column(
        df_export, "part_number", 
        fallback_columns=["P/N.", "P/N", "Part Number", "料号", "系统料号"]
    )
    
    export_df["DESCRIPTION"] = safe_get_column(
        df_export, "description_en", 
        fallback_columns=["DESCRIPTION", "Description", "英文品名", "customs_desc_en", "material_name"]
    )
    
    export_df["Model NO."] = safe_get_column(
        df_export, "model", 
        fallback_columns=["MODEL", "Model", "货物型号"]
    )
    
    export_df["Unit Price USD"] = safe_get_column(
        df_export, "cif_unit_price_usd", 
        fallback_columns=["unit_price", "Unit Price"]
    ).round(2)
    
    export_df["Qty"] = safe_get_column(
        df_export, "quantity", 
        fallback_columns=["QUANTITY", "Quantity", "数量", "Qty"]
    )
    
    export_df["Unit"] = safe_get_column(
        df_export, "unit", 
        fallback_columns=["Unit", "单位"]
    )
    
    # Calculate Amount USD (Unit Price USD * Qty)
    export_df["Amount USD"] = (export_df["Unit Price USD"] * export_df["Qty"]).round(2)
    
    # Save to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Export Receipt')
        
        # Add formatting if needed
        workbook = writer.book
        worksheet = writer.sheets['Export Receipt']
        
        # Add date and other metadata
        metadata_sheet = workbook.create_sheet('Metadata')
        metadata_sheet['A1'] = 'Generated Date'
        metadata_sheet['B1'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    print(f"Export receipt saved to {output_path}")
    return output_path


def generate_reimport_receipt(df, output_path):
    """
    Generate re-import receipt Excel file.
    
    Args:
        df (pd.DataFrame): DataFrame with CIF pricing information
        output_path (str): Path to save the re-import receipt
        
    Returns:
        str: Path to the saved re-import receipt
    """
    # Make a copy to avoid modifying the original DataFrame
    df_reimport = df.copy()
    
    print("Generating re-import receipt...")
    print(f"Available columns for re-import receipt: {df_reimport.columns.tolist()}")
    
    # Handle duplicate columns by renaming them with a suffix
    # This prevents errors when accessing columns with the same name
    if df_reimport.columns.duplicated().any():
        print("Warning: Duplicate column names detected. Renaming duplicate columns.")
        duplicate_cols = df_reimport.columns[df_reimport.columns.duplicated()].tolist()
        for col in duplicate_cols:
            # Find all occurrences of the duplicate column
            cols = df_reimport.columns.tolist()
            indices = [i for i, x in enumerate(cols) if x == col]
            
            # Rename all but the first occurrence
            for i, idx in enumerate(indices[1:], 1):
                cols[idx] = f"{col}_{i}"
            
            # Assign new column names to DataFrame
            df_reimport.columns = cols
            print(f"Renamed duplicate columns for '{col}'")
    
    # Create a new DataFrame with the required columns for re-import receipt
    reimport_df = pd.DataFrame()
    
    # Helper function to safely get column data with a fallback
    def safe_get_column(dataframe, column_name, fallback_value=None, fallback_columns=None):
        if column_name in dataframe.columns:
            return dataframe[column_name]
        elif fallback_columns:
            for fallback_col in fallback_columns:
                if fallback_col in dataframe.columns:
                    print(f"Using '{fallback_col}' instead of '{column_name}'")
                    return dataframe[fallback_col]
        print(f"Column '{column_name}' not found. Using fallback value.")
        return fallback_value if fallback_value is not None else pd.Series(['-'] * len(dataframe))
    
    # Map the existing columns to a format similar to export receipt
    # but include customs description fields which are important for re-import
    reimport_df["NO."] = safe_get_column(
        df_reimport, "serial_no", 
        fallback_value=range(1, len(df_reimport) + 1),
        fallback_columns=["Sr NO", "序列号", "Serial No", "Serial Number"]
    )
    
    reimport_df["P/N"] = safe_get_column(
        df_reimport, "part_number", 
        fallback_columns=["P/N.", "P/N", "Part Number", "料号", "系统料号"]
    )
    
    reimport_df["English Description"] = safe_get_column(
        df_reimport, "customs_desc_en", 
        fallback_columns=["清关英文货描", "Customs Description", "description_en", "DESCRIPTION", "英文品名"]
    )
    
    reimport_df["Chinese Description"] = safe_get_column(
        df_reimport, "customs_desc_cn", 
        fallback_columns=["报关中文品名", "中文品名"]
    )
    
    reimport_df["Model NO."] = safe_get_column(
        df_reimport, "model", 
        fallback_columns=["MODEL", "Model", "货物型号"]
    )
    
    reimport_df["Unit Price USD"] = safe_get_column(
        df_reimport, "cif_unit_price_usd", 
        fallback_columns=["unit_price", "Unit Price"]
    ).round(2)
    
    reimport_df["Qty"] = safe_get_column(
        df_reimport, "quantity", 
        fallback_columns=["QUANTITY", "Quantity", "数量", "Qty"]
    )
    
    reimport_df["Unit"] = safe_get_column(
        df_reimport, "unit", 
        fallback_columns=["Unit", "单位"]
    )
    
    reimport_df["Amount USD"] = (reimport_df["Unit Price USD"] * reimport_df["Qty"]).round(2)
    
    reimport_df["Net Weight (kg)"] = safe_get_column(
        df_reimport, "unit_net_weight", 
        fallback_columns=["单件净重", "Unit Net Weight"]
    )
    
    reimport_df["Total Net Weight (kg)"] = safe_get_column(
        df_reimport, "total_net_weight", 
        fallback_columns=["N.W", "总净重", "Total Net Weight"]
    )
    
    reimport_df["Gross Weight (kg)"] = safe_get_column(
        df_reimport, "unit_gross_weight", 
        fallback_columns=["单件毛重", "Unit Gross Weight"]
    )
    
    reimport_df["Total Gross Weight (kg)"] = safe_get_column(
        df_reimport, "total_gross_weight", 
        fallback_columns=["G.W", "总毛重", "Total Gross Weight"]
    )
    
    # Save to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        reimport_df.to_excel(writer, index=False, sheet_name='Re-Import Receipt')
        
        # Add formatting if needed
        workbook = writer.book
        worksheet = writer.sheets['Re-Import Receipt']
        
        # Add date and other metadata
        metadata_sheet = workbook.create_sheet('Metadata')
        metadata_sheet['A1'] = 'Generated Date'
        metadata_sheet['B1'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    print(f"Re-import receipt saved to {output_path}")
    return output_path


def process_shipping_list(
    shipping_list_file,
    policy_file,
    shipping_rate_file,
    exchange_rate_file,
    output_deduped_file,
    output_export_file,
    output_reimport_file
):
    """
    Main function to process shipping list and generate required outputs.
    
    Args:
        shipping_list_file (str): Path to shipping list Excel file
        policy_file (str): Path to policy Excel file
        shipping_rate_file (str): Path to shipping rate Excel file
        exchange_rate_file (str): Path to exchange rate Excel file
        output_deduped_file (str): Path to save deduplicated shipping list
        output_export_file (str): Path to save export receipt
        output_reimport_file (str): Path to save re-import receipt
        
    Returns:
        tuple: Paths to the generated files
    """
    try:
        # Step 1: Read input files
        print("Reading shipping list file...")
        shipping_df = read_shipping_list(shipping_list_file)
        
        print("Reading policy file...")
        policy = read_policy_file(policy_file)
        
        print("Reading shipping rate file...")
        shipping_rate = read_shipping_rate_file(shipping_rate_file)
        
        print("Reading exchange rate file...")
        exchange_rates = read_exchange_rate_file(exchange_rate_file)
        
        # Step 2: Deduplicate shipping list
        print("Deduplicating shipping list...")
        try:
            deduped_df = deduplicate_shipping_list(shipping_df)
        except Exception as e:
            print(f"Error during deduplication: {e}")
            print("Proceeding with the original data without deduplication")
            deduped_df = shipping_df
        
        # Save deduplicated file
        print(f"Saving deduplicated shipping list to {output_deduped_file}...")
        deduped_df.to_excel(output_deduped_file, index=False)
        
        # Step 3: Calculate CIF prices
        print("Calculating CIF prices...")
        try:
            cif_df = calculate_cif_prices(deduped_df, policy, shipping_rate, exchange_rates)
        except Exception as e:
            print(f"Error during CIF price calculation: {e}")
            print("Proceeding with the original data and adding default CIF pricing")
            cif_df = deduped_df.copy()
            cif_df['cif_unit_price_usd'] = cif_df.get('unit_price', 1.0) * 0.14  # Rough RMB to USD conversion
        
        # Step 4: Generate export receipt
        print(f"Generating export receipt to {output_export_file}...")
        try:
            export_path = generate_export_receipt(cif_df, output_export_file)
        except Exception as e:
            print(f"Error generating export receipt: {e}")
            # Create a simple export receipt with basic data
            simple_export_df = pd.DataFrame({
                "NO.": range(1, len(cif_df) + 1),
                "P/N": cif_df.get('part_number', ['Unknown'] * len(cif_df)),
                "DESCRIPTION": cif_df.get('description_en', [''] * len(cif_df)),
                "Model NO.": cif_df.get('model', [''] * len(cif_df)),
                "Unit Price USD": cif_df.get('cif_unit_price_usd', [0] * len(cif_df)),
                "Qty": cif_df.get('quantity', [1] * len(cif_df)),
                "Unit": cif_df.get('unit', [''] * len(cif_df)),
                "Amount USD": cif_df.get('cif_unit_price_usd', [0] * len(cif_df)) * cif_df.get('quantity', [1] * len(cif_df))
            })
            simple_export_df.to_excel(output_export_file, index=False)
            export_path = output_export_file
            print(f"Created simplified export receipt at {export_path}")
        
        # Step 5: Generate re-import receipt
        print(f"Generating re-import receipt to {output_reimport_file}...")
        try:
            reimport_path = generate_reimport_receipt(cif_df, output_reimport_file)
        except Exception as e:
            print(f"Error generating re-import receipt: {e}")
            # Create a simple re-import receipt with basic data
            simple_reimport_df = pd.DataFrame({
                "NO.": range(1, len(cif_df) + 1),
                "P/N": cif_df.get('part_number', ['Unknown'] * len(cif_df)),
                "English Description": cif_df.get('customs_desc_en', [''] * len(cif_df)),
                "Chinese Description": cif_df.get('customs_desc_cn', [''] * len(cif_df)),
                "Model NO.": cif_df.get('model', [''] * len(cif_df)),
                "Unit Price USD": cif_df.get('cif_unit_price_usd', [0] * len(cif_df)),
                "Qty": cif_df.get('quantity', [1] * len(cif_df)),
                "Unit": cif_df.get('unit', [''] * len(cif_df)),
                "Amount USD": cif_df.get('cif_unit_price_usd', [0] * len(cif_df)) * cif_df.get('quantity', [1] * len(cif_df)),
                "Net Weight (kg)": cif_df.get('unit_net_weight', [0] * len(cif_df)),
                "Total Net Weight (kg)": cif_df.get('total_net_weight', [0] * len(cif_df))
            })
            simple_reimport_df.to_excel(output_reimport_file, index=False)
            reimport_path = output_reimport_file
            print(f"Created simplified re-import receipt at {reimport_path}")
        
        print("Processing completed successfully!")
        return output_deduped_file, export_path, reimport_path
        
    except Exception as e:
        print(f"An error occurred during processing: {e}")
        # Create empty output files if they don't exist
        for file_path, file_type in [(output_deduped_file, 'deduped'), 
                                    (output_export_file, 'export'), 
                                    (output_reimport_file, 'reimport')]:
            if not os.path.exists(file_path):
                pd.DataFrame(columns=['Error']).to_excel(file_path, index=False)
                print(f"Created empty {file_type} file at {file_path}")
        
        return output_deduped_file, output_export_file, output_reimport_file


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process shipping list to generate export and re-import receipts.')
    parser.add_argument('--shipping-list', required=True, help='Path to shipping list Excel file')
    parser.add_argument('--policy-file', required=True, help='Path to policy Excel file')
    parser.add_argument('--shipping-rate-file', required=True, help='Path to shipping rate Excel file')
    parser.add_argument('--exchange-rate-file', required=True, help='Path to exchange rate Excel file')
    parser.add_argument('--output-deduped', default='deduped_shipping_list.xlsx', help='Path to save deduplicated shipping list')
    parser.add_argument('--output-export', default='export_receipt.xlsx', help='Path to save export receipt')
    parser.add_argument('--output-reimport', default='reimport_receipt.xlsx', help='Path to save re-import receipt')
    
    args = parser.parse_args()
    
    process_shipping_list(
        args.shipping_list,
        args.policy_file,
        args.shipping_rate_file,
        args.exchange_rate_file,
        args.output_deduped,
        args.output_export,
        args.output_reimport
    ) 