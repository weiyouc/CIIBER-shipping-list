#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Example script demonstrating how to use the shipping_processor.py
"""

import os
from shipping_processor import process_shipping_list

def main():
    """
    Example showing how to use the shipping processor with sample files.
    """
    # Define paths to sample input files
    # Assuming you have created or have these files available
    shipping_list_file = "testfiles/original-input-shippinglist.xlsx"
    
    # You would need to create these files based on your specific requirements
    policy_file = "sample_policy.xlsx"
    shipping_rate_file = "sample_shipping_rate.xlsx"
    exchange_rate_file = "sample_exchange_rate.xlsx"
    
    # Define output file paths
    output_deduped_file = "output_deduped_shipping_list.xlsx"
    output_export_file = "output_export_receipt.xlsx"
    output_reimport_file = "output_reimport_receipt.xlsx"
    
    # Check if input files exist
    if not os.path.exists(shipping_list_file):
        print(f"Error: Shipping list file not found at {shipping_list_file}")
        return
    
    # Create sample policy file if it doesn't exist
    if not os.path.exists(policy_file):
        print(f"Sample policy file doesn't exist. You would need to create it with the required format.")
        print(f"It should contain columns: 'markup_percentage', 'insurance_rate', and optionally 'insurance_coefficient'.")
        return
    
    # Create sample shipping rate file if it doesn't exist
    if not os.path.exists(shipping_rate_file):
        print(f"Sample shipping rate file doesn't exist. You would need to create it with the required format.")
        print(f"It should contain a column: 'shipping_rate'.")
        return
    
    # Create sample exchange rate file if it doesn't exist
    if not os.path.exists(exchange_rate_file):
        print(f"Sample exchange rate file doesn't exist. You would need to create it with the required format.")
        print(f"It should contain columns: 'RMB_USD', and optionally 'RMB_RUPEE' and 'USD_RUPEE'.")
        return
    
    # Process the shipping list
    try:
        print("Processing shipping list...")
        deduped_path, export_path, reimport_path = process_shipping_list(
            shipping_list_file,
            policy_file,
            shipping_rate_file,
            exchange_rate_file,
            output_deduped_file,
            output_export_file,
            output_reimport_file
        )
        
        print("\nProcessing completed successfully!")
        print(f"Deduplicated shipping list saved to: {deduped_path}")
        print(f"Export receipt saved to: {export_path}")
        print(f"Re-import receipt saved to: {reimport_path}")
        
    except Exception as e:
        print(f"Error occurred during processing: {e}")

if __name__ == "__main__":
    main() 