#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script to create sample policy, shipping rate, and exchange rate files for demonstration.
"""

import pandas as pd
import numpy as np
import os


def create_policy_file(file_path="sample_policy.xlsx"):
    """
    Create a sample policy file with markup percentage and insurance rate.
    
    Args:
        file_path (str): Path to save the policy file
    """
    # Example data for policy file
    data = {
        'markup_percentage': [15],  # 15% markup
        'insurance_rate': [2.5],   # 2.5% insurance rate
        'insurance_coefficient': [1.05]  # 1.05 insurance coefficient
    }
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel
    df.to_excel(file_path, index=False)
    print(f"Created sample policy file at: {file_path}")


def create_shipping_rate_file(file_path="sample_shipping_rate.xlsx"):
    """
    Create a sample shipping rate file.
    
    Args:
        file_path (str): Path to save the shipping rate file
    """
    # Example data for shipping rate file
    data = {
        'shipping_rate': [2.75],  # $2.75 per kg
        'effective_date': [pd.Timestamp('2023-01-01')],
        'expiry_date': [pd.Timestamp('2023-12-31')],
        'carrier': ['Sample Carrier'],
        'notes': ['Sample shipping rate for demonstration']
    }
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel
    df.to_excel(file_path, index=False)
    print(f"Created sample shipping rate file at: {file_path}")


def create_exchange_rate_file(file_path="sample_exchange_rate.xlsx"):
    """
    Create a sample exchange rate file with currency exchange rates.
    
    Args:
        file_path (str): Path to save the exchange rate file
    """
    # Example data for exchange rate file
    data = {
        'RMB_USD': [6.85],       # 1 USD = 6.85 RMB
        'RMB_RUPEE': [0.085],    # 1 Rupee = 0.085 RMB
        'USD_RUPEE': [82.5],     # 1 USD = 82.5 Rupee
        'effective_date': [pd.Timestamp('2023-01-01')],
        'notes': ['Sample exchange rates for demonstration']
    }
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel
    df.to_excel(file_path, index=False)
    print(f"Created sample exchange rate file at: {file_path}")


def main():
    """
    Main function to create all sample files.
    """
    print("Creating sample files for shipping processor demonstration...")
    
    # Create sample files
    create_policy_file()
    create_shipping_rate_file()
    create_exchange_rate_file()
    
    print("\nAll sample files created successfully!")
    print("You can now run the example.py script to test the shipping processor.")


if __name__ == "__main__":
    main() 