# Shipping List Processing Workflow

This document outlines the complete workflow for processing shipping lists and generating export and re-import receipts.

## Step-by-Step Workflow

### 1. Setup Environment

First, ensure that you have Python and the required packages installed:

```bash
pip install -r requirements.txt
```

### 2. Create Sample Configuration Files

If you don't have your own policy, shipping rate, and exchange rate files, you can create sample files:

```bash
python create_sample_files.py
```

This will generate:
- `sample_policy.xlsx` - Contains markup percentage and insurance rates
- `sample_shipping_rate.xlsx` - Contains shipping rates per kg
- `sample_exchange_rate.xlsx` - Contains currency exchange rates

### 3. Prepare Your Shipping List

Place your shipping list Excel file in an accessible location. The file should contain all required columns as described in the README.

If you want to use the sample file for testing:
- Use the file at `testfiles/original-input-shippinglist.xlsx`

### 4. Run the Processor

Run the shipping processor using the command-line interface:

```bash
python shipping_processor.py --shipping-list "testfiles/original-input-shippinglist.xlsx" \
                           --policy-file "sample_policy.xlsx" \
                           --shipping-rate-file "sample_shipping_rate.xlsx" \
                           --exchange-rate-file "sample_exchange_rate.xlsx" \
                           --output-deduped "output_deduped_shipping_list.xlsx" \
                           --output-export "output_export_receipt.xlsx" \
                           --output-reimport "output_reimport_receipt.xlsx"
```

Or use the example script for a quick demonstration:

```bash
python example.py
```

### 5. Review Generated Files

After successful execution, the following files will be generated:

1. `output_deduped_shipping_list.xlsx` - Deduplicated shipping list
2. `output_export_receipt.xlsx` - Export receipt with the following columns:
   - `NO.` - Sequential number
   - `P/N` - Part number
   - `DESCRIPTION` - Product description in English
   - `Model NO.` - Product model number
   - `Unit Price USD` - CIF unit price in USD
   - `Qty` - Quantity
   - `Unit` - Unit of measure
   - `Amount USD` - Total amount (Unit Price USD × Qty)
3. `output_reimport_receipt.xlsx` - Re-import receipt with customs information:
   - `NO.` - Sequential number
   - `P/N` - Part number
   - `English Description` - Customs description in English
   - `Chinese Description` - Customs description in Chinese
   - `Model NO.` - Product model number
   - `Unit Price USD` - CIF unit price in USD
   - `Qty` - Quantity
   - `Unit` - Unit of measure
   - `Amount USD` - Total amount (Unit Price USD × Qty)
   - Weight information for customs purposes

Review these files to ensure that they contain the expected data.

## Troubleshooting

### Common Issues

#### File Not Found Error
- Ensure that the paths to the input files are correct
- Check that the files exist in the specified locations

#### Column Names Mismatch
- Ensure that the input files have the expected column names
- If needed, modify the column mapping in the `read_shipping_list` function in `shipping_processor.py`

#### Calculation Errors
- Verify that the policy, shipping rate, and exchange rate files contain valid numeric data
- Check for missing or zero values that might cause division by zero errors

## Customization

### Adjusting Column Mappings

If your shipping list file has different column names, you can adjust the column mapping in the `read_shipping_list` function in `shipping_processor.py`.

### Modifying Output Formats

To modify the content or format of the output files, adjust the `generate_export_receipt` and `generate_reimport_receipt` functions in `shipping_processor.py`. 