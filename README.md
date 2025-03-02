# Shipping List Processor

This Python script processes shipping list Excel files to generate export and re-import receipts according to specified business rules.

## Features

- De-duplicates shipping list items where part number and unit price are the same
- Calculates CIF (Cost, Insurance, Freight) prices based on markup policy, insurance rates, and shipping rates
- Generates export and re-import receipts in Excel format
- Handles currency conversions based on exchange rates

## Requirements

- Python 3.6+
- Required Python packages:
  - pandas
  - numpy
  - openpyxl

## Installation

1. Clone this repository or download the script files.
2. Install the required packages:

```bash
pip install pandas numpy openpyxl
```

## Input Files

The script requires the following input files:

1. **Shipping List Excel File**: Contains details of items to be shipped.
2. **Policy Excel File**: Contains markup percentage and insurance rate settings.
3. **Shipping Rate Excel File**: Contains the current shipping rate.
4. **Exchange Rate Excel File**: Contains currency exchange rates.

### File Formats

#### Shipping List File
The shipping list file should contain columns with the following information (column names may vary):
- Part numbers
- Quantities
- Unit prices
- Weights
- Dimensions
- Other product details

#### Policy File
The policy file should contain:
- `markup_percentage`: The percentage markup to apply
- `insurance_rate`: The insurance rate to apply
- `insurance_coefficient`: (Optional) A coefficient for insurance calculations

#### Shipping Rate File
The shipping rate file should contain:
- `shipping_rate`: The shipping rate per kg

#### Exchange Rate File
The exchange rate file should contain:
- `RMB_USD`: The exchange rate from RMB to USD
- `RMB_RUPEE`: (Optional) The exchange rate from RMB to Rupee
- `USD_RUPEE`: (Optional) The exchange rate from USD to Rupee

## Usage

Run the script with the following command:

```bash
python shipping_processor.py --shipping-list "path/to/shipping_list.xlsx" \
                           --policy-file "path/to/policy.xlsx" \
                           --shipping-rate-file "path/to/shipping_rate.xlsx" \
                           --exchange-rate-file "path/to/exchange_rate.xlsx" \
                           --output-deduped "path/to/deduped_output.xlsx" \
                           --output-export "path/to/export_receipt.xlsx" \
                           --output-reimport "path/to/reimport_receipt.xlsx"
```

### Command-line Arguments

- `--shipping-list`: Path to the shipping list Excel file (required)
- `--policy-file`: Path to the policy Excel file (required)
- `--shipping-rate-file`: Path to the shipping rate Excel file (required)
- `--exchange-rate-file`: Path to the exchange rate Excel file (required)
- `--output-deduped`: Path to save the deduplicated shipping list (default: "deduped_shipping_list.xlsx")
- `--output-export`: Path to save the export receipt (default: "export_receipt.xlsx")
- `--output-reimport`: Path to save the re-import receipt (default: "reimport_receipt.xlsx")

## Output Files

The script generates three output files:

1. **Deduplicated Shipping List**: An Excel file containing the deduplicated shipping list.
2. **Export Receipt**: An Excel file containing the export receipt with CIF pricing. It includes the following columns:
   - `NO.`: Sequential number
   - `P/N`: Part number
   - `DESCRIPTION`: Product description in English
   - `Model NO.`: Product model number
   - `Unit Price USD`: CIF unit price in USD
   - `Qty`: Quantity
   - `Unit`: Unit of measure
   - `Amount USD`: Total amount (Unit Price USD × Qty)
3. **Re-import Receipt**: An Excel file containing the re-import receipt with additional customs information, including:
   - `NO.`: Sequential number
   - `P/N`: Part number
   - `English Description`: Customs description in English
   - `Chinese Description`: Customs description in Chinese
   - `Model NO.`: Product model number
   - `Unit Price USD`: CIF unit price in USD
   - `Qty`: Quantity
   - `Unit`: Unit of measure
   - `Amount USD`: Total amount (Unit Price USD × Qty)
   - Weight information for customs purposes

## Processing Steps

1. Read and parse all input files.
2. Deduplicate the shipping list based on part number and unit price.
3. Calculate CIF prices based on the specified formulas.
4. Generate export and re-import receipts.

## CIF Price Calculation

The CIF unit price is calculated as follows:

1. Calculate supplier cost: `unit_price * quantity`
2. Calculate adjusted cost with markup: `supplier_cost * (1 + markup_percentage)`
3. Calculate adjusted cost with insurance: `supplier_cost * insurance_coefficient * (1 + insurance_rate)`
4. Calculate shipping cost: `total_net_weight * shipping_rate`
5. Calculate final CIF unit price: `(adjusted_cost_with_insurance + shipping_cost) / quantity`
6. Convert to USD: `cif_unit_price_rmb / exchange_rate_RMB_USD`

## License

This project is licensed under the MIT License - see the LICENSE file for details. 