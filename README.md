# Multi-Worksheet Excel Report Generator from Snowflake

A Python CLI application that generates multi-worksheet Excel workbooks from Snowflake data, matching exact template structures with detail and summary tables.

## Features

- ✅ Single Python file containing all functionality
- ✅ Multi-worksheet support with configurable SQL queries
- ✅ Automatic summary table generation from detail records
- ✅ Exact spacing preservation between detail and summary sections
- ✅ Template-matching formatting (headers, filters, highlighting, borders)
- ✅ Dual authentication support (SSO and username/password)
- ✅ YAML-based configuration for flexibility

## Installation

1. Install Python 3.7 or higher

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Configuration

All worksheet structures are **hardcoded** in the Python file based on the template images. You only need to provide Snowflake connection details and table name.

## Usage

Run the script with Snowflake connection parameters and output path:

```bash
python excel_report_generator.py --account <account> --warehouse <warehouse> --database <database> --schema <schema> --table <table_name> --output report.xlsx
```

### Arguments

**Required:**
- `--account`: Snowflake account identifier (e.g., "xy12345.us-east-1")
- `--warehouse`: Snowflake warehouse name
- `--database`: Database name
- `--schema`: Schema name
- `--table`: Snowflake table name (used in all queries)
- `--output`: Path for output Excel file

**Optional:**
- `--user`: Snowflake username (required if using password auth)
- `--password`: Snowflake password (required if using password auth)
- `--authenticator`: Authentication method - `externalbrowser` (SSO, default) or `snowflake` (username/password)

### Examples

**Using SSO (External Browser):**
```bash
python excel_report_generator.py --account xy12345.us-east-1 --warehouse MY_WH --database MY_DB --schema MY_SCHEMA --table claims_table --output MCAS_Reporting_Year_2024_CCC_v1.xlsx
```

**Using Username/Password:**
```bash
python excel_report_generator.py --account xy12345.us-east-1 --warehouse MY_WH --database MY_DB --schema MY_SCHEMA --table claims_table --user myuser --password mypass --authenticator snowflake --output MCAS_Reporting_Year_2024_CCC_v1.xlsx
```

## Hardcoded Worksheet Structures

All worksheet definitions are hardcoded in the Python file and match the template structure:

- **Summary**: Aggregates data from all detail worksheets
- **1-001, 1-004, 1-006**: Policy/Claim details with Issue/Resident State summaries
- **2-001 through 2-005**: Policy/Claim details with summaries
- **3-001**: Decision details with TAT metrics and summaries
- **3-003, 3-004, 3-005, 3-006, 3-007**: Decision details with TAT metrics
- **5-001 through 5-004**: Policy/Claim details with summaries
- **6-001**: EOB/Benefit details
- **6-002 through 6-004**: Policy/Claim details with summaries

Each worksheet's column structure, summary configurations, and formatting are predefined based on the template images.

## Worksheet Structure

The generator creates worksheets matching your template structure:

1. **Detail Tables**: Fetched from Snowflake using configured SQL queries
2. **Summary Tables**: Automatically generated from detail records based on `summary_config`
3. **Spacing**: Empty columns maintained between detail and summary sections
4. **Formatting**: 
   - Headers with filters
   - Yellow highlighting on specified columns
   - Borders around tables
   - Auto-adjusted column widths

## Configuration Examples

### Worksheet with Detail and Summary Tables

```yaml
- name: "1-001"
  query: "SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM your_table WHERE Schedule_ID = '1-001'"
  detail_start_column: "A"
  detail_columns: ["Policy Num", "Claim Num", "Product", "Claim Status", "Company", "Issue Sta", "Resident Sta"]
  spacing_columns: ["H"]
  summary_config:
    - group_by: "Issue_State"
      aggregates:
        - field: "Policy_Num"
          function: "COUNT"
          label: "Count"
      start_column: "I"
      columns: ["Issue", "Count"]
  formatting:
    header_row: 1
    filters: true
    highlight_columns: []
```

### Worksheet with Highlighted Columns

```yaml
- name: "3-001"
  query: "SELECT Decision, Decision_Reason, Company, Issue_State FROM your_table WHERE Schedule_ID = '3-001'"
  detail_start_column: "A"
  detail_columns: ["Decision", "Decision Reason", "Company", "Issue State"]
  formatting:
    header_row: 1
    filters: true
    highlight_columns: ["A", "B"]  # Yellow highlight on Decision and Decision Reason
```

## Troubleshooting

### Connection Issues

- **SSO Authentication**: Ensure you're logged into your browser and can access Snowflake
- **Username/Password**: Verify credentials are correct
- **Account Format**: Use format like `xy12345.us-east-1` (without `https://` or `.snowflakecomputing.com`)

### Query Errors

- Verify table name matches your Snowflake schema
- Ensure your table has columns matching the hardcoded query structures:
  - `Policy_Num`, `Claim_Num`, `Product`, `Claim_Status`, `Company`, `Issue_State`, `Resident_State` (for standard worksheets)
  - `Schedule_ID` column for filtering
- Check that `Schedule_ID` values in your data match worksheet names (e.g., '1-001', '2-001', etc.)

### Excel Generation Issues

- Ensure output directory exists and is writable
- Verify that your Snowflake table structure matches the expected columns for each worksheet
- Check that data exists for the Schedule_IDs referenced in the worksheets

## Requirements

- Python 3.7+
- snowflake-connector-python >= 3.0.0
- openpyxl >= 3.1.0

## License

This project is provided as-is for internal use.

