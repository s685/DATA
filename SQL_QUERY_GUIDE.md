# Multi-Line SQL Query Configuration Guide

## Overview

You can now write custom multi-line SQL queries directly in `config.yaml` for each worksheet. The query will be executed directly, and results will be used for direct dump or summaries based on the template type.

## How It Works

1. **If `query` is provided**: The SQL query is executed directly, and results are used as-is
2. **If `query` is NOT provided**: A default query is generated based on `template_type` and `filter` (if provided)

## YAML Multi-Line SQL Syntax

### Option 1: Literal Block Scalar (`|`) - Preserves Newlines

```yaml
worksheets:
  1-004:
    table_name: my_table_1_004
    template_type: direct_dump
    query: |
      SELECT 
        Policy_Num,
        Claim_Num,
        Product,
        Claim_Status,
        Company,
        Issue_State,
        Resident_State
      FROM my_table_1_004
      WHERE Schedule_ID = '1-004'
        AND Status = 'Active'
        AND Year = 2024
      ORDER BY Policy_Num
```

### Option 2: Folded Block Scalar (`>-`) - Folds Newlines to Spaces

```yaml
worksheets:
  2-001:
    table_name: my_table_2_001
    template_type: direct_dump_state_summary
    query: >-
      SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company,
             Issue_State, Resident_State
      FROM my_table_2_001
      WHERE Schedule_ID = '2-001' AND Status = 'Active'
      ORDER BY Policy_Num
```

## Examples

### Example 1: Direct Dump with Custom SQL

```yaml
worksheets:
  1-004:
    table_name: my_table_1_004
    template_type: direct_dump
    query: |
      SELECT 
        Policy,
        Lapse_Da,
        Stati,
        Status_Reas,
        Company,
        Issue_St,
        Resident_St
      FROM my_table_1_004
      WHERE Schedule_ID = '1-004'
        AND Status = 'Active'
        AND Lapse_Da >= '2024-01-01'
      ORDER BY Policy
```

### Example 2: Direct Dump with State Summary and Custom SQL

```yaml
worksheets:
  2-001:
    table_name: my_table_2_001
    template_type: direct_dump_state_summary
    query: |
      SELECT 
        Policy_Num,
        Claim_Num,
        Product,
        Claim_Status,
        Company,
        Issue_State,
        Resident_State
      FROM my_table_2_001
      WHERE Schedule_ID = '2-001'
        AND Claim_Status IN ('Approved', 'Pending')
        AND Year = 2024
      ORDER BY Issue_State, Policy_Num
```

**Note:** For summaries to work, the query must include the required columns:
- For state summaries: `Issue_State`, `Resident_State`, `Policy_Num` (or similar count field)
- For TAT summaries: `TAT_in_Days` column
- For pay req summaries: `Year_Pay_Req_Received` column

### Example 3: Direct Dump with TAT Summary and Custom SQL

```yaml
worksheets:
  2-003:
    table_name: my_table_2_003
    template_type: direct_dump_tat_summary
    query: |
      SELECT 
        Policy_Num,
        Claim_Num,
        Product,
        Claim_Status,
        Company,
        Issue_State,
        Resident_State,
        TAT_in_Days
      FROM my_table_2_003
      WHERE Schedule_ID = '2-003'
        AND TAT_in_Days IS NOT NULL
        AND Claim_Status = 'Approved'
      ORDER BY TAT_in_Days DESC
```

### Example 4: Complex Query with Joins

```yaml
worksheets:
  3-001:
    table_name: my_table_3_001
    template_type: direct_dump_tat_summary
    query: |
      SELECT 
        d.Inq_Date,
        d.Stat_Start_Date,
        d.Decision,
        d.Decision_Reason,
        d.Company,
        d.Issue_State,
        d.Resident_State,
        d.Product,
        d.Date_of_Loss,
        d.Schedule_ID,
        d.TAT_in_Days
      FROM my_table_3_001 d
      INNER JOIN status_table s ON d.Status_ID = s.Status_ID
      WHERE d.Schedule_ID = '3-001'
        AND s.Status = 'Active'
        AND d.Decision_Date >= '2024-01-01'
      ORDER BY d.TAT_in_Days
```

### Example 5: Using Filter Instead of Full Query

If you don't want to write the full query, you can use `filter` with default query generation:

```yaml
worksheets:
  1-004:
    table_name: my_table_1_004
    template_type: direct_dump
    filter: "Schedule_ID = '1-004' AND Status = 'Active' AND Year = 2024"
```

This will generate: `SELECT Policy_Num, Claim_Num, ... FROM my_table_1_004 WHERE Schedule_ID = '1-004' AND Status = 'Active' AND Year = 2024`

## Important Notes

1. **Query Priority**: If `query` is provided, it is used directly and `filter` is ignored
2. **Column Requirements**: For summaries to work correctly, ensure your query includes required columns:
   - State summaries need: `Issue_State`, `Resident_State`, and a count field (e.g., `Policy_Num`)
   - TAT summaries need: `TAT_in_Days`
   - Pay req summaries need: `Year_Pay_Req_Received`
3. **Table Name**: The `table_name` field is still required in config, but if you provide a full `query`, you can reference any table(s) in your SQL
4. **SQL Injection**: Always validate your SQL queries. The code executes queries as-is, so ensure proper security practices

## Benefits

- ✅ **Full SQL Control**: Write any SQL query (SELECT, JOINs, subqueries, etc.)
- ✅ **Multi-line Support**: Use YAML literal blocks for readable SQL
- ✅ **Flexible Filtering**: Complex WHERE clauses, date ranges, etc.
- ✅ **Works with Summaries**: Custom queries work with all template types
- ✅ **Backward Compatible**: If no query provided, uses default generation

