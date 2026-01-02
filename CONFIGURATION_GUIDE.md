# Configuration Guide for Schedules 1, 2, and 3

## Overview

This guide shows how to configure Schedule 1, Schedule 2, and Schedule 3 worksheets in your `config.yaml` file.

## Configuration Format

You can use two formats for worksheets:

### Format 1: Simple (Just Table Name)
```yaml
worksheets:
  1-001: your_table_name
  2-001: your_table_name
```

### Format 2: Advanced (With Template Type and Custom Query)
```yaml
worksheets:
  1-001:
    table_name: your_table_name
    template_type: state_summary_only
    query: |
      SELECT Policy_Num, Issue_State, Resident_State
      FROM your_table_name
      WHERE Schedule_ID = '1-001'
```

---

## Schedule 1 Worksheets

### 1-001: Summary Only (No Detail Records)
**Template Type:** `state_summary_only`
**Structure:** Issue State summary + Resident State summary (no detail records)

**Simple Format:**
```yaml
worksheets:
  1-001: schedule1_table
```

**Advanced Format:**
```yaml
worksheets:
  1-001:
    table_name: schedule1_table
    template_type: state_summary_only
    # Optional: Custom query
    query: |
      SELECT Policy_Num, Issue_State, Resident_State
      FROM schedule1_table
      WHERE Schedule_ID = '1-001'
```

### 1-004: Detail Records Only (No Summaries)
**Template Type:** `direct_dump`
**Structure:** Detail records only, no summaries

**Simple Format:**
```yaml
worksheets:
  1-004: schedule1_table
```

**Advanced Format:**
```yaml
worksheets:
  1-004:
    table_name: schedule1_table
    template_type: direct_dump
    # Optional: Custom query
    query: |
      SELECT Policy, Lapse_Da, Stati, Status_Reas, Company, Issue_St, Resident_St
      FROM schedule1_table
      WHERE Schedule_ID = '1-004'
```

### 1-006: Summary Only (No Detail Records)
**Template Type:** `state_summary_only`
**Structure:** Issue State summary + Resident State summary (no detail records)

**Simple Format:**
```yaml
worksheets:
  1-006: schedule1_table
```

**Advanced Format:**
```yaml
worksheets:
  1-006:
    table_name: schedule1_table
    template_type: state_summary_only
```

---

## Schedule 2 Worksheets

### 2-001: Detail + Issue State/Resident State Summaries
**Template Type:** `direct_dump_state_summary`
**Structure:** Detail records + Issue State summary + Resident State summary

**Simple Format:**
```yaml
worksheets:
  2-001: schedule2_table
```

**Advanced Format:**
```yaml
worksheets:
  2-001:
    table_name: schedule2_table
    template_type: direct_dump_state_summary
    # Optional: Custom query
    query: |
      SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State
      FROM schedule2_table
      WHERE Schedule_ID = '2-001'
```

### 2-002: Same as 2-001
**Template Type:** `direct_dump_state_summary`

```yaml
worksheets:
  2-002: schedule2_table
```

### 2-003: Detail + TAT Summary
**Template Type:** `direct_dump_tat_summary`
**Structure:** Detail records + TAT summary (with % of Total)

**Simple Format:**
```yaml
worksheets:
  2-003: schedule2_table
```

**Advanced Format:**
```yaml
worksheets:
  2-003:
    table_name: schedule2_table
    template_type: direct_dump_tat_summary
    # Optional: Custom query (must include TAT_in_Days)
    query: |
      SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State, TAT_in_Days
      FROM schedule2_table
      WHERE Schedule_ID = '2-003'
```

### 2-004: Same as 2-001
**Template Type:** `direct_dump_state_summary`

```yaml
worksheets:
  2-004: schedule2_table
```

### 2-005: Same as 2-001
**Template Type:** `direct_dump_state_summary`

```yaml
worksheets:
  2-005: schedule2_table
```

---

## Schedule 3 Worksheets

### 3-001: Detail + TAT Summary
**Template Type:** `direct_dump_tat_summary`
**Structure:** Detail records + TAT summary (with % of Total)

**Simple Format:**
```yaml
worksheets:
  3-001: schedule3_table
```

**Advanced Format:**
```yaml
worksheets:
  3-001:
    table_name: schedule3_table
    template_type: direct_dump_tat_summary
    # Optional: Custom query (must include TAT_in_Days)
    query: |
      SELECT Inq_Date, Stat_Start_Date, Decision, Decision_Reason, Company, Issue_State, Resident_State, Product, Date_of_Loss, Schedule_ID, TAT_in_Days
      FROM schedule3_table
      WHERE Schedule_ID = '3-001'
```

### 3-003: Detail Records Only
**Template Type:** `direct_dump`
**Structure:** Detail records only, no summaries

**Simple Format:**
```yaml
worksheets:
  3-003: schedule3_table
```

**Advanced Format:**
```yaml
worksheets:
  3-003:
    table_name: schedule3_table
    template_type: direct_dump
    # Optional: Custom query
    query: |
      SELECT Decision_Date, Inq_Date, Stat_Start_Date, Decision, Decision_Reason, Company, Issue_State, Resident_State, Product, Date_of_Loss, Schedule_ID, TAT_in_Days
      FROM schedule3_table
      WHERE Schedule_ID = '3-003'
```

### 3-004: Detail Records Only
**Template Type:** `direct_dump`

```yaml
worksheets:
  3-004: schedule3_table
```

### 3-005: Detail Records Only
**Template Type:** `direct_dump`

```yaml
worksheets:
  3-005: schedule3_table
```

### 3-006: Detail Records Only
**Template Type:** `direct_dump`

```yaml
worksheets:
  3-006: schedule3_table
```

### 3-007: Detail Records Only
**Template Type:** `direct_dump`

```yaml
worksheets:
  3-007: schedule3_table
```

---

## Complete Example Configuration

Here's a complete `config.yaml` example with Summary, Schedule 1, Schedule 2, and Schedule 3:

```yaml
# Snowflake Connection (or use environment variables)
snowflake:
  account: <your_account>
  warehouse: <your_warehouse>
  database: <your_database>
  schema: <your_schema>
  authenticator: externalbrowser

# Summary Worksheet Configuration
summary:
  table_name: "summary_table"
  query: "SELECT ID, description, value FROM summary_table"
  schedule_titles:
    1: "Schedule 1 - General Information"
    2: "Schedule 2 - Claimants"
    3: "Schedule 3 - Claimant Requests Denied/Not Paid"
    4: "Schedule 4"
    5: "Schedule 5"
    6: "Schedule 6"

# Worksheet Configuration
worksheets:
  # Schedule 1
  1-001: schedule1_table  # Summary only
  1-004: schedule1_table  # Detail only
  1-006: schedule1_table  # Summary only
  
  # Schedule 2
  2-001: schedule2_table  # Detail + State summaries
  2-002: schedule2_table  # Detail + State summaries
  2-003: schedule2_table  # Detail + TAT summary
  2-004: schedule2_table  # Detail + State summaries
  2-005: schedule2_table  # Detail + State summaries
  
  # Schedule 3
  3-001: schedule3_table  # Detail + TAT summary
  3-003: schedule3_table  # Detail only
  3-004: schedule3_table  # Detail only
  3-005: schedule3_table  # Detail only
  3-006: schedule3_table  # Detail only
  3-007: schedule3_table  # Detail only
```

---

## Using Different Tables for Each Worksheet

If each worksheet uses a different table:

```yaml
worksheets:
  # Schedule 1
  1-001: schedule1_001_table
  1-004: schedule1_004_table
  1-006: schedule1_006_table
  
  # Schedule 2
  2-001: schedule2_001_table
  2-002: schedule2_002_table
  2-003: schedule2_003_table
  2-004: schedule2_004_table
  2-005: schedule2_005_table
  
  # Schedule 3
  3-001: schedule3_001_table
  3-003: schedule3_003_table
  3-004: schedule3_004_table
  3-005: schedule3_005_table
  3-006: schedule3_006_table
  3-007: schedule3_007_table
```

---

## Using Custom Queries

If you need custom queries with filters:

```yaml
worksheets:
  2-001:
    table_name: schedule2_table
    template_type: direct_dump_state_summary
    query: |
      SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State
      FROM schedule2_table
      WHERE Schedule_ID = '2-001'
        AND Status = 'Active'
        AND Year = 2024
      ORDER BY Policy_Num
```

**Note:** Table names in queries are automatically resolved to `database.schema.table_name` using your Snowflake configuration.

---

## Available Template Types

- `direct_dump`: Detail records only, no summaries
- `direct_dump_state_summary`: Detail + Issue State + Resident State summaries
- `direct_dump_tat_summary`: Detail + TAT summary (with % of Total)
- `direct_dump_state_tat_summary`: Detail + Issue State + Resident State + TAT summaries
- `state_summary_only`: Only summaries (Issue State + Resident State), no detail records
- `direct_dump_state_payreq_summary`: Detail + Issue State + Resident State + Year Pay Req Received summaries

---

## Important Notes

1. **Table Name Resolution**: If you use just the table name in queries (e.g., `FROM my_table`), it will automatically become `database.schema.my_table` based on your Snowflake config.

2. **Column Requirements**: 
   - For state summaries: Need `Issue_State` and `Resident_State` columns
   - For TAT summaries: Need `TAT_in_Days` column
   - For state summary only: Need `Policy_Num`, `Issue_State`, `Resident_State`

3. **Schedule_ID**: The code automatically filters by `Schedule_ID = 'worksheet_name'` unless you provide a custom query.

4. **Simple vs Advanced**: Use simple format if your table structure matches the default. Use advanced format for custom queries or different table structures.

