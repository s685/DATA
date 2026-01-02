# Template Type Configuration Guide

## Overview

You can now configure template types for each worksheet in `config.yaml`, allowing you to choose what summaries are included without modifying code.

## Available Template Types

1. **`direct_dump`** - Detail records only, no summaries
2. **`direct_dump_state_summary`** - Detail + Issue State + Resident State summaries
3. **`direct_dump_tat_summary`** - Detail + TAT summary (with % of Total)
4. **`direct_dump_state_tat_summary`** - Detail + Issue State + Resident State + TAT summaries
5. **`state_summary_only`** - Only summaries (Issue State + Resident State), no detail records
6. **`direct_dump_state_payreq_summary`** - Detail + Issue State + Resident State + Year Pay Req Received summaries

## Configuration Format

### Format 1: Simple (Backward Compatible)

Use the simple format to use hardcoded structures (existing behavior):

```yaml
worksheets:
  1-001: my_table_1_001
  2-001: my_table_2_001
```

### Format 2: Advanced (Configurable Template Types)

Use the advanced format to specify template types:

```yaml
worksheets:
  1-004:
    table_name: my_table_1_004
    template_type: direct_dump
  
  2-001:
    table_name: my_table_2_001
    template_type: direct_dump_state_summary
  
  2-003:
    table_name: my_table_2_003
    template_type: direct_dump_tat_summary
  
  5-003:
    table_name: my_table_5_003
    template_type: direct_dump_state_tat_summary
  
  1-001:
    table_name: my_table_1_001
    template_type: state_summary_only
  
  6-004:
    table_name: my_table_6_004
    template_type: direct_dump_state_payreq_summary
```

### Advanced Options

You can also specify custom queries and column names:

```yaml
worksheets:
  2-001:
    table_name: my_table_2_001
    template_type: direct_dump_state_summary
    query: "SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM my_table_2_001 WHERE Schedule_ID = '2-001'"
    detail_columns: ["Policy Num", "Claim Num", "Product", "Claim Status", "Company", "Issue State", "Resident State"]
```

**Note:** If `query` is not specified, a default query is generated based on the template type. If `detail_columns` is not specified, actual column names from the query results are used.

## Examples

### Example 1: All Direct Dump (No Summaries)

```yaml
worksheets:
  1-004:
    table_name: table_1_004
    template_type: direct_dump
  
  3-003:
    table_name: table_3_003
    template_type: direct_dump
  
  3-004:
    table_name: table_3_004
    template_type: direct_dump
```

### Example 2: Mix of Template Types

```yaml
worksheets:
  # Direct dump only
  1-004:
    table_name: table_1_004
    template_type: direct_dump
  
  # Direct dump + state summaries
  2-001:
    table_name: table_2_001
    template_type: direct_dump_state_summary
  
  # Direct dump + TAT summary
  2-003:
    table_name: table_2_003
    template_type: direct_dump_tat_summary
  
  # Direct dump + state + TAT summaries
  5-003:
    table_name: table_5_003
    template_type: direct_dump_state_tat_summary
  
  # Summary only
  1-001:
    table_name: table_1_001
    template_type: state_summary_only
  
  # Direct dump + state + pay req summary
  6-004:
    table_name: table_6_004
    template_type: direct_dump_state_payreq_summary
```

## Column Requirements

### For State Summaries
- Requires: `Issue_State`, `Resident_State`, `Policy_Num` (or similar count field)

### For TAT Summaries
- Requires: `TAT_in_Days` column in detail records

### For Pay Req Summaries
- Requires: `Year_Pay_Req_Received` column in detail records

## Backward Compatibility

The code maintains **full backward compatibility**:
- If you use simple format (`worksheet_name: table_name`), it uses hardcoded structures
- If you use advanced format but don't specify `template_type`, it falls back to hardcoded structures
- If `template_type` is invalid, it falls back to hardcoded structures with a warning

## Benefits

1. **Flexibility**: Choose template types per worksheet without code changes
2. **Maintainability**: Update configurations without modifying Python code
3. **Reusability**: Same template types work across different worksheets
4. **Backward Compatible**: Existing configurations continue to work

