# Schedule 2 Configuration Guide

## Overview

Schedule 2 has 5 worksheets with two different template types:
- **2-001, 2-002, 2-004, 2-005**: Detail records + Issue State/Resident State summaries
- **2-003**: Detail records + TAT summary (with % of Total)

---

## Worksheet 2-001, 2-002, 2-004, 2-005

### Template Type: `direct_dump_state_summary`

### Required Columns:

Your query must return these columns (or use aliases):

1. **`Policy_Num`** - Used for COUNT aggregation in summaries
2. **`Claim_Num`** - Claim number
3. **`Product`** - Product name/type
4. **`Claim_Status`** - Status of the claim
5. **`Company`** - Company name
6. **`Issue_State`** - State where policy was issued (used for Issue State summary)
7. **`Resident_State`** - State where policyholder resides (used for Resident State summary)

### Default Query:

If you don't provide a custom query, the code generates:
```sql
SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State 
FROM your_table 
WHERE Schedule_ID = '2-001'
```

### Example Configuration:

**If your columns match exactly:**
```yaml
2-001:
  table_name: schedule2_table
  template_type: direct_dump_state_summary
```

**If your columns have different names (use aliases):**
```yaml
2-001:
  table_name: schedule2_table
  template_type: direct_dump_state_summary
  query: |
    SELECT 
      policy_number AS Policy_Num,
      claim_number AS Claim_Num,
      product_name AS Product,
      claim_status AS Claim_Status,
      company_name AS Company,
      issue_state_code AS Issue_State,
      resident_state_code AS Resident_State
    FROM schedule2_table
    WHERE Schedule_ID = '2-001'
      AND Status = 'Active'
    ORDER BY Policy_Num
```

### What Gets Generated:

1. **Detail Table** (Columns A-G):
   - All detail records with the 7 columns listed above
   - Headers with filters enabled
   - Column H is a gap/spacing column

2. **Issue State Summary** (Columns I-J):
   - Groups by Issue_State
   - Counts Policy_Num per state
   - Shows: Issue State | Count

3. **Resident State Summary** (Columns L-M):
   - Groups by Resident_State
   - Counts Policy_Num per state
   - Shows: Resident State | Count
   - Column K is a gap between summaries

---

## Worksheet 2-003

### Template Type: `direct_dump_tat_summary`

### Required Columns:

Your query must return these columns (or use aliases):

1. **`Policy_Num`** - Policy number
2. **`Claim_Num`** - Claim number
3. **`Product`** - Product name/type
4. **`Claim_Status`** - Status of the claim
5. **`Company`** - Company name
6. **`Issue_State`** - State where policy was issued
7. **`Resident_State`** - State where policyholder resides
8. **`TAT_in_Days`** - Turnaround time in days (REQUIRED for TAT summary)

### Default Query:

If you don't provide a custom query, the code generates:
```sql
SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State, TAT_in_Days 
FROM your_table 
WHERE Schedule_ID = '2-003'
```

### Example Configuration:

**If your columns match exactly:**
```yaml
2-003:
  table_name: schedule2_table
  template_type: direct_dump_tat_summary
```

**If your columns have different names (use aliases):**
```yaml
2-003:
  table_name: schedule2_table
  template_type: direct_dump_tat_summary
  query: |
    SELECT 
      policy_number AS Policy_Num,
      claim_number AS Claim_Num,
      product_name AS Product,
      claim_status AS Claim_Status,
      company_name AS Company,
      issue_state_code AS Issue_State,
      resident_state_code AS Resident_State,
      turnaround_days AS TAT_in_Days
    FROM schedule2_table
    WHERE Schedule_ID = '2-003'
      AND TAT_in_Days IS NOT NULL
    ORDER BY TAT_in_Days
```

### What Gets Generated:

1. **Detail Table** (Columns A-H):
   - All detail records with the 8 columns listed above
   - Headers with filters enabled
   - Column I is a gap/spacing column

2. **TAT Summary** (Columns J-L):
   - Groups TAT_in_Days into ranges:
     - "-1 to <31" (days -1 to 30)
     - ">30 and <61" (days 31 to 60)
     - ">60 and <91" (days 61 to 90)
     - ">90" (days 91 and above)
   - Shows: (empty) | TAT COUNTS | % of Total
   - Includes Grand Total row

---

## Complete Schedule 2 Configuration Example

```yaml
worksheets:
  # Schedule 2 - All use same table
  2-001:
    table_name: schedule2_table
    template_type: direct_dump_state_summary
  
  2-002:
    table_name: schedule2_table
    template_type: direct_dump_state_summary
  
  2-003:
    table_name: schedule2_table
    template_type: direct_dump_tat_summary
    # Note: Must include TAT_in_Days column
  
  2-004:
    table_name: schedule2_table
    template_type: direct_dump_state_summary
  
  2-005:
    table_name: schedule2_table
    template_type: direct_dump_state_summary
```

---

## Column Mapping Reference

If your table uses different column names, here's the mapping:

| Expected Column | Your Column Name | Example Alias |
|----------------|------------------|--------------|
| `Policy_Num` | `policy_number`, `policy_no`, `PolicyNo` | `policy_number AS Policy_Num` |
| `Claim_Num` | `claim_number`, `claim_no`, `ClaimNo` | `claim_number AS Claim_Num` |
| `Product` | `product_name`, `product_type`, `ProductName` | `product_name AS Product` |
| `Claim_Status` | `claim_status`, `status`, `Status` | `status AS Claim_Status` |
| `Company` | `company_name`, `company`, `CompanyName` | `company_name AS Company` |
| `Issue_State` | `issue_state`, `issue_state_code`, `IssueState` | `issue_state_code AS Issue_State` |
| `Resident_State` | `resident_state`, `resident_state_code`, `ResidentState` | `resident_state_code AS Resident_State` |
| `TAT_in_Days` | `tat_days`, `turnaround_days`, `TATDays` | `turnaround_days AS TAT_in_Days` |

---

## Common Issues and Solutions

### Issue 1: Missing TAT_in_Days for 2-003
**Error**: TAT summary not showing or showing zeros
**Solution**: Make sure your query includes `TAT_in_Days` column (or alias it)

### Issue 2: Summaries showing zero counts
**Error**: Issue State/Resident State summaries show 0
**Solution**: 
- Check that `Policy_Num` column exists and has values
- Check that `Issue_State` and `Resident_State` columns exist and have values
- Verify your WHERE clause is correct

### Issue 3: Wrong column names
**Error**: Columns not displaying correctly
**Solution**: Use column aliases in your query to match expected names

---

## Testing Your Configuration

1. **Start with simple configuration** (no custom query):
   ```yaml
   2-001:
     table_name: schedule2_table
     template_type: direct_dump_state_summary
   ```

2. **Run the script** and check if data appears

3. **If columns don't match**, add a custom query with aliases:
   ```yaml
   2-001:
     table_name: schedule2_table
     template_type: direct_dump_state_summary
     query: |
       SELECT 
         your_col1 AS Policy_Num,
         your_col2 AS Claim_Num,
         ...
       FROM schedule2_table
       WHERE Schedule_ID = '2-001'
   ```

4. **Verify summaries** are calculating correctly

---

## Quick Reference

| Worksheet | Template Type | Key Requirement |
|-----------|--------------|----------------|
| 2-001 | `direct_dump_state_summary` | Need: Policy_Num, Issue_State, Resident_State |
| 2-002 | `direct_dump_state_summary` | Need: Policy_Num, Issue_State, Resident_State |
| 2-003 | `direct_dump_tat_summary` | Need: TAT_in_Days column |
| 2-004 | `direct_dump_state_summary` | Need: Policy_Num, Issue_State, Resident_State |
| 2-005 | `direct_dump_state_summary` | Need: Policy_Num, Issue_State, Resident_State |

