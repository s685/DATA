#!/usr/bin/env python3
"""
Test script for state_summary_with_company template
Creates sample data and tests the summary generation
"""

import sys
from collections import defaultdict
from dataclasses import dataclass
from typing import List, Dict, Any

# Import the necessary functions from excel_report_generator
sys.path.insert(0, '.')
from excel_report_generator import (
    SummaryConfig, AggregateConfig, generate_summary
)

def test_state_summary_with_company():
    """Test state_summary_with_company template with sample data"""
    
    print("=" * 80)
    print("Testing state_summary_with_company template")
    print("=" * 80)
    
    # Create sample data matching the expected structure
    # Data should have: Issue_State, Resident_State, Company, Count
    sample_data = [
        {'Issue_State': 'CA', 'Resident_State': 'CA', 'Company': 'Company A', 'Count': 10},
        {'Issue_State': 'CA', 'Resident_State': 'CA', 'Company': 'Company A', 'Count': 15},
        {'Issue_State': 'CA', 'Resident_State': 'NY', 'Company': 'Company A', 'Count': 5},
        {'Issue_State': 'NY', 'Resident_State': 'NY', 'Company': 'Company B', 'Count': 20},
        {'Issue_State': 'NY', 'Resident_State': 'NY', 'Company': 'Company B', 'Count': 25},
        {'Issue_State': 'NY', 'Resident_State': 'TX', 'Company': 'Company B', 'Count': 8},
        {'Issue_State': 'TX', 'Resident_State': 'TX', 'Company': 'Company C', 'Count': 12},
        {'Issue_State': 'TX', 'Resident_State': 'TX', 'Company': 'Company C', 'Count': 18},
        {'Issue_State': 'TX', 'Resident_State': 'CA', 'Company': 'Company C', 'Count': 7},
    ]
    
    print(f"\nSample data ({len(sample_data)} records):")
    for i, record in enumerate(sample_data, 1):
        print(f"  {i}. {record}")
    
    # Create summary config for Issue State
    issue_state_config = SummaryConfig(
        group_by='Issue_State',
        aggregates=[
            AggregateConfig(field='Count', function='SUM', label='Count'),
            AggregateConfig(field='Company', function='FIRST', label='Company')
        ],
        start_column='A',
        columns=['Issue State', 'Count', 'Company']
    )
    
    # Create summary config for Resident State
    resident_state_config = SummaryConfig(
        group_by='Resident_State',
        aggregates=[
            AggregateConfig(field='Count', function='SUM', label='Count'),
            AggregateConfig(field='Company', function='FIRST', label='Company')
        ],
        start_column='E',
        columns=['Resident State', 'Count', 'Company']
    )
    
    print("\n" + "=" * 80)
    print("Testing Issue State Summary")
    print("=" * 80)
    
    # Test Issue State summary
    issue_summary = generate_summary(sample_data, issue_state_config, include_grand_total=True)
    
    print(f"\nIssue State Summary ({len(issue_summary)} rows):")
    print(f"{'Issue State':<15} {'Count':<10} {'Company':<15}")
    print("-" * 40)
    for row in issue_summary:
        issue_state = row.get('Issue State', '') or ''
        count = row.get('Count', 0) or 0
        company = row.get('Company', '') or ''
        print(f"{str(issue_state):<15} {str(count):<10} {str(company):<15}")
    
    # Verify Issue State results
    print("\nExpected Issue State Summary:")
    print("  CA: Count = 10 + 15 + 5 = 30, Company = 'Company A' (first value)")
    print("  NY: Count = 20 + 25 + 8 = 53, Company = 'Company B' (first value)")
    print("  TX: Count = 12 + 18 + 7 = 37, Company = 'Company C' (first value)")
    print("  Grand Total: Count = 120")
    
    print("\n" + "=" * 80)
    print("Testing Resident State Summary")
    print("=" * 80)
    
    # Test Resident State summary
    resident_summary = generate_summary(sample_data, resident_state_config, include_grand_total=True)
    
    print(f"\nResident State Summary ({len(resident_summary)} rows):")
    print(f"{'Resident State':<15} {'Count':<10} {'Company':<15}")
    print("-" * 40)
    for row in resident_summary:
        resident_state = row.get('Resident State', '') or ''
        count = row.get('Count', 0) or 0
        company = row.get('Company', '') or ''
        print(f"{str(resident_state):<15} {str(count):<10} {str(company):<15}")
    
    # Verify Resident State results
    print("\nExpected Resident State Summary:")
    print("  CA: Count = 10 + 15 + 7 = 32, Company = 'Company A' (first value)")
    print("  NY: Count = 5 + 20 + 25 = 50, Company = 'Company A' (first value)")
    print("  TX: Count = 8 + 12 + 18 = 38, Company = 'Company B' (first value)")
    print("  Grand Total: Count = 120")
    
    print("\n" + "=" * 80)
    print("Test Results")
    print("=" * 80)
    
    # Check Issue State results
    issue_ca = next((r for r in issue_summary if r.get('Issue State') == 'CA'), None)
    issue_ny = next((r for r in issue_summary if r.get('Issue State') == 'NY'), None)
    issue_tx = next((r for r in issue_summary if r.get('Issue State') == 'TX'), None)
    issue_grand = next((r for r in issue_summary if r.get('Issue State') == 'Grand Total'), None)
    
    print("\nIssue State Verification:")
    if issue_ca and issue_ca.get('Count') == 30:
        print("  [PASS] CA count is correct (30)")
    else:
        print(f"  [FAIL] CA count is incorrect. Expected: 30, Got: {issue_ca.get('Count') if issue_ca else 'N/A'}")
    
    if issue_ny and issue_ny.get('Count') == 53:
        print("  [PASS] NY count is correct (53)")
    else:
        print(f"  [FAIL] NY count is incorrect. Expected: 53, Got: {issue_ny.get('Count') if issue_ny else 'N/A'}")
    
    if issue_tx and issue_tx.get('Count') == 37:
        print("  [PASS] TX count is correct (37)")
    else:
        print(f"  [FAIL] TX count is incorrect. Expected: 37, Got: {issue_tx.get('Count') if issue_tx else 'N/A'}")
    
    if issue_grand and issue_grand.get('Count') == 120:
        print("  [PASS] Grand Total count is correct (120)")
    else:
        print(f"  [FAIL] Grand Total count is incorrect. Expected: 120, Got: {issue_grand.get('Count') if issue_grand else 'N/A'}")
    
    # Check Company values
    if issue_ca and issue_ca.get('Company') == 'Company A':
        print("  [PASS] CA company is correct (Company A)")
    else:
        print(f"  [FAIL] CA company is incorrect. Expected: Company A, Got: {issue_ca.get('Company') if issue_ca else 'N/A'}")
    
    if issue_ny and issue_ny.get('Company') == 'Company B':
        print("  [PASS] NY company is correct (Company B)")
    else:
        print(f"  [FAIL] NY company is incorrect. Expected: Company B, Got: {issue_ny.get('Company') if issue_ny else 'N/A'}")
    
    if issue_tx and issue_tx.get('Company') == 'Company C':
        print("  [PASS] TX company is correct (Company C)")
    else:
        print(f"  [FAIL] TX company is incorrect. Expected: Company C, Got: {issue_tx.get('Company') if issue_tx else 'N/A'}")
    
    # Check Resident State results
    resident_ca = next((r for r in resident_summary if r.get('Resident State') == 'CA'), None)
    resident_ny = next((r for r in resident_summary if r.get('Resident State') == 'NY'), None)
    resident_tx = next((r for r in resident_summary if r.get('Resident State') == 'TX'), None)
    resident_grand = next((r for r in resident_summary if r.get('Resident State') == 'Grand Total'), None)
    
    print("\nResident State Verification:")
    if resident_ca and resident_ca.get('Count') == 32:
        print("  [PASS] CA count is correct (32)")
    else:
        print(f"  [FAIL] CA count is incorrect. Expected: 32, Got: {resident_ca.get('Count') if resident_ca else 'N/A'}")
    
    if resident_ny and resident_ny.get('Count') == 50:
        print("  [PASS] NY count is correct (50)")
    else:
        print(f"  [FAIL] NY count is incorrect. Expected: 50, Got: {resident_ny.get('Count') if resident_ny else 'N/A'}")
    
    if resident_tx and resident_tx.get('Count') == 38:
        print("  [PASS] TX count is correct (38)")
    else:
        print(f"  [FAIL] TX count is incorrect. Expected: 38, Got: {resident_tx.get('Count') if resident_tx else 'N/A'}")
    
    if resident_grand and resident_grand.get('Count') == 120:
        print("  [PASS] Grand Total count is correct (120)")
    else:
        print(f"  [FAIL] Grand Total count is incorrect. Expected: 120, Got: {resident_grand.get('Count') if resident_grand else 'N/A'}")
    
    # Check Company values for Resident State
    if resident_ca and resident_ca.get('Company') == 'Company A':
        print("  [PASS] CA company is correct (Company A)")
    else:
        print(f"  [FAIL] CA company is incorrect. Expected: Company A, Got: {resident_ca.get('Company') if resident_ca else 'N/A'}")
    
    if resident_ny and resident_ny.get('Company') == 'Company A':
        print("  [PASS] NY company is correct (Company A)")
    else:
        print(f"  [FAIL] NY company is incorrect. Expected: Company A, Got: {resident_ny.get('Company') if resident_ny else 'N/A'}")
    
    if resident_tx and resident_tx.get('Company') == 'Company B':
        print("  [PASS] TX company is correct (Company B)")
    else:
        print(f"  [FAIL] TX company is incorrect. Expected: Company B, Got: {resident_tx.get('Company') if resident_tx else 'N/A'}")
    
    print("\n" + "=" * 80)
    print("Test Complete")
    print("=" * 80)

if __name__ == '__main__':
    test_state_summary_with_company()

