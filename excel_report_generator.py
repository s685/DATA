#!/usr/bin/env python3
"""
Multi-Worksheet Excel Report Generator from Snowflake
Generates Excel workbooks matching template structure with detail and summary tables.
All worksheet structures are hardcoded based on template images.
"""

import argparse
import sys
import os
import yaml
from collections import defaultdict
from dataclasses import dataclass
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime

try:
    import snowflake.connector
except ImportError:
    print("Error: snowflake-connector-python not installed. Run: pip install snowflake-connector-python")
    sys.exit(1)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:
    print("Error: openpyxl not installed. Run: pip install openpyxl")
    sys.exit(1)


# ============================================================================
# Configuration Classes
# ============================================================================

@dataclass
class SnowflakeConfig:
    """Snowflake connection configuration"""
    account: str
    user: Optional[str] = None
    password: Optional[str] = None
    warehouse: str = ""
    database: str = ""
    schema: str = ""
    authenticator: str = "externalbrowser"  # 'externalbrowser' for SSO or 'snowflake' for user/pass


@dataclass
class AggregateConfig:
    """Aggregate function configuration"""
    field: str
    function: str  # COUNT, SUM, AVG, etc.
    label: str


@dataclass
class SummaryConfig:
    """Summary table configuration"""
    group_by: str
    aggregates: List[AggregateConfig]
    start_column: str
    columns: List[str]


@dataclass
class FormattingConfig:
    """Formatting configuration"""
    header_row: int = 1
    filters: bool = True
    highlight_columns: List[str] = None
    highlight_cells: List[str] = None


@dataclass
class WorksheetConfig:
    """Worksheet configuration"""
    name: str
    query: str
    detail_start_column: str = "A"
    detail_columns: List[str] = None
    spacing_columns: List[str] = None
    summary_config: List[SummaryConfig] = None
    formatting: FormattingConfig = None
    layout_type: str = "table"  # "table" or "form"
    form_layout: Dict[str, Any] = None


# ============================================================================
# Hardcoded Worksheet Definitions (Structure only - table names come from config)
# ============================================================================

def get_hardcoded_worksheet_structure(worksheet_name: str, table_name: str) -> Optional[WorksheetConfig]:
    """Return hardcoded worksheet structure based on template images. Table name is injected into query."""
    
    # Helper function to create standard worksheet configs
    def create_standard_worksheet(name, schedule_id, table_name, has_summary=True, highlight_cols=None):
        query = f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '{schedule_id}'"
        detail_cols = ['Policy Num', 'Claim Num', 'Product', 'Claim Status', 'Company', 'Issue Sta', 'Resident Sta']
        formatting = FormattingConfig(header_row=1, filters=True, highlight_columns=highlight_cols or [])
        
        summary_configs = None
        if has_summary:
            summary_configs = [
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='I',
                    columns=['Issue State', 'CountOfPolicy No']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='K',
                    columns=['Resident State', 'CountOfPolicy No']
                )
            ]
        
        return WorksheetConfig(
            name=name,
            query=query,
            detail_start_column='A',
            detail_columns=detail_cols,
            spacing_columns=['H'],
            summary_config=summary_configs,
            formatting=formatting
        )
    
    # Summary Worksheet - handled separately, not returned here
    if worksheet_name == 'Summary':
        return None  # Summary is handled separately in create_workbook
    
    # Worksheet 1-001 - Summary only (no detail records)
    elif worksheet_name == '1-001':
        return WorksheetConfig(
            name='1-001',
            query=f"SELECT Policy_Num, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '1-001'",
            detail_start_column='A',
            detail_columns=None,  # None means use actual column names from query results
            spacing_columns=[],
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='A',  # Start from column A
                    columns=['Issue State', 'CountOfPolicy No']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='D',  # One column gap after Issue State summary (columns A-B, gap C, start D)
                    columns=['Resident State', 'CountOfPolicy No']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=False)  # No filters for summary-only
        )
    
    # Worksheet 1-004 - Detail records only (no summaries)
    elif worksheet_name == '1-004':
        return WorksheetConfig(
            name='1-004',
            query=f"SELECT Policy, Lapse_Da, Stati, Status_Reas, Company, Issue_St, Resident_St FROM {table_name} WHERE Schedule_ID = '1-004'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=[],
            summary_config=None,  # No summaries - detail only
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 1-006 - Summary only (no detail records)
    elif worksheet_name == '1-006':
        return WorksheetConfig(
            name='1-006',
            query=f"SELECT Policy_Num, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '1-006'",
            detail_start_column='A',
            detail_columns=None,  # No detail columns - summary only
            spacing_columns=[],
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='A',  # Start from column A
                    columns=['Issue State', 'CountOfPolicy No']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='D',  # One column gap after Issue State summary (columns A-B, gap C, start D)
                    columns=['Resident State', 'CountOfPolicy No']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=False)  # No filters for summary-only
        )
    
    # Worksheet 2-001 - Detail records + Issue State/Resident State summaries
    elif worksheet_name == '2-001':
        return WorksheetConfig(
            name='2-001',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '2-001'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['H'],  # One column gap between detail (A-G) and summaries
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='I',  # Start after gap column H (Issue State: I-J)
                    columns=['Issue State', 'Count']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='L',  # One column gap after Issue State summary (gap K, Resident State: L-M)
                    columns=['Resident State', 'Count']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 2-002 - Same as 2-001
    elif worksheet_name == '2-002':
        return WorksheetConfig(
            name='2-002',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '2-002'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['H'],  # One column gap between detail (A-G) and summaries
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='I',  # Issue State: I-J
                    columns=['Issue State', 'Count']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='L',  # Gap K, Resident State: L-M
                    columns=['Resident State', 'Count']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 2-003 - Detail records + TAT summary by ranges with % of total
    elif worksheet_name == '2-003':
        return WorksheetConfig(
            name='2-003',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State, TAT_in_Days FROM {table_name} WHERE Schedule_ID = '2-003'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['I'],  # One column gap between detail (A-H) and TAT summary
            summary_config=[
                SummaryConfig(
                    group_by='TAT_Range',  # TAT summary by ranges (special handling)
                    aggregates=[AggregateConfig(field='TAT_in_Days', function='COUNT', label='TAT COUNTS')],
                    start_column='J',  # Start after gap column I (TAT Range: J-K-L)
                    columns=['', 'TAT COUNTS', '% of Total']  # Empty first column header (range values are in data rows)
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 2-004 - Same as 2-001
    elif worksheet_name == '2-004':
        return WorksheetConfig(
            name='2-004',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '2-004'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['H'],  # One column gap between detail (A-G) and summaries
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='I',  # Issue State: I-J
                    columns=['Issue State', 'Count']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='L',  # Gap K, Resident State: L-M
                    columns=['Resident State', 'Count']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 2-005 - Same as 2-001
    elif worksheet_name == '2-005':
        return WorksheetConfig(
            name='2-005',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '2-005'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['H'],  # One column gap between detail (A-G) and summaries
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='I',  # Issue State: I-J
                    columns=['Issue State', 'Count']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='L',  # Gap K, Resident State: L-M
                    columns=['Resident State', 'Count']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 3-001 - Detail records + TAT summary (same pattern as 2-003)
    elif worksheet_name == '3-001':
        return WorksheetConfig(
            name='3-001',
            query=f"SELECT Inq_Date, Stat_Start_Date, Decision, Decision_Reason, Company, Issue_State, Resident_State, Product, Date_of_Loss, Schedule_ID, TAT_in_Days FROM {table_name} WHERE Schedule_ID = '3-001'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['L'],  # One column gap between detail (A-K) and TAT summary
            summary_config=[
                SummaryConfig(
                    group_by='TAT_Range',  # TAT summary by ranges (same as 2-003)
                    aggregates=[AggregateConfig(field='TAT_in_Days', function='COUNT', label='TAT COUNTS')],
                    start_column='M',  # Start after gap column L (TAT Range: M-N-O)
                    columns=['', 'TAT COUNTS', '% of Total']  # Empty first column header
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True, highlight_columns=['C', 'D'])
        )
    
    # Worksheets 3-003, 3-004, 3-005 - Decision Details
    elif worksheet_name in ['3-003', '3-004', '3-005']:
        schedule_id = worksheet_name
        return WorksheetConfig(
            name=worksheet_name,
            query=f"SELECT Decision_Date, Inq_Date, Stat_Start_Date, Decision, Decision_Reason, Company, Issue_State, Resident_State, Product, Date_of_Loss, Schedule_ID, TAT_in_Days FROM {table_name} WHERE Schedule_ID = '{schedule_id}'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=[],
            summary_config=None,
            formatting=FormattingConfig(header_row=1, filters=True, highlight_columns=['D', 'E'])
        )
    
    # Worksheet 3-006
    elif worksheet_name == '3-006':
        return WorksheetConfig(
            name='3-006',
            query=f"SELECT Inq_Date, Stat_Start_Date, Decision, Decision_Reason, Company, Issue_State, Resident_State, Product, Date_of_Loss, Schedule_ID, TAT_in_Days FROM {table_name} WHERE Schedule_ID = '3-006'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=[],
            summary_config=None,
            formatting=FormattingConfig(header_row=1, filters=True, highlight_columns=['C', 'D'])
        )
    
    # Worksheet 3-007
    elif worksheet_name == '3-007':
        return WorksheetConfig(
            name='3-007',
            query=f"SELECT Policy, Claim_Number, Decision_Date, Inq_Date, Stat_Start_Date, Decision, Decision_Reason, Company, Issue_State, Resident_State, Product, Date_of_Loss, Schedule_ID_2 FROM {table_name} WHERE Schedule_ID = '3-007'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=[],
            summary_config=None,
            formatting=FormattingConfig(header_row=1, filters=True, highlight_columns=['F', 'G'])
        )
    
    # Worksheet 5-001 - Detail records + Issue State/Resident State summaries (like 2-002)
    elif worksheet_name == '5-001':
        return WorksheetConfig(
            name='5-001',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '5-001'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['H'],  # One column gap between detail (A-G) and summaries
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='I',  # Issue State: I-J
                    columns=['Issue State', 'CountOfPolicy No']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='L',  # Gap K, Resident State: L-M
                    columns=['Resident State', 'CountOfPolicy No']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 5-002 - Summary only (no detail records): Issue State/Resident State summaries with Company
    elif worksheet_name == '5-002':
        # Summary only: Issue State, Count, Company & Resident State, Count, Company
        return WorksheetConfig(
            name='5-002',
            query=f"SELECT Issue_State, Resident_State, Company, Count FROM {table_name} WHERE Schedule_ID = '5-002'",
            detail_start_column='A',
            detail_columns=None,  # No detail columns - summary only
            spacing_columns=[],  # No detail records, so no spacing needed
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[
                        AggregateConfig(field='Count', function='SUM', label='Count'),
                        AggregateConfig(field='Company', function='COUNT', label='Company')  # Count of distinct companies per state
                    ],
                    start_column='A',  # Issue State: A-B-C
                    columns=['Issue State', 'Count', 'Company']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[
                        AggregateConfig(field='Count', function='SUM', label='Count'),
                        AggregateConfig(field='Company', function='COUNT', label='Company')  # Count of distinct companies per state
                    ],
                    start_column='E',  # Gap D, Resident State: E-F-G
                    columns=['Resident State', 'Count', 'Company']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=False)  # No filters for summary-only
        )
    
    # Worksheet 5-003 - Detail records + TAT summary + Issue State/Resident State summaries
    elif worksheet_name == '5-003':
        return WorksheetConfig(
            name='5-003',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State, TAT_in_Days FROM {table_name} WHERE Schedule_ID = '5-003'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['I'],  # One column gap between detail (A-H) and summaries
            summary_config=[
                SummaryConfig(
                    group_by='TAT_Range',  # TAT summary by ranges (same as 2-003)
                    aggregates=[AggregateConfig(field='TAT_in_Days', function='COUNT', label='TAT COUNTS')],
                    start_column='J',  # TAT Range: J-K-L
                    columns=['', 'TAT COUNTS', '% of Total']
                ),
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='N',  # Gap M, Issue State: N-O
                    columns=['Issue State', 'CountOfPolicy No']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='Q',  # Gap P, Resident State: Q-R (one column gap after Issue State summary)
                    columns=['Resident State', 'CountOfPolicy No']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 5-004 - Detail records + Issue State/Resident State summaries
    elif worksheet_name == '5-004':
        return WorksheetConfig(
            name='5-004',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '5-004'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['H'],  # One column gap between detail (A-G) and summaries
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='I',  # Issue State: I-J
                    columns=['Issue State', 'CountOfPolicy No']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='CountOfPolicy No')],
                    start_column='L',  # Gap K, Resident State: L-M
                    columns=['Resident State', 'CountOfPolicy No']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 6-001 - Summary only (no detail records): Like 5-002 (Issue State, Count, Company & Resident State, Count, Company)
    elif worksheet_name == '6-001':
        return WorksheetConfig(
            name='6-001',
            query=f"SELECT Issue_State, Resident_State, Company, Count FROM {table_name} WHERE Schedule_ID = '6-001'",
            detail_start_column='A',
            detail_columns=None,  # No detail columns - summary only
            spacing_columns=[],  # No detail records, so no spacing needed
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[
                        AggregateConfig(field='Count', function='SUM', label='Count'),
                        AggregateConfig(field='Company', function='COUNT', label='Company')
                    ],
                    start_column='A',  # Issue State: A-B-C
                    columns=['Issue State', 'Count', 'Company']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[
                        AggregateConfig(field='Count', function='SUM', label='Count'),
                        AggregateConfig(field='Company', function='COUNT', label='Company')
                    ],
                    start_column='E',  # Gap D, Resident State: E-F-G
                    columns=['Resident State', 'Count', 'Company']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=False)  # No filters for summary-only
        )
    
    # Worksheet 6-002 - Detail records + Issue State/Resident State summaries (Count, not CountOfPolicy No)
    elif worksheet_name == '6-002':
        return WorksheetConfig(
            name='6-002',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '6-002'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['H'],  # One column gap between detail (A-G) and summaries
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='I',  # Issue State: I-J
                    columns=['Issue State', 'Count']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='L',  # Gap K, Resident State: L-M
                    columns=['Resident State', 'Count']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 6-003 - Same as 6-002
    elif worksheet_name == '6-003':
        return WorksheetConfig(
            name='6-003',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State FROM {table_name} WHERE Schedule_ID = '6-003'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['H'],  # One column gap between detail (A-G) and summaries
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='I',  # Issue State: I-J
                    columns=['Issue State', 'Count']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='L',  # Gap K, Resident State: L-M
                    columns=['Resident State', 'Count']
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    # Worksheet 6-004 - Same as 6-003 + Additional Summary (Counts, Year Pay Req Received)
    elif worksheet_name == '6-004':
        return WorksheetConfig(
            name='6-004',
            query=f"SELECT Policy_Num, Claim_Num, Product, Claim_Status, Company, Issue_State, Resident_State, Year_Pay_Req_Received FROM {table_name} WHERE Schedule_ID = '6-004'",
            detail_start_column='A',
            detail_columns=None,  # Use actual column names from query results
            spacing_columns=['I'],  # One column gap between detail (A-H) and summaries
            summary_config=[
                SummaryConfig(
                    group_by='Issue_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='J',  # Issue State: J-K
                    columns=['Issue State', 'Count']
                ),
                SummaryConfig(
                    group_by='Resident_State',
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Count')],
                    start_column='M',  # Gap L, Resident State: M-N
                    columns=['Resident State', 'Count']
                ),
                SummaryConfig(
                    group_by='Year_Pay_Req_Received',  # Additional summary by Year
                    aggregates=[AggregateConfig(field='Policy_Num', function='COUNT', label='Counts')],
                    start_column='P',  # Gap O, Year summary: P-Q (Year Pay Req Received, Counts)
                    columns=['Year Pay Req Received', 'Counts']  # Year value first, then Counts
                )
            ],
            formatting=FormattingConfig(header_row=1, filters=True)
        )
    
    return None


# ============================================================================
# Snowflake Connection Functions
# ============================================================================

def create_snowflake_connection(config: SnowflakeConfig) -> snowflake.connector.SnowflakeConnection:
    """Create Snowflake connection with SSO or username/password"""
    try:
        conn_params = {
            'account': config.account,
            'warehouse': config.warehouse,
            'database': config.database,
            'schema': config.schema,
            'authenticator': config.authenticator
        }
        
        if config.authenticator == 'snowflake':
            if not config.user or not config.password:
                raise ValueError("User and password required for 'snowflake' authenticator")
            conn_params['user'] = config.user
            conn_params['password'] = config.password
        
        connection = snowflake.connector.connect(**conn_params)
        print(f"Successfully connected to Snowflake account: {config.account}")
        return connection
    except Exception as e:
        print(f"Error connecting to Snowflake: {e}")
        sys.exit(1)


def execute_query(connection: snowflake.connector.SnowflakeConnection, query: str) -> List[Dict[str, Any]]:
    """Execute SQL query and return results as list of dictionaries"""
    try:
        cursor = connection.cursor()
        cursor.execute(query)
        
        # Get column names
        columns = [desc[0] for desc in cursor.description]
        
        # Fetch all rows
        rows = cursor.fetchall()
        
        # Convert to list of dictionaries
        results = []
        for row in rows:
            row_dict = {}
            for i, col in enumerate(columns):
                row_dict[col] = row[i]
            results.append(row_dict)
        
        cursor.close()
        return results
    except Exception as e:
        print(f"Error executing query: {e}")
        print(f"Query: {query[:200]}...")
        raise


def close_connection(connection: snowflake.connector.SnowflakeConnection):
    """Close Snowflake connection"""
    try:
        connection.close()
        print("Snowflake connection closed")
    except Exception as e:
        print(f"Error closing connection: {e}")


# ============================================================================
# Data Processing Functions
# ============================================================================

def fetch_detail_records(connection: snowflake.connector.SnowflakeConnection, query: str) -> List[Dict[str, Any]]:
    """Fetch detail records from Snowflake"""
    return execute_query(connection, query)


def get_tat_range(tat_value: Any) -> str:
    """Categorize TAT value into range buckets"""
    try:
        tat = float(tat_value) if tat_value is not None else 0
        if -1 <= tat < 31:  # -1 to <31 (i.e., -1 to 30)
            return "-1 to <31"
        elif 31 <= tat < 61:  # >30 and <61 (i.e., 31 to 60)
            return ">30 and <61"
        elif 61 <= tat < 91:  # >60 and <91 (i.e., 61 to 90)
            return ">60 and <91"
        elif tat >= 91:  # >90 (i.e., 91 and above)
            return ">90"
        else:
            return None  # Exclude from summary
    except (ValueError, TypeError):
        return None  # Exclude from summary


def generate_summary(detail_records: List[Dict[str, Any]], summary_config: SummaryConfig, include_grand_total: bool = True) -> List[Dict[str, Any]]:
    """Generate summary table from detail records with optional Grand Total row and percentage calculation"""
    # Group records by the specified field
    grouped = defaultdict(list)
    
    # Special handling for TAT_Range grouping
    if summary_config.group_by == 'TAT_Range':
        for record in detail_records:
            # Get TAT_in_Days value and categorize into range
            tat_value = record.get('TAT_in_Days', None)
            group_key = get_tat_range(tat_value)
            # Include all records with valid TAT ranges (including -1 to <31)
            if group_key is not None:
                grouped[group_key].append(record)
    else:
        # Standard grouping by field value
        for record in detail_records:
            group_key = record.get(summary_config.group_by, '')
            grouped[group_key].append(record)
    
    # Calculate aggregates for each group
    summary_rows = []
    grand_total_values = {}
    total_count = 0  # For percentage calculation
    
    # Define TAT range order for proper sorting
    tat_range_order = ["-1 to <31", ">30 and <61", ">60 and <91", ">90"]
    
    # Sort groups - use TAT range order if grouping by TAT_Range, otherwise alphabetical
    if summary_config.group_by == 'TAT_Range':
        sorted_groups = sorted(grouped.items(), key=lambda x: (
            tat_range_order.index(x[0]) if x[0] in tat_range_order else len(tat_range_order)
        ))
    else:
        sorted_groups = sorted(grouped.items())
    
    # First pass: calculate counts and totals
    for group_key, records in sorted_groups:
        # For TAT_Range grouping, use empty string for first column (users know row position = category)
        if summary_config.group_by == 'TAT_Range':
            summary_row = {summary_config.columns[0]: ''}  # Empty first column - users know row position
        else:
            summary_row = {summary_config.columns[0]: group_key}
        
        for agg in summary_config.aggregates:
            field_value = agg.field
            function = agg.function.upper()
            
            if function == 'COUNT':
                value = len(records)
                total_count += value  # Accumulate total for percentage
            elif function == 'SUM':
                value = sum(float(r.get(field_value, 0) or 0) for r in records)
            elif function == 'AVG' or function == 'AVERAGE':
                values = [float(r.get(field_value, 0) or 0) for r in records if r.get(field_value) is not None]
                value = sum(values) / len(values) if values else 0
            elif function == 'MIN':
                values = [r.get(field_value) for r in records if r.get(field_value) is not None]
                value = min(values) if values else None
            elif function == 'MAX':
                values = [r.get(field_value) for r in records if r.get(field_value) is not None]
                value = max(values) if values else None
            else:
                value = 0
            
            summary_row[agg.label] = value
            
            # Accumulate for grand total
            if include_grand_total:
                if agg.label not in grand_total_values:
                    grand_total_values[agg.label] = 0
                if isinstance(value, (int, float)):
                    grand_total_values[agg.label] += value
        
        summary_rows.append(summary_row)
    
    # Second pass: calculate percentages if '% of Total' column exists
    if '% of Total' in summary_config.columns and total_count > 0:
        for summary_row in summary_rows:
            # Find the count value (should be the aggregate label)
            count_value = None
            for agg in summary_config.aggregates:
                if agg.label in summary_row:
                    count_value = summary_row[agg.label]
                    break
            
            if count_value is not None and isinstance(count_value, (int, float)):
                percentage = (count_value / total_count) * 100
                # Format with % symbol
                summary_row['% of Total'] = f"{round(percentage, 2)}%"
            else:
                summary_row['% of Total'] = "0%"
    
    # Add Grand Total row if requested
    if include_grand_total and grand_total_values:
        # For TAT_Range, first column should be empty (no "Grand Total" label)
        if summary_config.group_by == 'TAT_Range':
            grand_total_row = {summary_config.columns[0]: ''}  # Empty for TAT range
        else:
            grand_total_row = {summary_config.columns[0]: 'Grand Total'}
        for agg in summary_config.aggregates:
            grand_total_row[agg.label] = grand_total_values.get(agg.label, 0)
        
        # Add percentage for grand total (should be 100%)
        if '% of Total' in summary_config.columns:
            grand_total_row['% of Total'] = "100%"
        
        summary_rows.append(grand_total_row)
    
    return summary_rows


def process_worksheet_data(connection: snowflake.connector.SnowflakeConnection, 
                           worksheet_config: WorksheetConfig) -> Tuple[List[Dict[str, Any]], List[List[Dict[str, Any]]]]:
    """Process worksheet data: fetch details and generate summaries"""
    # Fetch detail records
    detail_records = fetch_detail_records(connection, worksheet_config.query)
    
    # Generate summaries if configured
    summaries = []
    if worksheet_config.summary_config:
        for sum_config in worksheet_config.summary_config:
            # Include Grand Total for Schedule 1 worksheets (1-001, 1-004, 1-006)
            include_grand_total = worksheet_config.name in ['1-001', '1-004', '1-006']
            summary_data = generate_summary(detail_records, sum_config, include_grand_total=include_grand_total)
            summaries.append(summary_data)
    
    return detail_records, summaries


# ============================================================================
# Excel Generation Functions
# ============================================================================

def column_letter_to_index(column_letter: str) -> int:
    """Convert column letter (e.g., 'A', 'Z', 'AA') to 1-based index"""
    result = 0
    for char in column_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def index_to_column_letter(index: int) -> str:
    """Convert 1-based index to column letter"""
    return get_column_letter(index)


def apply_cell_formatting(ws, row: int, col: int, value: Any, is_header: bool = False):
    """Apply formatting to a cell"""
    cell = ws.cell(row=row, column=col)
    cell.value = value
    
    if is_header:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        # Apply Orange Accent 2 Light 60% fill (approximately #FFD966)
        orange_fill = PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')
        cell.fill = orange_fill
    else:
        cell.alignment = Alignment(vertical='top')
    
    # Apply borders
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    cell.border = thin_border


def apply_highlight(ws, row: int, col: int):
    """Apply yellow highlight to a cell"""
    cell = ws.cell(row=row, column=col)
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    cell.fill = yellow_fill

def apply_border(ws, row: int, col: int):
    """Apply thin border to a cell"""
    cell = ws.cell(row=row, column=col)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    cell.border = thin_border


def write_detail_table(ws, detail_records: List[Dict[str, Any]], 
                      worksheet_config: WorksheetConfig, start_row: int = 1):
    """Write detail table to worksheet"""
    if not detail_records:
        return start_row
    
    # Get column names from first record or use configured columns
    if worksheet_config.detail_columns and len(worksheet_config.detail_columns) > 0:
        # Use configured display column names
        columns = worksheet_config.detail_columns
        # Map display names to actual column names from query results
        actual_columns = list(detail_records[0].keys())
        column_map = {}
        for i, display_col in enumerate(columns):
            if i < len(actual_columns):
                # Try to match by position first
                column_map[display_col] = actual_columns[i]
            else:
                # If no match by position, try to find by name (case-insensitive)
                found = False
                for actual_col in actual_columns:
                    if actual_col.lower().replace('_', ' ') == display_col.lower():
                        column_map[display_col] = actual_col
                        found = True
                        break
                if not found:
                    column_map[display_col] = display_col
    else:
        # Use actual column names from query results
        columns = list(detail_records[0].keys())
        column_map = {col: col for col in columns}
    
    # Write headers
    start_col_idx = column_letter_to_index(worksheet_config.detail_start_column)
    header_row = worksheet_config.formatting.header_row if worksheet_config.formatting else start_row
    
    for col_idx, col_name in enumerate(columns):
        apply_cell_formatting(ws, header_row, start_col_idx + col_idx, col_name, is_header=True)
    
    # Write data rows
    data_start_row = header_row + 1
    for row_idx, record in enumerate(detail_records):
        for col_idx, col_name in enumerate(columns):
            actual_col = column_map.get(col_name, col_name)
            value = record.get(actual_col, '')
            cell = ws.cell(row=data_start_row + row_idx, column=start_col_idx + col_idx)
            cell.value = value
            cell.alignment = Alignment(vertical='top')
            apply_border(ws, data_start_row + row_idx, start_col_idx + col_idx)
            
            # Apply highlighting if configured
            col_letter = index_to_column_letter(start_col_idx + col_idx)
            if worksheet_config.formatting:
                if col_letter in (worksheet_config.formatting.highlight_columns or []):
                    apply_highlight(ws, data_start_row + row_idx, start_col_idx + col_idx)
    
    # Apply filters if configured
    if worksheet_config.formatting and worksheet_config.formatting.filters:
        end_col_idx = start_col_idx + len(columns) - 1
        ws.auto_filter.ref = f"{index_to_column_letter(start_col_idx)}{header_row}:{index_to_column_letter(end_col_idx)}{data_start_row + len(detail_records)}"
    
    return data_start_row + len(detail_records)


def write_summary_table(ws, summary_data: List[Dict[str, Any]], 
                        summary_config: SummaryConfig, start_row: int = 1):
    """Write summary table to worksheet with Grand Total row formatting"""
    if not summary_data:
        return
    
    start_col_idx = column_letter_to_index(summary_config.start_column)
    header_row = start_row
    
    # Write headers
    for col_idx, col_name in enumerate(summary_config.columns):
        apply_cell_formatting(ws, header_row, start_col_idx + col_idx, col_name, is_header=True)
    
    # Write data rows
    data_start_row = header_row + 1
    for row_idx, summary_row in enumerate(summary_data):
        is_grand_total = summary_row.get(summary_config.columns[0]) == 'Grand Total'
        
        for col_idx, col_name in enumerate(summary_config.columns):
            value = summary_row.get(col_name, '')
            cell = ws.cell(row=data_start_row + row_idx, column=start_col_idx + col_idx)
            cell.value = value
            
            # Format Grand Total row (bold, but not orange header)
            if is_grand_total:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='right' if col_idx > 0 else 'left', vertical='top')
            else:
                cell.alignment = Alignment(horizontal='right' if col_idx > 0 else 'left', vertical='top')
            
            # Apply borders
            apply_border(ws, data_start_row + row_idx, start_col_idx + col_idx)
        
        # Add Grand Total value in the next cell after CountOfPolicy No (if this is the last summary column)
        if is_grand_total and col_idx == len(summary_config.columns) - 1:
            # The Grand Total value is already in the CountOfPolicy No column
            # But if we need to add it in a separate cell, we can do it here
            pass
    
    return data_start_row + len(summary_data)


def format_date_for_reporting_period(date_str: str) -> str:
    """
    Format date string into format like "January 1, 2024"
    Accepts formats: YYYY-MM-DD, MM/DD/YYYY, DD-MM-YYYY, etc.
    """
    from datetime import datetime
    
    # Try common date formats
    date_formats = [
        '%Y-%m-%d',      # 2024-01-01
        '%m/%d/%Y',      # 01/01/2024
        '%d/%m/%Y',      # 01/01/2024 (European)
        '%Y/%m/%d',      # 2024/01/01
        '%m-%d-%Y',      # 01-01-2024
        '%d-%m-%Y',      # 01-01-2024 (European)
        '%B %d, %Y',     # January 1, 2024
        '%b %d, %Y',     # Jan 1, 2024
    ]
    
    for fmt in date_formats:
        try:
            dt = datetime.strptime(date_str.strip(), fmt)
            # Format as "January 1, 2024" (remove leading zero from day)
            day = dt.day
            month_name = dt.strftime('%B')
            year = dt.year
            return f"{month_name} {day}, {year}"
        except ValueError:
            continue
    
    # If no format matches, return as-is
    return date_str


def create_summary_worksheet(wb, summary_data, schedule_titles=None, reporting_period=None):
    """
    Create Summary worksheet with all 6 schedules, Data Source column, proper spacing
    summary_data: List of dicts with keys: Schedule_ID, Description, Value
    schedule_titles: Dict mapping schedule number to title
    reporting_period: String like "January 1, 2024 through December 31, 2024" or None to use default
    """
    ws = wb.create_sheet(title="Summary")
    
    # Font: Aptos Narrow, size 12
    default_font = Font(name='Aptos Narrow', size=12)
    bold_font = Font(name='Aptos Narrow', size=12, bold=True)
    
    # Default schedule titles if not provided
    if schedule_titles is None:
        schedule_titles = {
            1: "Schedule 1 - General Information",
            2: "Schedule 2 - Claimants",
            3: "Schedule 3 - Claimant Requests Denied/Not Paid",
            4: "Schedule 4",
            5: "Schedule 5",
            6: "Schedule 6"
        }
    
    # Default reporting period if not provided
    if reporting_period is None:
        reporting_period = "January 1, 2024 through December 31, 2024"
    
    # Leave first row empty (row 1)
    # Header Section - Start from row 2
    ws['A2'] = "Line of Business: Individual Long-Term Care"
    ws['A3'] = f"Reporting Period: {reporting_period}"
    ws['A4'] = "Filing Deadline: n/a"
    
    # Apply fonts and alignment (left aligned)
    ws['A2'].font = bold_font
    ws['A2'].alignment = Alignment(horizontal='left', vertical='top')
    
    ws['A3'].font = bold_font
    ws['A3'].alignment = Alignment(horizontal='left', vertical='top')
    apply_highlight(ws, 3, 1)  # Yellow highlight
    
    ws['A4'].font = bold_font
    ws['A4'].alignment = Alignment(horizontal='left', vertical='top')
    apply_highlight(ws, 4, 1)  # Yellow highlight
    
    # Group data by schedule number (1-xxx, 2-xxx, 3-xxx, etc.)
    schedule_groups = defaultdict(list)
    for row in summary_data:
        schedule_id = row.get('Schedule_ID', '')
        if '-' in schedule_id:
            schedule_num = int(schedule_id.split('-')[0])
            schedule_groups[schedule_num].append(row)
    
    # Sort groups by schedule number
    sorted_groups = sorted(schedule_groups.items())
    
    # Starting row for first schedule - row 6 (after header rows 2-4, with 1 empty row gap)
    current_row = 6
    
    for schedule_num, rows in sorted_groups:
        # Sort rows by Schedule_ID
        rows.sort(key=lambda x: x.get('Schedule_ID', ''))
        
        # Write Schedule title - left aligned, column A
        title_row = current_row
        title_text = schedule_titles.get(schedule_num, f"Schedule {schedule_num}")
        title_cell = ws.cell(row=title_row, column=1)
        title_cell.value = title_text
        title_cell.font = bold_font
        title_cell.alignment = Alignment(horizontal='left', vertical='top')
        
        # Write headers - Column A: ID, Column C: Description (B is gap), Column E: Value (D is gap), Column F: Data Source
        header_row = title_row + 1
        id_header = ws.cell(row=header_row, column=1)
        id_header.value = "ID"
        id_header.font = bold_font
        id_header.alignment = Alignment(horizontal='left', vertical='top')
        apply_border(ws, header_row, 1)
        
        desc_header = ws.cell(row=header_row, column=3)  # Column C (B is gap)
        desc_header.value = "Description"
        desc_header.font = bold_font
        desc_header.alignment = Alignment(horizontal='left', vertical='top')
        apply_border(ws, header_row, 3)
        
        value_header = ws.cell(row=header_row, column=5)  # Column E (D is gap)
        value_header.value = "Value"
        value_header.font = bold_font
        value_header.alignment = Alignment(horizontal='right', vertical='top')
        apply_border(ws, header_row, 5)
        
        datasource_header = ws.cell(row=header_row, column=6)  # Column F: Data Source
        datasource_header.value = "Data Source"
        datasource_header.font = bold_font
        datasource_header.alignment = Alignment(horizontal='left', vertical='top')
        apply_border(ws, header_row, 6)
        
        # Write data rows - 1 row after headers
        data_start_row = header_row + 1
        for idx, row_data in enumerate(rows):
            row_num = data_start_row + idx
            
            # Column A: Schedule_ID (e.g., "2-001")
            schedule_id = row_data.get('Schedule_ID', '')
            id_cell = ws.cell(row=row_num, column=1)
            id_cell.value = schedule_id
            id_cell.font = default_font
            id_cell.alignment = Alignment(horizontal='left', vertical='top')
            apply_border(ws, row_num, 1)
            
            # Column C: Description (Column B is gap)
            description = row_data.get('Description', '')
            desc_cell = ws.cell(row=row_num, column=3)
            desc_cell.value = description
            desc_cell.font = default_font
            desc_cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            apply_border(ws, row_num, 3)
            
            # Column E: Value (Column D is gap, highlighted in yellow)
            value = row_data.get('Value', '')
            value_cell = ws.cell(row=row_num, column=5)
            
            # Format numeric values with thousand separators (commas) - US format: 101,037
            if isinstance(value, (int, float)):
                # Format as string with US-style comma separators to ensure correct display
                # This ensures 101037 displays as 101,037 (not 1,01,037)
                value_cell.value = f"{int(value):,}" if isinstance(value, float) and value.is_integer() else f"{value:,}"
            else:
                # For non-numeric values like "N/A", keep as string
                value_cell.value = value
            
            value_cell.font = default_font
            value_cell.alignment = Alignment(horizontal='right', vertical='top')
            apply_border(ws, row_num, 5)
            apply_highlight(ws, row_num, 5)  # Yellow highlight
            
            # Column F: Data Source - always "Snowflake"
            datasource_cell = ws.cell(row=row_num, column=6)
            datasource_cell.value = "Snowflake"
            datasource_cell.font = default_font
            datasource_cell.alignment = Alignment(horizontal='left', vertical='top')
            apply_border(ws, row_num, 6)
        
        # Move to next schedule section - leave 1 empty row between schedules
        current_row = data_start_row + len(rows) + 2  # +2 for empty row gap
    
    # Adjust column widths for better appearance
    ws.column_dimensions['A'].width = 15  # ID column
    ws.column_dimensions['B'].width = 3   # Gap column (narrow)
    ws.column_dimensions['C'].width = 70  # Description column
    ws.column_dimensions['D'].width = 3   # Gap column (narrow)
    ws.column_dimensions['E'].width = 18  # Value column
    ws.column_dimensions['F'].width = 20  # Data Source column
    
    # Set row heights
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 20
    
    return ws


def create_worksheet(wb: Workbook, worksheet_config: WorksheetConfig,
                    detail_records: List[Dict[str, Any]], summaries: List[List[Dict[str, Any]]]):
    """Create a worksheet with detail and summary tables"""
    ws = wb.create_sheet(title=worksheet_config.name)
    
    # Note: Summary worksheet is handled separately in create_workbook
    # This function is only for detail worksheets
    if False:  # Placeholder - Summary is handled separately
        pass
    else:
        # Standard table layout
        # Determine header row
        header_row = worksheet_config.formatting.header_row if worksheet_config.formatting else 1
        
        # Write detail table
        last_detail_row = write_detail_table(ws, detail_records, worksheet_config, start_row=header_row)
        
        # Handle spacing columns (spacing is handled by column positioning in summary_config)
        
        # Write summary tables (aligned with detail table header row)
        if worksheet_config.summary_config and summaries:
            for summary_config, summary_data in zip(worksheet_config.summary_config, summaries):
                write_summary_table(ws, summary_data, summary_config, start_row=header_row)
    
    # Adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width


def aggregate_all_worksheets_data(all_worksheets_data: Dict[str, Tuple[List[Dict[str, Any]], List[List[Dict[str, Any]]]]]) -> Dict[str, Any]:
    """Aggregate data from all detail worksheets for Summary sheet"""
    aggregated = {
        'all_records': [],
        'issue_state_counts': defaultdict(int),
        'resident_state_counts': defaultdict(int),
        'schedule_counts': defaultdict(int)
    }
    
    # Collect all records and aggregate counts
    for ws_name, (detail_records, summaries) in all_worksheets_data.items():
        if ws_name == "Summary":
            continue  # Skip Summary worksheet itself
        
        # Add all detail records
        aggregated['all_records'].extend(detail_records)
        
        # Aggregate Issue State and Resident State counts
        for record in detail_records:
            issue_state = record.get('Issue_State') or record.get('Issue_St') or record.get('issue_state') or ''
            resident_state = record.get('Resident_State') or record.get('Resident_St') or record.get('resident_state') or ''
            
            if issue_state:
                aggregated['issue_state_counts'][issue_state] += 1
            if resident_state:
                aggregated['resident_state_counts'][resident_state] += 1
        
        # Count records per schedule (worksheet name)
        aggregated['schedule_counts'][ws_name] = len(detail_records)
    
    return aggregated


def create_workbook(connection: snowflake.connector.SnowflakeConnection,
                   worksheets_config: List[WorksheetConfig],
                   summary_config: Dict[str, Any],
                   reporting_period: Optional[str] = None) -> Workbook:
    """Create Excel workbook with all worksheets - Summary worksheet first"""
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    # Process Summary worksheet FIRST - query from summary table
    if summary_config and summary_config.get('table_name'):
        summary_table_name = summary_config['table_name']
        schedule_titles = summary_config.get('schedule_titles', {})
        
        print(f"Processing Summary worksheet from table: {summary_table_name}")
        # Basic table name validation (alphanumeric, underscore, dot for schema.table)
        if not all(c.isalnum() or c in ('_', '.') for c in summary_table_name):
            raise ValueError(f"Invalid table name format: {summary_table_name}")
        query = f"SELECT Schedule_ID, Description, Value FROM {summary_table_name} ORDER BY Schedule_ID"
        summary_data = execute_query(connection, query)
        print(f"  Fetched {len(summary_data)} summary records")
        
        # Convert schedule_titles from config (may be strings like "1", "2") to integers
        schedule_titles_dict = {}
        for key, value in schedule_titles.items():
            try:
                schedule_titles_dict[int(key)] = value
            except (ValueError, TypeError):
                schedule_titles_dict[key] = value
        
        create_summary_worksheet(wb, summary_data, schedule_titles_dict, reporting_period)
    
    # Process all detail worksheets AFTER Summary
    for ws_config in worksheets_config:
        print(f"Processing worksheet: {ws_config.name}")
        detail_records, summaries = process_worksheet_data(connection, ws_config)
        print(f"  Fetched {len(detail_records)} detail records")
        if summaries:
            print(f"  Generated {len(summaries)} summary table(s)")
        create_worksheet(wb, ws_config, detail_records, summaries)
    
    return wb


# ============================================================================
# Main CLI Function
# ============================================================================

# ============================================================================
# Configuration Loading Functions
# ============================================================================

def load_config(config_path: str) -> Dict[str, Any]:
    """Load and parse YAML configuration file"""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        return config
    except FileNotFoundError:
        print(f"Error: Configuration file '{config_path}' not found.")
        sys.exit(1)
    except yaml.YAMLError as e:
        print(f"Error: Invalid YAML in configuration file: {e}")
        sys.exit(1)


def parse_config(config: Dict[str, Any]) -> Tuple[SnowflakeConfig, Dict[str, str], List[WorksheetConfig], Dict[str, Any]]:
    """
    Parse configuration: Snowflake settings, worksheet-to-table mapping, and summary config
    Snowflake connection parameters are read from environment variables first, then fall back to config.yaml
    """
    # Get Snowflake config from environment variables (preferred) or config file
    # Environment variables take precedence for security
    account = os.getenv('SNOWFLAKE_ACCOUNT') or (config.get('snowflake', {}).get('account') if 'snowflake' in config else None)
    user = os.getenv('SNOWFLAKE_USER') or (config.get('snowflake', {}).get('user') if 'snowflake' in config else None)
    password = os.getenv('SNOWFLAKE_PASSWORD') or (config.get('snowflake', {}).get('password') if 'snowflake' in config else None)
    warehouse = os.getenv('SNOWFLAKE_WAREHOUSE') or (config.get('snowflake', {}).get('warehouse') if 'snowflake' in config else None)
    database = os.getenv('SNOWFLAKE_DATABASE') or (config.get('snowflake', {}).get('database') if 'snowflake' in config else None)
    schema = os.getenv('SNOWFLAKE_SCHEMA') or (config.get('snowflake', {}).get('schema') if 'snowflake' in config else None)
    authenticator = os.getenv('SNOWFLAKE_AUTHENTICATOR') or (config.get('snowflake', {}).get('authenticator', 'externalbrowser') if 'snowflake' in config else 'externalbrowser')
    
    # Validate required fields
    required_fields = {
        'account': account,
        'warehouse': warehouse,
        'database': database,
        'schema': schema,
        'authenticator': authenticator
    }
    
    missing_fields = [field for field, value in required_fields.items() if not value]
    if missing_fields:
        raise ValueError(
            f"Missing required Snowflake configuration. "
            f"Set environment variables or config.yaml for: {', '.join(missing_fields)}. "
            f"Environment variables: SNOWFLAKE_ACCOUNT, SNOWFLAKE_WAREHOUSE, SNOWFLAKE_DATABASE, SNOWFLAKE_SCHEMA, SNOWFLAKE_AUTHENTICATOR"
        )
    
    # For username/password auth, validate credentials are provided
    if authenticator == 'snowflake':
        if not user or not password:
            raise ValueError(
                "For 'snowflake' authenticator, SNOWFLAKE_USER and SNOWFLAKE_PASSWORD "
                "must be set (via environment variables or config.yaml)"
            )
    
    snowflake_cfg = SnowflakeConfig(
        account=account,
        user=user,
        password=password,
        warehouse=warehouse,
        database=database,
        schema=schema,
        authenticator=authenticator
    )
    
    # Get worksheet to table mapping (exclude Summary if present)
    worksheet_tables = {k: v for k, v in config.get('worksheets', {}).items() if k != 'Summary'}
    
    # Get summary configuration
    summary_config = config.get('summary', {})
    
    # Build worksheet configs using hardcoded structures
    worksheets_config = []
    for worksheet_name, table_name in worksheet_tables.items():
        ws_config = get_hardcoded_worksheet_structure(worksheet_name, table_name)
        if ws_config:
            worksheets_config.append(ws_config)
        else:
            print(f"Warning: Unknown worksheet name '{worksheet_name}', skipping...")
    
    return snowflake_cfg, worksheet_tables, worksheets_config, summary_config


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Generate multi-worksheet Excel report from Snowflake data'
    )
    parser.add_argument(
        '--config',
        type=str,
        required=True,
        help='Path to YAML configuration file with Snowflake settings and worksheet-to-table mapping'
    )
    parser.add_argument(
        '--output',
        type=str,
        required=True,
        help='Output Excel file path'
    )
    parser.add_argument(
        '--report-start-dt',
        type=str,
        required=True,
        help='Report start date (format: YYYY-MM-DD, MM/DD/YYYY, etc.)'
    )
    parser.add_argument(
        '--report-end-dt',
        type=str,
        required=True,
        help='Report end date (format: YYYY-MM-DD, MM/DD/YYYY, etc.)'
    )
    
    args = parser.parse_args()
    
    # Format dates and create reporting period string
    start_date_formatted = format_date_for_reporting_period(args.report_start_dt)
    end_date_formatted = format_date_for_reporting_period(args.report_end_dt)
    reporting_period = f"{start_date_formatted} through {end_date_formatted}"
    print(f"Reporting Period: {reporting_period}")
    
    # Load configuration
    print(f"Loading configuration from: {args.config}")
    config = load_config(args.config)
    
    # Parse configuration
    snowflake_cfg, worksheet_tables, worksheets_config, summary_config = parse_config(config)
    
    print(f"Found {len(worksheets_config)} detail worksheet(s) to process")
    if summary_config.get('table_name'):
        print(f"Summary worksheet will use table: {summary_config['table_name']}")
    
    # Create Snowflake connection
    print("Connecting to Snowflake...")
    connection = create_snowflake_connection(snowflake_cfg)
    
    try:
        # Generate workbook
        print("Generating Excel workbook...")
        wb = create_workbook(connection, worksheets_config, summary_config, reporting_period)
        
        # Save workbook
        print(f"Saving workbook to: {args.output}")
        wb.save(args.output)
        print(f"Successfully created Excel report: {args.output}")
        
    except Exception as e:
        print(f"Error generating workbook: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        close_connection(connection)


if __name__ == '__main__':
    main()

