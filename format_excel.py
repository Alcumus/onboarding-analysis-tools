#!/usr/bin/env python3
"""
Enhanced Excel Formatter for Parallel Processing Results
Matches ONCW.out.xlsx baseline structure with all required sheets
"""

import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

def apply_excel_formatting(input_file, output_file):
    """Apply proper Excel formatting to match ONCW.out.xlsx baseline structure"""
    
    print(f"Applying Excel formatting to {input_file} -> {output_file}")
    
    # Load the data
    df = pd.read_excel(input_file)
    
    # Create workbook with proper formatting
    wb = openpyxl.Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Define the complete sheet structure matching ONCW.out.xlsx baseline
    sheet_structure = [
        'all',                      # All records
        'onboarding',              # Action-specific sheets
        'association_fee',
        're_onboarding', 
        'subscription_upgrade',
        'ambiguous_onboarding',
        'restore_suspended',
        'activation_link',
        'already_qualified',
        'add_questionnaire',
        'missing_info',
        'follow_up_qualification',
        'Data to import',          # Special sheets
        'Existing Contractors',
        'Data for HS'
    ]
    
    sheets = []
    
    # Create all sheets in order
    for sheet_name in sheet_structure:
        ws = wb.create_sheet(title=sheet_name)
        sheets.append(ws)
        
        # Determine data for this sheet
        if sheet_name == 'all':
            # All data
            sheet_df = df.copy()
        elif sheet_name == 'Data for HS':
            # All data (same as 'all' sheet)
            sheet_df = df.copy()
        elif sheet_name == 'Existing Contractors':
            # Only contractors that were found/matched (not missing_info)
            sheet_df = df[df['action'] != 'missing_info'] if 'action' in df.columns else pd.DataFrame()
        elif sheet_name == 'Data to import':
            # Empty sheet for manual data entry (matching baseline)
            sheet_df = pd.DataFrame(columns=df.columns[:31] if len(df.columns) >= 31 else df.columns)
        else:
            # Action-specific sheet
            if 'action' in df.columns:
                sheet_df = df[df['action'] == sheet_name]
            else:
                sheet_df = pd.DataFrame()
        
        # Write headers
        headers = sheet_df.columns.tolist() if len(sheet_df) > 0 else df.columns.tolist()
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)
        
        # Write data (if any)
        if len(sheet_df) > 0:
            for row_idx, (_, row) in enumerate(sheet_df.iterrows(), 2):
                for col_idx, value in enumerate(row, 1):
                    # Clear analysis columns for missing_info records
                    column_name = headers[col_idx - 1] if col_idx <= len(headers) else ''
                    
                    # Check if this is an analysis column and record has missing_info action
                    is_analysis_column = column_name.lower() in ['analysis', 'hc_contractor_summary', 'cbx_contractor_summary']
                    is_missing_info = row.get('action', '') == 'missing_info'
                    
                    if is_analysis_column and is_missing_info:
                        # Keep analysis columns empty for missing_info records
                        ws.cell(row=row_idx, column=col_idx, value="")
                    else:
                        ws.cell(row=row_idx, column=col_idx, value=value)
    
    # Apply formatting to all sheets (matching baseline structure)
    style = TableStyleInfo(
        name="TableStyleMedium2", 
        showFirstColumn=False,
        showLastColumn=False, 
        showRowStripes=True, 
        showColumnStripes=False
    )
    
    for sheet in sheets:
        # Always create table structure, even for empty sheets
        if sheet.max_row >= 1:  # At least headers
            # Ensure we have minimum dimensions
            if sheet.max_row == 1:
                # Add empty row for table structure (matching baseline)
                for col_idx in range(1, sheet.max_column + 1):
                    sheet.cell(row=2, column=col_idx, value="")
            
            # Create table
            table_ref = f'A1:{get_column_letter(sheet.max_column)}{max(sheet.max_row, 2)}'
            tab = Table(
                displayName=f"Table_{sheet.title.replace(' ', '_').replace('-', '_')}",
                ref=table_ref
            )
            tab.tableStyleInfo = style
            sheet.add_table(tab)
            
            # Auto-adjust column widths
            for column in sheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                # Set reasonable width limits
                adjusted_width = min(max(max_length + 2, 12), 50)
                sheet.column_dimensions[column_letter].width = adjusted_width
            
            # Apply text wrapping to analysis columns if they exist
            analysis_columns = ['analysis', 'hc_contractor_summary', 'cbx_contractor_summary']
            for col_name in analysis_columns:
                if col_name in df.columns:
                    col_idx = df.columns.get_loc(col_name) + 1
                    col_letter = get_column_letter(col_idx)
                    
                    # Set wider width for analysis columns
                    sheet.column_dimensions[col_letter].width = 40
                    
                    # Apply text wrapping
                    for row in range(2, sheet.max_row + 1):
                        cell = sheet.cell(row=row, column=col_idx)
                        cell.alignment = Alignment(wrapText=True, vertical='top')
            
            # Note: Auto filters are handled by the table structure automatically
            # Baseline doesn't show explicit auto filters, so we skip this
    
    # Save the formatted workbook
    wb.save(output_file)
    print(f"✅ Applied formatting: {len(sheets)} sheets, tables, filters, and styling")

def main():
    import sys
    
    if len(sys.argv) != 3:
        print("Usage: python3 format_excel.py <input_file> <output_file>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    apply_excel_formatting(input_file, output_file)

if __name__ == "__main__":
    main()