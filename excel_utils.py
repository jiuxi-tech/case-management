# excel_utils.py
import xlsxwriter
import pandas as pd
import logging

logger = logging.getLogger(__name__)

def get_column_letter(df, column_name):
    """
    Gets the column index and converts it to an Excel column letter (1-based index).
    """
    try:
        col_idx = df.columns.get_loc(column_name) + 1
        return xlsxwriter.utility.xl_col_to_name(col_idx - 1)
    except KeyError:
        logger.warning(f"Column '{column_name}' not found in DataFrame for get_column_letter.")
        return None # Or handle more specifically if needed

def apply_format(worksheet, row_idx, col_letter, value, condition_met, format_obj):
    """
    Applies formatting based on a condition.
    row_idx is the DataFrame's index (0-based), Excel row number is row_idx + 2 (due to header and Pandas' default 0-row).
    """
    if col_letter is None: # Handle cases where column might not exist (from get_column_letter)
        return

    excel_row = row_idx + 2      # Excel rows start from 1, and there's a header, so add 2
    if condition_met:
        # If the condition is met, write the value with formatting
        worksheet.write(f'{col_letter}{excel_row}', value if pd.notna(value) else '', format_obj)
    else:
        # If the condition is not met, write the value without formatting
        worksheet.write(f'{col_letter}{excel_row}', value if pd.notna(value) else '')