import pandas as pd
import xlsxwriter
import logging
from config import Config
from excel_utils import get_column_letter, apply_format, apply_clue_table_formats, apply_case_table_formats, create_clue_issues_sheet, create_case_issues_sheet

logger = logging.getLogger(__name__)

import pandas as pd

def format_clue_excel(df, output_path, issues_list):
    """
    Formats the Excel file for clue data, coloring cells based on validation issues.
    df: Original DataFrame
    output_path: Path for the output Excel file
    issues_list: List of issues obtained from validation_core.py,
                 each element might be (original_df_index, clue_code_value, issue_description) (3 values)
    """
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
            df_str = df.fillna('').astype(str)
            df_str.to_excel(writer, sheet_name='Sheet1', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            for col in range(df_str.shape[1]):
                worksheet.set_column(col, col, None, workbook.add_format({'num_format': '@'}))

            red_format = workbook.add_format({'bg_color': Config.FORMATS["red"]})
            yellow_format = workbook.add_format({'bg_color': Config.FORMATS["yellow"]})

            for idx in range(len(df)):
                row = df.iloc[idx]
                apply_clue_table_formats(worksheet, df, row, idx, issues_list, False, yellow_format, red_format)

            create_clue_issues_sheet(writer, issues_list)

            logger.info(f"Clue Excel file formatted and saved successfully: {output_path}")
            return True

    except Exception as e:
        logger.error(f"Error formatting Clue Excel file: {e}", exc_info=True)
        return False

def format_case_excel(df, mismatch_indices, output_path, issues_list,
                      gender_mismatch_indices=set(), age_mismatch_indices=set(),
                      birth_date_mismatch_indices=set(), education_mismatch_indices=set(), ethnicity_mismatch_indices=set(),
                      party_member_mismatch_indices=set(), party_joining_date_mismatch_indices=set(),
                      brief_case_details_mismatch_indices=set(), filing_time_mismatch_indices=set(),
                      disciplinary_committee_filing_time_mismatch_indices=set(),
                      disciplinary_committee_filing_authority_mismatch_indices=set(),
                      supervisory_committee_filing_time_mismatch_indices=set(),
                      supervisory_committee_filing_authority_mismatch_indices=set(),
                      case_report_keyword_mismatch_indices=set(), disposal_spirit_mismatch_indices=set(),
                      voluntary_confession_highlight_indices=set(), closing_time_mismatch_indices=set(),
                      no_party_position_warning_mismatch_indices=set(),
                      recovery_amount_highlight_indices=set(),
                      trial_acceptance_time_mismatch_indices=set(),
                      trial_closing_time_mismatch_indices=set(),
                      trial_authority_agency_mismatch_indices=set(),
                      disposal_decision_keyword_mismatch_indices=set(),
                      trial_report_non_representative_mismatch_indices=set(),
                      trial_report_detention_mismatch_indices=set(),
                      confiscation_amount_indices=set(),
                      confiscation_of_property_amount_indices=set(),
                      compensation_amount_highlight_indices=set(),
                      registered_handover_amount_indices=set(),
                      disciplinary_sanction_mismatch_indices=set(),
                      administrative_sanction_mismatch_indices=set()
                      ):
    """
    Formats the Excel file for case data, coloring cells based on validation issues.
    df: Original DataFrame
    mismatch_indices: Set of indices for inconsistent rows
    output_path: Path for the output Excel file
    issues_list: List of issues obtained from case_validators.py,
                 each element might be (original_df_index, case_code_value, person_code_value, issue_description) (4 values)
    Other index sets: Used for red or yellow highlighting of specific fields
    """
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
            df_str = df.fillna('').astype(str)
            df_str.to_excel(writer, sheet_name='Sheet1', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            for col in range(df_str.shape[1]):
                worksheet.set_column(col, col, None, workbook.add_format({'num_format': '@'}))

            red_format = workbook.add_format({'bg_color': Config.FORMATS["red"]})
            yellow_format = workbook.add_format({'bg_color': Config.FORMATS["yellow"]})

            for idx in range(len(df)):
                row = df.iloc[idx]
                apply_case_table_formats(worksheet, df, row, idx, mismatch_indices, issues_list, True,
                                         gender_mismatch_indices, age_mismatch_indices, birth_date_mismatch_indices,
                                         education_mismatch_indices, ethnicity_mismatch_indices, party_member_mismatch_indices,
                                         party_joining_date_mismatch_indices, brief_case_details_mismatch_indices,
                                         filing_time_mismatch_indices, disciplinary_committee_filing_time_mismatch_indices,
                                         disciplinary_committee_filing_authority_mismatch_indices,
                                         supervisory_committee_filing_time_mismatch_indices,
                                         supervisory_committee_filing_authority_mismatch_indices,
                                         case_report_keyword_mismatch_indices, disposal_spirit_mismatch_indices,
                                         voluntary_confession_highlight_indices, closing_time_mismatch_indices,
                                         no_party_position_warning_mismatch_indices, recovery_amount_highlight_indices,
                                         trial_acceptance_time_mismatch_indices, trial_closing_time_mismatch_indices,
                                         trial_authority_agency_mismatch_indices, disposal_decision_keyword_mismatch_indices,
                                         trial_report_non_representative_mismatch_indices, trial_report_detention_mismatch_indices,
                                         confiscation_amount_indices, confiscation_of_property_amount_indices,
                                         compensation_amount_highlight_indices, registered_handover_amount_indices,
                                         disciplinary_sanction_mismatch_indices,
                                         administrative_sanction_mismatch_indices,
                                         yellow_format, red_format)

            create_case_issues_sheet(writer, issues_list)

            logger.info(f"Case Excel file formatted and saved successfully: {output_path}")
            return True

    except Exception as e:
        logger.error(f"Error formatting Case Excel file: {e}", exc_info=True)
        return False
