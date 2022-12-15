##########################################################################################
# Filename:     save_to_excel.py
# For:          Follow_Up_Creation_Tool
# Author:       Spyros Acheimastos (acheimastos@althom.eu)
# Date:         15/12/2022
##########################################################################################

import string
import regex as re
import numpy as np
import pandas as pd


COLOR_HEADER_YELLOW = '#FFD966'
COLOR_HEADER_BLUE = '#9BC2E6'
COLOR_HEADER_ORANGE = '#F4B084'
COLOR_HEADER_GREEN = '#A9D08E'
COLOR_HEADER_PINK = '#F28594'
COLOR_HEADER_GRAY = '#D9D9D9'
COLOR_LIGHT_GREEN = '#C6EFCE'
COLOR_DARK_GREEN = '#006100'
COLOR_LIGHT_RED = '#FFC7CE'
COLOR_DARK_RED = '#9C0006'
COLOR_PN_RED = '#FFCCCC'
COLOR_PN_BLUE = '#99CCFF'
COLOR_PN_YELLOW = '#FFF2CC'
COLOR_EA_N = '#C6EFCE'
COLOR_EA_R = '#FFEB9C'
COLOR_EA_D = '#FFC7CE'


# For Part Number Column Format
EFFECT_COLUMN_FOLLOW_UP = 'Part Number Effectivity'
EFFECT_COLUMN_DSOL = 'Effectivity'

# Drop Down Lists
DROP_LIST_TYPE = ['EFW', 'AIB', 'TBD']
DROP_LIST_STATUS = ['Not started', 'In progress', 'To be checked', 'Pending Info', 'Waiting Illu', 'Rework Needed', 'To be final checked', 'Finished']
DROP_LIST_RFT_WFT = ['RFT', 'WFT']
DROP_LIST_OTD = ['ON TIME', 'LATE']
DROP_LIST_WEIGHT = ['E', 'I', 'S']
DROP_LIST_MANUAL = ['IPC', 'SRM A321', 'SRM A320']
DROP_LIST_ILLU = ['NEW AIB', 'NEW EFW', 'REVISION']
DROP_LIST_YES_NO = ['Yes', 'No']


def add_formats_to_workbook(workbook):
    """
    Add format to an xlsxwriter workbook for the creation of the 'Follow-up' Project Managment Board

    Args:
    ----------
        workbook:
            An xlsxwriter workbook.

    Returns:
    ----------
        workbook:
            An xlsxwriter workbook with the new formats.

        formats:
            A dictionary with the created formats

    Formats:
    ----------
        - cell_left
        - cell_left_wrap
        - cell_center
        - header_blue
        - header_orange
        - header_green
        - header_yellow
        - header_pink
        - header_gray
        - PN_blue
        - PN_yellow
        - PN_red
        - start_date
        - cell_true
        - cell_false
        - EA_N
        - EA_R
        - EA_D
    """
    cell_left_format_dict = {
        'right':        2,
        'left':         2,
        'bottom':       1,
        'top':          1,
        'align':        'left',
        'valign':       'vcenter'
        # 'num_format':   '@'       # 02/12/22 Made every cell "General"
    }

    cell_left_wrap_format_dict = {
        'right':        2,
        'left':         2,
        'bottom':       1,
        'top':          1,
        'text_wrap':    True,
        'align':        'left',
        'valign':       'vcenter'
        # 'num_format':   '@'       # 02/12/22 Made every cell "General"
    }

    cell_center_format_dict = {
        'right':        2,
        'left':         2,
        'bottom':       1,
        'top':          1,
        'align':        'center',
        'valign':       'vcenter'
        # 'num_format':   '@'       # 02/12/22 Made every cell "General"
    }

    header_general_dict = {
        'right':        2,
        'left':         2,
        'bottom':       2,
        'top':          2,
        'text_wrap':    True,
        'bold':         True, 
        'align':        'center',
        'valign':       'vcenter',
    }

    # Dictionaries for formats
    header_blue_dict = header_general_dict.copy()
    header_orange_dict = header_general_dict.copy()
    header_green_dict = header_general_dict.copy()
    header_yellow_dict = header_general_dict.copy()
    header_pink_dict = header_general_dict.copy()
    header_gray_dict = header_general_dict.copy()
    header_blue_dict['bg_color'] = COLOR_HEADER_BLUE
    header_orange_dict['bg_color'] = COLOR_HEADER_ORANGE
    header_green_dict['bg_color'] = COLOR_HEADER_GREEN
    header_yellow_dict['bg_color'] = COLOR_HEADER_YELLOW
    header_pink_dict['bg_color'] = COLOR_HEADER_PINK
    header_gray_dict['bg_color'] = COLOR_HEADER_GRAY

    # Add formats to workbook
    cell_left = workbook.add_format(cell_left_format_dict)
    cell_left_wrap = workbook.add_format(cell_left_wrap_format_dict)
    cell_center = workbook.add_format(cell_center_format_dict)
    header_blue = workbook.add_format(header_blue_dict)
    header_orange = workbook.add_format(header_orange_dict)
    header_green = workbook.add_format(header_green_dict)
    header_yellow = workbook.add_format(header_yellow_dict)
    header_pink = workbook.add_format(header_pink_dict)
    header_gray = workbook.add_format(header_gray_dict)
    PN_blue = workbook.add_format({'bg_color': COLOR_PN_BLUE})
    PN_yellow = workbook.add_format({'bg_color': COLOR_PN_YELLOW})
    PN_red = workbook.add_format({'bg_color': COLOR_PN_RED})
    start_date = workbook.add_format({'bg_color': COLOR_LIGHT_RED, 'font_color': COLOR_DARK_RED})
    cell_true = workbook.add_format({'bg_color': COLOR_LIGHT_GREEN, 'font_color': COLOR_DARK_GREEN})
    cell_false = workbook.add_format({'bg_color': COLOR_LIGHT_RED, 'font_color': COLOR_DARK_RED})
    EA_N = workbook.add_format({'bg_color': COLOR_EA_N})
    EA_R = workbook.add_format({'bg_color': COLOR_EA_R})
    EA_D = workbook.add_format({'bg_color': COLOR_EA_D})


    formats = {
        "cell_left":         cell_left,
        "cell_left_wrap":    cell_left_wrap,
        "cell_center":       cell_center,
        "header_blue":       header_blue,
        "header_orange":     header_orange,
        "header_green":      header_green,
        "header_yellow":     header_yellow,
        "header_pink":       header_pink,
        "header_gray":       header_gray,
        "PN_blue":           PN_blue,
        "PN_yellow":         PN_yellow,
        "PN_red":            PN_red,
        "start_date":        start_date,
        "cell_true":         cell_true,
        "cell_false":        cell_false,
        "EA_N":              EA_N,
        "EA_R":              EA_R,
        "EA_D":              EA_D,
    }

    return workbook, formats

def format_sheet_Follow_Up(writer, formats: dict, prop_dict: dict, column_names: list, max_length: int, num_of_rows: int):
    """
    Format one of the 'Follow-up' sheets ('IPC Follow-up', 'SRM A321 Follow-up', 'SRM A320 Follow-up').

    Args:
    ----------
        writer:
            An xlsxwriter writer.

        formats:
            A dictionary with the created formats.

        prop_dict:
            A dictionary the properties of the sheet {'sheetname', 'color', 'header_format'}

        column_names:
            A list with the column names of the DataFrame.

        max_length:
            The maximum length of the string in Effectivity column.

        num_of_rows:
            Number of rows of the DataFrame.

    Returns:
    ----------
        writer:
            An xlsxwriter writer with the new format for the 'Follow-up' sheet.
    """
    worksheet = writer.sheets[prop_dict['sheetname']]
    worksheet.set_tab_color(prop_dict['color'])
    worksheet.freeze_panes(1, 0)

    # Set format for columns    
    worksheet.set_column('A:A', 18, formats['cell_left'])
    worksheet.set_column('B:B', 12.1, formats['cell_center'])
    worksheet.set_column('C:C', 7.4, formats['cell_center'])
    worksheet.set_column('D:D', 7.86, formats['cell_center'])
    worksheet.set_column('E:E', 56.4, formats['cell_left'])
    worksheet.set_column('F:F', 5*(len(column_names)-15), formats['cell_left'])
    worksheet.set_column('G:G', 14.5, formats['cell_center'])
    worksheet.set_column('H:H', 13.2, formats['cell_center'])
    worksheet.set_column('I:I', 11, formats['cell_center'])
    worksheet.set_column('J:J', 17.3, formats['cell_center'])
    worksheet.set_column('K:K', 9.3, formats['cell_center'])
    worksheet.set_column('L:L', 13.2, formats['cell_center'])
    worksheet.set_column('M:M', 9.3, formats['cell_center'])
    worksheet.set_column('N:N', 63.6, formats['cell_left_wrap'])
    worksheet.set_column('O:O', 9.3, formats['cell_center'])
    for idx in range(15, len(column_names)):
        worksheet.set_column(get_column_range(idx+1, mode=0), 10.3, formats['cell_center'])

    # Unfortunattely I cant format cells that already have values inside as I did for QBs

    # Set format for headers
    for idx, col in enumerate(column_names):
        worksheet.write(get_column_range(idx+1, mode=1), col, prop_dict['header_format'])
    worksheet.set_row(0, 30.75)

    # Conditional formatting for highlighting Part Numbers
    worksheet.conditional_format(f'A2:A{num_of_rows+1}', {'type': 'formula', 'criteria': f'=IF( AND({0}<LEN($F2), LEN($F2)<={4}), TRUE(), FALSE() )', 'format': formats['PN_red']})
    worksheet.conditional_format(f'A2:A{num_of_rows+1}', {'type': 'formula', 'criteria': f'=IF( AND({4}<LEN($F2), LEN($F2)<={max_length-6}), TRUE(), FALSE() )', 'format': formats['PN_yellow']})
    worksheet.conditional_format(f'A2:A{num_of_rows+1}', {'type': 'formula', 'criteria': f'=IF( AND({max_length-6}<LEN($F2), LEN($F2)<={max_length}), TRUE(), FALSE() )', 'format': formats['PN_blue']})

    # Conditional formatting for Start Date
    worksheet.conditional_format(f'I2:I{num_of_rows+1}', {'type': 'formula', 'criteria': '=$H2<>""', 'format': formats['start_date']})

    # Update 05/12/22: Add drop down lists for 'Type' column and 'Status' column
    worksheet.data_validation(f'D2:D{num_of_rows+1}', {'validate': 'list', 'source': DROP_LIST_TYPE})
    worksheet.data_validation(f'J2:J{num_of_rows+1}', {'validate': 'list', 'source': DROP_LIST_STATUS})

    # Add Authors
    if ('authors' in prop_dict) and (prop_dict['authors'] is not None):
        worksheet.data_validation(f'H2:H{num_of_rows+1}', {'validate': 'list', 'source': prop_dict['authors']})
        worksheet.data_validation(f'L2:L{num_of_rows+1}', {'validate': 'list', 'source': prop_dict['authors']})

    return writer

def format_sheet_DSOL(writer, formats: dict, prop_dict: dict, column_names: list, max_length: int):
    """
    Format the 'DSOL' sheet.

    Args:
    ----------
        writer:
            An xlsxwriter writer.

        formats:
            A dictionary with the created formats.

        prop_dict:
            A dictionary the properties of the sheet {'sheetname', 'color', 'header_format'}

        column_names:
            A list with the column names of the DataFrame.

        max_length:
            The maximum length of the string in Effectivity column.

    Returns:
    ----------
        writer:
            An xlsxwriter writer with the new format for the 'DSOL' sheet.
    """
    worksheet = writer.sheets[prop_dict['sheetname']]
    worksheet.set_tab_color(prop_dict['color'])
    worksheet.freeze_panes(1, 0)

    # Set format for columns
    worksheet.set_column('A:A', 18, formats['cell_left'])
    worksheet.set_column('B:B', 56.4, formats['cell_left'])
    worksheet.set_column('C:C', 4, formats['cell_center'])
    worksheet.set_column('D:D', 6, formats['cell_center'])
    worksheet.set_column('E:E', 5.43, formats['cell_center'])
    worksheet.set_column('F:F', 5*(len(column_names)-6), formats['cell_left'])
    for idx in range(6, len(column_names)):
        worksheet.set_column(get_column_range(idx+1, mode=0), 10.3, formats['cell_center'])

    # Set format for headers
    for idx, col in enumerate(column_names):
        worksheet.write(get_column_range(idx+1, mode=1), col, prop_dict['header_format'])
    worksheet.set_row(0, 30.75)

    # Conditional formatting for highlighting Part Numbers
    worksheet.conditional_format('A2:A1048576', {'type': 'formula', 'criteria': f'=IF( AND({0}<LEN($F2), LEN($F2)<={4}), TRUE(), FALSE() )', 'format': formats['PN_red']})
    worksheet.conditional_format('A2:A1048576', {'type': 'formula', 'criteria': f'=IF( AND({4}<LEN($F2), LEN($F2)<={max_length-6}), TRUE(), FALSE() )', 'format': formats['PN_yellow']})
    worksheet.conditional_format('A2:A1048576', {'type': 'formula', 'criteria': f'=IF( AND({max_length-6}<LEN($F2), LEN($F2)<={max_length}), TRUE(), FALSE() )', 'format': formats['PN_blue']})

    return writer

def format_sheet_PS(writer, formats: dict, prop_dict: dict, column_names: list):
    """
    Format the 'PS' sheet.

    Args:
    ----------
        writer:
            An xlsxwriter writer.

        formats:
            A dictionary with the created formats.

        prop_dict:
            A dictionary the properties of the sheet {'sheetname', 'color', 'header_format'}

        column_names:
            A list with the column names of the DataFrame.

    Returns:
    ----------
        writer:
            An xlsxwriter writer with the new format for the 'PS' sheet.
    """
    worksheet = writer.sheets[prop_dict['sheetname']]
    worksheet.set_tab_color(prop_dict['color'])
    worksheet.freeze_panes(1, 0)

    # Set format for columns
    worksheet.set_column('A:A', 18, formats['cell_left'])
    worksheet.set_column('B:B', 10, formats['cell_center'])
    worksheet.set_column('C:C', 18, formats['cell_left'])
    worksheet.set_column('D:D', 64, formats['cell_left'])
    worksheet.set_column('E:E', 5*(len(column_names)-5), formats['cell_left'])
    for idx in range(5, len(column_names)):
        worksheet.set_column(get_column_range(idx+1, mode=0), 10.3, formats['cell_center'])

    # Set format for headers
    for idx, col in enumerate(column_names):
        worksheet.write(get_column_range(idx+1, mode=1), col, prop_dict['header_format'])
    worksheet.set_row(0, 30.75)

    return writer

def format_sheet_NC(writer, formats: dict, prop_dict: dict, column_names: list):
    """
    Format the 'NC' sheet.

    Args:
    ----------
        writer:
            An xlsxwriter writer.

        formats:
            A dictionary with the created formats.

        prop_dict:
            A dictionary the properties of the sheet {'sheetname', 'color', 'header_format'}

        column_names:
            A list with the column names of the DataFrame.

    Returns:
    ----------
        writer:
            An xlsxwriter writer with the new format for the 'NC' sheet.
    """
    worksheet = writer.sheets[prop_dict['sheetname']]
    worksheet.set_tab_color(prop_dict['color'])
    worksheet.freeze_panes(1, 0)

    # Set format for columns
    worksheet.set_column('A:A', 23.1, formats['cell_left'])
    worksheet.set_column('B:B', 6, formats['cell_center'])
    worksheet.set_column('C:C', 33.1, formats['cell_left'])
    worksheet.set_column('D:D', 6, formats['cell_center'])
    worksheet.set_column('E:E', 73, formats['cell_left_wrap'])
    worksheet.set_column('F:F', 10, formats['cell_center'])
    worksheet.set_column('G:G', 10, formats['cell_center'])
    worksheet.set_column('H:H', 78, formats['cell_left'])
    worksheet.set_column('I:I', 5*(len(column_names)-9), formats['cell_left'])
    for idx in range(9, len(column_names)):
        worksheet.set_column(get_column_range(idx+1, mode=0), 10.3, formats['cell_center'])

    # Set format for headers
    for idx, col in enumerate(column_names):
        worksheet.write(get_column_range(idx+1, mode=1), col, prop_dict['header_format'])
    worksheet.set_row(0, 30.75)

    return writer

def add_sheet_QB(workbook, formats: dict, prop_dict: dict):
    """
    Add a Quality Board sheet to the workbook.

    Args:
    ----------
        workbook:
            An xlsxwriter workbook.

        formats:
            A dictionary with the created formats.

        prop_dict:
            A dictionary the properties of the sheet {'sheetname', 'color', 'header_format'}

    Returns:
    ----------
        workbook:
            An xlsxwriter workbook with the new Quality Board sheet.
    """
    QB_header_list = [
        'PN', 'CSN', 'FIGURE', 'Author  (SC)', 'SC Date',
        'Cross Checker (CC)', 'CC Date', 'RFT/WFT', 'OTD', 'Verification Number',
        'NC Number', 'NC Code', 'NC Quantity', 'Weight', 'Description of NC'
    ]
    
    worksheet = workbook.add_worksheet(prop_dict['sheetname'])
    worksheet.freeze_panes(1, 0)
    worksheet.set_tab_color(prop_dict['color'])
    for idx, header in enumerate(QB_header_list):
        worksheet.write(get_column_range(idx+1, mode=1), header, prop_dict['header_format'])

    # Set Header Row Height
    worksheet.set_row(0, 30)

    # Set Column Width
    worksheet.set_column('A2:A5000', 18)
    worksheet.set_column('B2:B5000', 10.5)
    worksheet.set_column('C2:C5000', 8.5)
    worksheet.set_column('D2:D5000', 13.2)
    worksheet.set_column('E2:E5000', 12.2)
    worksheet.set_column('F2:F5000', 13.2)
    worksheet.set_column('G2:G5000', 12.2)
    worksheet.set_column('H2:H5000', 9.3)
    worksheet.set_column('I2:I5000', 7.9)
    worksheet.set_column('J2:J5000', 10.7)
    worksheet.set_column('K2:K5000', 7.86)
    worksheet.set_column('L2:L5000', 7.86)
    worksheet.set_column('M2:M5000', 7.86)
    worksheet.set_column('N2:N5000', 7.86)
    worksheet.set_column('O2:O5000', 96)

    # Set Cell formatting for the first 1000 rows
    for i in range(2, 5001):
        worksheet.write_blank(f'A{i}', None, formats['cell_left'])
        worksheet.write_blank(f'B{i}', None, formats['cell_center'])
        worksheet.write_blank(f'C{i}', None, formats['cell_center'])
        worksheet.write_blank(f'D{i}', None, formats['cell_center'])
        worksheet.write_blank(f'E{i}', None, formats['cell_center'])
        worksheet.write_blank(f'F{i}', None, formats['cell_center'])
        worksheet.write_blank(f'G{i}', None, formats['cell_center'])
        worksheet.write_blank(f'H{i}', None, formats['cell_center'])
        worksheet.write_blank(f'I{i}', None, formats['cell_center'])
        worksheet.write_blank(f'J{i}', None, formats['cell_center'])
        worksheet.write_blank(f'K{i}', None, formats['cell_center'])
        worksheet.write_blank(f'L{i}', None, formats['cell_center'])
        worksheet.write_blank(f'M{i}', None, formats['cell_center'])
        worksheet.write_blank(f'N{i}', None, formats['cell_center'])
        worksheet.write_blank(f'O{i}', None, formats['cell_left_wrap'])

    # Add DropDown Lists
    worksheet.data_validation('H2:H5000', {'validate': 'list', 'source': DROP_LIST_RFT_WFT})
    worksheet.data_validation('I2:I5000', {'validate': 'list', 'source': DROP_LIST_OTD})
    worksheet.data_validation('N2:N5000', {'validate': 'list', 'source': DROP_LIST_WEIGHT})

    # Add Authors
    if ('authors' in prop_dict) and (prop_dict['authors'] is not None):
        worksheet.data_validation('D2:D5000', {'validate': 'list', 'source': prop_dict['authors']})
        worksheet.data_validation('F2:F5000', {'validate': 'list', 'source': prop_dict['authors']})

    return workbook

def add_sheet_QB_illu(workbook, formats: dict, prop_dict: dict):
    """
    Add the Illu Quality Board sheet to the workbook.

    Args:
    ----------
        workbook:
            An xlsxwriter workbook.

        formats:
            A dictionary with the created formats.

        prop_dict:
            A dictionary with the properties of the sheet {'sheetname', 'color', 'header_format'}

    Returns:
    ----------
        workbook:
            An xlsxwriter workbook with the new Illu Quality Board sheet.
    """
    illu_QB_header_list = [ 'Manual', 'PN', 'CSN', 'Fig', 'Sheet', 'Author', 'MCK Date', 'Type of Illu',
        'Illustrator', 'Time (minutes)', 'SC Date', 'Illustrator CC', 'CC Time (minutes)', 'CC Date',
        'RFT/WFT', 'OTD', 'Verification Number', 'NC Number', 'NC Code', 'NC Quantity', 'Weight',
        'Description of NC', 'Comment', 'IO Deleted', 'Incorporated'
    ]

    worksheet = workbook.add_worksheet(prop_dict['sheetname'])
    worksheet.freeze_panes(1, 0)
    worksheet.set_tab_color(prop_dict['color'])
    for idx, header in enumerate(illu_QB_header_list):
        worksheet.write(get_column_range(idx+1, mode=1), header, prop_dict['header_format'])

    # Set Header Row Height
    worksheet.set_row(0, 30)

    # Set Column Width
    worksheet.set_column('A:A', 9.5)
    worksheet.set_column('B:B', 18)
    worksheet.set_column('C:C', 10.5)
    worksheet.set_column('D:D', 8.43)
    worksheet.set_column('E:E', 6.9)
    worksheet.set_column('F:F', 13.2)
    worksheet.set_column('G:G', 12.2)
    worksheet.set_column('H:H', 9.5)
    worksheet.set_column('I:I', 16.3)
    worksheet.set_column('J:J', 8.6)
    worksheet.set_column('K:K', 12.2)
    worksheet.set_column('L:L', 16.3)
    worksheet.set_column('M:M', 8.6)
    worksheet.set_column('N:N', 12.2)
    worksheet.set_column('O:O', 9.3)
    worksheet.set_column('P:P', 8.5)
    worksheet.set_column('Q:Q', 10.7)
    worksheet.set_column('R:R', 12.1)
    worksheet.set_column('S:S', 8.43)
    worksheet.set_column('T:T', 8.43)
    worksheet.set_column('U:U', 8.43)
    worksheet.set_column('V:V', 25.0)
    worksheet.set_column('W:W', 43.5)
    worksheet.set_column('X:X', 10.14)
    worksheet.set_column('Y:Y', 11.3)

    # Set Cell formatting for the first 1500 rows
    for i in range(2, 1501):
        worksheet.write_blank(f'A{i}', None, formats['cell_center'])
        worksheet.write_blank(f'B{i}', None, formats['cell_left'])
        worksheet.write_blank(f'C{i}', None, formats['cell_center'])
        worksheet.write_blank(f'D{i}', None, formats['cell_center'])
        worksheet.write_blank(f'E{i}', None, formats['cell_center'])
        worksheet.write_blank(f'F{i}', None, formats['cell_center'])
        worksheet.write_blank(f'G{i}', None, formats['cell_center'])
        worksheet.write_blank(f'H{i}', None, formats['cell_center'])
        worksheet.write_blank(f'I{i}', None, formats['cell_center'])
        worksheet.write_blank(f'J{i}', None, formats['cell_center'])
        worksheet.write_blank(f'K{i}', None, formats['cell_center'])
        worksheet.write_blank(f'L{i}', None, formats['cell_center'])
        worksheet.write_blank(f'M{i}', None, formats['cell_center'])
        worksheet.write_blank(f'N{i}', None, formats['cell_center'])
        worksheet.write_blank(f'O{i}', None, formats['cell_center'])
        worksheet.write_blank(f'P{i}', None, formats['cell_center'])
        worksheet.write_blank(f'Q{i}', None, formats['cell_center'])
        worksheet.write_blank(f'R{i}', None, formats['cell_center'])
        worksheet.write_blank(f'S{i}', None, formats['cell_center'])
        worksheet.write_blank(f'T{i}', None, formats['cell_center'])
        worksheet.write_blank(f'U{i}', None, formats['cell_center'])
        worksheet.write_blank(f'V{i}', None, formats['cell_left_wrap'])
        worksheet.write_blank(f'W{i}', None, formats['cell_left_wrap'])
        worksheet.write_blank(f'X{i}', None, formats['cell_center'])
        worksheet.write_blank(f'Y{i}', None, formats['cell_center'])

    # Conditional formatting for Start Date and Incorporated
    worksheet.conditional_format('G2:G1500', {'type': 'formula', 'criteria': '=$F2<>""', 'format': formats['start_date']})
    worksheet.conditional_format('Y2:Y1500', {'type': 'text', 'criteria': 'containing', 'value': 'Yes', 'format': formats['cell_true']})
    worksheet.conditional_format('Y2:Y1500', {'type': 'text', 'criteria': 'containing', 'value': 'No', 'format': formats['cell_false']})

    # Add DropDown Lists
    worksheet.data_validation('A2:A1500', {'validate': 'list', 'source': DROP_LIST_MANUAL})
    worksheet.data_validation('H2:H1500', {'validate': 'list', 'source': DROP_LIST_ILLU})
    worksheet.data_validation('O2:O1500', {'validate': 'list', 'source': DROP_LIST_RFT_WFT})
    worksheet.data_validation('P2:P1500', {'validate': 'list', 'source': DROP_LIST_OTD})
    worksheet.data_validation('U2:U1500', {'validate': 'list', 'source': DROP_LIST_WEIGHT})
    worksheet.data_validation('X2:X1500', {'validate': 'list', 'source': DROP_LIST_YES_NO})
    worksheet.data_validation('Y2:Y1500', {'validate': 'list', 'source': DROP_LIST_YES_NO})

    # Add Authors
    if ('authors' in prop_dict) and (prop_dict['authors'] is not None):
        worksheet.data_validation('F2:F1500', {'validate': 'list', 'source': prop_dict['authors']})
    if ('illustrators' in prop_dict) and (prop_dict['illustrators'] is not None):
        worksheet.data_validation('I2:I1500', {'validate': 'list', 'source': prop_dict['illustrators']})
        worksheet.data_validation('L2:L1500', {'validate': 'list', 'source': prop_dict['illustrators']})

    return workbook


def divmod_excel(n):
    a, b = divmod(n, 26)
    if b == 0:
        return a - 1, b + 26
    return a, b

def get_column_range(num: str, mode: int = 0):
    """ Given the idx of a column, get the column range for Excel (1 -> A:A, 2 -> B:B etc)

    Args:
     ----------
        num:   
            The idx of the column.

        mode:  
            The mode, check below.

    Modes:
    ----------
    - 0 = Whole column
    - 1 = Just Header
    - 2 = All except header
    - 3 = Only second cell
    - 4 = Only letter
    """
    chars = []
    while num > 0:
        num, d = divmod_excel(num)
        chars.append(string.ascii_uppercase[d - 1])
    x = ''.join(reversed(chars))
    
    if mode == 0:
        return x + ':' + x
    elif mode == 1:
        return x + '1'
    elif mode == 2:
        return x + '2' + ':' + x + '1048576'
    elif mode == 3:
        return x + '2'
    elif mode == 4:
        return x
    else:
        print('Wrong mode given, it was set to 0')
        return x + ':' + x

# def pseudo_db_to_excel_OLD(df: pd.DataFrame, excelfilepath: str):
    """
    Save PseudoDataBase to Excel with custom formatting.
    
    Args:
    ----------
        df:
            The PseudoDataBase DataFrame.
        
        excelfilepath:
            The filepath of the Excel that will be created.
    """
    cell_left_format_dict = {
        'right':        2,
        'left':         2,
        'bottom':       1,
        'top':          1,
        'align':        'left',
        'valign':       'vcenter',
        'num_format':   '@'
    }
    cell_center_format_dict = {
        'right':        2,
        'left':         2,
        'bottom':       1,
        'top':          1,
        'align':        'center',
        'valign':       'vcenter',
        'num_format':   '@'
    }
    header_general_dict = {
        'right':        2,
        'left':         2,
        'bottom':       2,
        'top':          2,
        'text_wrap':    True,
        'bold':         True, 
        'align':        'center',
        'valign':       'vcenter',
    }

    # Create Dictionaries
    header_blue_dict = header_general_dict.copy()
    header_orange_dict = header_general_dict.copy()
    header_green_dict = header_general_dict.copy()
    header_yellow_dict = header_general_dict.copy()
    header_blue_dict['bg_color'] = COLOR_HEADER_BLUE
    header_orange_dict['bg_color'] = COLOR_HEADER_ORANGE
    header_green_dict['bg_color'] = COLOR_HEADER_GREEN
    header_yellow_dict['bg_color'] = COLOR_HEADER_YELLOW

    # Creater Writer and write DataFrame to Excel
    writer = pd.ExcelWriter(excelfilepath, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Pseudo_Data_Base')

    # Get the xlsxwriter workbook and worksheet objects
    workbook  = writer.book
    worksheet = writer.sheets['Pseudo_Data_Base']

    # Add formating to workbook
    cell_left_format = workbook.add_format(cell_left_format_dict)
    cell_center_format = workbook.add_format(cell_center_format_dict)
    header_blue_format = workbook.add_format(header_blue_dict)
    header_orange_format = workbook.add_format(header_orange_dict)
    header_green_format = workbook.add_format(header_green_dict)
    header_yellow_format = workbook.add_format(header_yellow_dict)
    
    # For Columns
    worksheet.set_column('A:A', 17, cell_left_format)
    worksheet.set_column('B:B', 11.5, cell_center_format)
    worksheet.set_column('C:C', 10, cell_center_format)
    worksheet.set_column('D:D', 10, cell_center_format)
    worksheet.set_column('E:E', 10, cell_center_format)
    worksheet.set_column('F:F', 64, cell_left_format)
    worksheet.set_column('G:G', 10, cell_center_format)
    worksheet.set_column('H:H', 10, cell_center_format)
    worksheet.set_column('I:I', 10, cell_center_format)

    # For Headers
    column_names = list(df)
    worksheet.freeze_panes(1, 0)
    worksheet.write('A1', column_names[0], header_yellow_format)
    worksheet.write('B1', column_names[1], header_yellow_format)
    worksheet.write('C1', column_names[2], header_yellow_format)
    worksheet.write('D1', column_names[3], header_yellow_format)
    worksheet.write('E1', column_names[4], header_yellow_format)
    worksheet.write('F1', column_names[5], header_yellow_format)
    worksheet.write('G1', column_names[6], header_blue_format)
    worksheet.write('H1', column_names[7], header_orange_format)
    worksheet.write('I1', column_names[8], header_green_format)
    worksheet.set_row(0, 20)

    # Conditional formatting for highlighting True/False
    true_cell_format = workbook.add_format({'bg_color': COLOR_LIGHT_GREEN})
    false_cell_format = workbook.add_format({'bg_color': COLOR_LIGHT_RED})
    worksheet.conditional_format('G2:I1048576', {'type': 'text', 'criteria': 'containing', 'value': 'TRUE', 'format': true_cell_format})
    worksheet.conditional_format('G2:I1048576', {'type': 'text', 'criteria': 'containing', 'value': 'FALSE', 'format': false_cell_format})

    # Save
    writer.save()

def pseudo_db_to_excel(df: pd.DataFrame, excelfilepath: str):
    """
    Save PseudoDataBase to Excel with custom formatting.
    
    Args:
    ----------
        df:
            The PseudoDataBase DataFrame.
        
        excelfilepath:
            The filepath of the Excel that will be created.
    """
    # Creater Writer and write DataFrame to Excel
    writer = pd.ExcelWriter(excelfilepath, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Pseudo_Data_Base')

    # Get the xlsxwriter workbook and worksheet objects, and add formats to workbook
    workbook  = writer.book
    worksheet = writer.sheets['Pseudo_Data_Base']
    workbook, formats = add_formats_to_workbook(workbook)

    # For Columns
    worksheet.set_column('A:A', 17, formats['cell_left'])
    worksheet.set_column('B:B', 11.5, formats['cell_center'])
    worksheet.set_column('C:C', 10, formats['cell_center'])
    worksheet.set_column('D:D', 10, formats['cell_center'])
    worksheet.set_column('E:E', 10, formats['cell_center'])
    worksheet.set_column('F:F', 64, formats['cell_left'])
    worksheet.set_column('G:G', 10, formats['cell_center'])
    worksheet.set_column('H:H', 10, formats['cell_center'])
    worksheet.set_column('I:I', 10, formats['cell_center'])

    # For Headers
    column_names = list(df)
    worksheet.freeze_panes(1, 0)
    worksheet.write('A1', column_names[0], formats['header_yellow'])
    worksheet.write('B1', column_names[1], formats['header_yellow'])
    worksheet.write('C1', column_names[2], formats['header_yellow'])
    worksheet.write('D1', column_names[3], formats['header_yellow'])
    worksheet.write('E1', column_names[4], formats['header_yellow'])
    worksheet.write('F1', column_names[5], formats['header_yellow'])
    worksheet.write('G1', column_names[6], formats['header_blue'])
    worksheet.write('H1', column_names[7], formats['header_orange'])
    worksheet.write('I1', column_names[8], formats['header_green'])
    worksheet.set_row(0, 20)

    # Conditional formatting for highlighting True/False
    true_cell_format = workbook.add_format({'bg_color': COLOR_LIGHT_GREEN})
    false_cell_format = workbook.add_format({'bg_color': COLOR_LIGHT_RED})
    worksheet.conditional_format('G2:I1048576', {'type': 'text', 'criteria': 'containing', 'value': 'TRUE', 'format': true_cell_format})
    worksheet.conditional_format('G2:I1048576', {'type': 'text', 'criteria': 'containing', 'value': 'FALSE', 'format': false_cell_format})

    # Save
    writer.save()

def initial_follow_up_to_excel(df_initial: pd.DataFrame, excelfilepath: str):
    """
    Generate the initial Follow-Up.
    """
    # excelfilepath = 'Follow_Up_test_From_OLD_DSOL.xlsx'
    # excelfilepath = 'Follow_Up_test_From_FINAL_DSOL_2.xlsx'
    writer = pd.ExcelWriter(excelfilepath, engine='xlsxwriter')
    df_initial.to_excel(writer, index=False, sheet_name='Follow-up Initial')

    # Get the xlsxwriter workbook and add formats
    workbook = writer.book
    workbook, formats = add_formats_to_workbook(workbook)

    # Change 'Follow-up' to 'Follow-up Initial' !!!
    prop_dict_Initial = {'sheetname': 'Follow-up Initial', 'color': COLOR_HEADER_PINK, 'header_format': formats['header_pink']}
    writer = format_sheet_Follow_Up(writer, formats, prop_dict_Initial, list(df_initial), max_length=df_initial[EFFECT_COLUMN_FOLLOW_UP].str.len().max(), num_of_rows=df_initial.shape[0])

    # Save and close
    writer.save()
    # writer.close()

def final_follow_up_to_excel(df_dsol: pd.DataFrame, df_ps: pd.DataFrame, df_nc: pd.DataFrame, excelfilepath: str, authors_dict: dict = None, add_QBs = True, **dict_with_follow_ups):
    """
    Create the final Follow-Up Excel.

    Args:
    ----------
        df_dsol:
            The DataFrame for DSOL.

        df_ps:
            The DataFrame for PS.

        df_nc:
            The DataFrame for NC.

        excelfilepath:
            The filepath for the Excel that will be created.

        authors_dict:
            Dictionary of Authors with keys "IPC", "SRM", "ILLU"

        add_QBs:
            Boolean to add empty Quality Boards. (Default=True)

        **dict_with_follow_ups:
            kwargs with possible keys: 'IPC', 'SRM_A321', 'SRM_A320' and DataFrames as values.
            This is to handle the case of Follow-up without 'SRM_A320'.
    """
    # Unpack "authors_dict"
    if authors_dict:
        list_of_authors_IPC = authors_dict['IPC']
        list_of_authors_SRM = authors_dict['SRM']
        list_of_illustrators = authors_dict['ILLU']
        list_of_authors_ALL = list_of_authors_IPC + list_of_authors_SRM
    else:
        list_of_authors_IPC = None
        list_of_authors_SRM = None
        list_of_illustrators = None
        list_of_authors_ALL = None

    # Get xlsx writer
    writer = pd.ExcelWriter(excelfilepath, engine='xlsxwriter')

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    workbook, formats = add_formats_to_workbook(workbook)

    # Properties For 
    prop_dict = {
        'IPC': {'sheetname': 'IPC Follow-up', 'color': COLOR_HEADER_BLUE, 'header_format': formats['header_blue'], 'authors': list_of_authors_IPC},
        'SRM_A321': {'sheetname': 'SRM A321 Follow-up', 'color': COLOR_HEADER_ORANGE, 'header_format': formats['header_orange'], 'authors': list_of_authors_SRM},
        'SRM_A320': {'sheetname': 'SRM A320 Foolow-up', 'color': COLOR_HEADER_GREEN, 'header_format': formats['header_green'], 'authors': list_of_authors_SRM},
        'DSOL': {'sheetname': 'DSOL', 'color': COLOR_HEADER_YELLOW, 'header_format': formats['header_yellow']},
        'PS': {'sheetname': 'PS', 'color': COLOR_HEADER_YELLOW, 'header_format': formats['header_yellow']},
        'NC': {'sheetname': 'NC', 'color': COLOR_HEADER_YELLOW, 'header_format': formats['header_yellow']}
    }

    prop_QB_dict = {
        'IPC': {'sheetname': 'IPC - QB', 'color': COLOR_HEADER_BLUE, 'header_format': formats['header_blue'], 'authors': list_of_authors_IPC},
        'SRM_A321': {'sheetname': 'SRM A321 - QB', 'color': COLOR_HEADER_ORANGE, 'header_format': formats['header_orange'], 'authors': list_of_authors_SRM},
        'SRM_A320': {'sheetname': 'SRM A320 - QB', 'color': COLOR_HEADER_GREEN, 'header_format': formats['header_green'], 'authors': list_of_authors_SRM},
        'ILLU': {'sheetname': 'ILLU - QB', 'color': COLOR_HEADER_PINK, 'header_format': formats['header_pink'], 'authors': list_of_authors_ALL, 'illustrators': list_of_illustrators}
    }

    # Add and format Follow-Up and Quality Board Sheets
    for key, df in dict_with_follow_ups.items():
        df.to_excel(writer, index=False, sheet_name=prop_dict[key]['sheetname'])
        writer = format_sheet_Follow_Up(writer, formats, prop_dict[key], list(df.columns), max_length=df[EFFECT_COLUMN_FOLLOW_UP].str.len().max(), num_of_rows=df.shape[0])
        if add_QBs is True:
            workbook = add_sheet_QB(workbook, formats, prop_QB_dict[key])
    
    # Add Illu Quality Board
    if add_QBs is True:
        workbook = add_sheet_QB_illu(workbook, formats, prop_QB_dict['ILLU'])

    # Write and formats sheets: DSOL / PS / NC 
    df_dsol.to_excel(writer, index=False, sheet_name='DSOL')
    df_ps.to_excel(writer, index=False, sheet_name='PS')
    df_nc.to_excel(writer, index=False, sheet_name='NC')
    writer = format_sheet_DSOL(writer, formats, prop_dict['DSOL'], list(df_dsol.columns), max_length=df_dsol[EFFECT_COLUMN_DSOL].str.len().max())
    writer = format_sheet_PS(writer, formats, prop_dict['PS'], list(df_ps.columns))
    writer = format_sheet_NC(writer, formats, prop_dict['NC'], list(df_nc.columns))

    # Save and close
    writer.save()
    # writer.close()