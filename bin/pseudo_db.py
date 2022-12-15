##########################################################################################
# Filename:     pseudo_db.py
# For:          Follow_Up_Creation_Tool
# Author:       Spyros Acheimastos (acheimastos@althom.eu)
# Date:         15/12/2022
##########################################################################################

import os
# import string
import warnings
import regex as re
import numpy as np
import pandas as pd
from functools import reduce

COLOR_HEADER_YELLOW = '#FFD966'
COLOR_HEADER_BLUE = '#9BC2E6'
COLOR_HEADER_ORANGE = '#F4B084'
COLOR_HEADER_GREEN = '#A9D08E'
COLOR_LIGHT_GREEN = '#C6EFCE'
COLOR_LIGHT_RED = '#FFC7CE'


KEEP_COLUMN_LIST = ['PART NUMBER', 'CSN', 'Fig', 'Type', 'BOM Parts', 'PART TITLE']
BOOK_COLUMN_LIST = ['IPC', 'SRM A321', 'SRM A320']
POSSIBLE_SHEET_NAMES = ['IPC Follow-up', 'SRM A321 Follow-up', 'SRM Follow-up', 'SRM A320 Follow-up']
PSEUDO_DB_COLUMNS = ['PART NUMBER', 'CSN', 'Fig', 'Type', 'BOM Parts', 'PART TITLE', 'IPC', 'SRM A321', 'SRM A320']

def follow_up_to_pseudo_db(filepath: str):
    """
    Read a Follow-up Excel and combine Sheets 'IPC Follow-up', 'SRM A321 Follow-up', 'SRM A320 Follow-up' 
    into a single PseudoDataBase DataFrame.

    Args:
    ----------
        filepath:
            The filepath of the Follow-up Excel.

    Returns:
    ----------
        df_merged:
            The PseudoDataBase DataFrame that was created.

    """
    # Catch Errors
    if not filepath.endswith('.xlsx'): raise Exception('The given file was not an Excel')

    # To ignore "UserWarning: Data Validation" and "UserWarning: Conditional Formatting"
    with warnings.catch_warnings():
        warnings.simplefilter(action='ignore', category=UserWarning)
        # Get sheet names
        xls = pd.ExcelFile(filepath, engine='openpyxl')
        current_sheet_names = xls.sheet_names
        sheet_list = [x for x in current_sheet_names if x in POSSIBLE_SHEET_NAMES]

        # Print missing Sheets for debugging
        missing_sheets = [x for x in POSSIBLE_SHEET_NAMES if x not in sheet_list]
        if missing_sheets:
            for sheet in missing_sheets:
                print(f'File {os.path.basename(filepath)} is missing sheet: "{sheet}"')
        
        # Catch Errors
        if not sheet_list: raise Exception(f'Sheets not found: "IPC Follow-up", "SRM A321 Follow-up", "SRM Follow-up", "SRM A320 Follow-up"')
        
        # Loop for reading. 
        df_list = []
        missing_type_column = False
        for sheet in sheet_list:
            column_name = sheet.replace(' Follow-up', '')
            if column_name == 'SRM': column_name = 'SRM A321' # EFW Follow-up R07 has Sheet "SRM Follow-up" instead of "SRM A321 Follow-up"
            df = pd.read_excel(filepath, dtype=str, sheet_name=sheet, usecols='A:P') 
            # Added "usecols" to avoid reading Connectors of Cosmin. This could create a problem if IPC_CSN is at O instead of P ?

            # If a Follow-up version doesnt have a "Type" column, it is added as NaN 
            if 'Type' not in list(df.columns):
                df['Type'] = np.nan
                missing_type_column = True
            
            # If a Follow-up version doesnt have a "BOM Parts" column, it is added as NaN 
            if 'BOM Parts' not in list(df.columns):
                df['BOM Parts'] = ''                    # Maybe this should be NaN, but is read as "float64" and there is a problem when merging

            # If a "TITLE" column exists, rename it to "PART TITLE" to avoid mistakes later
            if "TITLE" in list(df.columns): 
                df = df.rename(columns={'TITLE': 'PART TITLE'})

            df = df[KEEP_COLUMN_LIST]
            df = df.drop_duplicates().reset_index(drop=True)
            df[column_name] = True
            df_list.append(df)


        # This should be moved to when the combination happens with other versions
        # df_merged = df_merged.sort_values(by=KEEP_COLUMN_LIST).drop_duplicates(subset=KEEP_COLUMN_LIST[:-2]).reset_index(drop=True)


        # Combine IPC, SRM A321, SRM A320
        df_merged = reduce(lambda left, right: pd.merge(left, right, on=KEEP_COLUMN_LIST, how='outer'), df_list)
        
        # Add column in case one of the books is missing in this Follow-Up version (i.e. Follow-up R07 doesn't have SRM A320)
        missing_BOOK_COLUMN_LIST = [x for x in BOOK_COLUMN_LIST if x not in list(df_merged.columns)]
        for missing_book in missing_BOOK_COLUMN_LIST:
            df_merged[missing_book] = False

        # Fill NaN with appropriate values
        df_merged[BOOK_COLUMN_LIST] = df_merged[BOOK_COLUMN_LIST].fillna(value=False)
        df_merged[['CSN', 'Fig']] = df_merged[['CSN', 'Fig']].fillna('TBD')

        # # Every 'Fig' that contains a letter higher than 'S' will be characterized as an 'EFW' Figure
        # # If something has 'Fig' = TBD or anything else, will be sorted as 'AIB'
        # if missing_type_column is True:
        #     series_type = df_merged['Fig'].apply(lambda fig: 'EFW' if re.findall(r'^\d+[S-Z]$', fig) else 'AIB')
        #     df_merged['Type'] = series_type
        
        # # Removing letter from AIB figures
        # df_merged['Fig'] = df_merged.apply(lambda row: row['Fig'][:-1] if row['Type'] == 'AIB' and re.findall(r'^\d+[A-R]$', row['Fig']) else row['Fig'], axis=1)
        # df_merged = df_merged.sort_values(by=KEEP_COLUMN_LIST).drop_duplicates(subset=KEEP_COLUMN_LIST[:-2]).reset_index(drop=True)

        # Change 31/10/22: 
        if missing_type_column is True:
            # Every 'Fig' that contains a letter higher than 'S' will be characterized as an 'EFW' Figure
            # If something has 'Fig' = TBD or anything else, will be sorted as 'AIB'
            series_type = df_merged['Fig'].apply(lambda fig: 'EFW' if re.findall(r'^\d+[S-Z]$', fig) else 'AIB')
            df_merged['Type'] = series_type
        
            # Removing letter from AIB figures that end with any letter until 'S'
            df_merged['Fig'] = df_merged.apply(lambda row: row['Fig'][:-1] if row['Type'] == 'AIB' and re.findall(r'^\d+[A-R]$', row['Fig']) else row['Fig'], axis=1)

            # Sort and drop duplicates
            df_merged = df_merged.sort_values(by=KEEP_COLUMN_LIST).drop_duplicates(subset=KEEP_COLUMN_LIST[:-2]).reset_index(drop=True)

        else:
            # Removing letter from AIB figures
            df_merged['Fig'] = df_merged.apply(lambda row: row['Fig'][:-1] if row['Type'] == 'AIB' and re.findall(r'^\d+[A-Z]$', row['Fig']) else row['Fig'], axis=1)

            # Update 05/12/22: replace anything that is not 'EFW', 'AIB', 'TBD' with 'TBD'
            df_merged.loc[~df_merged['Type'].isin(['EFW', 'AIB', 'TBD']), 'Type'] = 'TBD'

            # Sort and drop duplicates     
            df_merged = df_merged.sort_values(by=KEEP_COLUMN_LIST).drop_duplicates(subset=KEEP_COLUMN_LIST[:-2]).reset_index(drop=True)


        # Vasilis reportred losing things from SRM A320/A321
        # Easy solution (09/11/22)
        # df_merged['SRM A321'] = pd.merge(df_merged['SRM A321'], df_merged['SRM A320'], how='outer')
        df_merged['SRM A321'] = df_merged[['SRM A321', 'SRM A320']].any(axis=1)
        df_merged['SRM A320'] = df_merged['SRM A321']

        return df_merged

def merge_pseudo_dbs(list_of_dfs: list):
    """
    Take a list of PseudoDataBases and combine them.

    Args:
    ----------
        list_of_dfs:
            A list of PseudoDataBase DataFrames.

    Returns:
    ----------
        df_merged:
            The PseudoDataBase DataFrame that was created by merging the ones in the list

    """
    # For sorting. We have descending order in "IPC", "SRM A321", "SRM A320" to put TRUE before FALSE and not lose any info!!
    final_sort_list = KEEP_COLUMN_LIST + BOOK_COLUMN_LIST
    final_sort_ascending_list = [True for x in KEEP_COLUMN_LIST] + [False for x in BOOK_COLUMN_LIST]

    # Merging and sorting
    # df = pd.merge(df_R07, df_R08, how='outer')
    # for x in list_of_dfs:
    #     print('The types of columns is:')
    #     print(x.dtypes)
    df_merged = reduce(lambda left, right: pd.merge(left, right, how='outer'), list_of_dfs)
    df_merged = df_merged.sort_values(by=final_sort_list, ascending=final_sort_ascending_list).drop_duplicates(subset=KEEP_COLUMN_LIST[:-2]).reset_index(drop=True)
    # df_merged = df_merged.drop_duplicates(subset=KEEP_COLUMN_LIST[:-2]).reset_index(drop=True)


    return df_merged

def create_pseudo_db_for_CC(df_initial: pd.DataFrame, df_pseudo_db: pd.DataFrame):
    """
    Get the New Part Numbers not in the PseudoDataBase and add them at the bottom of PseudoDataBase.
    
    Args:
    ----------
        df_initial:
            The initial Follow-up DataFrame.

        df_pseudo_db:
            The PseudoDataBase DataFrame.
    
    Returns:
    ----------
        df_new_pseudo_db:
            The New PseudoDataBase DataFrame that was created by adding the New Part Numbers at the bottom.

    """
    # Get New Part Numbers not in PseudoDB
    new_PNs = set(df_initial['PART NUMBER']) - set(df_pseudo_db['PART NUMBER'])
    
    # Check if there are no New Part Numbers
    if not new_PNs:
        print(f'No New Part Numbers where found. You can proceed.')
        return df_pseudo_db

    # DataFrame for New Part Numbers
    df_new_PNs = df_initial.copy()
    df_new_PNs = df_new_PNs[df_new_PNs['PART NUMBER'].isin(new_PNs)]
    df_new_PNs['CSN'] = 'TBD'
    df_new_PNs['Fig'] = 'TBD'
    df_new_PNs['Type'] = 'TBD'                                          # NOT SURE IF THIS BREAKS ANYTHING SOMEWHERE ELSE
    # df_new_PNs['Type'] = 'AIB'
    df_new_PNs['BOM Parts'] = ''
    df_new_PNs['IPC'] = 'TBD'
    df_new_PNs['SRM A321'] = 'TBD'
    df_new_PNs['SRM A320'] = 'TBD'
    df_new_PNs = df_new_PNs[list(df_pseudo_db.columns)].sort_values(by=['PART NUMBER'])

    # Add New Part Numbers at the bottom of PseudoDB
    df_new_pseudo_db = pd.concat([df_pseudo_db, df_new_PNs])

    # Save updated PseudoDB
    # excelfilepath_for_CC = excelfilepath.replace('.xlsx', '_for_CC.xlsx')
    # pseudo_db_to_excel(df_new_pseudo_db, excelfilepath_for_CC)

    # pseudo_db_to_excel(df_new_pseudo_db, excelfilepath)


    print(f'{len(new_PNs)} New Part Numbers where added at the bottom of PseudoDataBase. \
        \nBefore proceeding, please sort them to the correct manual: IPC, SRM A321, SRM A320')
    
    return df_new_pseudo_db

def read_pseudo_db(filepath: str):
    """
    Read a PseudoDataBase Excel and return the PseudoDataBase DataFrame.

    Args:
    ----------
        filepath:
            The filepath of the PseudoDataBase Excel.

    Returns:
    ----------
        df_pseudo_db:
            The PseudoDataBase DataFrame.
    """

    df_pseudo_db = pd.read_excel(filepath, dtype=str, sheet_name='Pseudo_Data_Base')

    # Check column names
    if set(df_pseudo_db.columns) != set(PSEUDO_DB_COLUMNS):
        raise Exception(f'Wrong column names inside: {filepath}')

    # Check for TBDs, Empty or Wrong values
    if (~df_pseudo_db[['IPC', 'SRM A321', 'SRM A320']].isin(['True', 'False'])).any().any():
        print(f'TBDs or wrong values found at columns "IPC"/"SRM A321"/"SRM A320" inside: {filepath}.\nWill be parsed as "TRUE"')
        # raise Exception(f'TBDs or wrong values found at columns "IPC" or "SRM A321" or "SRM A320" inside: {filepath}')
        
    # Convert to boolean and convert TBDs to 'True'
    df_pseudo_db['IPC'] = df_pseudo_db['IPC'].fillna(True).replace({'True': True, 'False': False, 'TBD': True})
    df_pseudo_db['SRM A321'] = df_pseudo_db['SRM A321'].fillna(True).replace({'True': True, 'False': False, 'TBD': True})
    df_pseudo_db['SRM A320'] = df_pseudo_db['SRM A320'].fillna(True).replace({'True': True, 'False': False, 'TBD': True})

    # Update 05/12/22: Check for Empty or Wrong values in 'Type' column
    if (~df_pseudo_db['Type'].isin(['EFW', 'AIB', 'TBD'])).any():
        print(f'Empty or wrong values found at column "Type" inside: {filepath}.\nWill be parsed as "TBD"')

    # Update 05/12/22: Replace anything that is not 'EFW', 'AIB', 'TBD' with 'TBD'
    df_pseudo_db.loc[~df_pseudo_db['Type'].isin(['EFW', 'AIB', 'TBD']), 'Type'] = 'TBD'     

    return df_pseudo_db



def gnrt_lines_and_split(df_initial: pd.DataFrame, df_pseudo_db: pd.DataFrame, A320_msn_list: list):
    """
     Generate new lines for the Follow-up using the PseudoDataBase, and then split them into 'IPC', 'SRM A321', 'SRM A320'.

    Args:
    ----------
        df_initial:
            The initial Follow-up DataFrame.

        df_pseudo_db:
            The PseudoDataBase DataFrame, preferably after 'TBDs' have been characterized as 'IPC', 'SRM A321', 'SRM A320'.

        A320_msn_list:
            A list with A320 MSNs.

    Returns:
    ----------
        df_IPC:                  
            The DataFrame for the IPC with the generated lines

         df_SRM_A321:                  
            The DataFrame for the SRM A321 with the generated lines

        df_SRM_A320:                  
            The DataFrame for the SRM A320 with the generated lines

    Columns of Returned:
    ----------
        - 'PART NUMBER'
        - 'CSN'
        - 'Fig'
        - 'Type'
        - 'PART TITLE'
        - MSN_MDL_#_1
        - MSN_MDL_#_2
        - ...
    """
    # Keep only some columns from df_initial
    mns_column_list = [x for x in list(df_initial.columns) if re.findall(r'^\d{4}', x)]
    keep_only_list = ['PART NUMBER', 'PART TITLE'] + mns_column_list
    df_initial = df_initial[keep_only_list]

    # Convert to boolean and convert TBDs to 'True'
    # Spyros 25/10/22: I added '.fillna(True)' just to be sure
    df_pseudo_db['IPC'] = df_pseudo_db['IPC'].fillna(True).replace({'True': True, 'False': False, 'TBD': True})
    df_pseudo_db['SRM A321'] = df_pseudo_db['SRM A321'].fillna(True).replace({'True': True, 'False': False, 'TBD': True})
    df_pseudo_db['SRM A320'] = df_pseudo_db['SRM A320'].fillna(True).replace({'True': True, 'False': False, 'TBD': True})

    # Update 05/12/22: Removed 'TBD' to 'EFW'
    # Maybe convert TBDs 'Type' to 'EFW' to not lose any info ?????????????
    # df_pseudo_db['Type'] = df_pseudo_db['Type'].replace({'TBD': 'EFW'})

    # Generate new lines
    df_gnrt = df_initial.merge(df_pseudo_db, how='left', on=['PART NUMBER'])

    # Check for different titles
    try:
        df_gnrt['Different Title'] = df_gnrt.apply(lambda row: True if row['PART TITLE_x'] != row['PART TITLE_y'] else False, axis=1)

        print('Some Part Numbers have a different Title in the MDLs and in the PseudoDataBase:')
        for idx, row in df_gnrt.loc[df_gnrt['Different Title'] == True].iterrows():
            print('PN: {} \t MDL: "{}" \t PDB: "{}"'.format(row['PART NUMBER'], row['PART TITLE_x'], row['PART TITLE_y']))

        # Rename and Drop x/y. Keep the title from the MDL
        df_gnrt = df_gnrt.rename({'PART TITLE_x': 'PART TITLE'}, axis=1)
        df_gnrt = df_gnrt.drop(columns=['Different Title', 'PART TITLE_y'])

    except KeyError:
        print('Part Numbers have the same Title in the MDLs and in the PseudoDataBase:')
        pass

    # Split gnrted lines into the correct manual
    df_IPC =  df_gnrt.loc[df_gnrt['IPC'] == True]
    df_SRM_A321 =  df_gnrt.loc[df_gnrt['SRM A321'] == True]
    df_SRM_A320 =  df_gnrt.loc[df_gnrt['SRM A320'] == True]

    # Sort and Keep only certain columns
    keep_only_list = ['PART NUMBER', 'CSN', 'Fig', 'Type', 'PART TITLE'] + mns_column_list
    df_IPC = df_IPC[keep_only_list].sort_values(by=['PART NUMBER', 'CSN', 'Fig', 'Type', 'PART TITLE']).reset_index(drop=True)
    df_SRM_A321 = df_SRM_A321[keep_only_list].sort_values(by=['PART NUMBER', 'CSN', 'Fig', 'Type', 'PART TITLE']).reset_index(drop=True)
    df_SRM_A320 = df_SRM_A320[keep_only_list].sort_values(by=['PART NUMBER', 'CSN', 'Fig', 'Type', 'PART TITLE']).reset_index(drop=True)

    # Subtract MSN columns from 'df_SRM_A321' and 'df_SRM_A320' based on A320_msn_list
    all_msn_column_names = [x for x in list(df_IPC.columns) if re.findall(r'^\d{4}', x)]
    # drop_columns_from_SRM_A321 = [x for x in all_msn_column_names if x[:4] in A320_msn_list]
    # drop_columns_from_SRM_A320 = [x for x in all_msn_column_names if x[:4] not in A320_msn_list]
    df_SRM_A321 = df_SRM_A321.drop(columns=[x for x in all_msn_column_names if x[:4] in A320_msn_list], axis=1)
    df_SRM_A320 = df_SRM_A320.drop(columns=[x for x in all_msn_column_names if x[:4] not in A320_msn_list], axis=1)
    if not A320_msn_list: df_SRM_A320 = None

    return df_IPC, df_SRM_A321, df_SRM_A320



if __name__ == '__main__':
    pass
    # Read older Follow-Up Excels and convert them to the PseudoDataBase format
    # filepath_R07 = 'D:\\August_setup_follow_up\\Older_Follow_Ups\\EFW Follow-up R07.xlsx'
    # filepath_R08 = 'D:\\August_setup_follow_up\\Older_Follow_Ups\\EFW Follow-up R08.xlsx'
    # df_R07 = follow_up_to_pseudo_db(filepath_R07)
    # df_R08 = follow_up_to_pseudo_db(filepath_R08)

    # # Merge PseudoDataBases into One
    # df_pseudo = merge_pseudo_dbs([df_R07, df_R08])



    #######################################################
    # Save Merged PseudoDataBases to Excel
    # excelfilepath = 'PseudoDataBase' + time.strftime('_%Y_%m_%d-%H_%M_%S') + '.xlsx'
    # excelfilepath = 'PseudoDataBase_4.xlsx'
    # pseudo_db_to_excel(df_pseudo, excelfilepath)

    







    # df_pseudo_2 = pd.merge(df_R07, df_R08, how='outer')
    # final_sort_list = KEEP_COLUMN_LIST + BOOK_COLUMN_LIST
    # final_sort_ascending_list = [True for x in KEEP_COLUMN_LIST] + [False for x in BOOK_COLUMN_LIST]
    # df_pseudo_2 = df_pseudo_2.sort_values(by=final_sort_list, ascending=final_sort_ascending_list).drop_duplicates(subset=KEEP_COLUMN_LIST[:-2]).reset_index(drop=True)
    # df_pseudo_2.to_excel('PseudoDataBase_3.xlsx', index=False)


    # df_pseudo = pd.read_excel('PseudoDataBase.xlsx', dtype=str)
