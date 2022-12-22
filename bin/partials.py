##########################################################################################
# Filename:     partials.py
# For:          Follow_Up_Creation_Tool
# Author:       Spyros Acheimastos (acheimastos@althom.eu)
# Date:         15/12/2022
##########################################################################################

import regex as re
import numpy as np
import pandas as pd


SHEET_NAMES = ['IPC Follow-up', 'SRM A321 Follow-up', 'SRM A320 Follow-up']


def compare_mdl_values(old_value: str, df_1_x_1):
    """
    Compare the Old Value with the New Value of a cell
    to determine if it should be: ['N', 'R', '-', 'D', NaN, 'PN', 'PD']

    Args:
    ----------
        old_value:
            The old value of a cell

            row[mdl_dict_old[msn]]

        df_1_x_1:
            1x1 DataFrame or Series to later get the 'new_value' (could be an empty DataFrame)

            df_new_unique.loc[df_new_unique['PART NUMBER'] == row['PART NUMBER']][mdl_dict_new[msn]]

    Returns:
    ----------  
                            New_Value
            #   N       R       -       D       NaN
        ###################################################
            #
    O   N   #   Swap    Swap    Swap    PN      PN
    L       #
    D   R   #   Swap    Swap    Swap    PN      PN
            #
    V   -   #   Swap    Swap    Swap    PN      PN
    A       #
    L   D   #   PD      PD      PD      D       NaN
    U       #
    E   NaN #   PD      PD      PD      D       NaN
            #
    """
    # Getting "new_value" from "df_1_x_1"
    try:
        new_value = df_1_x_1.values[0]
    except IndexError:
        return 'PD'                 # Phantom Deleted (should be marked with 'TRUE')


    if old_value in ['N', 'R', '-']:
        if new_value == 'D' or pd.isna(new_value):
            return 'PD'             # Phantom Deleted (should be marked with 'TRUE')
        else:
            return new_value        # Just change symbol (should be marked with 'FALSE')

    if pd.isna(old_value) or old_value == 'D':
        if pd.isna(new_value):
            return np.nan           # Became NaN (should be marked with 'FALSE')
        elif new_value == 'D':
            return 'D'              # Became Deleted (should be marked with 'FALSE')
        else:
            return 'PN'             # Phantom New (should be marked with 'TRUE')

    return 'WTF'                    # If something weird (should be marked with 'TRUE')


def get_MSNs_and_MDLs(columns_old: list, columns_new: list):
    """
    Find MDL that changed revision from OLD Follow-Up to NEW Follow-Up.

    Args:
    ----------
        columns_old:
            The column names of the OLD DataFrame: list(df_old)
        
        columns_new:
            The column names of the NEW DataFrame: list(df_old)
    
    Returns:
    ----------
        msn_list:
            A list of the MSNs that have a new MDL
        
        mdl_dict_old:
            A dict of the OLD MDLs (key=MSN: value=OLD_MDL)

        mdl_dict_new:
            A dict of the NEW MDLs (key=MSN: value=NEW_MDL)
    """
    all_mdl_list_old = [x for x in columns_old if re.findall(r'MDL', x)]
    all_mdl_list_new = [x for x in columns_new if re.findall(r'MDL', x)]
    mdl_list_old = [x for x in all_mdl_list_old if x not in all_mdl_list_new]
    mdl_list_new = [x for x in all_mdl_list_new if x not in all_mdl_list_old]
    msn_list = [x[:4] for x in mdl_list_old]
    msn_list_tmp = [x[:4] for x in mdl_list_new]
    msn_list.sort()
    msn_list_tmp.sort()
    if msn_list != msn_list_tmp:
        raise Exception("Some MSNs don't appear on both Follow-Ups")
    mdl_dict_old = dict(zip(msn_list, mdl_list_old))
    mdl_dict_new = dict(zip(msn_list, mdl_list_new))

    # Print MDLs that will be updated
    for msn in msn_list:
        print('These MSNs will be updated')
        print(f'MSN {msn} from {mdl_dict_old[msn]} to {mdl_dict_new[msn]}')

    return msn_list, mdl_dict_old, mdl_dict_new


def get_follow_ups(filepath: str):
    """
    Read sheets 'IPC Follow-up', 'SRM A321 Follow-up', 'SRM A320 Follow-up' from an Excel file
    and save them as DataFrame in a dict.

    Args:
    ----------
        filepath:
            Filepath to Old or New Excel with Follow-up.
        
        columns_new:
            The column names of the NEW DataFrame: list(df_old)
    
    Returns:
    ----------
        df_dict:
            Dict containing DataFrames. Keys: 'IPC', 'SRM A321', 'SRM A320'
    """
    df_dict = {}
    for sheet in SHEET_NAMES:
        try:
            # Read Excel
            df = pd.read_excel(filepath, dtype=str, sheet_name=sheet)

            # Drop extra columns to avoid problems later
            df = df.loc[:, ~df.columns.str.endswith('Change')]
            df = df.drop(['Part Number Effectivity', 'TASK'], axis=1)

            # Add to dict
            key = sheet.replace(' Follow-up', '').replace(' ', '_')
            df_dict[key] = df

        except ValueError:
            print(f'Sheet "{sheet}" not found')

    return df_dict


def add_PNs_to_df_old(df_old: pd.DataFrame, df_new: pd.DataFrame, mdl_dict: dict):
    """
    Add New Part Numbers to the OLD DataFrame.

    Args:
    ----------
        df_old:
            The DataFrame from the OLD Excel.
        
        df_new:
            The DataFrame from the NEW Excel.

        mdl_dict:
            Dict with MDL column names that have changed revision.
    
    Returns:
    ----------
        df_old:
            The original df_old with New Part Numbers added. The relevant MDL columns are turned to NaN.
    """
    # New Part Numbers that are not in Old Follow-Up
    df_not_in_old = df_new.copy().loc[~df_new['PART NUMBER'].isin(df_old['PART NUMBER'])].reset_index(drop=True)
    for mdl in mdl_dict.values():
        df_not_in_old[mdl] = np.nan

    # Add "df_not_in_old" at the bottom of "df_old" and then sort
    df_old = pd.concat([df_old, df_not_in_old], ignore_index=True)
    df_old = df_old.sort_values(by=['PART NUMBER', 'CSN', 'Fig', 'Type'])

    return df_old


def reduce_df_new(df_new: pd.DataFrame, mdl_dict: dict):
    """
    Keep only columns 'PART NUMBER' and MDLs that have changed revision in 'df_new'.
    
    Also keep only one line for each unique Part Number.

    Args:
    ----------
        df_new:
            The DataFrame from the NEW Excel.

        mdl_dict:
            Dict with MDL column names that have changed revision.
    
    Returns:
    ----------
        df_new:
            The reduced df_new.
    """
    # Keep unique values
    df_new = df_new[['PART NUMBER'] + list(mdl_dict.values())]
    df_new = df_new.drop_duplicates(subset='PART NUMBER', keep="first")

    return df_new


def update_MDLs_in_df_old(df_old: pd.DataFrame, df_new: pd.DataFrame, mdl_dict: dict):
    """
    Update MDL columns on 'df_old' based on 'df_new'. Also add columns 'Effectivity Change' and '{msn} Change'

    Args:
    ----------
        df_old:
            The DataFrame from the OLD Excel after it has gone through function "add_PNs_to_df_old".
        
        df_new:
            The DataFrame from the NEW Excel after it has gone through function "reduce_df_new".

        mdl_dict:
            Dict with MDL column names that have changed revision.
    
    Returns:
    ----------
        df_old:
            The original 'df_old' with updated MDL columns and columns 'Effectivity Change' and '{msn} Change' added.
    """
    # Initialize 'Effectivity Change' column
    df_old['Effectivity Change'] = False

    for mdl in mdl_dict.values():
        # Compare Old with New MDLs to find Phantom-New (PN) and Phantom-Deleted(PD) etc
        df_old[mdl] = df_old.apply(lambda row: compare_mdl_values(row[mdl], df_new.loc[df_new['PART NUMBER'] == row['PART NUMBER']][mdl]), axis=1)
        
        # Create column '{msn} Change' based on MDL Column of each MSN
        effect_col_name = f'{mdl[:4]} Change'
        df_old[effect_col_name] = df_old[mdl].apply(lambda x: True if x in ['PN', 'PD', 'WTF'] else False)
        
        # OR Operation on final column 'Effectivity Change'
        df_old['Effectivity Change'] = df_old['Effectivity Change'] | df_old[effect_col_name]
        
        # Change ['PN', 'PD'] to ['N', 'D']
        df_old[mdl] = df_old[mdl].replace(['PN', 'PD'], ['N', 'D'])

    return df_old


def read_PS_DSOL_NC(filepath: str):
    """
    Read sheets 'DSOL', 'PS', 'NC' from New Excel and return as DataFrames.

    Args:
    ----------
        filepath:
            Filepath to New Excel with Follow-up.
        
    Returns:
    ----------
        df_dsol:
            DataFrame of sheet 'DSOL'.
        
        df_ps:
            DataFrame of sheet 'PS'.
        
        df_nc:
            DataFrame of sheet 'NC'.
    """
    df_dsol = pd.read_excel(filepath, dtype=str, sheet_name='DSOL')
    df_ps = pd.read_excel(filepath, dtype=str, sheet_name='PS')
    df_nc = pd.read_excel(filepath, dtype=str, sheet_name='NC')
    return df_dsol, df_ps, df_nc



if __name__ == '__main__':
    pass