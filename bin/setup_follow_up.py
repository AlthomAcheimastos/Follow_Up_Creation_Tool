##########################################################################################
# Filename:     setup_follow_up.py
# For:          Follow_Up_Creation_Tool
# Author:       Spyros Acheimastos (acheimastos@althom.eu)
# Date:         15/12/2022
##########################################################################################

import os
import json
import warnings
import numpy as np
import regex as re
import pandas as pd
from functools import reduce


EFFECT_COLUMN = {
    'FOLLOW_UP': {'name': 'Part Number Effectivity', 'idx': 5},
    'FOLLOW_UP_INITIAL': {'name': 'Part Number Effectivity', 'idx': 2},
    'DSOL': {'name': 'Effectivity', 'idx': 5},
    'PS': {'name': 'Effectivity of the CHILD', 'idx': 4},
    'NC': {'name': 'Effectivity', 'idx': 5}
}


def read_JSON(filepath: str):
    """
    Read JSON file containg lists of strings for "all", "all_A320", "new" and "rev" MSNs.

    Args:
    ----------
        filepath:
            The path to the JSON file.

    Returns:
    ----------
        json_MSNs:
            Dictionary with keys "all", "all_A320", "new", "rev"
    """
    with open(filepath) as f:
        json_MSNs = json.load(f)

    # Check if keys are correct
    if list(set(['all', 'all_A320', 'new', 'rev']) - set(json_MSNs.keys())):
        raise Exception("The keys inside JSON should be 'all', 'all_A320', 'new', 'rev'")

    return json_MSNs

def read_JSON_authors(filepath: str):
    """
    Read JSON file containg lists of strings for "IPC", "SRM" and "ILLU" Authors.

    Args:
    ----------
        filepath:
            The path to the JSON file.

    Returns:
    ----------
        json_MSNs:
            Dictionary with keys "IPC", "SRM", "ILLU"
    """
    with open(filepath) as f:
        json_authors = json.load(f)

    # Check if keys are correct
    if list(set(['IPC', 'SRM', 'ILLU']) - set(json_authors.keys())):
        raise Exception("The keys inside JSON should be 'IPC', 'SRM', 'ILLU'")

    return json_authors

def read_MDLs_current(rootdir: str, current_msn_list: list):
    """
    Read only the current MDLs from 'rootdir', and create lists with Dataframes in order to create Follow-Up, DSOL and PS.

    Args:
    ----------
        rootdir:
            The path to the folder containing the MDLs in '.xlsx' format.
            
        current_msn_list:
            List of the MSNs for the current IPC/SRM revision.

    Returns:
    ----------
        mdl_msn_list:          
            List containing MSNs that where found in the MDLs.

        follow_up_list:     
            List containing DataFrames for each MSN in order to create Follow-Up.
    """
    mdl_msn_list = []
    follow_up_list = []
    for root, _, files in os.walk(rootdir, topdown=True):
        for file in files:
            if not file.endswith('xlsx'): continue
            msn = file[:4]
            mdl_msn_list.append(msn)
            if msn in ['0835', '2737']:
                mdl_column = file.replace('.xlsx','').replace('349-', '')     # The MDL filename follows the format "0835_349-MDL-0835-G.xlsx"
            else:                                                                  
                mdl_column = file.replace('.xlsx','').replace('EFW-E-', '')     # The MDL filename follows the format "3708_EFW-E-MDL-00243-C.xlsx"

            # To ignore "UserWarning: Data Validation" and "UserWarning: Conditional Formatting"
            with warnings.catch_warnings():
                warnings.simplefilter(action='ignore', category=UserWarning)
                if msn in current_msn_list:
                    # Read Sheet 'Applicable Part List' to create Follow-Up DataFrame (for New MSNs)
                    df = pd.read_excel(root + os.sep + file, dtype=str, sheet_name='Applicable Part List')
                    df = df[['PART NUMBER', 'PART TITLE', 'QTY', 'PART TYPE', 'PART ISSUE', 'DIFF']]    # Keep only these columns
                    df = df.rename(columns={'DIFF': mdl_column})
                    for column in df.columns: df[column] = df[column].str.strip()                       # Strip leading and trailing whitespaces
                    df = df.loc[ df['PART TYPE'] == 'DSOL']
                    df = df.loc[ df['PART NUMBER'].map(lambda x: True if re.findall(r'R0|R1|R3', x) else False) ]       # Keep only "R0", "R1" and "R3"
                    df = df.drop(['QTY', 'PART TYPE', 'PART ISSUE'], axis=1)
                    follow_up_list.append(df)

    return mdl_msn_list, follow_up_list

def read_MDLs(rootdir: str, current_msn_list: list):
    """
    Read all MDLs from 'rootdir', and create lists with Dataframes in order to create Follow-Up, DSOL and PS.

    Args:
    ----------
        rootdir:
            The path to the folder containing the MDLs in '.xlsx' format.
            
        current_msn_list:
            List of the MSNs for the current IPC/SRM revision.

    Returns:
    ----------
        mdl_msn_list:          
            List containing MSNs that where found in the MDLs.

        follow_up_list:     
            List containing DataFrames for each MSN in order to create Follow-Up.

        dsol_list:          
            List containing DataFrames for each MSN in order to create DSOL.

        ps_list:            
            List containing DataFrames for each MSN in order to create PS.

        nc_list:
            List containing DataFrames for each MSN in order to create NC.
    """
    mdl_msn_list = []
    dsol_list = []
    ps_list = []
    follow_up_list = []
    nc_list = []
    for root, _, files in os.walk(rootdir, topdown=True):
        for file in files:
            if not file.endswith('xlsx'): continue
            msn = file[:4]
            mdl_msn_list.append(msn)
            if msn in ['0835', '2737']:
                mdl_column = file.replace('.xlsx','').replace('349-', '')     # The MDL filename follows the format "0835_349-MDL-0835-G.xlsx"
            else:                                                                  
                mdl_column = file.replace('.xlsx','').replace('EFW-E-', '')     # The MDL filename follows the format "3708_EFW-E-MDL-00243-C.xlsx"

            # To ignore "UserWarning: Data Validation" and "UserWarning: Conditional Formatting"
            with warnings.catch_warnings():
                warnings.simplefilter(action='ignore', category=UserWarning)
                # Read Sheet 'Product Structure' (for All MSNs)
                df = pd.read_excel(root + os.sep + file, dtype=str, sheet_name='Product Structure')
                df = df[['PARENT NUMBER', 'LEVEL', 'CHILD NUMBER', 'CHILD TITLE', 'DIFF']]                          # Keep only these columns
                df = df.rename(columns={'DIFF': mdl_column})
                for column in df.columns: df[column] = df[column].str.strip()                                       # Strip leading and trailing whitespaces
                df = df.loc[ df['CHILD TITLE'].map(lambda x: False if re.findall(r'DELET|SALV', x) else True) ]     # Drop "Deleted" and "Salvage"
                df = df.loc[ df['CHILD NUMBER'].map(lambda x: False if re.findall(r'R6|R7', x) else True) ]         # Drop "R6" and "R7"
                ps_list.append(df)

                # Read Sheet 'Applicable Part List' to create DSOL (for All MSNs)
                df = pd.read_excel(root + os.sep + file, dtype=str, sheet_name='Applicable Part List')
                df = df[['PART NUMBER', 'PART TITLE', 'QTY', 'PART TYPE', 'PART ISSUE', 'DIFF']]    # Keep only these columns
                df = df.rename(columns={'DIFF': mdl_column})
                for column in df.columns: df[column] = df[column].str.strip()                       # Strip leading and trailing whitespaces
                dsol_list.append(df)

                # Create Follow-Up DataFrame from 'Applicable Part List' (only for New MSNs)
                if msn in current_msn_list:
                    df = df.loc[ df['PART TYPE'] == 'DSOL']
                    df = df.loc[ df['PART NUMBER'].map(lambda x: True if re.findall(r'R0|R1|R3', x) else False) ]       # Keep only "R0", "R1" and "R3"
                    df = df.drop(['QTY', 'PART TYPE', 'PART ISSUE'], axis=1)
                    follow_up_list.append(df)

                    # # Read Sheet 'Product Structure' (only for New MSNs)
                    # df = pd.read_excel(root + os.sep + file, dtype=str, sheet_name='Product Structure')
                    # df = df[['PARENT NUMBER', 'LEVEL', 'CHILD NUMBER', 'CHILD TITLE', 'DIFF']]                          # Keep only these columns
                    # df = df.rename(columns={'DIFF': mdl_column})
                    # for column in df.columns: df[column] = df[column].str.strip()                                       # Strip leading and trailing whitespaces
                    # df = df.loc[ df['CHILD TITLE'].map(lambda x: False if re.findall(r'DELET|SALV', x) else True) ]     # Drop "Deleted" and "Salvage"
                    # df = df.loc[ df['CHILD NUMBER'].map(lambda x: False if re.findall(r'R6|R7', x) else True) ]         # Drop "R6" and "R7"
                    # ps_list.append(df)

                    # Read Sheet 'Nonconformities' (only for New MSNs)
                    df = pd.read_excel(root + os.sep + file, dtype=str, sheet_name='Nonconformities')
                    df = df[['NUMBER', 'ISSUE', 'NC NUMBER', 'NC ISSUE', 'NC TITLE', 'DIFF']]           # Keep only these columns
                    df = df.rename(columns={'DIFF': mdl_column})
                    for column in df.columns: df[column] = df[column].str.strip()                       # Strip leading and trailing whitespaces
                    nc_list.append(df)

    return mdl_msn_list, follow_up_list, dsol_list, ps_list, nc_list

def merge_dfs(list_of_dfs: list):
    """
    Merge all the DataFrames from a list into a Single DataFrame.

    Args:
    ----------
        list_of_dfs:
            List containing the DataFrames to be merged.

    Returns:
    ----------
        df_merged:                  
            The merged DataFrame.
    """
    # Get "title_list"
    title_list = [x for x in list(list_of_dfs[0]) if not re.findall(r'^\d{4}_', x)]

    # Drop Duplicates here before df_merged becomes huge
    for i, _ in enumerate(list_of_dfs):
        list_of_dfs[i] = list_of_dfs[i].drop_duplicates() 
                            
    # Merge all Data frames from list (drop duplicates again just to be sure)
    df_merged = reduce(lambda left, right: pd.merge(left, right, on=title_list, how='outer'), list_of_dfs)
    df_merged = df_merged.drop_duplicates()                                                                     
    df_merged = df_merged.sort_values(by=title_list).fillna('')

    return df_merged

def add_effectivity_column(df_merged: pd.DataFrame, sheet: str, rev_msn_list: list = None, drop_empty_effectivity: bool = True):
    """
    Add effectivity column to a DataFrame that was created by "merge_dfs".

    Args:
    ----------
        df_merged:                  
            The merged DataFrame that was created by "merge_dfs".

        sheet:              
            Type of excel sheet to determine where to add effectivity column.
            Options: 'FOLLOW_UP', 'FOLLOW_UP_INITIAL', 'DSOL', 'PS', 'NC'

        rev_msn_list:               
            A list of the 90-Day Revivion MSNs. Add it only for 'NC' to keep the effectivity of the 'D' items.

        drop_empty_effectivity:     
            Set to 'False' only for 'NC' and when new MDLs are sent midway through the revision.
    
    Returns:
    ----------
        df_merged:                  
            The original DataFrame with the inserted effectivity column.
    """
    # Dict with 'name' and 'idx' of effectivity column to be added.
    if sheet not in ['FOLLOW_UP', 'FOLLOW_UP_INITIAL', 'DSOL', 'PS', 'NC']: sheet = 'DSOL'
    effect_column = EFFECT_COLUMN[sheet]

    # Get "title_list" and "msn_list"
    title_list = [x for x in list(df_merged.columns) if not re.findall(r'MDL', x)]
    msn_list = [x[:4] for x in list(df_merged.columns) if re.findall(r'MDL', x)]        

    # Replace '' with np.nan
    # Because it doesnt work correctly below?????????

    # Create Effectivity Column
    df_effect = df_merged.drop(title_list, axis=1)
    df_effect.columns = msn_list
    for msn in msn_list:
        if rev_msn_list and msn in rev_msn_list:
            # df_effect[msn] = df_effect[msn].apply(lambda x: msn if x != '' else np.nan)                                       # Original
            # df_effect[msn] = df_effect[msn].apply(lambda x: msn if ((x != '') or (pd.notna(x))) else np.nan)                  # One way to also work with NaN
            df_effect[msn] = df_effect[msn].apply(lambda x: msn if ['N', 'R', '-', 'D', 'WTF'] else np.nan)
        else:
            # df_effect[msn] = df_effect[msn].apply(lambda x: msn if x != '' and x != 'D' else np.nan)                          # Original
            # df_effect[msn] = df_effect[msn].apply(lambda x: msn if ((x != '') and (pd.notna(x))) and x != 'D' else np.nan)    # One way to also work with NaN
            df_effect[msn] = df_effect[msn].apply(lambda x: msn if x in ['N', 'R', '-', 'WTF'] else np.nan)
    df_effect = df_effect.apply(lambda x: ', '.join(x[x.notnull()]), axis = 1)

    # Insert Effectivity Column
    df_merged.insert(loc=effect_column['idx'], column=effect_column['name'], value=df_effect)

    # Drop rows with empty effectivity column
    # When New MDLs are sent midway through the revision, and we need to update the Follow-up,
    # we don't want to to drop these rows, espesially for NCs !!!!!!!!!!!!!!!
    if drop_empty_effectivity:
        df_merged = df_merged[df_merged[effect_column['name']].astype(bool)]

    return df_merged

def add_task_column(df: pd.DataFrame, rev_msn_list: list):
    """
    Add 'TASK' column with 'NEW MSNs' and 'REV OLD MSNs' to a DataFrame.

    Args:
    ----------
        df:
            One of the DataFrames: df_IPC, df_SRM_A321, df_SRM_A320.
        
        rev_msn_list:
            A list of the 90-Day Revivion MSNs.
    
    Returns:
    ----------
        df:
            The original DataFrame with the 'TASK' column added.
    """
    all_msn_column_names = [x for x in list(df.columns) if re.findall(r'^\d{4}', x)]
    new_msn_column_names = [x for x in all_msn_column_names if x[:4] not in rev_msn_list]
    task_column = df[new_msn_column_names].replace('', np.nan).any(axis=1)
    task_column = np.where(task_column, 'NEW MSNs', 'REV OLD MSNs')
    idx = 6 if 'CSN' in list(df.columns) else 3
    df.insert(loc=idx, column='TASK', value=task_column)
    return df

def add_columns_to_Follow_Up(df_follow_up: pd.DataFrame):
    """Add extra columns to Follow-up Dataframe"""

    empty_column_follow_up = ['' for i in range(df_follow_up.shape[0])]
    if 'CSN' not in list(df_follow_up.columns):
        for col in reversed(['CSN', 'Fig', 'Type']):
            df_follow_up.insert(loc=1, column=col, value=empty_column_follow_up)
    for col in reversed(['Author', 'Start Date', 'Status', 'Time (minutes)', 'Author CC', 'CC Time (minutes)', 'Comments', 'IPC CSN']):
        df_follow_up.insert(loc=7, column=col, value=empty_column_follow_up)

    return df_follow_up

def add_columns_to_PS(df_ps: pd.DataFrame):
    """Sort PS by Child. If there something that is both 'D' and 'N', keep the 'N'"""
    
    final_sort_list = ['CHILD NUMBER', 'PARENT NUMBER'] + list(df_ps)[5:]
    final_sort_ascending_list = [True, True] + [False for x in list(df_ps)[5:]]
    df_ps = df_ps.sort_values(by=final_sort_list, ascending=final_sort_ascending_list)
    df_ps = df_ps.drop_duplicates(subset=['PARENT NUMBER', 'LEVEL', 'CHILD NUMBER', 'CHILD TITLE'])
    df_ps = df_ps.reset_index(drop=True)

    return df_ps

def add_columns_to_NC(df_nc: pd.DataFrame):
    """Add extra columns to NC Dataframe"""
    empty_column_nc = ['' for i in range(df_nc.shape[0])]
    for col in reversed(['Author Check', 'Initial Status', 'Author Comment']):                        # I had an extra space like this before: ' Author Comment'
        df_nc.insert(loc=5, column=col, value=empty_column_nc)

    return df_nc
