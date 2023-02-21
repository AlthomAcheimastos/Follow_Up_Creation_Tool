import os
import warnings
import regex as re
import numpy as np
import pandas as pd

# import sys
# # sys.path.append('D:\\August_PySide2\\bin')
# sys.path.append('D:\\\Follow_Up_Creation_Tool\\bin')
# from setup_follow_up import read_JSON, merge_dfs, add_effectivity_column
# # from save_to_excel import add_formats_to_workbook, format_sheet_ALL_NCs, get_column_range
# from save_to_excel import all_NCs_to_excel


def read_MDLs_for_NCs(rootdir: str):
    """
    Read all MDLs from 'rootdir', and create lists with Dataframes in order to create NC.

    Args:
    ----------
        rootdir:
            The path to the folder containing the MDLs in '.xlsx' format.

    Returns:
    ----------
        mdl_msn_list:          
            List containing MSNs that where found in the MDLs.

        nc_dict:
            Dict containing DataFrames for each MSN in order to create NC.
    """
    mdl_msn_list = []
    nc_dict = {}
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

                # Read Sheet 'Nonconformities' (only for New MSNs)
                df = pd.read_excel(root + os.sep + file, dtype=str, sheet_name='Nonconformities')
                df = df[['NUMBER', 'ISSUE', 'NC NUMBER', 'NC ISSUE', 'NC TITLE', 'DIFF']]           # Keep only these columns
                for column in df.columns: df[column] = df[column].str.strip()                       # Strip leading and trailing whitespaces
                ### df['DIFF'] = df['DIFF'].replace(' ', np.nan)                                        # Read as NaN, this is already NaN
                df = df.rename(columns={'DIFF': mdl_column})
                nc_dict[msn] = df

    return mdl_msn_list, nc_dict



def update_90_day_rev(nc_dict_new: dict, nc_dict_old: dict, rev_msn_list: list):
    """
    Update the DataFrames of the 90-Day Revision MSNs inside 'nc_dict_new' based on 'nc_dict_old'.

    Find Phantom-New (PN) and Phantom-Deleted (PD)

    Args:
    ----------
        nc_dict_new:
            Dict containing DataFrames for each MSN from All MDLs.

        nc_dict_old:
            Dict containing DataFrames for each MSN from Old MDLs (only for those that will be revised).

    Returns:
    ----------
        nc_dict_new:
            Updated dict containing DataFrames for each MSN from All MDLs.
    """
    # Get title list
    title_list = [x for x in list(nc_dict_old[rev_msn_list[0]]) if not re.findall(r'^\d{4}_', x)]

    for msn in rev_msn_list:
        # Get DataFrames from dicts
        df_old = nc_dict_old[msn]
        df_new = nc_dict_new[msn]
        
        # Get column names
        mdl_old = list(df_old.columns)[-1]
        mdl_new = list(df_new.columns)[-1]

        # In case a 90-Day MDL has not changed revision, then skip it
        if mdl_old == mdl_new: continue

        # Merge "df_old" and "df_new"
        df_merged = pd.merge(df_old, df_new, on=title_list, how='outer')
        df_merged = df_merged.drop_duplicates()                                                                     
        df_merged = df_merged.sort_values(by=title_list)

        # Find "Phantom-New" and "Phantom-Deleted" EAs and DCNs
        df_merged[mdl_new] = df_merged[mdl_new].fillna('PD')
        df_merged[mdl_new] = df_merged.apply(lambda row: 'PN' if (row[mdl_new] == 'R' or row[mdl_new] == '-') and (row[mdl_new] == 'D') else row[mdl_new], axis=1)
        df_merged[mdl_new] = df_merged.apply(lambda row: 'PN' if (row[mdl_new] == 'R' or row[mdl_new] == '-') and (pd.isna(row[mdl_old])) else row[mdl_new], axis=1)

        # Replace "nc_dict_new" with the df_merged
        nc_dict_new[msn] = df_merged.drop([mdl_old], axis=1)

    return nc_dict_new


def replace_letters_with_MSNs(df_nc: pd.DataFrame, mdl_list: list):
    """
    Replace Letters with MSN and letter. i.e. 'N' -> '1207 (N)'
    
    Args:
    ----------
        df_nc:
            The DataFrame with ALL the NCs.
        
        mdl_list:
            List of MDLs.
    
    Returns:
    ----------
        df_nc:
            The initial DataFrame the replaced letters.
    """
    for mdl in mdl_list:
        msn = mdl[:4]
        df_nc[mdl] = df_nc[mdl].replace('N', f'{msn} (N)')
        df_nc[mdl] = df_nc[mdl].replace('-', f'{msn} (-)')
        df_nc[mdl] = df_nc[mdl].replace('R', f'{msn} (R)')
        df_nc[mdl] = df_nc[mdl].replace('D', f'{msn} (D)')
        df_nc[mdl] = df_nc[mdl].replace('PD', f'{msn} (PD)')
        df_nc[mdl] = df_nc[mdl].replace('PN', f'{msn} (PN)')
    
    return df_nc


# def main(filepath_mdl_new, filepath_mdl_old, filepath_json, revision):
#     # Read old MDLs for 90-Day Revisions
#     # Check that 90-Day Revisions have both old and new MDLs

#     # Read JSON with MSNs
#     json_MSNs = read_JSON(filepath_json)
#     new_msn_list = json_MSNs['new']
#     rev_msn_list = json_MSNs['rev']

#     # Read ALL Latest MDLs
#     mdl_msn_list_new, nc_dict_new = read_MDLs_for_NCs(filepath_mdl_new)

#     # Read OLD MDLs for 90-Day Revisions
#     mdl_msn_list_old, nc_dict_old = read_MDLs_for_NCs(filepath_mdl_old)
#     nc_dict_new = update_90_day_rev(nc_dict_new, nc_dict_old, rev_msn_list)

#     # Merge DataFrames
#     nc_list = list(nc_dict_new.values())
#     df_nc = merge_dfs(nc_list)

#     # Get lists
#     mdl_list = [x for x in list(df_nc.columns) if re.findall(r'^\d{4}_', x)]
#     current_mdl_list = [x for x in list(df_nc.columns) if (x[:4] in new_msn_list) or (x[:4] in rev_msn_list)]
#     old_mdl_list = list(set(mdl_list) - set(current_mdl_list))

#     # Replace Letters with MSN and letter. i.e. 'N' -> '1207 (N)'
#     df_nc = replace_letters_with_MSNs(df_nc, mdl_list)

#     # Get only Current NCs
#     df_nc_RXX = df_nc.drop(old_mdl_list, axis=1)
#     df_nc_RXX = df_nc_RXX.replace('', np.nan).dropna(subset=current_mdl_list, how='all')

#     # Save to excel
#     all_NCs_to_excel(df_nc, df_nc_RXX, new_msn_list, rev_msn_list, revision)



# if __name__ == '__main__':
#     # When new MDLs came for MSN 1199/1207/1713
#     # filepath_json = 'D:\\August_PySide2\\__Follow-up R10-12_12_22_New_MDLs\\MSN_INPUT_R10.json'
#     # filepath_mdl_new = 'D:\\August_PySide2\\__Follow-up R10-12_12_22_New_MDLs\\R10_MDLs_NEW_1199_1207_1713'
#     # filepath_mdl_old = 'D:\\August_PySide2\\__Follow-up R10-12_12_22_New_MDLs\\90-Day_Previous_MDLs'
#     # revision = 'R10_v2_TESTs'

#     filepath_json = 'D:\\August_PySide2\\__Follow-up R10-12_12_22_New_MDLs\\MSN_INPUT_R10.json'
#     filepath_mdl_new = 'D:\\Follow_Up_Creation_Tool\\__Follow-up R10-21_12_22\\R10_MDLs_NEW_1094_2060'
#     filepath_mdl_old = 'D:\\Follow_Up_Creation_Tool\\__Follow-up R10-21_12_22\\90-Day_Previous_MDLs'
#     revision = 'R10_v3'

#     main(filepath_mdl_new, filepath_mdl_old, filepath_json, revision)
#     pass