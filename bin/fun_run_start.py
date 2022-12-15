##########################################################################################
# Filename:     fun_run_start.py
# For:          Follow_Up_Creation_Tool
# Author:       Spyros Acheimastos (acheimastos@althom.eu)
# Date:         15/12/2022
##########################################################################################

import os
from PySide2.QtCore import *
from bin.pseudo_db import (
    follow_up_to_pseudo_db,
    merge_pseudo_dbs,
    create_pseudo_db_for_CC,
    gnrt_lines_and_split,
    read_pseudo_db
)
from bin.setup_follow_up import (
    read_JSON,
    read_JSON_authors,
    read_MDLs_current, 
    read_MDLs, 
    merge_dfs, 
    add_effectivity_column, 
    add_task_column, 
    add_columns_to_Follow_Up, 
    add_columns_to_PS, 
    add_columns_to_NC
)
from bin.save_to_excel import (
    pseudo_db_to_excel,
    initial_follow_up_to_excel,
    final_follow_up_to_excel
)

from bin.partials import (
    get_follow_ups,
    get_MSNs_and_MDLs, 
    add_PNs_to_df_old, 
    reduce_df_new,
    update_MDLs_in_df_old,
    read_PS_DSOL_NC
)

from bin.create_json import (
    create_json_msns,
    create_json_authors
)

SCRIPT_DIRECTORY = os.path.dirname(os.path.abspath(__file__))

# I can save values I want to store inside Object
# if I return them as dictionaries and use the function "save_result"
# return {'result_1': df_one_follow_up, 'result_2': (2, 3), 'result_3': {'my_key': 'my_value'}}

def fun_generate_authors_start(console: Signal = Signal('')):
    console.emit('Generating JSON with Authors.')
    create_json_authors()
    console.emit('---> Finished.')


def fun_generate_msns_start(console: Signal = Signal('')):
    console.emit('Generating JSON with MSNs.')
    create_json_msns()
    console.emit('---> Finished.')


def fun_run_0_start(filepath_one_follow_up: str, console: Signal = Signal('')):
    """
    Call the functions for Step-0/Create PseudoDB of 'Extra'
    """
    console.emit('Converting Follow-up to PseudoDataBase format.')
    df_one_follow_up = follow_up_to_pseudo_db(filepath_one_follow_up)

    console.emit('Saving New PseudoDataBase.')
    pseudo_db_to_excel(df_one_follow_up, 'NEW_PSEUDODATABASE.xlsx')

    console.emit('---> Finished.')


def fun_run_1_start(filepath_latest_follow_up: str, filepath_pseudo_db_1: str, console: Signal = Signal('')):
    """
    Call the functions for Step-1 of 'Create Follow-up'
    """
    console.emit('Converting Latest Follow-up to PseudoDataBase format.')
    df_latest_follow_up = follow_up_to_pseudo_db(filepath_latest_follow_up)

    console.emit('Reading Latest PseudoDataBase.')
    df_pseudo_db_1 = read_pseudo_db(filepath_pseudo_db_1)

    console.emit('Merging PseudoDataBases.')
    df_pseudo_db = merge_pseudo_dbs([df_latest_follow_up, df_pseudo_db_1])

    console.emit('Saving New PseudoDataBase.')
    pseudo_db_to_excel(df_pseudo_db, filepath_pseudo_db_1.replace('.xlsx', '_merged.xlsx'))

    console.emit('---> Finished.')


def fun_run_2_start(filepath_json: str, filepath_mdl: str, filepath_pseudo_db_2: str, console: Signal = Signal('')):
    """
    Call the functions for Step-2 of 'Create Follow-up' and Step-1 of 'Update Follow-up'
    """
    # Read JSON
    console.emit('Reading JSON file with MSNs.')
    json_MSNs = read_JSON(filepath_json)
    all_msn_list = json_MSNs['all']
    current_msn_list = json_MSNs['new'] + json_MSNs['rev']
    rev_msn_list = json_MSNs['rev']
    current_A320_msn_list = [x for x in current_msn_list if x in json_MSNs['all_A320']]

    # Read Current MDLs
    console.emit('Reading current MDLs.')
    mdl_msn_list, follow_up_list = read_MDLs_current(filepath_mdl, current_msn_list)

    # Read PseudoDataBase
    console.emit('Reading PseudoDataBase.')
    df_pseudo_db_2 = read_pseudo_db(filepath_pseudo_db_2)

    # Check that MDL names and MSNs inside JSON match
    missing_mdl_list = [x for x in all_msn_list if x not in mdl_msn_list]
    missing_json_list = [x for x in mdl_msn_list if x not in all_msn_list]
    if missing_mdl_list:
        return console.emit('The MDLs of the MSNs ' + ', '.join(missing_mdl_list) + ' are missing. Fix this error and run again.')
    if missing_json_list:
        return console.emit('The MSNs ' + ', '.join(missing_json_list) + ' are missing from the JSON file. Fix this error and run again.')

    # Merge Follow-up DataFrames
    console.emit('Merging Initial Follow-Up.')
    df_initial = merge_dfs(follow_up_list)     

    # Create and Save new PseudoDataBase
    console.emit('Saving new PseudoDataBase.')
    df_new_pseudo_db = create_pseudo_db_for_CC(df_initial, df_pseudo_db_2)
    pseudo_db_to_excel(df_new_pseudo_db, filepath_pseudo_db_2.replace('.xlsx', '_for_CC.xlsx'))

    # Create and Save Initial Follow-up
    console.emit('Saving Initial Follow-up.')
    df_initial_for_excel = add_effectivity_column(df_initial, 'FOLLOW_UP_INITIAL')
    df_initial_for_excel = add_task_column(df_initial_for_excel, rev_msn_list)
    df_initial_for_excel = add_columns_to_Follow_Up(df_initial_for_excel)
    initial_follow_up_to_excel(df_initial_for_excel, 'Follow-up_Initial.xlsx')

    console.emit('---> Finished.')
    console.emit('> Categorize new Part Numbers (TBDs) that were added to the PseudoDataBase using "Follow-up_Initial.xlsx", and then continue.')


def fun_run_3_start(filepath_json: str, filepath_mdl: str, filepath_pseudo_db: str, excelfilepath: str, filepath_json_authors: str = None, add_QBs: bool = True, console: Signal = Signal('')):
    """
    Call the functions for Step-3 of 'Create Follow-up' and Step-2 of 'Update Follow-up'
    """
    # Read JSON MSNs
    console.emit('Reading JSON file with MSNs.')
    json_MSNs = read_JSON(filepath_json)
    all_msn_list = json_MSNs['all']
    current_msn_list = json_MSNs['new'] + json_MSNs['rev']
    rev_msn_list = json_MSNs['rev']
    current_A320_msn_list = [x for x in current_msn_list if x in json_MSNs['all_A320']]

    # Read JSON Authors
    if filepath_json_authors:
        console.emit('Reading JSON file with Authors.')
        authors_dict = read_JSON_authors(filepath_json_authors)
    else:
        authors_dict = None
        
    # Read MDLs
    console.emit('Reading all MDLs.')
    mdl_msn_list, follow_up_list, dsol_list, ps_list, nc_list = read_MDLs(filepath_mdl, current_msn_list)

    # Read PseudoDataBase
    console.emit('Reading PseudoDataBase after human Cross Check.')
    # df_pseudo_db = read_pseudo_db(filepath_pseudo_db.replace('.xlsx', '_for_CC.xlsx'))
    df_pseudo_db = read_pseudo_db(filepath_pseudo_db)

    # Check that MDL names and MSNs inside JSON match
    missing_mdl_list = [x for x in all_msn_list if x not in mdl_msn_list]
    missing_json_list = [x for x in mdl_msn_list if x not in all_msn_list]
    if missing_mdl_list:
        return console.emit('The MDLs of the MSNs ' + ', '.join(missing_mdl_list) + ' are missing. Fix this error and run again.')
    if missing_json_list:
        return console.emit('The MSNs ' + ', '.join(missing_json_list) + ' are missing from the JSON file. Fix this error and run again.')

    # Merge DataFrames
    console.emit('Merging Initial Follow-Up')
    df_initial = merge_dfs(follow_up_list)  
    console.emit('Merging DSOL')
    df_dsol = merge_dfs(dsol_list)
    console.emit('Merging PS')
    df_ps = merge_dfs(ps_list)
    console.emit('Merging NC')
    df_nc = merge_dfs(nc_list)

    # Generate new lines
    console.emit('Generating new lines and splitting Follow-Up into IPC and SRM.')
    df_IPC, df_SRM_A321, df_SRM_A320 = gnrt_lines_and_split(df_initial, df_pseudo_db, current_A320_msn_list)

    # Adding additional columns
    console.emit('Adding effectivity and additional columns to "IPC", "SRM", "DSOL", "PS" and "NC"')

    # Add Effectivity column
    df_IPC = add_effectivity_column(df_IPC, 'FOLLOW_UP')
    df_SRM_A321 = add_effectivity_column(df_SRM_A321, 'FOLLOW_UP')
    df_dsol = add_effectivity_column(df_dsol, 'DSOL')
    df_ps = add_effectivity_column(df_ps, 'PS')
    df_nc = add_effectivity_column(df_nc, 'NC', rev_msn_list)

    # Add TASK column
    df_IPC = add_task_column(df_IPC, rev_msn_list)
    df_SRM_A321 = add_task_column(df_SRM_A321, rev_msn_list)

    # Add extra columns
    df_IPC = add_columns_to_Follow_Up(df_IPC)
    df_SRM_A321 = add_columns_to_Follow_Up(df_SRM_A321)
    df_ps = add_columns_to_PS(df_ps)
    df_nc = add_columns_to_NC(df_nc)

    console.emit('Saving final Follow-up.')

    # Dict with kwargs for 'final_follow_up_to_excel'
    dict_with_follow_ups = {
        'IPC': df_IPC,
        'SRM_A321': df_SRM_A321
    }

    # If there are any A320 add sheet for 'SRM_A320'
    if current_A320_msn_list:
        df_SRM_A320 = add_effectivity_column(df_SRM_A320, 'FOLLOW_UP')
        df_SRM_A320 = add_task_column(df_SRM_A320, rev_msn_list)
        df_SRM_A320 = add_columns_to_Follow_Up(df_SRM_A320)
        dict_with_follow_ups['SRM_A320'] = df_SRM_A320

    # Save to Excel
    # excelfilepath = f'EFW Follow-up R{revision}.xlsx'
    final_follow_up_to_excel(df_dsol, df_ps, df_nc, excelfilepath, authors_dict=authors_dict, add_QBs=add_QBs, **dict_with_follow_ups)

    console.emit('---> Finished.')
    if add_QBs is True:
        console.emit('> Be carefull with cell ranges if you manually add drop down lists.')
        console.emit('> Manually set formatting of dates to DD/MM/YYYY.')
        console.emit('> Manually add any other Sheets.')
    else:
        console.emit('> Do not change anything inside the TEMPORARY Follow-up.')
        console.emit('> Just use it for the next step.')


def fun_run_8_start(filepath_json: str, filepath_json_authors: str, filepath_old: str, filepath_new: str, excelfilepath: str, add_QBs: bool = False, console: Signal = Signal('')):
    """
    Call the functions for Step-3 of 'Update Follow-up'
    """
    # Read JSON
    console.emit('Reading JSON file with MSNs.')
    json_MSNs = read_JSON(filepath_json)
    # all_msn_list = json_MSNs['all']
    # current_msn_list = json_MSNs['new'] + json_MSNs['rev']
    rev_msn_list = json_MSNs['rev']
    # current_A320_msn_list = [x for x in current_msn_list if x in json_MSNs['all_A320']]

    # Read JSON Authors
    console.emit('Reading JSON file with Authors.')
    authors_dict = read_JSON_authors(filepath_json_authors)

    # Read old and new Follow-up
    console.emit('Reading Old and New Follow-up.')
    df_dict_old = get_follow_ups(filepath_old)
    df_dict_new = get_follow_ups(filepath_new)

    # Check that keys match
    if df_dict_old.keys() != df_dict_new.keys():
        console.emit('---> Error: Sheet names dont match between Excel files.')
        return

    # Initialize 
    dict_with_follow_ups = {}

    # Loop for 'IPC', 'SRM A321', 'SRM A320'
    for (k1, df_old), (k2, df_new) in zip(df_dict_old.items(), df_dict_new.items()):
        if k1 != k2:
            console.emit(f'---> Error: k1="{k1}" but k2="{k2}"')
            return

        console.emit(f'Merging {k1}')
        
        # Find MDLs that have changed revision
        msn_list, mdl_dict_old, mdl_dict_new = get_MSNs_and_MDLs(list(df_old), list(df_new))

        # Rename the MDL columns of "df_old"
        for msn in msn_list:
            df_old = df_old.rename(columns={mdl_dict_old[msn]: mdl_dict_new[msn]})

        # For simplicity
        mdl_dict = mdl_dict_new

        # Add New Part Numbers added at the bottom of the "df_old"
        df_old = add_PNs_to_df_old(df_old, df_new, mdl_dict)

        # Keep only unique Part Numbers
        df_new = reduce_df_new(df_new, mdl_dict)

        # #
        # # To test Phantom Deleted 
        # df_old['PART NUMBER'] = df_old['PART NUMBER'].replace('D113R1202-004-00', 'D999R9999-999-99')       # FOR TESTING
        # #

        # Get 
        df_final = update_MDLs_in_df_old(df_old, df_new, mdl_dict)
        
        # Add Effectivity and Task column
        df_final = add_effectivity_column(df_final, 'FOLLOW_UP', drop_empty_effectivity=False)          # Keeping empty effectivity just in case
        df_final = add_task_column(df_final, rev_msn_list)

        # Add DataFrame to Dict
        dict_with_follow_ups[k1] = df_final

    # Read PS, DSOL, NC from New Follow-up
    console.emit('Reading PS, DSOL, NC from New Follow-up')
    df_dsol, df_ps, df_nc = read_PS_DSOL_NC(filepath_new)

    # Save to excel
    console.emit('Saving final Follow-up.')
    final_follow_up_to_excel(df_dsol, df_ps, df_nc, excelfilepath, authors_dict=authors_dict, add_QBs=add_QBs, **dict_with_follow_ups)
    console.emit('---> Finished.')
    console.emit('> Be carefull with cell ranges if you manually add drop down lists.')
    console.emit('> Manually set formatting of dates to DD/MM/YYYY.')
    console.emit('> Manually add any other Sheets.')
    console.emit('> Use "MSN Change" columns at the far right to manually colour the cells.')