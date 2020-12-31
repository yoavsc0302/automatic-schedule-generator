# This module contains all the excel modification functions

import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter
import pickle
import math
import tkinter as tk

# Show entire df when printed
pd.set_option("display.max_rows", None, "display.max_columns", None)

# Read every sheet in the justice board file and make a df out of it
makel_officer_df = pd.read_excel('justice_board.xlsx',
                                 sheet_name='Makel Officer',
                                 engine='openpyxl', index_col=0)
makel_operator_df = pd.read_excel('justice_board.xlsx',
                                  sheet_name='Makel Operator',
                                  engine='openpyxl', index_col=0)
manager_df = pd.read_excel('justice_board.xlsx', sheet_name='Manager',
                           engine='openpyxl', index_col=0)
samba_df = pd.read_excel('justice_board.xlsx', sheet_name='Samba',
                         engine='openpyxl', index_col=0)


def get_justice_sheets_as_df():
    """
    Read all the sheets in the justice board file and make df's out of
    each one
    :return: dict of all the df's
    """
    makel_officer_df = pd.read_excel('justice_board.xlsx',
                                     sheet_name='Makel Officer',
                                     engine='openpyxl', index_col=0)
    makel_operator_df = pd.read_excel('justice_board.xlsx',
                                      sheet_name='Makel Operator',
                                      engine='openpyxl', index_col=0)
    manager_df = pd.read_excel('justice_board.xlsx', sheet_name='Manager',
                               engine='openpyxl', index_col=0)
    samba_df = pd.read_excel('justice_board.xlsx', sheet_name='Samba',
                             engine='openpyxl', index_col=0)
    return {'Makel Officer': makel_officer_df,
            'Makel Operator': makel_operator_df,
            'Manager': manager_df,
            'Samba': samba_df}


def get_ilutzim_sheets_as_df():
    """
    Read all the sheets in the ilutzim file and make df's out of
    each one
    :return: dict of all the df's
    """
    ilutzim_makel_officer_df = pd.read_excel('ilutzim.xlsx',
                                             sheet_name='Makel Officer',
                                             engine='openpyxl', index_col=0,
                                             header=[0, 1])
    ilutzim_makel_operator_df = pd.read_excel('ilutzim.xlsx',
                                              sheet_name='Makel Operator',
                                              engine='openpyxl', index_col=0,
                                              header=[0, 1])
    ilutzim_manager_df = pd.read_excel('ilutzim.xlsx', sheet_name='Manager',
                                       engine='openpyxl', index_col=0)
    ilutzim_samba_df = pd.read_excel('ilutzim.xlsx', sheet_name='Samba',
                                     engine='openpyxl', index_col=0)

    return {'Makel Officer': ilutzim_makel_officer_df,
            'Makel Operator': ilutzim_makel_operator_df,
            'Manager': ilutzim_manager_df,
            'Samba': ilutzim_samba_df}


def create_ilutzim_excel():
    """
    Create the ilutzim excel file as a 'Multiply indexed DataFrame'
    source:https://jakevdp.github.io/PythonDataScienceHandbook/03.05-
    hierarchical-indexing.html for each population
    :param makel_names: list that contains the names of every 'makel'
    :param manager_names: list that contains the names of every 'manaager'
    :param samba_names: list that contains the names of every 'samab'
    """

    # Makel Officer:
    index_tuples = []
    for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday']:
        for team in ['1', '2', '3+4']:
            index_tuples.append([day, team])

    index = pd.MultiIndex.from_tuples(index_tuples, names=["Day", "Team"])
    makel_officer_df = pd.DataFrame(columns=index)
    makel_officer_df['Name'] = ''
    makel_officer_df.set_index('Name', inplace=True)

    # Makel Operator:
    # Hierarchical indices and columns
    makel_operator_df = pd.DataFrame(columns=index)
    makel_operator_df['Name'] = ''
    makel_operator_df.set_index('Name', inplace=True)

    # Manager df:
    columns = ['Name', 'Sunday', 'Monday', 'Tuesday', 'Wednesday']
    manager_df = pd.DataFrame(columns=columns)

    # Samba df:
    samba_df = pd.DataFrame(columns=columns)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('ilutzim.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    makel_officer_df.to_excel(writer, sheet_name='Makel Officer')
    makel_operator_df.to_excel(writer, sheet_name='Makel Operator')
    manager_df.to_excel(writer, sheet_name='Manager')
    samba_df.to_excel(writer, sheet_name='Samba')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()


def create_justice_board_excel():
    """
    Create the justice board excel file
    """

    # Makel officer df:
    columns = ['Name', '1', '2', '3+4']
    makel_officer_df = pd.DataFrame(columns=columns)

    # Makel operator df:
    makel_operate_df = pd.DataFrame(columns=columns)

    # Manager df:
    columns = ['Name', 'Sum']
    manager_df = pd.DataFrame(columns=columns)

    # Samba df:
    columns = ['Name', 'Sum', 'Samba', 'Fast caller and Toran']
    samba_df = pd.DataFrame(columns=columns)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('justice_board.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    makel_officer_df.to_excel(writer, sheet_name='Makel Officer')
    makel_operate_df.to_excel(writer, sheet_name='Makel Operator')
    manager_df.to_excel(writer, sheet_name='Manager')
    samba_df.to_excel(writer, sheet_name='Samba')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()


def create_file_location_csv():
    """
    Create csv thats stores the ilutzim and the justice board files locations'
    """
    files_location_df = pd.DataFrame({'ilutzim': ['i location'],
                                      'justice_board': ['jb location']})
    files_location = files_location_df.to_csv('files_location.csv')


def create_tzevet_conan_excel():
    """
    Create an excel file of the tzevet conan
    """

    # Define columns and index names:
    columns = ['Sunday', 'Monday', 'Tuesday', 'Wednesday']
    index = ['Manager', 'Samba', 'Fast caller', 'Toran',
             'Officer 1', 'Officer 2', 'Officer 3', 'Officer 4',
             'Operator 1', 'Operator 2', 'Operator 3', 'Operator 4']
    tzevet_conan_df = pd.DataFrame(columns=columns, index=index, data='empty')

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('tzevet_conan.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    tzevet_conan_df.to_excel(writer, sheet_name='Tzevet Conan')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()


def add_new_person(name, manager_var, makel_officer_var, makel_operator_var,
                   samba_var, fast_and_toran_var, warning_label):
    """
    Add a new person to the right sheets according to the jobs he can do:
    manager/makel officer/makel opperator/samba/fast and toran
    :param name: the name of the new person
    :param manager_var: equals 1 if the manager job checkbox is checked
    :param makel_officer_var: equals 1 if the makel officer job checkbox
     is checked
    :param makel_operator_var: equals 1 if the makel operator job checkbox
     is checked
    :param samba_var: equals 1 if the samba job checkbox is checked
    :param fast_and_toran_var: equals 1 if the fast and toran job checkbox
     is checked
    :param warning_label: the warning label in the window that it's text
     will apperat in case of some kind of error
    """

    try:
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        with pd.ExcelWriter('justice_board.xlsx', engine='openpyxl', mode='a') \
                as writer:
            workbook = writer.book

        # Read each sheet in the justice board file and make a df out of iy
        makel_officer_df = pd.read_excel('justice_board.xlsx',
                                         sheet_name='Makel Officer',
                                         engine='openpyxl', index_col=0)
        makel_operator_df = pd.read_excel('justice_board.xlsx',
                                          sheet_name='Makel Operator',
                                          engine='openpyxl', index_col=0)
        manager_df = pd.read_excel('justice_board.xlsx',
                                   sheet_name='Manager',
                                   engine='openpyxl', index_col=0)
        samba_df = pd.read_excel('justice_board.xlsx',
                                 sheet_name='Samba',
                                 engine='openpyxl', index_col=0)

        # If the new person is a 'makel officer', insert his name into
        # the 'makel officer' sheet and set his sum to the average of
        # everybodies' sum
        if makel_officer_var.get() == 1:
            workbook.remove(workbook['Makel Officer'])
            try:
                sum_to_be_set_1 = math.floor(
                    makel_officer_df['1'].mean())
                sum_to_be_set_2 = math.floor(
                    makel_officer_df['2'].mean())
                sum_to_be_set_3_4 = math.floor(
                    makel_officer_df['3+4'].mean())
                makel_officer_df = makel_officer_df.append(
                    {'Name': name,
                     '1': sum_to_be_set_1,
                     '2': sum_to_be_set_2,
                     '3+4': sum_to_be_set_3_4},
                    ignore_index=True)

            except:  # If this is the first person in the sheet
                makel_officer_df = makel_officer_df.append(
                    {'Name': name,
                     '1': 0,
                     '2': 0,
                     '3+4': 0},
                    ignore_index=True)

            add_new_person_to_ilutzim('Makel Officer', name)

            # Write dataframe to the worksheet.
            makel_officer_df.to_excel(writer, sheet_name='Makel Officer')

        # If the new person is a 'makel operator', insert his name into
        # the 'makel operator' sheet and set his sum to the average of
        # everybodies' sum
        if makel_operator_var.get() == 1:

            workbook.remove(workbook['Makel Operator'])
            try:
                sum_to_be_set_1 = math.floor(
                    makel_operator_df['1'].mean())
                sum_to_be_set_2 = math.floor(
                    makel_operator_df['2'].mean())
                sum_to_be_set_3_4 = math.floor(
                    makel_operator_df['3+4'].mean())
                makel_operator_df = makel_operator_df.append(
                    {'Name': name,
                     '1': sum_to_be_set_1,
                     '2': sum_to_be_set_2,
                     '3+4': sum_to_be_set_3_4},
                    ignore_index=True)

            except:  # If this is the first person in the sheet
                makel_operator_df = makel_operator_df.append(
                    {'Name': name,
                     '1': 0,
                     '2': 0,
                     '3+4': 0},
                    ignore_index=True)

            add_new_person_to_ilutzim(job='Makel Operator', name=name)

            # Write dataframe to the worksheet.
            makel_operator_df.to_excel(writer, sheet_name='Makel Operator')

        # If the new person is a 'manager', insert his name into
        # the 'manager' sheet and set his sum to the average of everybodies' sum
        if manager_var.get() == 1:
            workbook.remove(workbook['Manager'])
            try:
                sum_to_be_set = math.floor(manager_df['Sum'].mean())
            except:  # If this is the first person in the sheet
                sum_to_be_set = 0
            manager_df = manager_df.append({'Name': name, 'Sum': sum_to_be_set},
                                           ignore_index=True)
            # Write dataframe to the worksheet.
            manager_df.to_excel(writer, sheet_name='Manager')

            # add_new_person_to_ilutzim('Manager', name)
            add_new_person_to_ilutzim(job='Manager', name=name)

        # If the new person is either a 'samba' or 'fast and toran',
        # insert his name into the 'samba' sheet and specify in the 'samba' and
        # 'fast and toran' columns what he is by 'TRUE' and 'False' values
        # Toran and Toran+Samba share the same mean
        if fast_and_toran_var.get() == 1:
            try:
                sum_to_be_set = math.floor(samba_df
                                           [samba_df['Fast caller and Toran']
                                            == True]['Sum'].mean())

            except:  # If this is the first person of this kind in the sheet
                sum_to_be_set = 0

            if samba_var.get() == 0:
                samba_df = samba_df.append({'Name': name, 'Sum': sum_to_be_set,
                                            'Samba': False,
                                            'Fast caller and Toran': True},
                                           ignore_index=True)
            else:
                samba_df = samba_df.append({'Name': name, 'Sum': sum_to_be_set,
                                            'Samba': True,
                                            'Fast caller and Toran': True},
                                           ignore_index=True)
            # Write dataframe to the worksheet.
            workbook.remove(workbook['Samba'])
            samba_df.to_excel(writer, sheet_name='Samba')

            add_new_person_to_ilutzim(job='Samba', name=name)
        else:
            if samba_var.get() == 1:
                try:
                    sum_to_be_set = math.floor(
                        samba_df[(samba_df['Fast caller and Toran'] == False)
                                 & (samba_df['Samba'] == True)]['Sum'].mean())
                except:  # If this is the first person of this kind in the sheet
                    sum_to_be_set = 0
                samba_df = samba_df.append(
                    {'Name': name, 'Sum': sum_to_be_set, 'Samba': True,
                     'Fast caller and Toran': False}, ignore_index=True)
                workbook.remove(workbook['Samba'])
                samba_df.to_excel(writer, sheet_name='Samba')

                add_new_person_to_ilutzim(job='Samba', name=name)

        # Save the justice board file
        writer.save()
        warning_label['text'] = ''
        warning_label['bg'] = None


    except:

        # Show a warning in the edit people window
        warning_label['text'] = 'אזהרה: הקובץ של \nלוח הצדק פתוח.\n אנא סגור ' \
                                'אותו \nכדי שיתאפשר \nלשמור את השינויים!'
        warning_label['bg'] = 'red'


def add_new_person_to_ilutzim(job, name):
    """
    Add the new person to the ilutzim file
    :param job: the jobs that the person does ('Makel Officer, Makel Operator,
    Manager, Samba)
    :param name: the name of the new person
    """

    # Dictionairy containing the df of each sheet in the ilutzim file
    dict_of_df = get_ilutzim_sheets_as_df()

    # Getting the df's from the dicionairy
    makel_officer_df_ilutzim = dict_of_df['Makel Officer']
    makel_operator_df_ilutzim = dict_of_df['Makel Operator']
    manager_df_ilutzim = dict_of_df['Manager']
    samba_df_ilutzim = dict_of_df['Samba']


    with pd.ExcelWriter('ilutzim.xlsx', engine='openpyxl', mode='a') \
            as writer:
        workbook = writer.book

    if job == 'Makel Officer':
        makel_officer_df_ilutzim.loc[name, :] = '0'

    if job == 'Makel Operator':
        makel_operator_df_ilutzim.loc[name, :] = '0'

    if job == 'Manager':
        manager_df_ilutzim = manager_df_ilutzim.append({'Name': name,
                                                        'Sunday': '0',
                                                        'Monday': '0',
                                                        'Tuesday': '0',
                                                        'Wednesday': '0'},
                                                       ignore_index=True)
        dict_of_df['Manager'] = manager_df_ilutzim

    if job == 'Samba':
        samba_df_ilutzim = samba_df_ilutzim.append({'Name': name,
                                                    'Sunday': '0',
                                                    'Monday': '0',
                                                    'Tuesday': '0',
                                                    'Wednesday': '0'},
                                                   ignore_index=True)
        dict_of_df['Samba'] = samba_df_ilutzim

    workbook.remove(workbook[job])
    dict_of_df[job].to_excel(writer, sheet_name=job)
    writer.save()


def get_list_of_all_people():
    """
    Go over all sheets in the justice board file, grab all names and remove
    duplicated names
    :return: list of names in the justice board file
    """
    dict_of_df = get_justice_sheets_as_df()
    list_of_all_df = [dict_of_df['Makel Officer'],
                      dict_of_df['Makel Operator'],
                      dict_of_df['Manager'],
                      dict_of_df['Samba']]
    names_of_all_people = []
    for df in list_of_all_df:
        names_of_all_people += df['Name'].values.tolist()
    names_of_all_people = list(set(names_of_all_people))
    return names_of_all_people


def delete_person(name_of_person, warning_label, chosen_option,
                  edit_people_window, list_if_empty):
    """
    Delete the person from the ilutzim and the justice board file,
    by calling the functions:
        delete_person_from_justice_board(name_of_person)
        delete_person_from_ilutzim(name_of_person)
    :param name_of_person: the person to delete
    :param warning_label: the warning label that will pop if there was a
    problem by executing the functions (the files are open)
    :param chosen_option: option in the drop down menu
    :param edit_people_window: the gui window of editing people
    :param list_if_empty: list that contain the value 'List is empty', in case
    the files are empty
    """

    try:
        delete_person_from_justice_board(name_of_person)
        delete_person_from_ilutzim(name_of_person)

        try:
            # Refresh the list in the drop down menu
            chosen_option.set(get_list_of_all_people()[0])  # default value
            dropped_down_menu = tk.OptionMenu(edit_people_window, chosen_option,
                                              *get_list_of_all_people())
        except:
            # Set the drop down menu values with the 'empty list' values
            chosen_option.set(list_if_empty[0])  # If the file is empty
            dropped_down_menu = tk.OptionMenu(edit_people_window, chosen_option,
                                              *list_if_empty)
        dropped_down_menu.grid(row=1, column=1)
        warning_label['text'] = ''
        warning_label['bg'] = None

    except: # If the files are open
        warning_label['text'] = 'אזהרה: הקובץ של \nלוח הצדק פתוח.\n אנא סגור ' \
                                'Iאותו \nכדי שיתאפשר \nלשמור את השינויים!'
        warning_label['bg'] = 'red'

def delete_person_from_justice_board(name_of_person):
    """
    Delete the given person from every sheet in the justice board file
    :param name_of_person: the person to delete
    :param warning_label: a label that will warn the user if the justice
     board file is open
    """

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    with pd.ExcelWriter('justice_board.xlsx', engine='openpyxl',
                        mode='a') as writer:
        workbook = writer.book

    dict_of_df = get_justice_sheets_as_df()

    # Run over each DF and delete the person from it if the person is in it
    # and update the sheet
    for key in dict_of_df:
        df = dict_of_df[key]
        if name_of_person in df['Name'].values:
            index_of_name_to_remove = df[df['Name'] == name_of_person].index
            removed_name_df = df.drop(index_of_name_to_remove) \
                .reset_index(drop=True)
            workbook.remove(workbook[key])
            removed_name_df.to_excel(writer, sheet_name=key)

    writer.save()


def delete_person_from_ilutzim(name_of_person):

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    with pd.ExcelWriter('ilutzim.xlsx', engine='openpyxl',
                        mode='a') as writer:
        workbook = writer.book
    dict_of_df = get_ilutzim_sheets_as_df()

    # Run over each DF and delete the person from it if the person is
    # in it and update the sheet
    for key in ['Samba', 'Manager']:
        df = dict_of_df[key]
        if name_of_person in df['Name'].values:
            index_of_name_to_remove = df[
                df['Name'] == name_of_person].index
            removed_name_df = df.drop(index_of_name_to_remove) \
                .reset_index(drop=True)
            workbook.remove(workbook[key])
            removed_name_df.to_excel(writer, sheet_name=key)

    # Run over each DF and delete the person from it if the person is
    # in it and update the sheet
    for key in ['Makel Officer', 'Makel Operator']:
        df = dict_of_df[key]
        if name_of_person in df.index.to_list():
            removed_name_df = df.drop(name_of_person)
            workbook.remove(workbook[key])
            removed_name_df.to_excel(writer, sheet_name=key)
    writer.save()

