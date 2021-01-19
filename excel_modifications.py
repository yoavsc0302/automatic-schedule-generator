# This module contains all the excel modification functions

import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter
import pickle
import math
import tkinter as tk
from tkinter import *
from tkinter import filedialog

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
    toran_df = pd.read_excel('justice_board.xlsx', sheet_name='Toran',
                             engine='openpyxl', index_col=0)
    driver_df = pd.read_excel('justice_board.xlsx', sheet_name='Driver',
                             engine='openpyxl', index_col=0)
    return {'Makel Officer': makel_officer_df,
            'Makel Operator': makel_operator_df,
            'Manager': manager_df,
            'Samba': samba_df,
            'Toran': toran_df,
            'Driver': driver_df}


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
    ilutzim_toran_df = pd.read_excel('ilutzim.xlsx', sheet_name='Toran',
                                     engine='openpyxl', index_col=0)
    ilutzim_driver_df = pd.read_excel('ilutzim.xlsx', sheet_name='Driver',
                                     engine='openpyxl', index_col=0)

    return {'Makel Officer': ilutzim_makel_officer_df,
            'Makel Operator': ilutzim_makel_operator_df,
            'Manager': ilutzim_manager_df,
            'Samba': ilutzim_samba_df,
            'Toran': ilutzim_toran_df,
            'Driver': ilutzim_driver_df}


def get_tzevet_conan_location():
    """
    Get the ilutzim file location
    :return: the ilutzim file location as a string
    """
    files_location_df = pd.read_csv('files_location.csv')
    tzevet_conan_file = files_location_df['tzevet_conan'][0]
    return tzevet_conan_file


def get_ilutzim_location():
    """
    Get the ilutzim file location
    :return: the ilutzim file location as a string
    """
    files_location_df = pd.read_csv('files_location.csv')
    ilutzim_file = files_location_df['ilutzim'][0]
    return ilutzim_file


def get_justice_board_location():
    """
    Get the justice board file location
    :return: the justice board file location as a string
    """
    files_location_df = pd.read_csv('files_location.csv')
    justice_board_file = files_location_df['justice_board'][0]
    return justice_board_file


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

    # Driver df:
    driver_df = pd.DataFrame(columns=columns)

    # Toran df:
    toran_df = pd.DataFrame(columns=columns)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('ilutzim.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    makel_officer_df.to_excel(writer, sheet_name='Makel Officer')
    makel_operator_df.to_excel(writer, sheet_name='Makel Operator')
    manager_df.to_excel(writer, sheet_name='Manager')
    samba_df.to_excel(writer, sheet_name='Samba')
    driver_df.to_excel(writer, sheet_name='Driver')
    toran_df.to_excel(writer, sheet_name='Toran')

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

    # Driver df:
    driver_df = pd.DataFrame(columns=columns)

    # Samba df:
    samba_df = pd.DataFrame(columns=columns)

    # Toran df:
    toran_df = pd.DataFrame(columns=columns)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('justice_board.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    makel_officer_df.to_excel(writer, sheet_name='Makel Officer')
    makel_operate_df.to_excel(writer, sheet_name='Makel Operator')
    manager_df.to_excel(writer, sheet_name='Manager')
    samba_df.to_excel(writer, sheet_name='Samba')
    driver_df.to_excel(writer, sheet_name='Driver')
    toran_df.to_excel(writer, sheet_name='Toran')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()


def create_file_location_csv():
    """
    Create csv thats stores the ilutzim and the justice board files locations'
    """
    files_location_df = pd.DataFrame({'ilutzim': ['i location'],
                                      'justice_board': ['jb location'],
                                      'tzevet_conan': ['t location']})
    files_location_df.to_csv('files_location.csv')


def create_tzevet_conan_excel():
    """
    Create an excel file of the tzevet conan
    """

    # Define columns and index names:
    columns = ['Sunday', 'Monday', 'Tuesday', 'Wednesday']
    index = ['Manager', 'Samba', 'Driver', 'Toran',
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
                   samba_var, toran_var, driver_var, warning_label):
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
    :param toran_var: equals 1 if the fast and toran job checkbox
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
        toran_df = pd.read_excel('justice_board.xlsx',
                                 sheet_name='Toran',
                                 engine='openpyxl', index_col=0)
        driver_df = pd.read_excel('justice_board.xlsx',
                                 sheet_name='Driver',
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


        # If the new person is a 'samba', insert his name into
        # the 'samba' sheet and set his sum to the average of everybodies' sum
        if samba_var.get() == 1:
            workbook.remove(workbook['Samba'])
            try:
                sum_to_be_set = math.floor(samba_df['Sum'].mean())
            except:  # If this is the first person in the sheet
                sum_to_be_set = 0
            samba_df = samba_df.append({'Name': name, 'Sum': sum_to_be_set},
                                           ignore_index=True)
            # Write dataframe to the worksheet.
            samba_df.to_excel(writer, sheet_name='Samba')

            # add_new_person_to_ilutzim('Manager', name)
            add_new_person_to_ilutzim(job='Samba', name=name)


        # If the new person is a 'Toran', insert his name into
        # the 'Toran' sheet and set his sum to the average of everybodies' sum
        if toran_var.get() == 1:
            workbook.remove(workbook['Toran'])
            try:
                sum_to_be_set = math.floor(toran_df['Sum'].mean())
            except:  # If this is the first person in the sheet
                sum_to_be_set = 0
            toran_df = toran_df.append({'Name': name, 'Sum': sum_to_be_set},
                                           ignore_index=True)
            # Write dataframe to the worksheet.
            toran_df.to_excel(writer, sheet_name='Toran')

            # add_new_person_to_ilutzim('Toran', name)
            add_new_person_to_ilutzim(job='Toran', name=name)


        # If the new person is a 'Driver', insert his name into
        # the 'Driver' sheet and set his sum to the average of everybodies' sum
        if driver_var.get() == 1:
            workbook.remove(workbook['Driver'])
            try:
                sum_to_be_set = math.floor(samba_df['Sum'].mean())
            except:  # If this is the first person in the sheet
                sum_to_be_set = 0
            driver_df = driver_df.append({'Name': name, 'Sum': sum_to_be_set},
                                           ignore_index=True)
            # Write dataframe to the worksheet.
            driver_df.to_excel(writer, sheet_name='Driver')

            # add_new_person_to_ilutzim('Manager', name)
            add_new_person_to_ilutzim(job='Driver', name=name)



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
    toran_df_ilutzim = dict_of_df['Toran']
    driver_df_ilutzim = dict_of_df['Driver']


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

    if job == 'Toran':
        toran_df_ilutzim = toran_df_ilutzim.append({'Name': name,
                                                    'Sunday': '0',
                                                    'Monday': '0',
                                                    'Tuesday': '0',
                                                    'Wednesday': '0'},
                                                   ignore_index=True)
        dict_of_df['Toran'] = toran_df_ilutzim

    if job == 'Driver':
        driver_df_ilutzim = driver_df_ilutzim.append({'Name': name,
                                                    'Sunday': '0',
                                                    'Monday': '0',
                                                    'Tuesday': '0',
                                                    'Wednesday': '0'},
                                                   ignore_index=True)
        dict_of_df['Driver'] = driver_df_ilutzim

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
                      dict_of_df['Samba'],
                      dict_of_df['Toran'],
                      dict_of_df['Driver']]
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
    for key in ['Samba', 'Manager', 'Driver', 'Toran']:
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


def convert_ilutzim_format_to_conan_makel(name, ilutzim_df, df):
    """
    Read the ilutzim file and according to the data in the df, insert the
    ilutzim entered by the people in the df into the:
    'generating_makel_officers_df' while converting days and teams to tuples
    :param name: the person to convert his ilutzim to the
    'generating_makel_officers_df'
    :param ilutzim_df: the ilutzim df
    :param df: the updated 'generating_makel_officers_df'
    """

    # Dictionaries that will assist the function to convert days and team
    # to numbers that will be set in the tuples
    dict_days_convert = {'Sunday': 0, 'Monday': 1, 'Tuesday': 2, 'Wednesday': 3}
    dict_team_convert = {'1': 0, '2': 1}

    # For each day and team, check that the person doesn't have an ilutz in the
    # ilutzim file.
    # If he has, in the 'generating_makel_officers_df' suiting tuples change
    # the value to '1'. else, change it to '0'

    for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday']:
        for team in ['1', '2', '3+4']:

            # The value in the cell of the day + team combination
            value_in_cell = ilutzim_df.loc[name][day][team]

            # Check that the person has an ilutz of being that team in that day
            if (value_in_cell != '0') and (value_in_cell != 0):
                if team == '3+4':
                    df.at[name, (2, dict_days_convert[day])] = '1'
                    df.at[name, (3, dict_days_convert[day])] = '1'

                else:
                    df.at[name, (dict_team_convert[team],
                                 dict_days_convert[day])] = '1'

            else:
                if team == '3+4':
                    df.at[name, (2, dict_days_convert[day])] = '0'
                    df.at[name, (3, dict_days_convert[day])] = '0'

                else:
                    df.at[name, (dict_team_convert[team],
                                 dict_days_convert[day])] = '0'


def convert_ilutzim_format_to_conan_manager(name, ilutzim_df, df):
    """
    Read the ilutzim file and according to the data in the df, insert the
    ilutzim entered by the people in the df into the:
    'generating_managers_df' while converting days to tuples
    :param name: the person to convert his ilutzim to the
    'generating_managers_df'
    :param ilutzim_df: the ilutzim df
    :param df: the updated 'generating_managers_df'
    """

    # Dictionaries that will assist the function to convert days and team
    # to numbers that will be set in the tuples
    dict_days_convert = {'Sunday': 0, 'Monday': 1, 'Tuesday': 2, 'Wednesday': 3}
    dict_manager_convert = {'Manager': 0}

    # For each day, check that the person doesn't have an ilutz in the
    # ilutzim file.
    # If he has, in the 'generating_ilutzim_df' suiting tuples change
    # the value to '1'. else, change it to '0'

    for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday']:

        # The value in the cell of the day
        value_in_cell = ilutzim_df.loc[name][day]

        # Check that the person has an ilutz of being a toran in that day
        if (value_in_cell != '0') and (value_in_cell != 0):
            df.at[name, (0, dict_days_convert[day])] = '1'

        else:
            df.at[name, (0, dict_days_convert[day])] = '0'


def define_df_for_generating_makel(job):
    """
    Create a df that contains:
    - Name of Makel
    - How many shifts he has done in each team (1/2/3+4)
    - The combination of team + day (team, day), that the person can't do in
      the comming week according to the ilutzim file
    :param job: officer or an operator
    :return: generate_makel_df - the df explained above
    """

    # Get the 'Makel Officer/operator' sheet from the ilutzim and the justice
    # board files
    dict_ilutzim = get_ilutzim_sheets_as_df()
    dict_justice = get_justice_sheets_as_df()
    ilutzim_df = dict_ilutzim[job]
    justice_df = dict_justice[job]

    # Create a list containing tuples of locations (0, 0), (1, 0)..
    list_col = []
    for i in range(4):
        for j in range(4):
            list_col.append((i, j))

    # Create a df which it's columns are the location tuples
    locations_pd = pd.DataFrame(columns=list_col)

    # Concat the justice board df ['Name', '1', '2', '3+4'] with the
    # locations df [(0, 0), (0, 1), (0, 2)...]
    generate_makel_df = pd.concat([justice_df, locations_pd])

    generate_makel_df.set_index('Name', inplace=True, drop=True)

    # Execute this function for each name in df
    for name in generate_makel_df.index.values.tolist():
        convert_ilutzim_format_to_conan_makel(name, ilutzim_df, generate_makel_df)

    return generate_makel_df


def define_df_for_generating_manager():
    """
    Create a df that contains:
    - Name of Manager
    - How many shifts he has done
    - The combination of manager + day (manager, day), that the person can't
     do in the comming week according to the ilutzim file
    :return: generate_manager_df - the df explained above
    """
    # Get the 'Manager' sheet from the ilutzim and the justice
    # board files
    dict_ilutzim = get_ilutzim_sheets_as_df()
    dict_justice = get_justice_sheets_as_df()
    ilutzim_df = dict_ilutzim['Manager']
    ilutzim_df.set_index('Name', inplace=True, drop=True)
    justice_df = dict_justice['Manager']


    # Create a list containing tuples of locations (0, 0), (1, 0)..
    list_col = []
    for i in range(4):
        list_col.append((0, i))

    # Create a df which it's columns are the location tuples
    locations_df = pd.DataFrame(columns=list_col)

    # Concat the justice board df ['Name', 'Sum'] with the
    # locations df [(0, 0), (0, 1), (0, 2), (0, 3)]
    generate_manager_df = pd.concat([justice_df, locations_df])

    generate_manager_df.set_index('Name', inplace=True, drop=True)

    # Execute this function for each name in df
    for name in generate_manager_df.index.values.tolist():
        convert_ilutzim_format_to_conan_manager(name, ilutzim_df, generate_manager_df)

    return generate_manager_df


def define_df_for_generating_samba():
    """
    Create a df that contains:
    - Name of Manager
    - How many shifts he has done
    - The combination of manager + day (manager, day), that the person can't
     do in the comming week according to the ilutzim file
    :return: generate_manager_df - the df explained above
    """
    # Get the 'Samba' sheet from the ilutzim and the justice
    # board files
    dict_ilutzim = get_ilutzim_sheets_as_df()
    dict_justice = get_justice_sheets_as_df()
    ilutzim_df = dict_ilutzim['Samba']
    ilutzim_df.set_index('Name', inplace=True, drop=True)
    justice_df = dict_justice['Samba']


    # Create a list containing tuples of locations (0, 0), (1, 0)..
    list_col = []
    for i in range(4):
        list_col.append((0, i))

    # Create a df which it's columns are the location tuples
    locations_df = pd.DataFrame(columns=list_col)

    # Concat the justice board df ['Name', 'Sum'] with the
    # locations df [(0, 0), (0, 1), (0, 2), (0, 3)]
    generate_samba_df = pd.concat([justice_df, locations_df])

    generate_samba_df.set_index('Name', inplace=True, drop=True)

    # Execute this function for each name in df
    for name in generate_samba_df.index.values.tolist():
        convert_ilutzim_format_to_conan_manager(name, ilutzim_df, generate_samba_df)

    return generate_samba_df


def define_df_for_generating_toranim():
    """
    Create a df that contains:
    - Name of toran / fast caller
    - How many shifts he has done
    - The combination of toran + day (toran, day), that the person can't
     do in the comming week according to the ilutzim file
    :return: generate_toranim_df - the df explained above
    """
    # Get the 'Toran' sheet from the ilutzim and the justice
    # board files
    dict_ilutzim = get_ilutzim_sheets_as_df()
    dict_justice = get_justice_sheets_as_df()
    ilutzim_df = dict_ilutzim['Toran']
    ilutzim_df.set_index('Name', inplace=True, drop=True)
    justice_df = dict_justice['Toran']


    # Create a list containing tuples of locations (0, 0), (1, 0)..
    list_col = []
    for i in range(4):
        list_col.append((0, i))

    # Create a df which it's columns are the location tuples
    locations_df = pd.DataFrame(columns=list_col)

    # Concat the justice board df ['Name', 'Sum'] with the
    # locations df [(0, 0), (0, 1), (0, 2), (0, 3)]
    generate_toranim_df = pd.concat([justice_df, locations_df])

    generate_toranim_df.set_index('Name', inplace=True, drop=True)

    # Execute this function for each name in df
    for name in generate_toranim_df.index.values.tolist():
        convert_ilutzim_format_to_conan_manager(name, ilutzim_df, generate_toranim_df)

    return generate_toranim_df


def define_df_for_generating_driver():
    """
    Create a df that contains:
    - Name of toran / fast caller
    - How many shifts he has done
    - The combination of toran + day (toran, day), that the person can't
     do in the comming week according to the ilutzim file
    :return: generate_toranim_df - the df explained above
    """
    # Get the 'Toran' sheet from the ilutzim and the justice
    # board files
    dict_ilutzim = get_ilutzim_sheets_as_df()
    dict_justice = get_justice_sheets_as_df()
    ilutzim_df = dict_ilutzim['Driver']
    ilutzim_df.set_index('Name', inplace=True, drop=True)
    justice_df = dict_justice['Driver']


    # Create a list containing tuples of locations (0, 0), (1, 0)..
    list_col = []
    for i in range(4):
        list_col.append((0, i))

    # Create a df which it's columns are the location tuples
    locations_df = pd.DataFrame(columns=list_col)

    # Concat the justice board df ['Name', 'Sum'] with the
    # locations df [(0, 0), (0, 1), (0, 2), (0, 3)]
    generate_driver_df = pd.concat([justice_df, locations_df])

    generate_driver_df.set_index('Name', inplace=True, drop=True)

    # Execute this function for each name in df
    for name in generate_driver_df.index.values.tolist():
        convert_ilutzim_format_to_conan_manager(name, ilutzim_df, generate_driver_df)

    return generate_driver_df


def arrange_df_by_availability_and_justice(df, pos):
    """
    Takes the df and return the df of all available names in that position
    and sort the names in an order of max to min shifts needed to be taken by
    that person
    :param df: the df that the function will sort
    :param pos: the position in the tzevet conan (0,0), (0, 1)...
    :return: the df of the availible people sorted from max to min
    """

    dict_loc_to_team = {'0': '1', '1': '2', '2': '3+4', '3': '3+4'}
    team = dict_loc_to_team[f'{pos[0]}']

    # Check if its makel, manager or toranim
    # manager and toranim

    # Managers /
    if (len(list(df.columns)) == 5) or \
            (len(list(df.columns)) == 9) or \
            (len(list(df.columns)) == 7):
        available = df[df[pos] == '0']['Sum']

    # Makel
    else:
        available = df[df[pos] == '0'][team]

    available_and_sorted_df = available.sort_values()


    return available_and_sorted_df


def insert_generated_tzevet_conan_df_to_tzevet_conan_file(tzevet_conan_df):
    with pd.ExcelWriter('tzevet_conan.xlsx', engine='openpyxl', mode='a') \
            as writer:
        workbook = writer.book

    workbook.remove(workbook['Tzevet Conan'])
    tzevet_conan_df.to_excel(writer, sheet_name='Tzevet Conan')
    writer.save()


def insert_sum_to_justice_board(tzevet_conan_df):
    """
    Go over the enitre tzevet conan file that was generated and according to it,
    insert into the justice board file, for each name in each job, the number
    of times that the person had a shift in this job
    :param tzevet_conan_df: the tzevet conan df that was generated
    """
    dict_of_all_jobs = {'Manager': {}, 'Samba': {}, 'Driver': {},
                        'Toran': {},
                        'Officer 1': {}, 'Officer 2': {}, 'Officer 3': {},
                        'Officer 4': {},
                        'Operator 1': {}, 'Operator 2': {}, 'Operator 3': {},
                        'Operator 4': {}}
    dict_of_sum = {}

    # For every job in the tzevet conan (manager, samba, officer..), and for
    # every day of the week
    for job in dict_of_all_jobs.keys():
        for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday']:

            # The name of the person that is in the cell
            name = tzevet_conan_df.loc[job][day]

            # If the name is not in any of the sheets, that mean he is
            # miluim / hatzach, so continue:
            if name not in get_list_of_all_people():
                continue

            # Check if he is already in the dict (all ready in the row)
            if name not in dict_of_sum.keys():

                # Add him to the dict and set his sum to 1
                dict_of_sum[name] = 1

            else:

                # Add 1 to he's sum
                dict_of_sum[name] += 1

        # Add the dict of people and their sum, into the dict of all jobs, as
        # the values of the keys (the jobs)
        dict_of_all_jobs[job] = dict_of_sum
        dict_of_sum = {}

    # Dictionairy containing the df of each sheet in the ilutzim file
    dict_of_justice_dfs = get_justice_sheets_as_df()

    # Getting the df's from the dicionairy
    makel_officer_df_justice = dict_of_justice_dfs['Makel Officer']
    makel_operator_df_justice = dict_of_justice_dfs['Makel Operator']
    manager_df_justice = dict_of_justice_dfs['Manager']
    samba_df_justice = dict_of_justice_dfs['Samba']
    toran_df_justice = dict_of_justice_dfs['Toran']
    driver_df_justice = dict_of_justice_dfs['Driver']

    with pd.ExcelWriter('justice_board.xlsx', engine='openpyxl', mode='a') \
            as writer:
        workbook = writer.book

    # For every job in the tzevet conan (manager, samba, officer..), and for
    # every name in the dict_of_sum that is being the value of the
    # dict_of_all_jobs
    for job in dict_of_all_jobs.keys():
        for name in dict_of_all_jobs[job].keys():

            # Get the sum of that person's
            sum_of_name = dict_of_all_jobs[job][name]

            if job == 'Manager':
                index = manager_df_justice[manager_df_justice['Name'] == name]['Sum'].index[0]
                manager_df_justice.at[index, 'Sum'] += sum_of_name

            if job == 'Samba':
                index = samba_df_justice[samba_df_justice['Name'] == name][
                    'Sum'].index[0]
                samba_df_justice.at[index, 'Sum'] += sum_of_name

            if job == 'Driver':
                index = driver_df_justice[driver_df_justice['Name'] == name][
                    'Sum'].index[0]
                driver_df_justice.at[index, 'Sum'] += sum_of_name

            if job == 'Toran':
                index = toran_df_justice[toran_df_justice['Name'] == name][
                    'Sum'].index[0]
                toran_df_justice.at[index, 'Sum'] += sum_of_name

            elif job in ['Officer 1', 'Officer 2']:
                index = makel_officer_df_justice[makel_officer_df_justice['Name'] == name][job[-1]].index[0]
                print('Index')
                print(index)
                print('col')
                print(job[-1])
                makel_officer_df_justice.at[index, job[-1]] += sum_of_name

            elif job in ['Officer 3', 'Officer 4']:
                index = makel_officer_df_justice[makel_officer_df_justice['Name'] == name]['3+4'].index[0]
                makel_officer_df_justice.at[index, '3+4'] += sum_of_name

            elif job in ['Operator 1', 'Operator 2']:
                index = makel_operator_df_justice[makel_operator_df_justice['Name'] == name][job[-1]].index[0]
                makel_operator_df_justice.at[index, job[-1]] += sum_of_name

            elif job in ['Operator 3', 'Operator 4']:
                index = makel_operator_df_justice[makel_operator_df_justice['Name'] == name]['3+4'].index[0]
                makel_operator_df_justice.at[index, '3+4'] += sum_of_name

    dict_of_updated_df = {'Manager': manager_df_justice,
                          'Samba': samba_df_justice,
                          'Makel Officer': makel_officer_df_justice,
                          'Makel Operator': makel_operator_df_justice,
                          'Driver': driver_df_justice,
                          'Toran': toran_df_justice}

    for sheet in ['Manager', 'Samba', 'Makel Officer', 'Makel Operator',
                  'Driver', 'Toran']:
        workbook.remove(workbook[sheet])
        dict_of_updated_df[sheet].to_excel(writer, sheet_name=sheet)
        writer.save()


def browse_file(location_in_data_base):
    filename = filedialog.askopenfile(filetypes=(("All Files", "*.*"),))
    return (location_in_data_base, filename)


def save_files_new_locations(type_of_file, file_path):
    """
    Saves the new file location to the database
    """

    df = pd.read_csv('files_location.csv', index_col=0)

    if type_of_file == 'justice_board':
        df.at[0, 'justice_board'] = file_path

    elif type_of_file == 'ilutzim':
        df.at[0, 'ilutzim'] = file_path

    elif type_of_file == 'tzevet_conan':
        df.at[0, 'tzevet_conan'] = file_path

    df.to_csv('files_location.csv')
