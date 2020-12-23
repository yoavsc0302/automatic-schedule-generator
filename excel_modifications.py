# This module contains all the excel modification functions

import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter

# Show entire df when printed
pd.set_option("display.max_rows", None, "display.max_columns", None)


def create_ilutzim_excel(makel_names, manager_names, samba_names):
    """
    Create the ilutzim excel file as a 'Multiply indexed DataFrame'
    source:https://jakevdp.github.io/PythonDataScienceHandbook/03.05-
    hierarchical-indexing.html
    :param makel_names: list that contains the names of every 'makel'
    :param manager_names: list that contains the names of every 'manaager'
    :param samba_names: list that contains the names of every 'samab'
    """

    # Makel df:
    # Hierarchical indices and columns
    index = pd.MultiIndex.from_product([makel_names], names=['name'])
    columns = pd.MultiIndex.from_product([['Sunday', 'Monday', 'Tuesday',
                                           'Wednesday'], ['1', '2', '3+4']],
                                         names=['Day', 'Team'])
    makel_df = pd.DataFrame(index=index, columns=columns, data='0')

    # Manager df:
    index = manager_names
    columns = ['Sunday', 'Monday', 'Tuesday', 'Wednesday']
    manager_df = pd.DataFrame(index=index, columns=columns)

    # Samba df:
    index = samba_names
    columns = ['Sunday', 'Monday', 'Tuesday', 'Wednesday']
    samba_df = pd.DataFrame(index=index, columns=columns)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('ilutzim.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    makel_df.to_excel(writer, sheet_name='Makel')
    manager_df.to_excel(writer, sheet_name='Manager')
    samba_df.to_excel(writer, sheet_name='Samba')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

