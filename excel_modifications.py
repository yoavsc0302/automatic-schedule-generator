# This module contains all the excel modification functions

import pandas as pd
import numpy as np
import openpyxl

# Show entire df when printed
pd.set_option("display.max_rows", None, "display.max_columns", None)


def create_ilutzim_excel(list_of_names):
    """
    Create the ilutzim excel file as a 'Multiply indexed DataFrame'
    source:https://jakevdp.github.io/PythonDataScienceHandbook/03.05-
    hierarchical-indexing.html
    :param list_of_names: list that contains the names of everyone
    """

    # Hierarchical indices and columns
    index = pd.MultiIndex.from_product([list_of_names],
                                       names=['name'])
    columns = pd.MultiIndex.from_product(
        [['Sunday', 'Monday', 'Tuesday', 'Wednesday'], ['1', '2', '3+4']],
        names=['Day', 'Team'])

    ilutzim_df = pd.DataFrame(index=index, columns=columns, data='0')
    ilutzim_df.to_excel('ilutzim.xlsx')
