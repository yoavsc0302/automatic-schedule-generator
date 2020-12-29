# This module contains the backtracking algorithms with the functions it uses
import pandas as pd

tzevet_conan = pd.read_excel('tzevet_conan.xlsx',
                                 sheet_name='Tzevet Conan',
                                  engine='openpyxl', index_col=0)
makel_officer_df = tzevet_conan.loc[['Officer 1',
                                     'Officer 2',
                                     'Officer 3',
                                     'Officer 4']]

index_list = makel_officer_df.index.values.tolist() # List of the names of
# the indexes
columns_list = makel_officer_df.columns.values.tolist() # List of the names of
# the columns

def find_empty(df, index_list, col_list):
    """
    Get the location of the empty df's cells
    :param df: the df that the function will work on
    :return: the location of the empty df's cells
    """

    for i in range(len(index_list)):
        for j in range(len(col_list)):

            # If the cell is empty, get it's location
            if df.at[index_list[i], col_list[j]] == 'empty':
                return (i, j) # row, col

    return None # If all the cells in the df are set with valid names


def valid(df, name, pos, index_list, columns_list, list_of_names):
    """
    Chek if the opttional name in the cell is a valid name
    :param df: the df we fill it's cells with names
    :param name: the name we want to check if it creates a valid df
    :param pos: the position in the df which we are trying to find a valid name
    for it
    :param index_list: a list of the rows' names (used to find the length)
    :param columns_list: a list of the cols' names (used to find the length)
    :param list_of_names: a list of opptional names to be inserted to the df
    :return: boolean: True if a valid name for the position was found
    """

    # Set the optional name to the cell in the df which we are currently looking
    # for a match to it
    df.iat[pos[0], pos[1]] = name
    current = df.iat[pos[0], pos[1]] # The current optional name

    # Check that the same name doesn't repeat itself day after day for the
    # same row
    #if(pos[0] < 2): # Only for Officer 1 and 2
    for i in range(len(columns_list)):
            other_cell = df.iat[pos[0], i] # A cell in the row

            # Check if the to cells conatin the same name and
            # are 2 DIFFERENT cells
            if (current == other_cell) and (pos[1] != i):

                # Check if other cell is X steps from the current cell
                # X = the ammount of optional names.
                # This is done in order to use the max ammount of opptional
                # names and not only the first names in the list over and over
                if(abs(pos[1]-i) < len(list_of_names)):
                    return False

    # Check that the same name doesn't show up more than once in the col
    for i in range(len(index_list)):
            other_cell_col = df.iat[i, pos[1]] # A cell in the col

            # Check if the to cells conatin the same name and
            # are 2 DIFFERENT cells
            if (current == other_cell_col) and (pos[0] != i):

                # Check if other cell is X steps from the current cell
                # X = the ammount of optional names.
                # This is done in order to use the max ammount of opptional
                # names and not only the first names in the list over and over
                if(abs(pos[0]-i) < len(list_of_names)):
                    return False

    return True # If the name is valid


def generate(df):
    """
    Recursive function that generate names to the df
    :param df: the df which names will be seted to
    :return: boolean: True if suited names found
    """
    list_of_names = ['yoad', 'shelly', 'afek', 'yuval', 'mike']
    find = find_empty(df, index_list, columns_list)
    if not find:
        return True
    else:
        row, col = find

        # Check if the previews row completed succesfuly
        if (col == 0) and (row != 0):

            # Get the used names, move them to the end of the list to let
            # unused names to be used as well
            used_names = df.iloc[row-1].to_list()
            #print(used_names)
            for name in used_names:
                list_of_names.remove(name)
                list_of_names.append(name)

            #print(list_of_names)


    for name in list_of_names:
        #print(find)
        #print(df)
        #print('------------------------------------------')
        if valid(df, name, (row, col), index_list, columns_list, list_of_names):
            df.iat[row, col] = name

            if generate(df):
                return True

        df.iat[row, col] = 'empty'

    return False

print('Old df:')
print(makel_officer_df)
print('---------------------------------------')
generate(makel_officer_df)
print('New df:')
print(makel_officer_df)