# This module contains the backtracking algorithms with the functions it uses
import pandas as pd
import excel_modifications as em


def get_tzevet_conan():
    tzevet_conan = pd.read_excel('tzevet_conan.xlsx',
                                 sheet_name='Tzevet Conan',
                                 engine='openpyxl', index_col=0)
    return tzevet_conan


tzevet_conan = get_tzevet_conan()

print('First')
print(tzevet_conan)


# (1) Backtracking functions ---------------------------------------------------


def create_list(dict_of_available):
    """
    Insert into the list all the names of the available people from the dict
    :param dict_of_available: the dictionary containing the name of each
    available person and how many shifts he has done
    :return: the list
    """
    list_of_available = []
    for key in dict_of_available:
        list_of_available.append(key)

    return list_of_available


def update_list(dict_of_available_people, list_to_update):
    """
    Move the person with the highest number of 'shifts behind' to
    the first place in the list
    :param dict_of_available_people: the dictionary containing the name of
    each available person andhow many shifts is he behind
    :param list_to_update: the list we are updating
    :return: list_to_update
    """
    # Flags for keeping the lowest person name
    person_with_min_shifts = ''
    min_num_shifts = 9000000000

    # Go over each person 'count', and check if it's the highest till that
    # moment
    for key in dict_of_available_people:
        if dict_of_available_people[key] <= min_num_shifts:
            min_num_shifts = dict_of_available_people[key]
            person_with_min_shifts = key

    popped_item = list_to_update.pop(
        list_to_update.index(person_with_min_shifts))
    list_to_update.insert(0, popped_item)

    return list_to_update


def find_empty(df, index_list, col_list):
    """
    Get the location of the empty df's cells
    :param col_list: list containing the names of the columns
    :param index_list: list containing the names of the indexes
    :param df: the df that the function will work on
    :return: the location of the empty df's cells
    """

    for i in range(len(index_list)):
        for j in range(len(col_list)):

            # If the cell is empty, get it's location
            if df.at[index_list[i], col_list[j]] == 'empty':
                return (i, j)  # row, col

    # If all the cells in the df are set with valid names
    return None


def valid(df, name, pos, index_list, columns_list, list_of_names, job):
    """
    Chek if the optional name in the cell is a valid name
    :param df: the df we fill it's cells with names
    :param name: the name we want to check if it creates a valid df
    :param pos: the position in the df which we are trying to find a valid name
    for it
    :param index_list: a list of the rows' names (used to find the length)
    :param columns_list: a list of the cols' names (used to find the length)
    :param list_of_names: a list of optional names to be inserted to the df
    :return: boolean: True if a valid name for the position was found
    """

    row = pos[0]
    col = pos[1]

    # if index_list[0] == 'Officer 1':
    #     job = 'Makel Officer'
    # elif index_list[0] == 'Operator 1':
    #     job = 'Makel Operator'
    # elif index_list[0] == 'Manager':
    #     job = 'Manager'
    # elif index_list[0] == 'Fast caller':
    #     job = 'Fast caller'
    # elif index_list[0] == 'Samba':
    #     job = 'Samba'

    # Set the optional name to the cell in the df which we are currently looking
    # for a match to it
    df.iat[row, col] = name
    current = df.iat[row, col]  # The current optional name

    # Check that the same name doesn't repeat itself day after day in
    # the same row Only for rows 1 + 2:
    # Makel officer/manager it's: Officer/Operator 1 + 2
    # Toran it's: Toran + operator 1
    # Samba it's: samba + operator 1
    # Manager it's: manager + officer 1
    # Driver it's: driver + operator 1
    if row < 2:
        for i in range(len(columns_list)):
            # A cell in the row
            other_cell = df.iat[row, i]

            # Check if the to cells contain the same name and
            # are 2 DIFFERENT cells
            if (current == other_cell) and (col != i):

                # Check if other cell is 1 steps from the current cell
                if abs(col - i) == 1:
                    return False

    # For makel operator: check that a person isn't team 1 a day after
    # being team 2, to prevent him being on a double shift of 24 hours
    if job in ['Makel Operator']:

        # Check if we are in the second row (Operator 2), and on the third day
        # (Tuesday)
        if (row == 1) and (col != 3):
            team_one_a_day_after = df.iat[0, col + 1]
            if current == team_one_a_day_after:
                return False

    # Check that the same name doesn't show up more than once in the col
    for i in range(len(index_list)):
        other_cell_col = df.iat[i, col]  # A cell in the col

        # Check if the to cells contain the same name and
        # are 2 DIFFERENT cells
        if (current == other_cell_col) and (row != i):

            # Check if other cell is X steps from the current cell
            # X = the amount of optional names.
            # This is done in order to use the max amount of optional
            # names and not only the first names in the list over and over
            if job in ['Makel Officer', 'Makel Operator']:
                if abs(row - i) < len(list_of_names):
                    return False
            else:
                return False

    return True  # If the name is valid


def generate(df, makel_officers_gen_df, index_list, cols_list, job):
    """
    Recursive function that generate names to the df
    :param df: the df which names will be set to
    :param makel_officers_gen_df: the df constructed by the:
    'define_df_for_generating_makel' function from the
     excel_modifications module
    :return: boolean: True if suited names found
    """

    # Check if there are still empty cells in the df
    find = find_empty(df, index_list, cols_list)
    if not find:
        return True
    else:
        row, col = find

    # Get a series of available and sorted people
    names_se = em.arrange_df_by_availability_and_justice(makel_officers_gen_df,
                                                         (row, col))
    list_of_names = []
    dict_of_names = {}
    name_in_se = ''
    sum_of_name = 0

    # For every row in the series, set into the dictionary keys as the names,
    # and values as the 'how many shifts this person is behind'
    for i in range(len(names_se)):
        name_in_se = names_se.index[i]
        sum_of_name = names_se[i]
        dict_of_names[name_in_se] = sum_of_name

    # Create the list and update it by moving the person with the
    # lowest number of times being team (1/2/3+4) to the start
    list_of_names = create_list(dict_of_names)
    list_of_names = update_list(dict_of_names, list_of_names)

    # Check validation in the cell for each name in the available people list
    for name in list_of_names:
        if valid(df, name, (row, col), index_list, cols_list, list_of_names, job):

            # Set the name to the cell
            df.iat[row, col] = name

            # Update the df about this person being a team
            dict_loc_to_team = {'0': '1', '1': '2', '2': '3+4', '3': '3+4'}
            team = dict_loc_to_team[f'{row}']

            if job in ['Makel Officer', 'Makel Operator']:
                makel_officers_gen_df.at[name, team] += 1

            if job in ['Manager', 'Samba', 'Toran', 'Driver']:
                makel_officers_gen_df.at[name, 'Sum'] += 1

            if generate(df, makel_officers_gen_df, index_list, cols_list, job):
                return True

        df.iat[row, col] = 'empty'

    return False


# (1) Generate Makel Officers --------------------------------------------------

def generate_makel_officer():
    print('-------------------------------------')
    print('Officer')
    print(tzevet_conan)
    makel_officers_conanim_df = tzevet_conan.loc[['Officer 1',
                                                  'Officer 2',
                                                  'Officer 3',
                                                  'Officer 4']]

    # List of the names of the indexes
    index_list_makel_officers = makel_officers_conanim_df.index.values.tolist()

    # List of the names of the columns
    cols_list_makel_officers = makel_officers_conanim_df.columns.values.tolist()

    makel_officers_gen_df = em.define_df_for_generating_makel('Makel Officer')

    generate(makel_officers_conanim_df,
             makel_officers_gen_df,
             index_list_makel_officers,
             cols_list_makel_officers,
             'Makel Officer')

    # Insert into the tzevet conan file the generated makel officers
    for index_name in ['Officer 1', 'Officer 2', 'Officer 3', 'Officer 4']:
        tzevet_conan.loc[index_name] = makel_officers_conanim_df.loc[index_name]


# (2) Generate Makel Operators -------------------------------------------------

def generate_makel_operators():
    print('-------------------------------------')
    print('Operatpr')
    print(tzevet_conan)

    makel_operators_conanim_df = tzevet_conan.loc[['Operator 1',
                                                   'Operator 2',
                                                   'Operator 3',
                                                   'Operator 4']]

    # Makel operators' list of the names of the indexes
    index_list_makel_operators = makel_operators_conanim_df.index.values.tolist()

    # Makel operators' List of the names of the columns
    columns_list_makel_operators = makel_operators_conanim_df.columns.values.tolist()

    makel_operators_gen_df = em.define_df_for_generating_makel('Makel Operator')

    generate(makel_operators_conanim_df,
             makel_operators_gen_df,
             index_list_makel_operators,
             columns_list_makel_operators,
             'Makel Operator')

    # Insert into the tzevet conan file the generated makel officers
    for index_name in ['Operator 1', 'Operator 2', 'Operator 3', 'Operator 4']:
        tzevet_conan.loc[index_name] = makel_operators_conanim_df.loc[
            index_name]


# (3) Generate Managers --------------------------------------------------------

def generate_managers():
    print('-------------------------------')
    print('manager')
    print(tzevet_conan)
    managers_conanim_df = tzevet_conan.loc[['Manager',
                                            'Officer 1',
                                            'Officer 2']]

    # Makel operators' list of the names of the indexes
    index_list_managers = managers_conanim_df.index.values.tolist()

    # Makel operators' List of the names of the columns
    columns_list_managers = managers_conanim_df.columns.values.tolist()

    managers_gen_df = em.define_df_for_generating_manager()

    generate(managers_conanim_df,
             managers_gen_df,
             index_list_managers,
             columns_list_managers,
             'Manager')

    # Insert into the tzevet conan file the generated managers
    tzevet_conan.loc['Manager'] = managers_conanim_df.loc['Manager']


# (4) Generate Toran ---------------------------------------------

def generate_toranim():
    print('-------------------------------')
    print('toranim')
    print(tzevet_conan)
    toranim_conanim_df = tzevet_conan.loc[['Toran',
                                           'Operator 1',
                                           'Operator 2']]

    # Makel operators' list of the names of the indexes
    index_list_toranim = toranim_conanim_df.index.values.tolist()

    # Makel operators' List of the names of the columns
    columns_list_toranim = toranim_conanim_df.columns.values.tolist()

    toranim_gen_df = em.define_df_for_generating_toranim()

    generate(toranim_conanim_df,
             toranim_gen_df,
             index_list_toranim,
             columns_list_toranim,
             'Toran')

    # Insert into the tzevet conan file the generated toran
    tzevet_conan.loc['Toran'] = toranim_conanim_df.loc['Toran']


# (5) Generate Samba -----------------------------------------------------------

def generate_samba():
    print('-------------------------------')
    print('samba')
    print(tzevet_conan)
    samba_conanim_df = tzevet_conan.loc[['Samba',
                                         'Operator 1',
                                         'Operator 2']]

    # Makel operators' list of the names of the indexes
    index_list_samba = samba_conanim_df.index.values.tolist()

    # Makel operators' List of the names of the columns
    columns_list_samba = samba_conanim_df.columns.values.tolist()

    samba_gen_df = em.define_df_for_generating_samba()
    generate(samba_conanim_df,
             samba_gen_df,
             index_list_samba,
             columns_list_samba,
             'Samba')

    # Insert into the tzevet conan file the generated samba
    tzevet_conan.loc['Samba'] = samba_conanim_df.loc['Samba']


# (6) Generate Driver ----------------------------------------------------------

def generate_driver():
    print('----------------------------------')
    print('Driver')
    print(tzevet_conan)
    driver_conanim_df = tzevet_conan.loc[['Driver',
                                          'Operator 1',
                                          'Operator 2']]

    # Driver's list of the names of the indexes
    index_list_driver = driver_conanim_df.index.values.tolist()

    # Driver List of the names of the columns
    columns_list_driver = driver_conanim_df.columns.values.tolist()

    driver_gen_df = em.define_df_for_generating_driver()
    generate(driver_conanim_df,
             driver_gen_df,
             index_list_driver,
             columns_list_driver,
             'Driver')

    # Insert into the tzevet conan file the generated samba
    tzevet_conan.loc['Driver'] = driver_conanim_df.loc['Driver']


# (7) Generate All -------------------------------------------------------------

def generate_all():
    generate_makel_officer()
    generate_makel_operators()
    generate_managers()
    generate_toranim()
    generate_samba()
    generate_driver()
