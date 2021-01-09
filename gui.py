# Create the gui for the app
import tkinter as tk
import pandas as pd
import webbrowser
import excel_modifications as em
import os


def open_url(url):
    """
    Opens the given url
    :param url: the url which the function will open
    """
    os.startfile(url)


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


def save_files_new_locations(files_location_df, ilutzim_entry, justice_entry):
    """
    Saves the new files location of ilutzim and justice board
    :param files_location_df: the df which contain the files locations
    :param ilutzim_entry: The new ilutzim file location given by the user
     through the entry
    :param justice_entry: The new justice board location given by the user
     through the entry
    """
    files_location_df.at[0, 'ilutzim'] = ilutzim_entry.get()
    files_location_df.at[0, 'justice_board'] = justice_entry.get()
    files_location_df.to_csv('files_location.csv', index=False)


def open_edit_people_window():
    """
    Create a new window for editing people
    """
    # Create a new windows and set size and title
    edit_people_window = tk.Toplevel(window)
    edit_people_window.title("משבץ צוות אוטומטי")
    edit_people_window.geometry("400x300")

    # Divide the windows into 7x7 frames
    for i in range(7):
        edit_people_window.columnconfigure(i, weight=1, minsize=20)
        edit_people_window.rowconfigure(i, weight=1, minsize=20)
        for j in range(7):
            frame = tk.Frame(
                master=edit_people_window,
                relief=tk.RAISED,
                borderwidth=0,
            )
            frame.grid(row=i, column=j, sticky="nsew")

    # Add a person title
    new_person_headline = tk.Label(master=edit_people_window, text="אדם חדש")
    new_person_headline.grid(row=0, column=5, sticky="nsew")
    new_person_headline.config(font=("calibri", 12))

    # Delete a person title
    delete_person_headline = tk.Label(master=edit_people_window, text="מחק אדם")
    delete_person_headline.grid(row=0, column=1, sticky="nsew")
    delete_person_headline.config(font=("calibri", 12))

    # Get the name of the new person
    name = tk.Entry(edit_people_window)
    name.grid(row=1, column=5)

    # Get person rolls from checkboxes
    manager_var = tk.IntVar()
    makel_officer_var = tk.IntVar()
    makel_operator_var = tk.IntVar()
    samba_var = tk.IntVar()
    fast_and_toran_var = tk.IntVar()

    manager = tk.Checkbutton(edit_people_window, text="מנהל",
                             variable=manager_var).grid(
        row=2, column=5, sticky="e")
    makel_officer = tk.Checkbutton(edit_people_window,
                                   text="קצין מקל",
                                   variable=makel_officer_var) \
        .grid(row=3, column=5, sticky="e")
    makel_operator = tk.Checkbutton(edit_people_window, text="מפעיל מקל",
                                    variable=makel_operator_var) \
        .grid(row=4, column=5, sticky="e")
    samba = tk.Checkbutton(edit_people_window, text="סמבצ",
                           variable=samba_var).grid(row=2, column=4, sticky="e")
    fast_caller_and_toran = tk.Checkbutton(edit_people_window,
                                           text="ת. יחידתי + ק.מ",
                                           variable=fast_and_toran_var) \
        .grid(row=3, column=4, sticky="e")

    # list of people
    list_of_people = em.get_list_of_all_people()
    list_if_empty = ['The file is empty']
    chosen_option = tk.StringVar(edit_people_window)

    try:
        chosen_option.set(list_of_people[0])  # default value
        dropped_down_menu = tk.OptionMenu(edit_people_window, chosen_option,
                                          *list_of_people)
    except:
        chosen_option.set(list_if_empty[0])  # If the file is empty
        dropped_down_menu = tk.OptionMenu(edit_people_window, chosen_option,
                                          *list_if_empty)

    dropped_down_menu.grid(row=1, column=1)

    # Add a warning alert
    warning_label = tk.Label(edit_people_window, text='')
    warning_label.grid(row=2, column=1)

    # Add person button
    add_person = tk.Button(edit_people_window, text="הוסף בן אדם", bg="#ff677d",
                           command=lambda: em.add_new_person(name.get(),
                                                             manager_var,
                                                             makel_officer_var,
                                                             makel_operator_var,
                                                             samba_var,
                                                             fast_and_toran_var,
                                                             warning_label))
    add_person.grid(row=5, column=5, sticky="nsew")

    # Delete person button
    delete_person = tk.Button(edit_people_window, text="מחק בן אדם",
                              bg="#ff677d",
                              command=lambda: em.delete_person(
                                  chosen_option.get(),
                                  warning_label,
                                  chosen_option,
                                  edit_people_window,
                                  list_if_empty))
    delete_person.grid(row=5, column=1, sticky="nsew")


def open_change_file_loc_window():
    """
    Create a new window for changing ilutzim and justice board files locations'
    """

    # Create a new windows and set size and title
    change_file_loc_windows = tk.Toplevel(window)
    change_file_loc_windows.title("משבץ צוות אוטומטי")
    change_file_loc_windows.geometry("400x300")

    # Divide the windows into 5x5 frames
    for i in range(5):
        change_file_loc_windows.columnconfigure(i, weight=1, minsize=20)
        change_file_loc_windows.rowconfigure(i, weight=1, minsize=20)
        for j in range(5):
            frame = tk.Frame(
                master=change_file_loc_windows,
                relief=tk.RAISED,
                borderwidth=0,
            )
            frame.grid(row=i, column=j, sticky="nsew")

    # Justice board headline
    justice_board_headline = tk.Label(master=change_file_loc_windows,
                                      text="לוח צדק")
    justice_board_headline.grid(row=0, column=1, sticky="nsew")
    justice_board_headline.config(font=("calibri", 12))

    # Ilutzim headline
    ilutzim_headline = tk.Label(master=change_file_loc_windows, text="אילוצים")
    ilutzim_headline.grid(row=0, column=3, sticky="nsew")
    ilutzim_headline.config(font=("calibri", 12))

    # Get the location of the justice board file
    justice_board_file_loc = tk.Entry(change_file_loc_windows)
    justice_board_file_loc.grid(row=1, column=1, sticky="ew")
    justice_board_file_loc.insert(0, get_justice_board_location())

    # Get the location of the ilutzim file
    ilutzim_file_loc = tk.Entry(change_file_loc_windows)
    ilutzim_file_loc.grid(row=1, column=3, sticky="ew")
    ilutzim_file_loc.insert(0, get_ilutzim_location())

    # Save files locations button
    save_files_locations = tk.Button(change_file_loc_windows,
                                     text="שמור מיקומים", bg="#ff677d",
                                     command=lambda:
                                     save_files_new_locations(
                                         pd.read_csv('files_location.csv'),
                                         ilutzim_file_loc,
                                         justice_board_file_loc))
    save_files_locations.grid(row=3, column=1, sticky="nsew")


# Create the homepage window and define it's size
window = tk.Tk()
window.geometry("550x450")
window.title('משבץ צוות כונן אוטומטי')

# Divide the windows into 7x7 frames
for i in range(6):
    window.columnconfigure(i, weight=1, minsize=75)
    window.rowconfigure(i, weight=1, minsize=50)
    for j in range(7):
        frame = tk.Frame(
            master=window,
            relief=tk.RAISED,
            borderwidth=1,
        )
        frame.grid(row=i, column=j, sticky="nsew")

# Set window's headline
window_headline = tk.Label(master=window, text="דף הבית")
window_headline.grid(row=0, column=3, sticky="nsew")
window_headline.config(font=("calibri", 20))

# Set buttons

# all_button
generate_all = tk.Button(text="שבץ צוות שלם", bg="blue")
generate_all.grid(row=2, column=3, sticky="nsew")

# makel_button
generate_makel = tk.Button(text="מקל", bg="blue")
generate_makel.grid(row=3, column=4, sticky="nsew")

# manager_button
generate_manager = tk.Button(text="מנהלים", bg="blue")
generate_manager.grid(row=3, column=3, sticky="nsew")

# samba_button
generate_samba = tk.Button(text="סמבץ/ק.מ/ת.י", bg="blue")
generate_samba.grid(row=3, column=2, sticky="nsew")

# go to files location button
go_to_files_loc = tk.Button(text="שנה קבצים", bg="blue",
                            command=open_change_file_loc_window)
go_to_files_loc.grid(row=0, column=6, sticky="nsew")

# go to edit_people_button
go_to_edit_people = tk.Button(text="ערוך אנשים", bg="blue",
                              command=open_edit_people_window)
go_to_edit_people.grid(row=1, column=6, sticky="nsew")

# open ilutzim file button
open_ilutzim = tk.Button(text="פתח אילוצים", bg="blue",
                         command=lambda: open(get_ilutzim_location()))
open_ilutzim.grid(row=0, column=0, sticky="nsew")

# open justice board file button
open_justice_board = tk.Button(text="פתח לוח צדק", bg="blue")
open_justice_board.grid(row=1, column=0, sticky="nsew")

window.mainloop()
