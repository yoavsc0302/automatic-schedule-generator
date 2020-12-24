# Create the gui for the app
import tkinter as tk

# Create a window and define it's size
window = tk.Tk()
window.geometry("700x600")
window.title('משבץ צוות כונן אוטומטי')


def openEditPeopleWindow():
    """
    This function create a new window
    """
    # Create a new windows and set size and title
    edit_people_window = tk.Toplevel(window)
    edit_people_window.title("משבץ צוות אוטומטי")
    edit_people_window.geometry("400x300")

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
    delete_person_healine = tk.Label(master=edit_people_window, text="מחק אדם")
    delete_person_healine.grid(row=0, column=1, sticky="nsew")
    delete_person_healine.config(font=("calibri", 12))

    # Get the name of the new person
    name = tk.Entry(edit_people_window).grid(row=1, column=5)

    # Get person rolls
    manager = tk.Checkbutton(edit_people_window, text="מנהל").grid(row=2, column=5)
    makel_officer = tk.Checkbutton(edit_people_window, text="קצין מקל").grid(row=3, column=5)
    makel_operator = tk.Checkbutton(edit_people_window, text="מפעיל מקל").grid(row=4, column=5)
    samba = tk.Checkbutton(edit_people_window, text="סמבצ").grid(row=2, column=4)
    fast_caller = tk.Checkbutton(edit_people_window, text="קריאה מהירה").grid(row=3, column=4)
    toran_unit = tk.Checkbutton(edit_people_window, text="ת. יחידתי").grid(row=4, column=4)

    #list of people
    list_of_people = ["Jan","Feb","Mar"]
    variable = tk.StringVar(edit_people_window)
    variable.set(list_of_people[0])  # default value
    w = tk.OptionMenu(edit_people_window, variable, *list_of_people).grid(row=1, column=1)

    # Add person button
    add_person = tk.Button(edit_people_window, text="הוסף בן אדם", bg="blue")
    add_person.grid(row=5, column=5, sticky="nsew")

    # Delete person button
    add_person = tk.Button(edit_people_window, text="מחק בן אדם", bg="blue")
    add_person.grid(row=5, column=1, sticky="nsew")

def openChangeFileLocWindows():
    """
    This function opens a window of change file location
    """
    # Create a new windows and set size and title
    change_file_loc_windows = tk.Toplevel(window)
    change_file_loc_windows.title("משבץ צוות אוטומטי")
    change_file_loc_windows.geometry("400x300")

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
    justice_board_headline = tk.Label(master=change_file_loc_windows, text="לוח צדק")
    justice_board_headline.grid(row=0, column=1, sticky="nsew")
    justice_board_headline.config(font=("calibri", 12))

    # Ilutzim headline
    ilutzim_headline = tk.Label(master=change_file_loc_windows, text="אילוצים")
    ilutzim_headline.grid(row=0, column=3, sticky="nsew")
    ilutzim_headline.config(font=("calibri", 12))

    # Get the location of the justice board file
    justice_board_file_loc = tk.Entry(change_file_loc_windows).grid(row=1, column=1, sticky="ew")

    # Get the location of the ilutzim file
    ilutzim_file_loc = tk.Entry(change_file_loc_windows).grid(row=1, column=3, sticky="ew")

    # Save justice board file location
    save_justic_board_loc = tk.Button(change_file_loc_windows, text="שמור מיקום לוח צדק", bg="blue")
    save_justic_board_loc.grid(row=3, column=1, sticky="nsew")

    # Save ilutzim file location
    save_ilutzim_loc = tk.Button(change_file_loc_windows, text="שמור מיקום קובץ אילוצים", bg="blue")
    save_ilutzim_loc.grid(row=3, column=3, sticky="nsew")


# Create grid of frames
for i in range(7):
    window.columnconfigure(i, weight=1, minsize=75)
    window.rowconfigure(i, weight=1, minsize=50)
    for j in range(7):
        frame = tk.Frame(
            master=window,
            relief=tk.RAISED,
            borderwidth=1,
        )
        frame.grid(row=i, column=j,sticky="nsew")

# Set window's headline
window_headline = tk.Label(master=window, text="דף הבית")
window_headline.grid(row=0, column=3, sticky="nsew")
window_headline.config(font=("calibri", 20))

# Set buttons

#all_button
generate_all = tk.Button(text="שבץ צוות שלם",bg="blue")
generate_all.grid(row=2, column=3, sticky="nsew")

#makel_button
generate_makel = tk.Button(text="מקל",bg="blue")
generate_makel.grid(row=3, column=4, sticky="nsew")

#manager_button
generate_manager = tk.Button(text="מנהלים",bg="blue")
generate_manager.grid(row=3, column=3, sticky="nsew")

#samba_button
generate_samba = tk.Button(text="סמבץ/ק.מ/ת.י",bg="blue")
generate_samba.grid(row=3, column=2, sticky="nsew")

#go to files location button
go_to_files_loc = tk.Button(text="שנה קבצים", bg="blue", command = openChangeFileLocWindows)
go_to_files_loc.grid(row=0, column=6, sticky="nsew")

#go to edit_people_button
go_to_edit_people = tk.Button(text="ערוך אנשים", bg="blue", command = openEditPeopleWindow)
go_to_edit_people.grid(row=1, column=6, sticky="nsew")

#open ilutzim file button
open_ilutzim = tk.Button(text="פתח אילוצים",bg="blue")
open_ilutzim.grid(row=0, column=0, sticky="nsew")

#open justice board file button
open_justice_board = tk.Button(text="פתח לוח צדק",bg="blue")
open_justice_board.grid(row=1, column=0, sticky="nsew")




window.mainloop()

