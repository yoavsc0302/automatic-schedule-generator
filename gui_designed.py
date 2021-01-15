import tkinter as tk
import tkinter.font as tkFont
import excel_modifications as em
import pandas as pd
import os
import tzevet_backtracking as tb

class App:
    def __init__(self, root):
        #setting title
        root.title("שיבוץ צוות כונן אוטומטי")

        #setting window size
        width=612
        height=460
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        reset_tzevet_conan_btn=tk.Button(root)
        reset_tzevet_conan_btn["anchor"] = "center"
        reset_tzevet_conan_btn["bg"] = "#ff677d"
        reset_tzevet_conan_btn["cursor"] = "arrow"
        ft = tkFont.Font(family='Calibri',size=12)
        reset_tzevet_conan_btn["font"] = ft
        reset_tzevet_conan_btn["fg"] = "#000000"
        reset_tzevet_conan_btn["justify"] = "center"
        reset_tzevet_conan_btn["text"] = "אפס צוות כונן"
        reset_tzevet_conan_btn["relief"] = "flat"
        reset_tzevet_conan_btn.place(x=460,y=140,width=110,height=35)
        reset_tzevet_conan_btn["command"] = self.reset_tzevet_conan

        generate_tzevet_conan_btn=tk.Button(root)
        generate_tzevet_conan_btn["anchor"] = "center"
        generate_tzevet_conan_btn["bg"] = "#ff677d"
        ft = tkFont.Font(family='Calibri',size=12)
        generate_tzevet_conan_btn["font"] = ft
        generate_tzevet_conan_btn["fg"] = "#000000"
        generate_tzevet_conan_btn["justify"] = "center"
        generate_tzevet_conan_btn["text"] = "שבץ צוות כונן"
        generate_tzevet_conan_btn["relief"] = "flat"
        generate_tzevet_conan_btn.place(x=320,y=140,width=110,height=35)
        generate_tzevet_conan_btn["command"] = self.generate_tzevet_conan

        view_generated_tzevet_conan_btn=tk.Button(root)
        view_generated_tzevet_conan_btn["bg"] = "#ff677d"
        ft = tkFont.Font(family='Calibri',size=12)
        view_generated_tzevet_conan_btn["font"] = ft
        view_generated_tzevet_conan_btn["fg"] = "#000000"
        view_generated_tzevet_conan_btn["justify"] = "center"
        view_generated_tzevet_conan_btn["text"] = "צפה בצוות שנוצר"
        view_generated_tzevet_conan_btn["relief"] = "flat"
        view_generated_tzevet_conan_btn.place(x=180,y=140,width=110,height=35)
        view_generated_tzevet_conan_btn["command"] = self.view_generated_tzevet_conan

        approve_tzevet_conan_btn=tk.Button(root)
        approve_tzevet_conan_btn["bg"] = "#ff677d"
        ft = tkFont.Font(family='Calibri',size=12)
        approve_tzevet_conan_btn["font"] = ft
        approve_tzevet_conan_btn["fg"] = "#000000"
        approve_tzevet_conan_btn["justify"] = "center"
        approve_tzevet_conan_btn["text"] = "אשר את הצוות"
        approve_tzevet_conan_btn["relief"] = "flat"
        approve_tzevet_conan_btn.place(x=40,y=140,width=110,height=35)
        approve_tzevet_conan_btn["command"] = self.approve_tzevet_conan


        GLabel_755=tk.Label(root)
        ft = tkFont.Font(family='Calibri Bold',size=12)
        GLabel_755["font"] = ft
        GLabel_755["fg"] = "#333333"
        GLabel_755["justify"] = "center"
        GLabel_755["text"] = ":שלב 1"
        GLabel_755.place(x=520,y=100,width=50,height=35)

        GLabel_167=tk.Label(root)
        ft = tkFont.Font(family='Calibri Bold',size=12)
        GLabel_167["font"] = ft
        GLabel_167["fg"] = "#333333"
        GLabel_167["justify"] = "center"
        GLabel_167["text"] = ":שלב 2"
        GLabel_167.place(x=380,y=100,width=50,height=35)

        GLabel_630=tk.Label(root)
        ft = tkFont.Font(family='Calibri Bold',size=12)
        GLabel_630["font"] = ft
        GLabel_630["fg"] = "#333333"
        GLabel_630["justify"] = "center"
        GLabel_630["text"] = ":שלב 3"
        GLabel_630.place(x=240,y=100,width=50,height=35)

        GLabel_638=tk.Label(root)
        ft = tkFont.Font(family='Calibri Bold',size=12)
        GLabel_638["font"] = ft
        GLabel_638["fg"] = "#333333"
        GLabel_638["justify"] = "center"
        GLabel_638["text"] = ":שלב 4"
        GLabel_638.place(x=100,y=100,width=50,height=35)

        GLabel_183=tk.Label(root)
        ft = tkFont.Font(family='Calibri Bold',size=28)
        GLabel_183["font"] = ft
        GLabel_183["fg"] = "#333333"
        GLabel_183["justify"] = "center"
        GLabel_183["text"] = "דף הבית"
        GLabel_183.place(x=240,y=40,width=120,height=35)

        GLabel_706=tk.Label(root)
        ft = tkFont.Font(family='Calibri Bold',size=12)
        GLabel_706["font"] = ft
        GLabel_706["fg"] = "#333333"
        GLabel_706["justify"] = "center"
        GLabel_706["text"] = ":פתח קבצים"
        GLabel_706.place(x=400,y=240,width=75,height=30)

        open_ilutzim_file_btn=tk.Button(root)
        open_ilutzim_file_btn["bg"] = "#cd6684"
        ft = tkFont.Font(family='Calibri',size=12)
        open_ilutzim_file_btn["font"] = ft
        open_ilutzim_file_btn["fg"] = "#000000"
        open_ilutzim_file_btn["justify"] = "center"
        open_ilutzim_file_btn["text"] = "קובץ אילוצים"
        open_ilutzim_file_btn["relief"] = "flat"
        open_ilutzim_file_btn.place(x=370,y=280,width=110,height=35)
        open_ilutzim_file_btn["command"] = self.open_ilutzim_file

        open_justice_board_file_btn=tk.Button(root)
        open_justice_board_file_btn["bg"] = "#cd6684"
        ft = tkFont.Font(family='Calibri',size=12)
        open_justice_board_file_btn["font"] = ft
        open_justice_board_file_btn["fg"] = "#000000"
        open_justice_board_file_btn["justify"] = "center"
        open_justice_board_file_btn["text"] = "לוח צדק"
        open_justice_board_file_btn["relief"] = "flat"
        open_justice_board_file_btn.place(x=370,y=340,width=110,height=35)
        open_justice_board_file_btn["command"] = self.open_justice_board_file

        GLabel_266=tk.Label(root)
        ft = tkFont.Font(family='Calibri Bold',size=12)
        GLabel_266["font"] = ft
        GLabel_266["fg"] = "#333333"
        GLabel_266["justify"] = "center"
        GLabel_266["text"] = ":עריכה"
        GLabel_266.place(x=190,y=240,width=75,height=30)

        open_edit_people_window_btn=tk.Button(root)
        open_edit_people_window_btn["bg"] = "#cd6684"
        ft = tkFont.Font(family='Calibri',size=12)
        open_edit_people_window_btn["font"] = ft
        open_edit_people_window_btn["fg"] = "#000000"
        open_edit_people_window_btn["justify"] = "center"
        open_edit_people_window_btn["text"] = "הוסף/מחק אדם"
        open_edit_people_window_btn["relief"] = "flat"
        open_edit_people_window_btn.place(x=140,y=280,width=110,height=35)
        open_edit_people_window_btn["command"] = self.open_edit_people_window

        open_change_file_loc_window_btn=tk.Button(root)
        open_change_file_loc_window_btn["bg"] = "#cd6684"
        ft = tkFont.Font(family='Calibri',size=12)
        open_change_file_loc_window_btn["font"] = ft
        open_change_file_loc_window_btn["fg"] = "#000000"
        open_change_file_loc_window_btn["justify"] = "center"
        open_change_file_loc_window_btn["text"] = "שנה מיקום קבצים"
        open_change_file_loc_window_btn["relief"] = "flat"
        open_change_file_loc_window_btn.place(x=140,y=340,width=110,height=35)
        open_change_file_loc_window_btn["command"] = self.open_change_file_loc_window

        GLabel_318=tk.Label(root)
        ft = tkFont.Font(family='Calibri Bold',size=12)
        GLabel_318["font"] = ft
        GLabel_318["fg"] = "#333333"
        GLabel_318["justify"] = "center"
        GLabel_318["text"] = "נבנה על ידי יואב שוומנטל באהבה למ.ע גף סייבר"
        GLabel_318.place(x=110,y=410,width=400,height=25)

    def reset_tzevet_conan(self):
        em.create_tzevet_conan_excel()
        tb.tzevet_conan = tb.get_tzevet_conan()


    def generate_tzevet_conan(self):
        tb.generate_all()
        tzevet = tb.tzevet_conan
        em.insert_generated_tzevet_conan_df_to_tzevet_conan_file(tzevet)


    def view_generated_tzevet_conan(self):
        file_path = "C:\\Users\\User\\Desktop\\Python\\Automatic schedule generator\\tzevet_conan.xlsx"
        os.startfile(file_path)


    def approve_tzevet_conan(self):
        em.insert_sum_to_justice_board(tb.get_tzevet_conan())


    def open_ilutzim_file(self):
        file_path = em.get_ilutzim_location()
        os.startfile(file_path)


    def open_justice_board_file(self):
        file_path = em.get_justice_board_location()
        os.startfile(file_path)


    def open_edit_people_window(self):
        """
        Create a new window for editing people
        """
        # Create a new windows and set size and title
        edit_people_window = tk.Toplevel(root)
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
        new_person_headline = tk.Label(master=edit_people_window,
                                       text="אדם חדש")
        new_person_headline.grid(row=0, column=5, sticky="nsew")
        new_person_headline.config(font=("calibri", 12))

        # Delete a person title
        delete_person_headline = tk.Label(master=edit_people_window,
                                          text="מחק אדם")
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
        toran_var = tk.IntVar()
        driver_var = tk.IntVar()

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
                               variable=samba_var).grid(row=2, column=4,
                                                        sticky="e")
        toran = tk.Checkbutton(edit_people_window,
                                               text="תורן יחידתי",
                                               variable=toran_var) \
            .grid(row=3, column=4, sticky="e")

        driver = tk.Checkbutton(edit_people_window,
                                text="נהג",
                                variable=driver_var).grid(row=4, column=4,
                                                          sticky="e")

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
        add_person = tk.Button(edit_people_window, text="הוסף בן אדם",
                               bg="#ff677d",
                               command=lambda: em.add_new_person(name.get(),
                                                                 manager_var,
                                                                 makel_officer_var,
                                                                 makel_operator_var,
                                                                 samba_var,
                                                                 toran_var,
                                                                 driver_var,
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


    def open_change_file_loc_window(self):
        """
        Create a new window for changing ilutzim and justice board files locations'
        """

        # Create a new windows and set size and title
        change_file_loc_windows = tk.Toplevel(root)
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
        ilutzim_headline = tk.Label(master=change_file_loc_windows,
                                    text="אילוצים")
        ilutzim_headline.grid(row=0, column=3, sticky="nsew")
        ilutzim_headline.config(font=("calibri", 12))

        # Get the location of the justice board file
        justice_board_file_loc = tk.Entry(change_file_loc_windows)
        justice_board_file_loc.grid(row=1, column=1, sticky="ew")
        justice_board_file_loc.insert(0, em.get_justice_board_location())

        # Get the location of the ilutzim file
        ilutzim_file_loc = tk.Entry(change_file_loc_windows)
        ilutzim_file_loc.grid(row=1, column=3, sticky="ew")
        ilutzim_file_loc.insert(0, em.get_ilutzim_location())

        # Save files locations button
        save_files_locations = tk.Button(change_file_loc_windows,
                                         text="שמור מיקומים", bg="#ff677d",
                                         command=lambda:
                                         em.save_files_new_locations(
                                             pd.read_csv('files_location.csv'),
                                             ilutzim_file_loc,
                                             justice_board_file_loc))
        save_files_locations.grid(row=3, column=1, sticky="nsew")


# ------------------------------------------------------------------------------

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


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
