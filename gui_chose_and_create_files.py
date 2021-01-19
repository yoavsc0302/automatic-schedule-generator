import tkinter as tk
import tkinter.font as tkFont
import excel_modifications as em
import re

class ChoseAndCreateFiles:
    def __init__(self, root):
        #setting title
        root.title("שיבוץ צוות כונן אוטומטי")
        #setting window size
        width=500
        height=250
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        browse_files_label=tk.Label(root)
        ft = tkFont.Font(family='Calibri Bold',size=13)
        browse_files_label["font"] = ft
        browse_files_label["fg"] = "#333333"
        browse_files_label["justify"] = "center"
        browse_files_label["text"] = "בחר מיקומי קבצים"
        browse_files_label.place(x=325,y=20,width=124,height=25)

        choose_tzevet_conan_btn=tk.Button(root)
        choose_tzevet_conan_btn["bg"] = "#ff677d"
        ft = tkFont.Font(family='Calibri',size=10)
        choose_tzevet_conan_btn["font"] = ft
        choose_tzevet_conan_btn["fg"] = "#000000"
        choose_tzevet_conan_btn["justify"] = "center"
        choose_tzevet_conan_btn["text"] = "צוות כונן"
        choose_tzevet_conan_btn["relief"] = "flat"
        choose_tzevet_conan_btn.place(x=350,y=80,width=70,height=25)
        choose_tzevet_conan_btn["command"] = self.choose_tzevet_conan_btn

        choose_justice_board_btn=tk.Button(root)
        choose_justice_board_btn["bg"] = "#ff677d"
        ft = tkFont.Font(family='Calibri',size=10)
        choose_justice_board_btn["font"] = ft
        choose_justice_board_btn["fg"] = "#000000"
        choose_justice_board_btn["justify"] = "center"
        choose_justice_board_btn["text"] = "לוח צדק"
        choose_justice_board_btn["relief"] = "flat"
        choose_justice_board_btn.place(x=350,y=130,width=70,height=25)
        choose_justice_board_btn["command"] = self.choose_justice_board_btn

        choose_ilutzim_btn=tk.Button(root)
        choose_ilutzim_btn["bg"] = "#ff677d"
        ft = tkFont.Font(family='Calibri',size=10)
        choose_ilutzim_btn["font"] = ft
        choose_ilutzim_btn["fg"] = "#000000"
        choose_ilutzim_btn["justify"] = "center"
        choose_ilutzim_btn["text"] = "אילוצים"
        choose_ilutzim_btn["relief"] = "flat"
        choose_ilutzim_btn.place(x=350,y=180,width=70,height=25)
        choose_ilutzim_btn["command"] = self.choose_ilutzim_btn

        create_new_files_label=tk.Label(root)
        ft = tkFont.Font(family='Calibri Bold',size=12)
        create_new_files_label["font"] = ft
        create_new_files_label["fg"] = "#333333"
        create_new_files_label["justify"] = "center"
        create_new_files_label["text"] = "צור קבצים מחדש במקרה שנמחקו או התבלגנו"
        create_new_files_label.place(x=20,y=20,width=290,height=25)

        create_new_files_sub_label = tk.Label(root)
        ft = tkFont.Font(family='Calibri',size=10)
        create_new_files_sub_label["font"] = ft
        create_new_files_sub_label["fg"] = "#333333"
        create_new_files_sub_label["justify"] = "center"
        create_new_files_sub_label["text"] = "(ייתכן שימחקו הקבצים הקיימים)"
        create_new_files_sub_label.place(x=20,y=45,width=260,height=15)

        create_new_tzevet_conan_btn=tk.Button(root)
        create_new_tzevet_conan_btn["bg"] = "#cd6684"
        ft = tkFont.Font(family='Calibri',size=10)
        create_new_tzevet_conan_btn["font"] = ft
        create_new_tzevet_conan_btn["fg"] = "#000000"
        create_new_tzevet_conan_btn["justify"] = "center"
        create_new_tzevet_conan_btn["text"] = "צור קובץ צוות כונן חדש"
        create_new_tzevet_conan_btn["relief"] = "flat"
        create_new_tzevet_conan_btn.place(x=80,y=80,width=120,height=50)
        create_new_tzevet_conan_btn["command"] = self.create_new_tzevet_conan_btn

        create_new_justice_board_btn=tk.Button(root)
        create_new_justice_board_btn["bg"] = "#cd6684"
        ft = tkFont.Font(family='Calibri',size=10)
        create_new_justice_board_btn["font"] = ft
        create_new_justice_board_btn["fg"] = "#000000"
        create_new_justice_board_btn["justify"] = "center"
        create_new_justice_board_btn["text"] = "צור קבצי לוח צדק\n ואילוצים חדשים"
        create_new_justice_board_btn["relief"] = "flat"
        create_new_justice_board_btn.place(x=80,y=155,width=120,height=50)
        create_new_justice_board_btn["command"] = self.create_new_justice_board_and_ilutzim_btn

        # create_new_ilutzim_btn=tk.Button(root)
        # create_new_ilutzim_btn["bg"] = "#cd6684"
        # ft = tkFont.Font(family='Calibri',size=10)
        # create_new_ilutzim_btn["font"] = ft
        # create_new_ilutzim_btn["fg"] = "#000000"
        # create_new_ilutzim_btn["justify"] = "center"
        # create_new_ilutzim_btn["text"] = "צור קובץ אילוצים חדש"
        # create_new_ilutzim_btn["relief"] = "flat"
        # create_new_ilutzim_btn.place(x=80,y=180,width=120,height=25)
        # create_new_ilutzim_btn["command"] = self.create_new_ilutzim_btn

    def choose_tzevet_conan_btn(self):
        location, filename = em.browse_file("Tzevet conan:")
        filename = filename.name
        em.save_files_new_locations(type_of_file='tzevet_conan',
                                    file_path=filename)


    def choose_justice_board_btn(self):
        location, filename = em.browse_file("Justice board:")
        filename = filename.name
        em.save_files_new_locations(type_of_file='justice_board',
                                    file_path=filename)

    def choose_ilutzim_btn(self):
        location, filename = em.browse_file("Ilutzim:")
        filename = filename.name
        em.save_files_new_locations(type_of_file='ilutzim',
                                    file_path=filename)

    def create_new_tzevet_conan_btn(self):
        em.create_tzevet_conan_excel()


    def create_new_justice_board_and_ilutzim_btn(self):
        em.create_justice_board_excel()
        em.create_ilutzim_excel()




if __name__ == "__main__":
    root = tk.Tk()
    app = ChoseAndCreateFiles(root)
    root.mainloop()

