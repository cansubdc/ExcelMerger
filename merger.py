from tkinter import *
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import Listbox
from utils import merge
import pandas as pd


class Window(tk.Tk):
    def __init__(self):
        # TODO https://stackoverflow.com/questions/19196130/select-all-in-a-tkinter-listbox
        super().__init__()
        self.turn = 0
        self.files = []
        self.output = pd.DataFrame()

        self.img_window_background = tk.PhotoImage(file='assets/background.png')
        self.img_banner = tk.PhotoImage(file='assets/banner.png', height=100, width=800)

        self.geometry('780x490+600+150')
        self.resizable(width=False, height=False)

        self.can = tk.Canvas(self, width=790, height=490, bg='lightGrey')
        self.can.grid()

        self.can.create_image(390, 245, image=self.img_window_background)
        self.can.create_image(500, 70, image=self.img_banner)

        self.title("Excel File Merge Program")

        self.button_files = tk.Button(text="Browse", bg='White', command=self.select_file)
        self.button_files.place(relx=0.75, rely=0.3, relheight=0.1, relwidth=0.15)

        self.button_files = tk.Button(text="Delete File", bg='White', command=self.delete_file)
        self.button_files.place(relx=0.75, rely=0.43, relheight=0.1, relwidth=0.15)

        self.button_files = tk.Button(text="Delete All Files", bg='White', command=self.delete_all_files)
        self.button_files.place(relx=0.75, rely=0.56, relheight=0.1, relwidth=0.15)

        self.button_merge = tk.Button(text="Merge", width=200, pady=20, bg='White', command=self.merge_file,
                                      font=('Helvetica', '30'))
        self.button_merge.place(relx=0.05, rely=0.77, relheight=0.15, relwidth=0.9)

        self.listbox_list_of_files = Listbox(listvariable=self.files)
        self.listbox_list_of_files.place(relx=0.1, rely=0.3, relheight=0.36, relwidth=0.6)

        self.button_filepath = tk.Button(text="Show Full Path", bg='White', command=self.change_path_mode)
        self.button_filepath.place(relx=0.1, rely=0.67, relheight=0.06, relwidth=0.20)

    def change_path_mode(self):
        self.button_filepath.config(text='Show Full Path' if self.turn else 'Hide Path')
        self.listbox_list_of_files.delete(0, END)
        if self.turn:
            for file in self.files:
                file = file.rsplit("/", 1)[1]
                self.listbox_list_of_files.insert(END, file)
        else:
            for file in self.files:
                self.listbox_list_of_files.insert(END, file)
        self.turn = not self.turn

    def select_file(self):
        files_select = None
        try:
            files_select = filedialog.askopenfilenames(
                filetypes=[("Default Excel file", "*.xlsx"), ("Excel file 97-2003", "*.xls")])
            for file in files_select:
                self.files.append(file)
                file = file.rsplit("/", 1)[1]
                self.listbox_list_of_files.insert(END, file)
        except FileNotFoundError:
            messagebox.showerror("Error Message", "File not found")
        except Exception as e:
            messagebox.showerror('Awww', 'Unknown Error Occurred.')

    def merge_file(self):
        self.output = merge(self.files)

        if self.files.__len__() == 0:
            messagebox.showerror("Error Message", "You did not select source file")

        if self.output.empty:
            messagebox.showerror("Error Message", "Column Values Is Different")
            return
        else:
            f = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                             filetypes=[("Default Excel file", "*.xlsx"),
                                                        ("Excel file 97-2003", "*.xls")])
            if not f.rsplit("/", 1)[1].endswith('.xlsx' or '.xls'):
                messagebox.showerror("Error Message", "Only xlsx and xls File Types")
                return

            try:
                writer = pd.ExcelWriter(f, engine='xlsxwriter')
                self.output.to_excel(writer)
                writer.save()
                self.listbox_list_of_files.delete(0, END)
                messagebox.showinfo("", "File Merge Completed")

            except Exception as e:
                messagebox.showerror('Awww', 'Unknown Error Occurred.')

    def delete_file(self):
        selection_file = self.listbox_list_of_files.curselection()
        if selection_file == ():
            messagebox.showwarning("",
                                   "You did not select the file to delete")
        else:
            self.listbox_list_of_files.delete(ACTIVE)

    def delete_all_files(self):
        self.listbox_list_of_files.delete(0, END)
        del self.files[:]


if __name__ == '__main__':
    window = Window()
    window.mainloop()
