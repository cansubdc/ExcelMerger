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

        self.img = tk.PhotoImage(file='assets/background.gif')

        self.w = self.img.width()
        self.h = self.img.height()
        self.label_background = tk.Label(image=self.img, width=800, height=1000)
        self.label_background.place(relx=0.00, rely=0.00, relheight=1, relwidth=1)

        self.geometry('690x490+600+150')
        self.resizable(width=False, height=False)
        self.configure()
        self.title("Merge 0.0.1")

        self.label = tk.Label(self, text="Excel File Merge Program", cursor='fleur', bg='ghost white')
        self.label.place(relx=0.05, rely=0.05, relheight=0.05, relwidth=0.3)

        self.button_files = tk.Button(text="Select Files", bg='White', command=self.select_file)
        self.button_files.place(relx=0.65, rely=0.4, relheight=0.1, relwidth=0.15)

        self.button_merge = tk.Button(text="Merge", width=200, pady=20, bg='White', command=self.merge_file)
        self.button_merge.place(relx=0.1, rely=0.8, relheight=0.13, relwidth=0.8)

        self.listbox_list_of_files = Listbox(listvariable=self.files)
        self.listbox_list_of_files.place(relx=0.20, rely=0.3, relheight=0.3, relwidth=0.4)

        self.button_filepath = tk.Button(text="Show Full Path", bg='White', command=self.change_path_mode)
        self.button_filepath.place(relx=0.3, rely=0.62, relheight=0.05, relwidth=0.20)

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
        try:
            files_select = filedialog.askopenfilenames(filetypes=[("Default Excel file", "*.xlsx"), ("Excel file 97-2003", "*.xls")])

        except FileNotFoundError:
            messagebox.showerror("Error Message", "File not found")
        except Exception as e:
            messagebox.showerror('Awww', 'Unknown Error Occurred.')

        for file in files_select:
            self.files.append(file)
            file = file.rsplit("/", 1)[1]

            self.listbox_list_of_files.insert(END, file)

    def merge_file(self):
        self.output = merge(self.files)
        if self.files.__len__() == 0:
            messagebox.showerror("Error Message", "You did not select source file")
        else:
            f = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                         filetypes=[("Default Excel file", "*.xlsx"), ("Excel file 97-2003", "*.xls")])

            if not f.rsplit("/", 1)[1].endswith('.xlsx' or '.xls'):
                messagebox.showerror("Error Message", "Only xlsx and xls File Types")
                return
            self.listbox_list_of_files.insert(END, "Your Selected Target Files " + f.rsplit("/", 1)[1])
            try:
                writer = pd.ExcelWriter(f, engine='xlsxwriter')
                self.output.to_excel(writer)
                writer.save()
            except Exception as e:
                messagebox.showerror('Awww', 'Unknown Error Occurred.')
            self.listbox_list_of_files.insert(END, "Merge Completed to " + f.rsplit("/", 1)[1])


if __name__ == '__main__':
    window = Window()
    window.mainloop()