from tkinter import *
import tkinter as tk
from tkinter import filedialog
from tkinter import Listbox
from utils import merge
import pandas as pd


class Window(tk.Tk):
    def __init__(self):
        # TODO https://stackoverflow.com/questions/19196130/select-all-in-a-tkinter-listbox
        super().__init__()
        self.files = []
        self.output = pd.DataFrame()

        self.wm_attributes('-alpha', 0.7)
        self.img = tk.PhotoImage(file='background.gif')
        self.w = self.img.width()
        self.h = self.img.height()
        self.label_background = tk.Label(image=self.img, width=800, height=1000)
        self.label_background.place(relx=0.00, rely=0.00, relheight=1, relwidth=1)

        self.geometry('690x490+600+150')
        self.resizable(width=False, height=False)
        self.configure()
        self.title("Merge 0.0.1")

        self.label = tk.Label(self, image=self.img, text="Excel File Merge Program", cursor='fleur')
        self.label.place(relx=0.05, rely=0.05, relheight=0.05, relwidth=0.3)

        self.button_files = tk.Button(text="Select Files", bg='White', command=self.select_file)
        self.button_files.place(relx=0.65, rely=0.4, relheight=0.1, relwidth=0.15)

        self.button_merge = tk.Button(text="Merge", width=200, pady=20, bg='White', command=self.merge_file)
        self.button_merge.place(relx=0.1, rely=0.8, relheight=0.1, relwidth=0.8)

        self.listbox_list_of_files = Listbox(listvariable=self.files)
        self.listbox_list_of_files.place(relx=0.20, rely=0.3, relheight=0.3, relwidth=0.4)

    def select_file(self):
        files_select = filedialog.askopenfilenames()
        for file in files_select:
            self.listbox_list_of_files.insert(END, file)

    def merge_file(self):
        self.output = merge(self.listbox_list_of_files.get(0, END))
        f = filedialog.asksaveasfilename(defaultextension='.xlsx')
        writer = pd.ExcelWriter(f, engine='xlsxwriter')
        self.output.to_excel(writer, sheet_name='RESULT')
        writer.save()
        print("completed")


if __name__ == '__main__':
    window = Window()
    window.mainloop()