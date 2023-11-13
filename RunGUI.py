# Import packages

from tkinter import *
from tkinter import messagebox, ttk
from Main import Main
from upload_sqlserver import Upload_sqls
import os
import pandas as pd


# Define functions

class Interface:

    def __init__(self):
        self.max_date = None
        self.integer_error_message ="***Number of Days entry is not a positive integer. Please enter a positive integer value.***"
    
    def data_activate(self, days_var):
        # Runs all the data formatting functions from "Main.py" when button is pressed to generate data
        self.days_var = days_var
        self.get_daysvar()
        Run = Main(self.days_var)
        Run.Main_processor()
        self.max_date = Run.max_date

    def upload_activate(self, days_var):
        # Runs all the upload functions from upload_sqlserver.py when button is pressed to upload
        self.days_var = days_var
        self.get_daysvar()
        Upload_sqls(self.days_var).upload_and_format()

    def get_daysvar(self):
        # Method function that also checks input days are valid positive integer and flags error if not.
        self.days_var = self.days_var.get()
        try:
            if self.days_var == '':
                self.days_var = 1
                return

            elif int(self.days_var) < 0:
                messagebox.showerror(title="ValueError", message=f"{self.integer_error_message}")
                raise ValueError(f"{self.integer_error_message}")
                return 1
        except ValueError:
            messagebox.showerror(title="ValueError", message=f"{self.integer_error_message}")
            raise ValueError(f"{self.integer_error_message}")


    def Run_GUI(self):
        # Initialise GUI

        GUI = Tk()

        GUI.geometry("400x200")
        GUI.title("data Automation Tool")

        # Set heading

        heading = Label(text = '''Enter number of days worth of data you intend to capture.
                If this field is left blank then the number will default to 1 day.''',
                         bg = "black", fg = "white", height = "3", width = "600")

        heading.pack()

        # Create number of days input

        days_var = StringVar()

        days_func_text = Label(text = "Number of Days:")

        days_entry = ttk.Entry(textvariable = days_var)

        days_func_text.place(x = "40", y = "65")

        days_entry.place(x="140", y = "66")


        # Create and place button to generate data

        gen_data = ttk.Button(text = "Generate data", command = lambda: self.data_activate(days_var))

        gen_data.place(x= "4", y = "120", width = "392")

        # Create and place a button to transfer data from the data spreadsheet

        run_Upload = ttk.Button(text = "Upload data (SQL Server Only)", command = lambda: self.upload_activate(days_var))

        run_Upload.place(x= "70", y = "150", width = "250")

        # Run GUI Loop

        GUI.mainloop()

Run_program = Interface()

Run_program.Run_GUI()
