import shutil
from datetime import datetime
import numpy as np
import pandas as pd
from tkinter import messagebox
import os
import warnings
import getpass


class Main:
    
    def __init__(self, days=None):
        self.days = days
        pd.set_option('display.max_rows', 500)

    def Main_processor(self):
        
        # Initialise pre-file reading class to access pre-file reading methods
        
        pre_file_reading = Pre_file_reading()

        # Copy files

        pre_file_reading.copy_files()

        # Initialise instance of file manipulation class as FM
        
        FM = File_manipulation()

        # Read in data

        self.df_ce, self.df = FM.read_xl()        

        # Intialise instances of classes for df_ce and df

        proc_tables = Table_processing(self.days)

        # Process dates in the ce table

        self.df_ce, self.df = proc_tables.main_table_processor(self.df_ce, self.df)

        ### Converting to string functions should come after table manipulation functions to avoid type errors as much as possible ###
        ### Any string amendment cleaning functions will require the data to be formatted as string so these should come after ###

        # turn all columns into strings
        
        self.df_ce, self.df = proc_tables.cols2string(self.df_ce, self.df)

        # Initialise cleaning class

        clean = Cleaning()

        # Remove N/As

        self.df_ce, self.df = clean.remove_na(self.df_ce, self.df)

        # Remove whitespace from all columns.

        self.df_ce, self.df = clean.remove_whitespace(self.df_ce, self.df)

        # capitalize nouns and block capitalize town name so formatting complies with Post Office standard

        self.df_ce, self.df = clean.capitalize_nouns(self.df_ce, self.df)

        # Make sure all emails are lowercase

        self.df_ce, self.df = clean.lowercase_emails(self.df_ce, self.df)

        # Remove DM URN
        
        self.df_ce, self.df = clean.remove_DM_URN(self.df_ce, self.df)

        # Format telephone numbers

        self.df_ce, self.df = clean.format_telephone(self.df_ce, self.df)

        # Remove "Householder" entries

        self.df_ce, self.df = clean.remove_householder(self.df_ce, self.df)

        # Handle business names -- NOT YET IMPLEMENTED

        # clean.handle_business_names(self.df_ce, self.df)

        # Description / functions

        RP = Reason_Processing()

        self.df_ce, self.df = RP.goneaway_proc(self.df_ce, self.df)

        # Select recent entries only and dedupe

        self.df_ce = proc_tables.get_recent_entries(self.df_ce, self.df, self.days)

        ### Note that any flagging functions should come after the get_recent_entries functions so that it can't affect the joins ###

        # Flag invalid entries class intitialised

        FIE = Flag_Invalid_Entries()

        # Flag invalid account numbers

        self.df_ce = FIE.flag_accounts_not_8(self.df_ce)

        # Flag invalid emails

        self.df_ce = FIE.flag_invalid_emails(self.df_ce)

        # Append the data to MasterV2

        FM.append_data(self.df_ce)
        
        # Retrieve the max dates from df_ce and update if higher than df

        max_dates = [pd.to_datetime(self.df['Date'].max(), dayfirst=True), pd.to_datetime(self.df_ce['Date'].max(), dayfirst=True)]

        
        self.max_date = max(max_dates)

 
            


class Pre_file_reading:

    def __init__(self):
        self.mv2_in_folder = "C:\\Users\\{}\\OneDrive - Novamedia\\Data & Campaign Team\\2023\\Player Issues & data\\data\\New data Process\\data MasterV2.xlsx".format(getpass.getuser())

    def copy_files(self):
        '''Copy all the original files to protect originals from corruption / error
        Program processes these copies only'''
        try:
            shutil.copy(r"\file\path\\data Automation Tool\data_CE.csv".format(getpass.getuser()), r"inputs\data_CE.csv")
            shutil.copy(self.mv2_in_folder, r"inputs\data MasterV2.xlsx")
            shutil.copy(self.mv2_in_folder, r"outputs\data MasterV2.xlsx")
            print("File copies made...")
        except PermissionError:
            file_permissions_error = '''***This application does not have permission to read the original files ('data_CE.csv' OR 'data MasterV2.xlsx').
            Please check no-one currently has these files open.'''
            raise PermissionError(f"{file_permissions_error}")
            messagebox.showerror(title="PermissionError", message=f"{file_permissions_error}")
        

class File_manipulation:
    

    def __init__(self):
        pass

    def read_xl(self):
        # catch warnings required due to warning shown by Openpyxl relating
        # to xls data validation feature not being supported. This does not
        # affect functionality of this application so warning suppressed.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            self.df_ce = pd.read_csv(r"inputs\data_CE.csv")
            self.df = pd.read_excel(r"inputs\data MasterV2.xlsx")
            print("Files read...")
            return self.df_ce, self.df

    @staticmethod
    def append_data(df):
        # Append data to end of spreadsheet and export
        try:
            with pd.ExcelWriter(
            r"outputs\data MasterV2.xlsx",
            mode="a",
            if_sheet_exists="overlay",
            ) as writer:
                df.to_excel(writer, sheet_name="data",startrow=writer.sheets['data'].max_row, engine="openpyxl", index=False, header=None)
            print("Excel file is now exported!")
            messagebox.showinfo(title="File successfully exported", message="The data data has been successfully appended to the MasterV2 file!")
            os.system('start EXCEL.EXE "outputs\data MasterV2.xlsx"')
        except BaseException:
            appending_data_error = "An unknown error has occured when trying to append the new data to the MasterV2 spreadsheet."
            messagebox.showerror(title="BaseException", message=f"{appending_data_error}")
            raise BaseException(f"{appending_data_error}")
        
        

class Table_processing:

    
    def __init__(self, days):
        # Initialisation
        self.days = days



    def main_table_processor(self, df_ce, df):
        self.df_ce, self.df = df_ce, df
        # Select only the relevant columns within each dataframe
        self.df_ce, self.df = self.select_columns(df_ce, df)
        # Replace column names so that both match (ce table inherits df column names)
        self.df_ce = self.rename_columns(self.df_ce, self.df)
        # Formats dates as dates
        self.df_ce, self.df = self.dates_proc(self.df_ce, self.df)
        return self.df_ce, self.df

        ### TO DO: Try to package all processing here

    def cols2string(self, df_ce, df, reverse=None):
        # take all columns and convert these to strings
        df_ce.columns.tolist()
        if reverse == None:
            self.df_ce = df_ce.astype("string")
            self.df = df.astype("string")
        # take all columns and infer the data-types
        if reverse == True:
            self.df_ce = df_ce.infer_objects()
            self.df = df.infer_objects()
        return self.df_ce, self.df
    
    
    def select_columns(self, df_ce, df):
        # A function to select only the data that is required. first 22 columns, all rows. See comment below.
        ### WILL BREAK IF 22 COLUMNS ARE EXCEEDED IN FUTURE VERSIONS OF SPREADSHEET. ##
        self.df_ce = df_ce.iloc[:,:22]
        self.df = df.iloc[:,:22]
        return self.df_ce, self.df


    
    def rename_columns(self, df_ce, df):
        '''Function to ensures both dataframes have the same column names so they can be brought into UNION.
        Could be depreciated as column names don't actually require changing for Excel Writer but note that
        Cleaning.capitalize_nouns method will break (or any other methods that reference column names directly)'''
        
        replace_cols = dict(zip(df_ce.columns, df.columns))
        
        self.df_ce = df_ce.rename(replace_cols, axis=1)
        
        return self.df_ce
        

                    

    def dates_proc(self, df_ce, df, process=None):
        self.df_ce, self.df = df_ce, df
        # Function to ensure that the dataframe dates are formatted correctly in both dataframes.
        if process == None:
            try:
                self.df_ce['Date'] = pd.to_datetime(self.df_ce['Date'], dayfirst=True)
                self.df['Date'] = pd.to_datetime(self.df['Date'], dayfirst=True)
                return self.df_ce, self.df
                
            except ValueError:
                date_error_message = "***There are currently some invalid dates present in the data data('Date' column)... Please amend these to ensure this program captures all the data.***"
                messagebox.showerror(title="Value Error", message=f"{date_error_message}")
                raise ValueError(f"{date_error_message}")
                return 1
            
        elif process == 'to_trunc':

            self.df_ce['Date'] =  pd.to_datetime(df_ce['Date']).dt.strftime('%d/%m/%Y').astype("string")
            self.df['Date'] =  pd.to_datetime(df['Date']).dt.strftime('%d/%m/%Y').astype("string")

            print("Dates truncated and converted to text format...")
            return self.df_ce, self.df

        else:
            raise ValueError("Invalid 'process' argument given to dates_proc() function. Please pass 'to_text' or None")

    

    def get_recent_entries(self, df_ce ,df , days):
        self.days = days
        
        if self.days == '':
            self.days = 1
        # Today minus days inputted by user = last update date.
        
        last_update_date = datetime.now().date() - pd.Timedelta(f'{self.days} days')

        print(f"Select all dates on/after: {last_update_date}")

        self.df_ce = df_ce[pd.to_datetime(df_ce['Date']).dt.date >= last_update_date]

        
        self.df = df[pd.to_datetime(df['Date']).dt.date >= last_update_date]

        # Process the dates to the correct format

        proc_tables = Table_processing(self.days)

        self.df_ce, self.df = proc_tables.dates_proc(self.df_ce.copy(deep=True), self.df.copy(deep=True), process='to_trunc')
        
        # Dedupe using pandas join


        keys = ['Date','Account Holder Number', 'First Name', 'Surname', 'Postcode', 'Email', 'Address1', 'Marketing Received']
        
        print("Number of data: " + str(len(self.df_ce)))


        # Deal with UTF - encoding discrepency in account numbers. Some of the account numbers show an invisible decimal place that prevents a merge.
        
        self.df_ce['Account Holder Number'] = self.df_ce['Account Holder Number'].astype(float)
        self.df['Account Holder Number'] = self.df['Account Holder Number'].astype(float)
        self.df_ce['Account Holder Number'] = self.df_ce['Account Holder Number'].round().astype("Int64").astype("string")
        self.df['Account Holder Number'] = self.df['Account Holder Number'].round().astype("Int64").astype("string")

        
        self.df_ce = pd.merge(self.df_ce, self.df, on=keys, how='left', indicator=True, suffixes=('', '_drop'))

        # Create two tables: 1 that shows rows that already exist and don't need to be added and one that shows all the rows that need adding

        already_exist_rows = self.df_ce[self.df_ce['_merge'] == 'both']

        self.df_ce = self.df_ce[self.df_ce['_merge'] == 'left_only']


        # Drop unnecessary columns ie. indicator and "_drop" cols
        
        self.df_ce.drop([col for col in self.df_ce if ('drop') in col], axis=1, inplace=True)

        self.df_ce = self.df_ce.drop('_merge', axis=1)
        

        # Show duplicates to help with mismatch investigation

        
        # Show all duplicated rows based on the keys

        duped_rows = self.df_ce[self.df_ce.duplicated(subset=keys)]

        duped_rows = pd.concat([already_exist_rows, duped_rows])

        if len(duped_rows) != 0:
            print("\nDuplicates or already present in MasterV2: \n", duped_rows[['Account Holder Number', 'Surname', 'Postcode']], "\n")
        
        self.df_ce = self.df_ce.drop_duplicates(subset=keys)


        # Show total without duplicates

        print("Without Duplicates: " + str(len(self.df_ce)))

        return self.df_ce


class Cleaning:

    def __init__(self):
        ...

    def remove_na(self, df_ce, df):
        # Some NAs are automatically removed when data was processed by pandas
        # This function should remove the rest
        self.dfs = [df_ce, df]
        for dfi in self.dfs:
            for column in dfi.columns:   
                dfi[column] = dfi[column].str.replace('N/A', '', case=False)

        print("'N/A' Values Removed...")
        return self.dfs

    def remove_whitespace(self, df_ce, df):
        self.df_ce, self.df = df_ce, df
        # Remove whitespaces from all columns
        try:
            for column in self.df_ce.columns:
                self.df_ce[column] = self.df_ce[column].str.strip()
            for column in self.df.columns:
                self.df[column] = self.df[column].str.strip()
            print("Trailing and leading spaces removed...")
            return self.df_ce, self.df
        
        except BaseException:
            raise
            raise BaseException("An error has occured while trying to remove whitespaces from the data...")
        

    def capitalize_nouns(self, df_ce, df):
        self.df_ce, self.df = df_ce, df
        # Capitalize nouns and converts town name to block capitals as per post-office standard format
        noun_columns = ["First Name", "Surname", "Address1", "Address2", "Address3", "County", "Title", "Address Line 1", "Address Line 2", "Address Line 3", "Last Name"]
        for colname in noun_columns:
            try:
                self.df_ce[colname] = df_ce[colname].str.title()
                self.df[colname] = df[colname].str.title()
            except KeyError:
                continue
                
        self.df_ce[['Town', 'Postcode']] = self.df_ce[['Town', 'Postcode']].apply(lambda x: x.str.upper())
        
        print("Capitalized proper nouns...")
        return self.df_ce, self.df

    def lowercase_emails(self, df_ce, df):
        self.dfs = [df_ce, df]
        for dfi in self.dfs:
            dfi['Email'] = dfi['Email'].str.lower()
        print("Standardised Email casing...")
        return self.dfs
    

    def remove_DM_URN(self, df_ce, df):
        self.df_ce, self.df = df_ce, df
        
        self.df_ce['Direct Mail URN Number'] = pd.Series(dtype=int)
        self.df['Direct Mail URN Number'] = pd.Series(dtype=int)
       
        print("Removed DM URN Numbers...")
        
        return self.df_ce, self.df


    def format_telephone(self, df_ce, df):
        self.df_ce, self.df = df_ce, df
        
        # a function to format telephone numbers consistently with our database (i.e. 44...)
        phone_columns = ['Landline','Mobile']

        for column in phone_columns:
            self.df_ce[column] = self.df_ce[column].str.replace("^(0|7|1)", "44", regex = True)
            self.df_ce[column] = self.df_ce[column].str.replace("\+44", "44", regex = True)
            self.df_ce[column] = self.df_ce[column].str.replace(" ", "")

            self.df[column] = self.df[column].str.replace("^(0|7|1)", "44", regex = True)
            self.df[column] = self.df[column].str.replace("\+44", "44", regex = True)
            self.df[column] = self.df[column].str.replace(" ", "")
            # Convert to float, int and then string
            self.df[column] = pd.to_numeric(self.df[column], errors="coerce").round().astype("Int64").astype("string")
            self.df_ce[column] = pd.to_numeric(self.df_ce[column], errors="coerce").round().astype("Int64").astype("string")

        print("Telephone numbers formatted...")
        return self.df_ce, self.df

    def remove_householder(self, df_ce, df):
        self.df_ce, self.df = df_ce, df
        name_columns = ['First Name', 'Last Name', 'Surname']
        for column in name_columns:
            try:
                self.df_ce[column] = self.df_ce[column].str.replace("(House?holder|Housheolder)", "", regex=True)
                self.df[column] = self.df[column].str.replace("(House?holder|Housheolder)", "", regex=True)
            except KeyError:
                continue
        print("'Householder' removed from the name columns...")
        return self.df_ce, self.df


    ## WORK IN PROGRESS, FUNCTION TO HANDLE BUSINESS NAMES ##
    def handle_business_names(self, df_ce, df):
        # I don't believe that the below implementation will work maybe use apply for this...
        self.df_ce, self.df = df_ce, df
        # if ~self.df_ce["Title"].str.contains("Mr|Ms|Ms|Mrs|^$", case=False)):
        print(self.df_ce.index) 
        #self.df_ce["Address1"]
        
        
    

class Reason_Processing:

    def __init__(self):
        ...

    def goneaway_proc(self, df_ce, df):
        # A function to replace entries with "Gone away" where they have been incorrectly labelled.
        self.dfs = [df_ce, df]
        # Iterate through dataframes and mask to replace text to 'Gone away' where relevant.
        for dfi in self.dfs:
            moved = dfi['Description'].str.contains('moved', case=False)
            dfi['data Type'] = dfi['data Type'].mask(moved, 'Gone away')
        print("'Moved' descriptions assigned 'Gone away' data reason...")
        
        return self.dfs


class Flag_Invalid_Entries:

    def __init__(self):
        ...

    def flag_accounts_not_8(self, df_ce):
        self.df_ce = df_ce
        # feature to pick up accounts numbers with less than 8 but not 0
        # Function to flag when account numbers don't have 8 digits
        self.df_ce['Account Holder Number'][(self.df_ce['Account Holder Number'].str.len() != 8) & (self.df_ce['Account Holder Number'].str.len() != 0)] = "*INVALIDACCOUNTNO*: " + self.df_ce['Account Holder Number']
        print("Flagged invalid Account Numbers...")
        return self.df_ce
        
    def flag_invalid_emails(self, df_ce):
        self.df_ce = df_ce
        self.df_ce['Email'][(~self.df_ce['Email'].str.contains('@', na=False)) & (self.df_ce['Email'] != '')] = "*INVALIDEMAIL*: " + self.df_ce['Email']
        print("Flagged invalid Emails...")
        return self.df_ce
   




