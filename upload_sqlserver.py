import pyodbc
import pandas as pd
import sqlalchemy as sa
from datetime import date, timedelta
from tkinter import messagebox
import getpass

class Upload_sqls:

    def __init__(self, days):
        self.days = int(days)
        self.date_range = []
        for i in range(self.days):
            self.date_range.append(str(date.today() - timedelta(days = i+1)))

    def upload_and_format(self):
        
        # Print running upload to notify user that upload process has begun

        print("Running upload scripts...")

        # Read the excel file as df
        try:
            df = pd.read_excel(r"\file\path\data MasterV2.xlsx".format(getpass.getuser()), dtype=object)
        except:
            raise OSError("\n\n ***Please close the master file when trying to upload the data.***")
        
        # Initialise SQL engine
        
        engine = sa.create_engine("*********//*******/*********?driver=ODBC+Driver+17+for+SQL+Server") # "ODBC+Driver+17+for+SQL+Server" Try if error SQL+Server+Native+Client+11.0

        # Read in schema.data

        df_db = self.read_db(engine)

        print("MasterV2 and schema.data_FINAL read...")

        # Before doing anything else, check if the new data you wish to append are not already in the database

        insert_date_today = df_db[df_db["insert_date"].astype("string").str.contains(str(date.today()))]
        
        
        if len(insert_date_today) == 0:
            
            # Upload
            self.process_and_upload(df, df_db, engine)

        else:
            ovrwrite_msg = "The data entries you are trying to upload have already been uploaded today. Do you want to overwrite these?"
            confirm_overwrite = messagebox.askquestion(title="data already been uploaded today", message=ovrwrite_msg, icon='warning')
            if confirm_overwrite == 'yes':
                # Upload but delete previous entries uploaded that day.
                self.process_and_upload(df, df_db, engine, already_present=True)
            else:
                messagebox.showinfo(title='No upload', message='data have not been uploaded')
                
    def read_db(self, engine):
        with engine.connect() as connection:
            # retrieves the last 2 weeks worth of data data from final table
            df_db = pd.read_sql(sa.text("SELECT * FROM schema.data_FINAL WHERE insert_date BETWEEN '{}' AND '{}' ORDER BY insert_date".format(date.today() - timedelta(days = 14), date.today())), con=connection)
        return df_db


    def process_and_upload(self, df, df_db, engine, already_present=None):
        if already_present == True:
            with engine.connect() as connection:
                connection.execute(sa.text("DELETE FROM schema.data_FINAL WHERE insert_date = '{}';".format(date.today())))
                connection.commit()
                print("Entries previously entered today have been deleted...")
        # Select only the relevant columns

        df = df.iloc[:,:22]

        df = df.drop(columns="County")


        # upload the data append

        self.upload_append_data(df, engine)

        # Format the dates as appropriate and append the data to the retention_data_final table

        self.append_append_data(engine)

        # check the final rows, print them out

        df_db = self.read_db(engine)

        print(df_db.tail(10))

        print("{} rows added to the retention_data table..".format(len(df_db[df_db['new_date'].isin(self.date_range)])))

        print("New entries for today have been added. DONE.")
        

    def upload_append_data(self, df, engine):
        # Insert dataframe into SQL Server (Campaigns.data_append):
        with engine.connect() as connection:
            connection.execute(sa.text("Truncate table campaigns.data_append;"))
            connection.commit()
            print("==Truncation complete==")
            print("Upload has begun, please wait...")
            df.to_sql(r'data_append', con=engine, schema="campaigns", if_exists='append', index=False)
            print("==Upload completed!==")
            




    def append_append_data(self, engine):
        # Connect to the database
        with engine.connect() as connection:
            for i in range(self.days):
            # Iterate over the dates, format them as required and then execute the SQL.
                date_dash_format = date.today() - timedelta(days = (i+1))

                date_slash_format = date_dash_format.strftime("%d/%m/%Y")

                sql = '''Declare @DateofNewRecords varchar(50);  
                SET @DateofNewRecords ='{}';
                Declare @DifferentDateFormat varchar(50);
                SET @DifferentDateFormat = '{}';

                Insert into schema.data_FINAL
                select [Date],[Agent Name],[Contact Method],[Marketing Received] ,[Account Holder Number],[Direct Mail URN Number],[Title],[First Name],[Surname]
                      ,[Address1],[Address2],[Address3],[Town],[Postcode],[email],[Landline],[Mobile],[data Type],[Description],[Recipient of Communication]
                          ,[Suppressed on Beehive],[Additional Information],[new_date],[insert_date] 
                from campaigns.data_append
                where 
                [Date] = @DateofNewRecords

                update schema.data_FINAL
                Set insert_date=CONVERT(date,GETDATE(),105)
                where
                 [Date]=@DateofNewRecords

                update schema.data_FINAL
                Set new_date=@DifferentDateFormat
                where 
                [Date]= @DateofNewRecords
                
                '''.format(date_slash_format, date_dash_format)

                connection.execute(sa.text(sql))
                connection.commit()
            print("Retention_data_final table has been fully updated!")

        

#Upload_sqls(1).upload_and_format()
