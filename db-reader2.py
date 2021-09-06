import pandas as pd
import pyodbc
import os
from openpyxl import *

user = ""# Add your username between the
passw = ""# Add your password between the

tables = ["1. A_DH_ASSAY_NUM", "2. A_DH_ASSAY_NUM_CACHED", "3. A_DH_ASSAY_NUM_NAME",
          "4. A_DH_ASSAY_NUM_NAME_CACHED", "5. A_DH_ASSAY_NUM_TOT",
          "6. A_DH_ASSAY_NUM_TOT_CACHED", "7. A_DH_ASSAY_SEL_NUM", "8. A_DH_ASSAY_SEL_NUM_CACHED",
          "9. A_DH_ASSAY_TXT", "10. A_DH_ASSAY_TXT_CACHED", "11. A_DH_ASSAY_TXT_NAME",
          "12. A_DH_ASSAY_TXT_NAME_CACHED", "13. A_DH_ASSAY_TXT_TOT", "14. A_DH_ASSAY_TXT_TOT_CACHED",
          "15. A_DH_BULKDENS_NUM", "16. A_DH_COLLAR", "17. A_DH_DIAMETER", "18. A_DH_GEOLOGY", "19. A_DH_GEOTECH",
          "20. A_DH_MAGSUSC", "21. A_DH_RECOVERY", "22. A_DH_STRUCTURE", "23. A_DH_SURVEY", "24. A_DS_RETURN",
          "25. A_DS_SEND"]

def save_to_file():     
    
    filename = ""

    save_file = input("Do you want to save the file? y/n: ")
    if save_file == "y":
        save_file_y = input("Do you want to save it as: (1) Excel or (2) CSV? ")
        if save_file_y == "1":
            filename = "db_"+ddh_in+"_"+table_in+".xlsx"
            df.to_excel(filename)
        elif save_file_y == "2":
            filename = "db_"+ddh_in+"_"+table_in+".csv"
            df.to_csv(filename)
            
    path = os.path.abspath(filename)
    directory = os.path.dirname(path)

    if save_file == "y":
        print("File saved to " + directory)
    
def get_db_no_ddh():

    global df
    
    cnxn = pyodbc.connect(r"Driver={SQL Server Native Client 11.0};"
                          "Server=;" # Add server here
                          "Database=acQ_Copperstone;"
                          "uid={%s};pwd={%s}" % (user, passw))

    df = pd.read_sql_query("select * from %s" % (table_in), cnxn)
    print("\nDataframe saved successfully!\n")
    return df

def get_db_ddh():

    global df
    
    cnxn = pyodbc.connect(r"Driver={SQL Server Native Client 11.0};"
                          "Server=;" # Add server here
                          "Database=acQ_Copperstone;"
                          "uid={%s};pwd={%s}" % (user, passw))

    df = pd.read_sql_query("select * from %s where HOLEID='%s'" % (table_in, ddh_in), cnxn)
    print("\nDataframe saved successfully!\n")
    return df

if not user:
    user = input("Username: ")
else:
    pass
if not passw:
    passw = input("Password: ")

print(*tables, sep = "\n")

print("\nLogged in as: " + user)

table_in = input("\nChoose a table: ")
ddh_in = "all_ddh"

usr_input = input("\nDo you need a specific ddh, y/n? ")
if usr_input == "y":
    ddh_in = input("\nWhich ddh do you need?: ")
    get_db_ddh()
elif usr_input == "n":
    get_db_no_ddh()

print("Dataframe head: \n")
print(df.head(), "\n")

save_to_file()
