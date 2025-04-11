# Created by David N Preiss

# Create Customers Folder

# Open CustomersPools excel document

# iterate through the CustomerID column in the xlsx
# for each CustomerID
    # create a CustomerID folder in the Customers folder
    # create a CustomerID.html file in that CustomerID folder
    # iterate through the row of the CustomerID in the xlsx
    # for each PoolID
        # create a PoolID folder in the CustomerID folder
        # inside the PoolID folder
            # create a PoolID.html file
            # for each of the last 52 pdfs of PoolID
                # import and rename the pdf into the PoolID folder
                # add a link to the PoolID.html that opens PoolDate.pdf
            # save PoolID.html
        # add a link to the CustomerID.html that leads to PoolID.html
    # save CustomerID.html
# close CustomersPools excel document

## Import Statements

import shutil
import os
import subprocess
# System call
os.system("")

# Class of different styles
class style():
    BLACK = '\033[30m'
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    MAGENTA = '\033[35m'
    CYAN = '\033[36m'
    WHITE = '\033[37m'
    UNDERLINE = '\033[4m'
    RESET = '\033[0m'

try:
    import openpyxl
except ImportError as e:
    print(style.RED + f"!--ERROR:{e}\nopenpyxl is not installed. Installing..." + style.RESET)
    subprocess.check_call(["pip", "install", "openpyxl"])
    print("Installation complete. You can now run the script.")
    # exit()

## GLOBAL VALUES

MAX_WEEKS = 52
DUMP_FOLDER_PATH = "../BNRPools/print_populated_pools/Dump Folder"
CUSTOMERS_POOLS_XLSX = "Book1.xlsx"
C_ID_COL = 2
C_ID_START_ROW = 2
C_ID_END_ROW = 1000
P_ID_START_COL = 3
P_ID_END_COL = 1000

## CODE START


# Open CustomersPools excel document
workbook = openpyxl.load_workbook(CUSTOMERS_POOLS_XLSX)
worksheet = workbook.active


# Create Customers Folder
output_dir = "./Customers"
os.makedirs(output_dir, exist_ok=True)
print(f"Created output_dir:{output_dir}")


# iterate through the CustomerID column in the xlsx
for row in range(C_ID_START_ROW, C_ID_END_ROW):
    CustomerID = worksheet.cell(row=row, column=C_ID_COL).value
    if CustomerID is None:
        break
    print(f"\n CustomerID: {CustomerID}")
    # for each CustomerID
    # create a CustomerID folder in the Customers folder
    cidFolder = "./Customers/"+CustomerID
    if not os.path.exists(cidFolder):
        os.makedirs(cidFolder, exist_ok=True)
        print(f"Created cidFolder:{cidFolder}")
    # create a CustomerID.html file in that CustomerID folder
    
    # iterate through the row of the CustomerID in the xlsx
    for col in range(P_ID_START_COL, P_ID_END_COL):
        PoolID = worksheet.cell(row=row, column=col).value
        if PoolID is None:
            break
        print(f"   PoolID: {PoolID}")
        # for each PoolID
        # create a PoolID folder in the CustomerID folder
        pidFolder = cidFolder+"/"+PoolID
        if not os.path.exists(pidFolder):
            os.makedirs(pidFolder, exist_ok=True)
            print(f"Created pidFolder:{pidFolder}")
        
        # create a PoolID.html file in that PoolID folder
        # iterate through the PoolID folder in the dumpfile
        myPath = DUMP_FOLDER_PATH+"/"+str(PoolID)
        fnames = os.listdir(myPath)
        for file in range(0,MAX_WEEKS):
            try:
                print(f"\t{os.path.basename(fnames[file])}")
                # import and rename the pdf into the PoolID folder
                # add a link to the PoolID.html that opens PoolDate.pdf
            except IndexError as e:
                # print(f"{e}") # debug
                break
            except Exception as e:
                print(style.RED + f"!--ERROR:{e}" + style.RESET)
            break
        # save PoolID.html
        # add a link to the CustomerID.html that leads to PoolID.html
    # save CustomerID.html
    


# iterate through the CustomerID column in the xlsx
# for each CustomerID
    # create a CustomerID folder in the Customers folder
    # create a CustomerID.html file in that CustomerID folder
    # iterate through the row of the CustomerID in the xlsx
    # for each PoolID
        # create a PoolID folder in the CustomerID folder
        # inside the PoolID folder
            # create a PoolID.html file
            # for each of the last 52 pdfs of PoolID
                # import and rename the pdf into the PoolID folder
                # add a link to the PoolID.html that opens PoolDate.pdf
            # save PoolID.html
        # add a link to the CustomerID.html that leads to PoolID.html
    # save CustomerID.html

# close CustomersPools excel document
workbook.close()

# Prompt the user to press Enter before closing
input("Press Enter to close...")