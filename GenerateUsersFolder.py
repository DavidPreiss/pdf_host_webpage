
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
import os
import subprocess

try:
    import openpyxl
except ImportError as e:
    print(style.RED + f"!--ERROR:{e}\nopenpyxl is not installed. Installing..." + style.RESET)
    subprocess.check_call(["pip", "install", "openpyxl"])
    print("Installation complete. You can now run the script.")
    # exit()
MAX_WEEKS = 52
DUMP_FOLDER_PATH = "C:/Users/David/Desktop/code/BNRPools/print_populated_pools/Dump Folder"
CUSTOMERS_POOLS_XLSX = "Book1.xlsx"
C_ID_COL = 2
C_ID_START_ROW = 2
C_ID_END_ROW = 1000
P_ID_START_COL = 3
P_ID_END_COL = 1000


workbook = openpyxl.load_workbook(CUSTOMERS_POOLS_XLSX)
worksheet = workbook.active

for row in range(C_ID_START_ROW, C_ID_END_ROW):
    CustomerID = worksheet.cell(row=row, column=C_ID_COL).value
    if CustomerID is None:
        break
    print(f"\n CustomerID: {CustomerID}")
    # for each CustomerID
    # create a CustomerID folder in the Customers folder
    # create a CustomerID.html file in that CustomerID folder
    
    # iterate through the row of the CustomerID in the xlsx
    for col in range(P_ID_START_COL, P_ID_END_COL):
        PoolID = worksheet.cell(row=row, column=col).value
        if PoolID is None:
            break
        print(f"   PoolID: {PoolID}")
        # for each PoolID
        # create a PoolID folder in the CustomerID folder
        # create a PoolID.html file in that PoolID folder
        # iterate through the PoolID folder in the dumpfile
        myPath = DUMP_FOLDER_PATH+"/"+str(PoolID)
        fnames = os.listdir(myPath)
        for file in range(0,MAX_WEEKS):
            try:
                print(f"\t{os.path.basename(fnames[file])}")
                # import and rename the pdf into the PoolID folder
                # add a link to the PoolID.html that opens PoolDate.pdf
            except Exception as e:
                # print(f"done")
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