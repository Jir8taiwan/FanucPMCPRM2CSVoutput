#!/usr/bin/python3
#vim:fileencoding=utf-8
import pandas as pd
#import update
import csv
import os, sys
import openpyxl
import time
#from tqdm import tqdm

#####################
# Author: JIR
# Version. 2022.10.24
#
# Using PYTHON3 language to open and convert at FANUC controller system of parameter backup.
# Please copy "PMC1_PRM.TXT" in program folder together.
# It will output CSV and EXCEL files for studying in a formatted data.
#
#####################
print("Fanuc PMCPRM to CSV format output by JIR made")
print()
## Configure
TXTinputPATH = 'PMC1_PRM.TXT'           #參數檔案
CSVoutputPATH2 = '1_PMC1_PRM.CSV'         #CSV輸出，無篩選。
CSVoutputPATH3 = '2_PMC1_PRM_trimed.CSV'  #CSV輸出，篩選空資料。
XLSXoutputPATH2 = '3_PMC1_PRM.xlsx'       #EXCEL輸出，無篩選。
XLSXoutputPATH = '4_PMC1_PRM_trimed.xlsx' #EXCEL輸出，篩選空資料。

## Compare and calculate 以下計算用
TEMPouputPATH = 'tempoutput.txt'
TEMPouputPATH2 = 'tempoutput2.txt'
CSVoutputPATH = 'OutputTemp.CSV'

## DEF
def filter_rows_by_values(df, col, values):
    return df[~df[col].isin(values)]


## Check PMC1_PRM.TXT, if no existed is exit
if not os.path.exists(TXTinputPATH):
    print()
    print("Please confirm \"", TXTinputPATH, "\" file that is correct filename or really existed in folder.")
    print()
    time.sleep(5)
    sys.exit(0)
    

## Go Loop and data formating
print("Converting and processing!!")
txt = ""
with open(TXTinputPATH, 'r', encoding='UTF-8') as file:
    while (line := file.readline().rstrip()):
        if "%" in line.strip():
            line = ""
        if "(PMC=0I-F,MSID=1)" in line.strip():
            line = ""

        #paramNUM = str(line[:6])
        paramNUM = str(line[:8])
        #print(paramNUM)
        checkNUM = str(paramNUM[1:3]).strip()
        try:
            PARAMnum = int(checkNUM)
        except ValueError:
            PARAMnum = 0
        
        #print(PARAMnum)
        #print(checkNUM.isspace())
        #print(type(PARAMnum))

        if (PARAMnum == 60):
            Txxx = str(paramNUM[4:7]).strip()
            line = line.replace(" P", ",TIMER,") + ",T" + Txxx
        if (PARAMnum == 61):
            Cxxxx = str(paramNUM[3:7]).strip()
            line = line.replace(" P", ",COUNTER,") + ",C" + Cxxxx
        if (PARAMnum == 62):
            Kxxx = str(paramNUM[4:7]).strip()
            line = line.replace(" P", ",KEEP RELAY,") + ",K" + Kxxx
        if (PARAMnum == 63):
            line = line.replace(" P", ",DATA TABLE,")
        if (PARAMnum == 64):
            line = line.replace(" P", ",DATA,")

        line = line.replace(" P", ",,")
        
        #print("Formating:", line)
        txt += line + "\n"

        ##develop usage
        #print(txt)
#print("force exit")
#sys.exit(0)

f = open(TEMPouputPATH,'w',encoding='utf-8')
#txt = txt.strip()
f.write(txt)
f.close()

file.close()

#print()
#print(txt) 

## Output file
print("Outputing new format CSV file")
f = open(TEMPouputPATH,'w', encoding='utf-8')
txt = txt.strip()
f.write(txt)
f.close()

with open(TEMPouputPATH, 'r', encoding='utf8') as in_file:
    stripped = (line.strip() for line in in_file)
    lines = (line.split(",") for line in stripped if line)
    with open(CSVoutputPATH, 'w') as out_file:
        writer = csv.writer(out_file, dialect="excel")
        writer.writerow(['ParamNO', 'DataName', 'Value', 'Note'])
        writer.writerows(lines)
        out_file.close()


with open(CSVoutputPATH, newline='', encoding='utf8') as in_file:
    with open(CSVoutputPATH2, 'w', newline='') as out_file:
        writer = csv.writer(out_file)
        for row in csv.reader(in_file):
            if row:
                writer.writerow(row)
        out_file.close()
        
# check temp file and exited is to delete
if os.path.exists(CSVoutputPATH):
    os.remove(CSVoutputPATH)

## Make a CSV format file without 0 value of data
print("Outputing CSV file without 0 value data")
df = pd.read_csv(CSVoutputPATH2, encoding='utf8', dtype=str)
if os.path.exists(CSVoutputPATH3):
    os.remove(CSVoutputPATH3)
new_df = filter_rows_by_values(df, "Value", ["0","00000000", "0.0"])
#print(new_df)
new_df.to_csv(CSVoutputPATH3, index=False, encoding='utf8')

## Make a EXCEL format file with and without 0 value of data
print("Outputing EXCEL file from CSV data")
read_file = pd.read_csv(CSVoutputPATH3, encoding='utf8', dtype=str)
read_file.to_excel(XLSXoutputPATH, index=None, header=True)
read_file = pd.read_csv(CSVoutputPATH2, encoding='utf8', dtype=str)
read_file.to_excel(XLSXoutputPATH2, index=None, header=True)

## Finish to know
print("Conveting is DONE!!")
print()
print("It will close in 5 seconds automatically...")
time.sleep(5)


      
