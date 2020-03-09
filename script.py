#IMPORTANT BEFORE YOU RUN
#==>You Need To Have Python 3 Installed Then Run These Commands On The CMD (for windows)
#==>pip install tkinter
#==>pip install csv
#==>pip install pathlib
#==>pip install xlrd

#--------------------------------------------------------------#
#==>The Script Is Very Easy To Use => Just Run and Don't Bother Your Self With Editing The Code Or Making Custom Inputs <==#
#--------------------------------------------------------------#

#==>Every Step Is Will Explained As Follows <==#
#--1->Choose The Exel File Where Variables Are Found ===> File NOT Folder
#--2->Choose The Folder Where Text Files Are Found ===> Folder NOT File
#--3->Choose The Final Save Path To Save The Final CSV File ===> Folder NOT File

#You Only Need To Change These Values First ↓IMPORTANT↓
num_of_cols = 4     #Number Of Columns Inside The Exel File (Variables File)
num_of_rows = 3     #Number Of Rows Inside The Exel File (Variables File)
intitialFileName = 'Final File'     #The Default Name Of The Final Result File ==>Don't Include Extension ==> .i.e 'FinalFile.csv' should be 'FinalFile'

import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import csv
import pathlib
import xlrd

############################################################

print("Select exel file to get variables from...")
csv_path = filedialog.askopenfilename()

csv_file = os.fspath(csv_path)
wb = xlrd.open_workbook(csv_file)
sheet = wb.sheet_by_index(0)
bigVarList = []

for x in range(1,num_of_rows):
    varList = []
    for y in range(num_of_cols):
        varList.append(sheet.cell_value(x,y))
    if isinstance(varList[0],(float,int)):
        varList[0] = str(int(varList[0]))
    if isinstance(varList[1],(float,int)):
        varList[1] = str(int(varList[1]))   
    bigVarList.append(varList)

############################################################

print("Select Text Files Directory....")
txt_dir_path = filedialog.askdirectory()

print("Select where to save the result....")
ftypes = [('Csv File','.csv'),('All Files','*')]
save_path = filedialog.asksaveasfilename(filetypes = ftypes,defaultextension='.csv',initialfile = intitialFileName)

final_file = open(save_path,'w')
for txtFile in pathlib.Path(txt_dir_path).iterdir():
    txtFileName = os.path.basename(txtFile)
    flag = 0
    if txtFile.is_file() and txtFileName.endswith('.txt'):
        o_txtFile = open(txtFile,'r')
        num_of_line = -1
        bigListMatch = []
        for line in o_txtFile:
            if num_of_line == -1:
                num_of_line =+ 1
                continue
            item_list = line.split()
            f_item_list = item_list[0:2] + item_list[3:5]
                        
            for var in bigVarList:
                if var == f_item_list:
                    flag = 1
                    final_file.write(txtFileName)
                    final_file.write('\n')
                    final_file.write(','.join(item_list))
                    final_file.write('\n')
            num_of_line += 1
    if f_item_list not in bigVarList and flag == 0:
            final_file.write(txtFileName)
            final_file.write('\n')

#For Furhter Information About The Code ==> Please Contact Me :)



    


        
