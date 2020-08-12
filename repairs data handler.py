# Python 3.8.3, Openpyxl 3.0.4 
# Name: Michael Tunduli
# Email: wafutu@gmail.com 
# Date : 2nd July 2020

""" This script is to help in verification 
and data entry of UNSOS Equipment repairs in EMU.
It should compare what has been submitted by the
contractor with the data of what we already have.
It should show the new serials for the equipment 
and record them in appropriate worksheets
at the same time updating the old worksheet
with the necessary details from the submitted work sheets.

"""
#import the necessary modules required for the script to work
import os
import sys
import  openpyxl

from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
from openpyxl.utils.cell import coordinate_from_string 
from openpyxl.utils.cell import column_index_from_string

#load the desired xlsx workbook
def open_ac_data_file():
    global wb1
    global ws1
    global ac_data_file
    ac_data_file = "A:\\0Excel\\acdata.xlsx"
    print("\nOpening file_1.........")
    wb1 = load_workbook(ac_data_file)
    print("\n" + str(wb1.sheetnames))

    ws1 = wb1['Sheet1'] # open the specific sheet with data
    print("\n" + "max_row : " + str(ws1.max_row) + "\nmax_column :" + str(ws1.max_column))
    
open_ac_data_file()    

# load the desired second  xlsx workbook that is to be compared
def open_rai_ac_serials_list (): 
    global rai_ac_serials_list
    global ws2   
    rai_ac_serials_list = "A:\\0Excel\\acraidata.xlsx"
    print("\nOpening file_2..........")
    wb2 = load_workbook(rai_ac_serials_list)
    print("\n" + str(wb2.sheetnames))
    ws2 = wb2['Sheet1'] # open the specific sheet with data
    print("\n" + "max_row : " + str(ws2.max_row) + "\nmax_column :" + str(ws2.max_column))
    
open_rai_ac_serials_list()

def ac_serials_list():
    global ac_serials_list
    ac_serials_list = []

    for col in ws1.iter_cols(min_row = 2, min_col = 4, max_col = 4, max_row = 3295+1):
        for cell in col:
            if cell.value != None and cell.value != '':
                ac_serials_list.append(cell.value)
    print("\n----------------\n")        
    print("THIS IS ALL SERIALS LIST:\n")
    print("THE NUMBER OF ALL SERIALS FOUND IS :" + str(len(ac_serials_list))+"\n")        
    print(ac_serials_list)
    
ac_serials_list()
        
print("\n----------------")

def rai_serials_list():
    global rai_serials_list
    rai_serials_list = []
    for col in ws2.iter_cols(min_row = 5, min_col = 3, max_col = 3, max_row = 37):
        for cell in col:
            if cell.value != None and cell.value != '': 
                rai_serials_list.append(cell.value)
    print("\nTHIS IS RAI SERIALS LIST:\n") 
    print("THE NUMBER OF RAI SERIALS FOUND IS :" + str(len(rai_serials_list))+"\n")       
    print(rai_serials_list)
    
rai_serials_list()
        
print("\n----------------")

# compare the serial numbers of all the equipment we have with the serials of the submitted ones.
def compare_lists():
    global new_serial
    new_serial = []        
    for i in rai_serials_list:
        if i not in ac_serials_list:
            new_serial.append(i)
    print("\n-----------------\n")          
    print("THIS ARE NEW SERIALS:")
    print("THE NUMBER OF NEW SERIALS FOUND IS :" + str(len(new_serial))+"\n")
    print(new_serial)

compare_lists()

# get the name of all the rooms that we have and make a list
def all_my_rooms_list():
    global all_my_rooms_list
    all_my_rooms_list = []
    for col in ws1.iter_cols(min_row = 5, min_col = 9, max_col = 9, max_row = ws1.max_row +1):
        for cell in col:
            all_my_rooms_list.append(cell.value)
    
    print("\n-----------------\n")         
    print("THIS IS ALL ROOMS LIST:")
    print("THE NUMBER OF ALL ROOMS FOUND IS :" + str(len(all_my_rooms_list))+"\n")
    print(all_my_rooms_list)

all_my_rooms_list() 

# get the new rooms fromm he submitted works summary
def rai_rooms_list():
    global rai_rooms_list
    rai_rooms_list = []
    for col in ws2.iter_cols(min_row=5, min_col= 2, max_col=2, max_row = 45):
        for cell in col:
            rai_rooms_list.append(cell.value)
    
    print("\n-----------------")         
    print("THIS IS RAI ROOMS LIST:")
    print("THE NUMBER OF RAI ROOMS FOUND IS :" + str(len(rai_rooms_list))+"\n")
    print(rai_rooms_list)

rai_rooms_list()        

# compare the rooms from the current submitted works summary with all the number of rooms that we have
def compare_rooms():
    global new_rooms
    new_rooms = []        
    for i in rai_rooms_list:
        if i not in all_my_rooms_list:
            new_rooms.append(i)
    print("\n-----------------")        
    print("THIS ARE NEW ROOMS:")
    print("THE NUMBER OF ROOMS FOUND IS :" + str(len(new_rooms))+"\n")  
    print(new_rooms)
    
compare_rooms()

# copy the specific data to the required places in different worksheets

#getting data from the new rooms and copying to all rooms data:


    
