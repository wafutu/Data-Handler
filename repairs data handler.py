# Python 3.8.3, Openpyxl 3.0.4 
# Name: Michael Tunduli
# Email: wafutu@gmail.com 

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

#load the desired xlsx workbook
def open_file1():
    global wb1
    global ws
    global file1
    file1 = "A:\\0Excel\\acdata.xlsx"
    wb1 = load_workbook(file1)
    print(wb1.sheetnames)

    ws = wb1['Sheet1'] # open the specific sheet with data

    print('\n')
    print(ws.max_row)
    print(ws.max_column)
    
open_file1()    

def my_list():
    global my_list
    my_list = []

    for col in ws.iter_cols(min_row=2, min_col=8, max_col=8, max_row=3295):
        for cell in col:
            my_list.append(cell.value)
    print("\n----------------")        
    print("THIS IS LIST 1:\n")        
    print(my_list)
    
my_list()
        
print("\n----------------")
print("\n----------------\n")

# load the desired second  xlsx workbook that is to be compared
file2 = "A:\\0Excel\\acraidata.xlsx"
wb2 = load_workbook(file2)
print(wb2.sheetnames)

ws2 = wb2['Sheet1'] # open the specific sheet with data

print('\n')
print(ws2.max_row)
print(ws2.max_column)

def my_list2():
    global my_list2
    my_list2 = []
    for col in ws2.iter_cols(min_row=5, min_col= 3, max_col=3, max_row=37):
        for cell in col:
            my_list2.append(cell.value)
    print("\nTHIS IS LIST 2:\n")        
    print(my_list2)
    
my_list2()
        
print("\n----------------")
print("\n----------------\n")

# compare the serial numbers of all the equipment we have with the serials of the submitted ones.
def compare_lists():
    global new_serial
    new_serial = []        
    for i in my_list2:
        if i not in my_list:
            new_serial.append(i)
    print("\n-----------------")          
    print("THIS ARE NEW SERIALS:")
    print(new_serial)

compare_lists()

# get the name of all the rooms that we have and make a list keep a list of them
def all_my_rooms_list():
    global all_my_rooms_list
    all_my_rooms_list = []
    for col in ws.iter_cols(min_row=5, min_col= 5, max_col=5, max_row = ws.max_row +1):
        for cell in col:
            all_my_rooms_list.append(cell.value)
    
    print("\n-----------------")         
    print("THIS IS ALL ROOMS LIST:")
    print("THE NUMBER OF ROOMS FOUND IS :" + str(len(all_my_rooms_list))+"\n")
    print(all_my_rooms_list)

all_my_rooms_list() 

# get the new rooms fromm he submitted works summary
def new_rooms_list():
    global new_rooms_list
    new_rooms_list = []
    for col in ws2.iter_cols(min_row=5, min_col= 2, max_col=2, max_row = 45):
        for cell in col:
            new_rooms_list.append(cell.value)
    
    print("\n-----------------")         
    print("THIS IS NEW ROOMS LIST:")
    print("THIS IS THE NUMBER OF ROOMS FOUND :" + str(len(new_rooms_list))+"\n")
    print(new_rooms_list)

new_rooms_list()        

# compare the rooms from the current submitted works summary with all the number of rooms that we have
def compare_rooms():
    global new_rooms
    new_rooms = []        
    for i in new_rooms_list:
        if i not in all_my_rooms_list:
            new_rooms.append(i)
    print("\n-----------------")        
    print("THIS ARE NEW ROOMS:")
    print("THIS IS THE NUMBER OF ROOMS FOUND :" + str(len(new_rooms))+"\n")  
    print(new_rooms)
    
compare_rooms()

# copy the specific data to the required places in different worksheets


    
