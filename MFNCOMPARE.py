# -*- coding: utf-8 -*-
"""
Created on Wed Sep 16 12:27:30 2020

@author: arao
"""
from xlwt import Workbook 
import xlrd 
  
# Give the location of the file 
#loc1 = ("E:\\Comparexls\\checkingxls.xlsx") 
loc1=input("Please input Cadence Generated BOM Path: ")
#loc2 = ("E:\\Comparexls\\checkingxlstest.xlsx")
loc2=input("Please input the path of BOM downloaded from Digikey/Mouser: ")
# To open Workbook 
wb = xlrd.open_workbook(loc1) 
sheet = wb.sheet_by_index(0) 
wb1 = xlrd.open_workbook(loc2) 
sheet1 = wb1.sheet_by_index(0) 
MFN1=0
Quantity1=0
DIC1={}
MFN2=0
Quantity2=0
DIC2={}
DIC1_L = []
z=0
x=0
for i in range(sheet.ncols):
    if(sheet.cell_value(0,i)=="Manufacturer Part Number"):
        MFN1=i
for i in range(sheet.ncols):
    if(sheet.cell_value(0,i)=="Quantity"):
        Quantity1=i
for i in range(1,sheet.nrows):
    DIC1_L.append(sheet.cell_value(i,MFN1))
    #print(sheet.cell_value(i,MFN1))
    DIC1[sheet.cell_value(i,MFN1)]=sheet.cell_value(i,Quantity1)
#print(DIC1_L)   
#print(DIC1)
for i in range( sheet1.ncols):
    if(  sheet1.cell_value(0,i)=="Manufacturer Part Number"):
        MFN2=i
for i in range( sheet1.ncols):
    if( sheet1.cell_value(0,i)=="Quantity"):
        Quantity2=i
for i in range(1, sheet1.nrows):
    DIC2[ sheet1.cell_value(i,MFN1)]= sheet1.cell_value(i,Quantity1)
#print(DIC2)

heading_l=["Manufacturer Part Number","Quantity Required","Quantity Ordered","Shortage","Surplus"]  
# Workbook is created 
wb = Workbook() 
sheet1 = wb.add_sheet('Sheet 1') 
for i,j in zip(range(len(heading_l)),heading_l):
    #print(0,i,j)
    sheet1.write(0,i,j) 
    #sheet1.write(c, i,i)
#sheet1.write(0, 0,"Manufacturer Part Number")   
#sheet1.write(0, 1,"Quantity")   
# add_sheet is used to create sheet. 
c=1
print("{:<30}{:<12}{:<12}{:<12}{:<12}".format("Manufacturer Part Number"," Quantity Required "," Quantity Ordered "," Shortage "," Surplus "))
for z in DIC1:
    if z in DIC2:
        #print(i)
        if(DIC1[z]>DIC2[z]):
            x = int(DIC1[z])-int(DIC2[z])
            s=0.0
        else:
            s= int(DIC2[z])-int(DIC1[z])
            x=0.0
        print("{:<30}{:12}{:12}".format(z,DIC1[z],DIC2[z]), "{:12}".format(x),"{:12}".format(s))
        data_l=[z,DIC1[z],DIC2[z],x,s]
        for y,q in zip(range(len(heading_l)),data_l):
            sheet1.write(c, y,q)
        #sheet1.write(c, 0,i)
        c+=1
    else:
        print("{:<30}{:12}{:>12}".format(z,DIC1[z],"0.0"), "{:12}".format(DIC1[z]),"{:12}".format(0.0))
        data_l=[z,DIC1[z],0.0,DIC1[z],0.0]
        for y,q in zip(range(len(heading_l)),data_l):
            sheet1.write(c, y,q)
        #print(DIC2[i]-DIC1[i])
        c+=1
for z in DIC2:
    if z not in DIC1:
        data_l=[z,0.0,DIC2[z],0.0,DIC2[z]]
        print("{:<30}{:>12}{:>12}".format(z,"0.0",DIC2[z]), "{:12}".format(0.0),"{:12}".format(DIC2[z]))
        for y,q in zip(range(len(heading_l)),data_l):
            sheet1.write(c, y,q)
        c+=1
wb.save('ResultSheet.xls') 

#    print(sheet.cell_value(i,MFN1),sheet.cell_value(i,Quantity1) )
    
        
        #print("Got it boss")
        
    #print(sheet.cell_value(0,i))
  
# For row 0 and column 0 
#print(sheet.cell_value(0,First_xls_c_n))