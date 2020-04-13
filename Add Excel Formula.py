# -*- coding: utf-8 -*-
"""
Created on Mon Apr 13 12:29:29 2020

@author: Róisín Anglim
"""

#Import Packages
import openpyxl
from openpyxl.formula.translate import Translator

#Reads in excel workbook
workbook = openpyxl.load_workbook(r"C:\Users\Harvey Norman\Desktop\Git Excel\Report Data.xlsx")

#Add formula to cell in workbook
wb1 = workbook["Summary"]
wb1["A2"]="=Sum(1,1)"

#Inserts column into workbook
wb1.insert_cols(61)
wb1.move_range("z1:z10",rows=0,cols=1,translate=True)

#Save updated workbook
workbook.save(r"C:\Users\Harvey Norman\Desktop\Git Excel\Report Data Jan.xlsx")



