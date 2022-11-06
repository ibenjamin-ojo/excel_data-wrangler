# -*- coding: utf-8 -*-
"""
Created on Mon Oct 31 14:31:59 2022

@author: Benjamin Ojo
"""
# Import packages. 
import pandas as pd 
import os 

# Excel packages.
from openpyxl import Workbook as wb 
from openpyxl import load_workbook as lwb

# File name.
cont_folder = 'C:/Users/User1/Documents/Data_Project/excel_data-wrangler/excel_files/container_file'

# Print list of file in a folder. 
files = os.listdir(cont_folder)

# Container code 
codes = []

for file in files:
    extr = file.split(' ')[1].split('.')[0]
    codes.append(extr)

print(codes)
# loaing packages. 
#xl = pd.read_excel('extract_2022-11.xlsx', sheet_name = '367-369',engine='openpyxl', na_values=['nan'])
