# -*- coding: utf-8 -*-
"""
Created on Mon Nov  7 14:17:33 2022

@author: User1
"""
# Import packages. 
import pandas as pd 
import os 

# File name.
cont_folder = 'C:/Users/User1/Documents/Data_Project/excel_data-wrangler/csv_files'

# Print list of file in a folder. 
files = os.listdir(cont_folder)
print(files)

excel_append = pd.DataFrame()
for file in files[:50]:
    origi = pd.read_csv(f'csv_files/{file}')
    excel_append = excel_append.append(origi, ignore_index =  True)

excel_append = excel_append.iloc[:, :12]

excel_append1 = pd.DataFrame()
for file in files[50:100]:
    origi = pd.read_csv(f'csv_files/{file}')
    excel_append1 = excel_append1.append(origi, ignore_index =  True)

excel_append1 = excel_append1.iloc[:, :12]

excel_append2 = pd.DataFrame()
for file in files[100:150]:        
    origi = pd.read_csv(f'csv_files/{file}')
    excel_append2 = excel_append2.append(origi, ignore_index = True)

excel_append2 = excel_append2.iloc[:, :12]

excel_append3 = pd.DataFrame()
for file in files[150:240]:
    origi = pd.read_csv(f'csv_files/{file}')
    excel_append3 = excel_append3.append(origi, ignore_index = True)

excel_append3 = excel_append3.iloc[:, :12]

excel_merge = excel_append.append(excel_append1, ignore_index = True)

excel_merge.append(excel_append2, ignore_index = True)

excel_merge.append(excel_append3, ignore_index = True)

excel_merge

excel_merge.to_excel('excel_merge.xlsx')

