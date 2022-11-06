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

print(codes, '\n')


for num in range(3): 
    # loaing packages.
    xl = pd.read_excel('extract_2022-11.xlsx', sheet_name = codes[num],
                       engine='openpyxl', na_values=['nan'])
    
    # Getting the container numbers.
    c_list = []
    
    for i in range(4): 
        c = xl.iloc[i, :].dropna().to_list()
        c = [str(d) for d in c]
        c_list.append(c)
        
    for j in range(len(c_list)): 
        if len(c_list[j])>=1 and c_list[j][0][:3] == 'CON':
            code = c_list[j][0]
            code_idx = j
            header_idx = j + 1
            break
    print(f'Container {code[num]}: {code}', '\n')
    
    # Column Header. 
    columns = xl.iloc[header_idx,:].to_list()
    print(f"Container {code[num]} columns: {columns}", '\n')
    
    # # converting first row to header. 
    xl.columns = columns
    print('Replacing None Columns', '\n')
    
    # Deleting rows
    xl.drop([i for i in range(header_idx+1)], axis=0, inplace=True)
    print('Dropping rows with headers', '\n')
    
    # Reset_index
    xl.reset_index(drop = True)
    
    # Creating columns.
    xl['CONTAINER_CODE'] = code
    xl['SUPPLIER_NAME'] = ' '
    print('Creating Container_code and Supplier name columns', '\n')
    
    # Drop s/n column.
    del xl[xl.columns[0]]
    
    # Supplier and item code list 
    supplier = xl['DESCRIPTION'].fillna('re').to_list()
    item_code = xl['CODE'].fillna('re').to_list()
    quantity = xl['QUANTITY'].fillna('re').to_list()
    
    # Extracting Supplier name for row. 
    supplier_name = []
    
    for i in range(len(supplier)):
        if supplier[i] != 're' and item_code[i] == 're' and quantity[i] == 're':
            supplier_name.append(supplier[i])
        elif supplier[i] == 're' and item_code[i] == 're'and quantity[i] == 're':
            supplier_name.append('delete_row')
        else:
            cont=supplier_name[-1].replace('~', '')
            supplier_name.append(f'{cont}~')
    print('Identifing individual puchures with supplier name', '\n')

    
    # Defining supplier name
    xl['SUPPLIER_NAME'] = supplier_name
    print('Adding Supplier name to rows')
    
    # Rows to Delete.
    delete_row = xl[xl['SUPPLIER_NAME'] == 'delete_row'].index.to_list()

    # Drop rows. 
    xl.drop(delete_row, axis = 0, inplace = True)
    print('Deleting rows thare are none')
    
    # Reset index. 
    xl = xl.reset_index(drop = True)
    print('Resetting Dataset')
    
    # Rows that aren't needed.
    sup_del = xl[xl['SUPPLIER_NAME'] == xl['DESCRIPTION']].index
    # Drop rows. 
    xl.drop(sup_del, axis = 0, inplace = True)
    print('Deleting supplier name row')
    
    # Reset index. 
    xl = xl.reset_index(drop = True)
    
    # Reordering columns 
    xl = xl.loc[:, xl.columns[-2:].to_list() + xl.columns[:-2].to_list()]
    print('Reorder Columns', '\n')
    
    # Saving file
    xl.to_csv(f'csv_files/{codes[num]}.csv', index = False)
    
    



    
    

    
    
    
