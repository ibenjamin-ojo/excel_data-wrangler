# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 22:55:44 2022

@author: Benjamin Ojo
"""
import pandas as pd
import os

'''
This script combines the pervious script (excel_extractor, excel_wrangler, 
and excel_marger) into one unifered pipeline that extract, clean and formate
the data, and merge excel sheet into one workbook. 
'''

# Creating a data wrangling pipeline. 

Class Excel_Pipeline:
    """ 
    This combine the last three python script into one that can 
    performe the three functions.
    """
    def __init__(self, folder_path, save_path): 
        
        self.folder_path = folder_path
        self.save_path = save_path
        
    def file_code(self):
        """Extracting file code and return them as list"""
        
        # Print list of file in a folder. 
        files = os.listdir(folder_path)
        
        codes = []

        for file in files:
            extr = file.split(' ')[1].split('.')[0]
            codes.append(extr)
            
        self.codes = codes
        
        return self.codes
    
    def excel_workbook(self): 
        """The create an excel workbook, and code sheet"""
        
        # Creating excel sheet. 
        wb = openpyxl.workbook.Workbook()

        for i in range(len(self.codes)):
            i = wb.create_sheet(codes[i])
            
    def sheet_copier(self):
        """
        This function copies each excel sheet from all files to combined 
        file with only the interest sheet we want to work with and export out
        a saved excel file. 
        """
        
        # loading copy file
        print('Loading coping sheet.', '\n')
        for code in self.codes: 
            copy_path = self.save_path + f'/CONTAINER {code}.xlsx'

            copy_wb = openpyxl.load_workbook(copy_path)
            copy_ws = copy_wb[code]

            past_ws = wb[code]
            
            # Get maximum row and columns for colume
            print(f'Getting Maximum row & column from data. Container-{code}', '\n')
            rm = copy_ws.max_row
            cm = copy_ws.max_column

            # coping sheet to extract folder. 
            print(f'Coping data from copy sheet. Container-{code}', '\n')
            for i in range(1, rm + 1): 
                for j in range(1, cm + 1):
                    c = copy_ws.cell(row = i , column = j)
                    
                    past_ws.cell(row = i, column = j).value = c.value
            print(f'Coping sheet complete. Container-{code}', '\n\n')


        wb.save(saving_path)
        print('Saving Pasting sheet', '\n')
    
    def excel_wrangler(self, exctrack_file):
        """
        This function will clean the dataset and combine it into a sigle
        excel file. 
        """
        
        for num in range(len(self.codes)): 
            
            # loaing packages.
            xl = pd.read_excel('extract_2022-11.xlsx', sheet_name = self.codes[num],
                               engine='openpyxl', na_values=['nan'])
            
            print(f"FORMATING: {self.codes[num]},\n")
            
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
            
            # Column Header. 
            columns = xl.iloc[header_idx,:].to_list()
            
            # # converting first row to header. 
            xl.columns = columns
            
            # Deleting rows
            xl.drop([i for i in range(header_idx+1)], axis=0, inplace=True)
            
            # Reset_index
            xl.reset_index(drop = True)
            
            # Creating columns.
            xl['CONTAINER_CODE'] = code
            xl['SUPPLIER_NAME'] = ' '
            
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
            
            # Defining supplier name
            xl['SUPPLIER_NAME'] = supplier_name
            
            # Rows to Delete.
            delete_row = xl[xl['SUPPLIER_NAME'] == 'delete_row'].index.to_list()

            # Drop rows. 
            xl.drop(delete_row, axis = 0, inplace = True)
           
            # Reset index. 
            xl = xl.reset_index(drop = True)
            
            # Rows that aren't needed.
            sup_del = xl[xl['SUPPLIER_NAME'] == xl['DESCRIPTION']].index
            
            # Drop rows. 
            xl.drop(sup_del, axis = 0, inplace = True)

            # Reset index. 
            xl = xl.reset_index(drop = True)
            
            # Reordering columns 
            xl = xl.loc[:, xl.columns[-2:].to_list() + xl.columns[:-2].to_list()]
            
            # Saving file
            xl.to_csv(self.save_path + f'/csv_files/{self.codes[num]}.csv', index = False)
            print(f'SAVING: {self.codes[num]}.csv', '\n\n')
    
    def excel_marger(self):
        """
        We will be merging all the csv files we created when cleaning our data
        into one excel file
        """
        csv_path = self.save_path + f"/csv_files"
        
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
        
            
            
        