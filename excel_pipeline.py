# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 22:55:44 2022

@author: User1
"""
import pandas as pd
import os

'''

This script combines the pervious script (excel_extractor, excel_wrangler, 
and excel_marger) into one unifered pipeline that extract, clean and formate
the data, and merge excel sheet into one workbook. 

'''

# Creating a data wrangling pipeline. 

Class excel_pipeline:
    """ 
    This combine the last three python script into one that can 
    performe the three functions.
    """
    def __init__(self, file_path, save_path): 
        
        self.file_path = file_path
        self.save_path = save_path
        
    def file_code(self):
        """Extracting file code and return them as list"""
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
        
        