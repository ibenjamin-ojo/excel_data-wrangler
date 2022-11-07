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

df = pd.concat(map(pd.read_csv, files), ignore_index = True)

df