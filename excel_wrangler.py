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

# loaing packages. 
xl = pd.read_excel('extract_2022-11.xlsx', sheet_name = '367-369',engine='openpyxl', na_values=['nan'])
