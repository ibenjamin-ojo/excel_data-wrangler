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
        
    def file_code(self, )