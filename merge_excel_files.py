# -*- coding: utf-8 -*-
"""
Created on Mon Feb  3 13:44:28 2020

@author: Róisín Anglim 

"""
# Combine multiple files into one based on the same column names

import pandas as pd
# Read in CSV files into dataframes
marketingfile = pd.read_excel("MarketingAnalystNames.xlsx")
salesfile = pd.read_excel("SalesRepNames.xlsx")
seniorfiles = pd.read_excel("SeniorLeadershipNames.xlsx")

#create a list for the files
all_files = [marketingfile,salesfile,seniorfiles]

# Pandas will append files based on column names 
append_df = pd.concat(all_files)
append_df.to_excel ("all_files.xlsx",index =False)