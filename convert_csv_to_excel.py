# -*- coding: utf-8 -*-
"""
Created on Thu Jul 16 11:51:53 2020

@author: Austin.Schrader
"""

from openpyxl import Workbook
import csv

input_file = r'C:\Users\austin.schrader\Desktop\My_Desktop_Documents\Python_Tools\reading_emails_parse_attachment\Attachments\Billing_Report_2020-07-10.csv'
output_file = r'C:\Users\austin.schrader\Desktop\My_Desktop_Documents\Python_Tools\reading_emails_parse_attachment\Attachments\Billing_Report_2020-07-10.xlsx'

wb = Workbook()
ws = wb.active
with open(input_file, 'r') as f:
    for row in csv.reader(f):
        ws.append(row)
wb.save(output_file)



# =============================================================================
# import os
# import win32com.client
# from openpyxl import Workbook
# import csv
# 
# for root, dirs, files in os.walk(r'C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\Python_Tools\\reading_emails_parse_attachment\\Attachments'):
#     for f in files:
# 
#         if f.endswith(".csv"):
#             print(f)
#             wb = Workbook()
#             ws = wb.active
#             with open(f, 'r') as f:
#                 for row in csv.reader(f):
#                     ws.append(row)
#             wb.save(f)
#             
# =============================================================================
        
# =============================================================================
# for root, dirs, files in os.walk(r'C:\Users\austin.schrader\Desktop\My_Desktop_Documents\Python_Tools\reading_emails_parse_attachment\Attachments\'):
#     for f in files:
# 
#         if f.endswith(".csv"):
#             print(f)
#             import pandas as pd
#             
#             read_file = pd.read_csv (f)
#             read_file.to_excel ('File name.xlsx', index = None, header=True)
# =============================================================================
