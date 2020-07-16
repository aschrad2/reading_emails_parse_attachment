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