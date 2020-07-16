# -*- coding: utf-8 -*-
"""
Created on Thu Jul 16 09:12:08 2020

@author: Austin.Schrader
"""

from win32com import client
import win32api
input_file = r'C:\Users\austin.schrader\Desktop\My_Desktop_Documents\Python_Tools\reading_emails_parse_attachment\Attachments\Report.xls'
#give your file name with valid path 
output_file = r'C:\Users\austin.schrader\Desktop\My_Desktop_Documents\Python_Tools\reading_emails_parse_attachment\Attachments\Report.pdf'
#give valid output file name and path
app = client.DispatchEx("Excel.Application")
app.Interactive = False
app.Visible = False
Workbook = app.Workbooks.Open(input_file)
try:
    Workbook.ActiveSheet.ExportAsFixedFormat(0, output_file)
except Exception as e:
    print("Failed to convert in PDF format.Please confirm environment meets all the requirements  and try again")
    print(str(e))
finally:
    Workbook.Close()
    app.Quit()