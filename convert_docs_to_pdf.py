# -*- coding: utf-8 -*-
"""
Created on Thu Jul 16 09:08:48 2020

@author: Austin.Schrader
"""

import os
import win32com.client
from openpyxl import Workbook
import csv

wdFormatPDF = 17

for root, dirs, files in os.walk(r'C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\Python_Tools\\reading_emails_parse_attachment\\Attachments'):
    for f in files:

        if f.endswith(".csv"):
            
            in_file = os.path.join(root, f)
            wb = Workbook()
            ws = wb.active
            with open(in_file, 'r') as f:
                for row in csv.reader(f):
                    ws.append(row)
            #output_file = os.path.join(root,f[:-4])
            wb.save(output_file)
        elif  f.endswith(".doc")  or f.endswith(".odt") or f.endswith(".rtf"):
            try:
                print(f)
                in_file=os.path.join(root,f)
                word = win32com.client.Dispatch('Word.Application')
                word.Visible = False
                doc = word.Documents.Open(in_file)
                doc.SaveAs(os.path.join(root,f[:-4]), FileFormat=wdFormatPDF)
                doc.Close()
                word.Quit()
                word.Visible = True
                print ('done')
                os.remove(os.path.join(root,f))
                pass
            except:
                print('could not open')
        elif f.endswith(".docx") or f.endswith(".dotm") or f.endswith(".docm"):
            try:
                print(f)
                in_file=os.path.join(root,f)
                word = win32com.client.Dispatch('Word.Application')
                word.Visible = False
                doc = word.Documents.Open(in_file)
                doc.SaveAs(os.path.join(root,f[:-5]), FileFormat=wdFormatPDF)
                doc.Close()
                word.Quit()
                word.Visible = True
                print ('done')
                os.remove(os.path.join(root,f))
                pass
            except:
                print('could not open')
        elif f.endswith(".xlsx") or f.endswith(".xlsm"):
            try:
                print(f)
                in_file=os.path.join(root,f)
                #give your file name with valid path 
                output_file = os.path.join(root,f[:-5])
                #give valid output file name and path
                app = win32com.client.DispatchEx("Excel.Application")
                app.Interactive = False
                app.Visible = False
                app.DisplayAlerts = False
                Workbook = app.Workbooks.Open(in_file)
                Workbook.ActiveSheet.ExportAsFixedFormat(0, output_file)
                Workbook.Close()
                app.Quit()
                os.remove(os.path.join(root, f))
                pass
            except:
                print('could not open')
        elif f.endswith(".xls"):
            try:
                print(f)
                in_file=os.path.join(root,f)
                #give your file name with valid path 
                output_file = os.path.join(root,f[:-4])
                #give valid output file name and path
                app = win32com.client.DispatchEx("Excel.Application")
                app.Interactive = False
                app.Visible = False
                app.DisplayAlerts = False
                Workbook = app.Workbooks.Open(in_file)
                Workbook.ActiveSheet.ExportAsFixedFormat(0, output_file)
                Workbook.Close()
                app.Quit()
                os.remove(os.path.join(root, f))
                pass
            except:
                print('could not open')
        elif f.endswith("csv"):
            import pypandoc
            
            output = pypandoc.convert_file('somefile.csv', 'pdf', outputfile="somefile.pdf")
            assert output == ""
            
        else:
            pass