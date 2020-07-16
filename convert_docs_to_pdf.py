# -*- coding: utf-8 -*-
"""
Created on Thu Jul 16 09:08:48 2020

@author: Austin.Schrader
"""

import os
import win32com.client

wdFormatPDF = 17

for root, dirs, files in os.walk(r'C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\Python_Tools\\reading_emails_parse_attachment\\Attachments'):
    for f in files:

        if  f.endswith(".doc")  or f.endswith(".odt") or f.endswith(".rtf"):
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
                # os.remove(os.path.join(root,f))
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
                # os.remove(os.path.join(root,f))
        if f.endswith(".xlsx") or f.endswith(".xls"):
            try:
                print(f)
                in_file=os.path.join(root,f)
                app = win32com.client.DispatchEx("Excel.Application")
                app.Visible = False
                app.Interactive = False
                Workbook = app.Workbooks.Open(in_file)
                Workbook.SaveAs(os.path.join(0, root,f[:-5]), FileFormat=57)
                Workbook.Close()
                app.Quit()
                app.Exit()
                app.Visible = True
                print ('done')
                os.remove(os.path.join(root,f))
                
# =============================================================================
#                 in_file=os.path.join(root,f)
#                 app = win32com.client.DispatchEx("Excel.Application")
#                 app.Interactive = False
#                 app.Visible = False
#                 Workbook = app.Workbooks.Open(in_file)
#                 Workbook.ActiveSheet.ExportAsFixedFormat(os.path.join(root,f[:-5]), FileFormat=57)
#                 print(os.path.join(root,f[:-5]))
#                 Workbook.Close()
#                 app.Quit()
# =============================================================================
            except:
                print('could not open')
        else:
            pass