# -*- coding: utf-8 -*-
"""
Created on Thu Jul 16 09:08:48 2020

@author: Austin.Schrader
"""

import os
import win32com.client
from openpyxl import Workbook
import csv
import img2pdf 
from PIL import Image 
import fitz

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
                Workbook2 = app.Workbooks.Open(in_file)
                Workbook2.ActiveSheet.ExportAsFixedFormat(0, output_file)
                Workbook2.Close()
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
                Workbook3 = app.Workbooks.Open(in_file)
                Workbook3.ActiveSheet.ExportAsFixedFormat(0, output_file)
                Workbook3.Close()
                app.Quit()
                os.remove(os.path.join(root, f))
                pass
            except:
                print('could not open')
        elif f.endswith(".png"):
            try:
                imglist = []
                imglist.append(os.path.join(root, f))
                output_file = os.path.join(root,f[:-4])
                #print(imglist)
                doc = fitz.open()
                for f in imglist:
                    img = fitz.open(f)
                    rect = img[0].rect
                    pdfbytes = img.convertToPDF()
                    img.close()
                    imgPDF = fitz.open("pdf", pdfbytes)
                    page = doc.newPage(width = rect.width, height = rect.height)
                    page.showPDFpage(rect, imgPDF, 0)
                doc.save(output_file +".pdf")
                os.remove(os.path.join(root, f))
                
                pass
            except:
                pass
        else:
            pass
