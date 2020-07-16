# -*- coding: utf-8 -*-
"""
Created on Tue Jul 14 09:06:19 2020

@author: Austin.Schrader
"""

from win32com import client
xlApp = client.Dispatch("Excel.Application")
books = xlApp.Workbooks.Open('C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\Python_Tools\\reading_emails_parse_attachment\\Attachments\\Daily Tax Check OLD.xlsm')
ws = books.Worksheets[0]
ws.Visible = 1
ws.ExportAsFixedFormat(0, 'C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\Python_Tools\\reading_emails_parse_attachment\\Attachments\\Daily Tax Check OLD.pdf')