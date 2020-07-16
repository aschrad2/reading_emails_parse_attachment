# -*- coding: utf-8 -*-
"""
Created on Thu Jul 16 12:50:03 2020

@author: Austin.Schrader
"""

import fitz
import os

imglist = [r'C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\Python_Tools\\reading_emails_parse_attachment\\Attachments\\MI Payment from ACH Report.png']
pdf_path = r'C:\Users\austin.schrader\Desktop\My_Desktop_Documents\Python_Tools\reading_emails_parse_attachment\Attachments\MI Payment from ACH Report.pdf'

for root, dirs, files in os.walk(r'C:\\Users\\austin.schrader\\Desktop\\My_Desktop_Documents\\Python_Tools\\reading_emails_parse_attachment\\Attachments'):
    for f in files:
        if f.endswith(".png"):
            try:
                doc = fitz.open()                            # PDF with the pictures
                for f in imglist:
                    img = fitz.open(f) # open pic as document
                    rect = img[0].rect                       # pic dimension
                    pdfbytes = img.convertToPDF()            # make a PDF stream
                    img.close()                              # no longer needed
                    imgPDF = fitz.open("pdf", pdfbytes)      # open stream as PDF
                    page = doc.newPage(width = rect.width,   # new page with ...
                                       height = rect.height) # pic dimension
                    page.showPDFpage(rect, imgPDF, 0) 
                           # image fills the page
                doc.save(pdf_path)
                os.remove(os.path.join(root, f))
            except:
                pass
