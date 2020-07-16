# -*- coding: utf-8 -*-
"""
Created on Thu Jul 16 10:47:20 2020

@author: Austin.Schrader
"""

import pypandoc

input_file = r'C:\Users\austin.schrader\Desktop\My_Desktop_Documents\Python_Tools\reading_emails_parse_attachment\Attachments\Billing_Report_2020-07-10.csv'
output_file = r'C:\Users\austin.schrader\Desktop\My_Desktop_Documents\Python_Tools\reading_emails_parse_attachment\Attachments\Billing_Report_2020-07-10.pdf'


# =============================================================================
# pandoc billing.csv -o billing.pdf
# =============================================================================
# =============================================================================
# 
# output = pypandoc.convert_file(input_file, 'pdf', outputfile=output_file)
# assert output == ""
# =============================================================================

from tabula import convert_into
# =============================================================================
# 
# convert_into(input_file, output_file, output_format="pdf")
# =============================================================================


import tabula

df = tabula.read_pdf(input_file, encoding='utf-8', pages='1-6041')