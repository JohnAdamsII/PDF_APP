#Converts docx files to pdf

import sys
import os
import comtypes.client
import glob2,re

def Docx_to_PDF(filename):
    wdFormatPDF = 17
    newstr = filename[0:-5]
    in_file = os.path.abspath(filename)
    out_file = os.path.abspath(newstr)
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

if __name__ == '__main__':
    pass
