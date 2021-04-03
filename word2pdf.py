import sys
import os
import comtypes.client

wdFormatPDF = 17
location = os.path.dirname(os.path.realpath(__file__))
print(location)
word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(location + "\\resume.docx")
doc.SaveAs(location + "\\resume.pdf", FileFormat=wdFormatPDF)
doc.Close()
word.Quit()