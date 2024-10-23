import os, sys
from docxtpl import DocxTemplate

os.chdir(sys.path[0])

doc = DocxTemplate('temp.docx')
context = {'name': 'Ben'}

for x in range(1, 6):
    gg = str(x)
    context['name'] = 'Ben'+gg
    doc.render(context)
    doc.save('temp'+gg+'.docx')
    
#Run in command line using: Python docGen.py
#Will create the documents in same file
