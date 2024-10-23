import tkinter as tk

root = tk.Tk()

canvas = tk.Canvas(root, width=300, height=300)
canvas.pack()

import os, sys
from docxtpl import DocxTemplate

os.chdir(sys.path[0])

doc = DocxTemplate('temp.docx')
context = {'name': 'Ben'}

numOfDoc = 6
numOfDoc+=1

def docgen():
    for x in range(1, numOfDoc):
        gg = str(x)
        context['name'] = 'Ben'+gg
        doc.render(context)
        doc.save('temp'+gg+'.docx')


button = tk.Button(text="Generate Documents", command=docgen, bg="brown", fg="white")
canvas.create_window(150, 150, window=button)

root.mainloop()

#interactive version - not completed, I want to make an input section