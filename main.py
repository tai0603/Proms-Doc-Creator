import os
import xlrd
import pandas as pd
from docx import Document
from docx.shared import Mm
from docx.shared import Inches

# Load proms data from Excel file
proms = pd.read_excel('PROM.xlsx')
page_num = proms.shape[0]
print(proms)

# Load proms pictures from folder
pic_path = './dresses on models'
pics = os.listdir(pic_path)

# Create Word document
document = Document()

# Change page orientation with standard A4 format
section = document.sections[0]
section.page_height = Mm(210)
section.page_width = Mm(297)

for i in range(0, page_num - 3):
	content = ''

	# Add prom STYLE # and prom colors
	style = int(proms.loc[i].at['STYLE #']) # Get prom STYLE #
	content += str(style) + "\n"

	# Get prom colors
	prom = proms.loc[i]
	prom = prom.drop(labels=['Unnamed: 0', 'DEV #', 'STYLE #', 'FACTORY', 'COLOR REQUEST', 'NOTES', 'PRICE'])
	for j in prom:
		if pd.isna(j) == False:
			content += j + ", "

	content = content[:-1] + '\n\n' # Remove last element in string (',')

	paragraph = document.add_paragraph(content) # Insert prom STYLE # and prom colors into file
	paragraph.alignment = 1 # for left, 1 for center, 2 right, 3 justify ....

	# Add prom pics
	r = paragraph.add_run()

	dev = proms.loc[i].at['DEV #'] # Get prom DEV #
	dev = dev.split('-')[1]

	for j in pics:
		tmp = j.split()
		if tmp[0] == dev:
			r.add_picture('./dresses on models/' + j, width=Inches(2.24), height=Inches(4))

	# Add blank page
	document.add_page_break()

	print("Row " + str(i + 1) + " Done..")

document.save('demo.docx')
print("Document Created...")
