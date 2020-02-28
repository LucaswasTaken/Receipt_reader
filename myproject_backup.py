import xlwings as xw
# Import libraries 
from PIL import Image 
import pytesseract 
import sys 
from pdf2image import convert_from_path 
import os 
import glob
from difflib import SequenceMatcher
from tika import parser

#Function to read all pdf files in folder
def readfiles():
	os.chdir("./")
	pdfs = []
	for file in glob.glob("*.pdf"):
		print(file)
		pdfs.append(file)
	return pdfs

#Find Similarity between strings
def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

#Find pdf info
def compare_info(read_name,cnpj,money,bill,i):
	file = open(read_name)
	for line in file:
		line_split = line.split()
		if(len(line_split)>0):
			for j in range(0,len(line_split)):
				cnpj_cell = "B" + i
				money_cell = "C" + i
				bill_cell = "D" + i

				#Compare CNPJ
				if((xw.Range(cnpj_cell).color ==None)&(len(line_split[j])>0.8*len(cnpj))&(len(line_split[j])<1.5*len(cnpj))):
					cnpj_n_compare = line_split[j].replace('.','')
					cnpj_n_compare = cnpj_n_compare.replace('/','')
					cnpj_n_compare = cnpj_n_compare.replace('-','')
					if(similar(cnpj_n_compare,cnpj)>0.8):
						xw.Range(cnpj_cell).color = (0, 255, 0)

					elif((similar(cnpj_n_compare,cnpj)>0.6)&(xw.Range(cnpj_cell).color ==None)):
						xw.Range(cnpj_cell).color = (100, 255, 0)

				#Compare bill number
				if((xw.Range(bill_cell).color ==None)&(len(line_split[j])>0.8*len(bill))):
					n_compare = line_split[j].replace('.','')
					if(RepresentsInt(n_compare)):
						n_compare = int(n_compare)
						n_compare = str(n_compare)
						bill = str(bill)
						if(similar(n_compare,bill)>0.8):
							xw.Range(bill_cell).color = (0, 255, 0)
						elif((similar(n_compare,bill)>0.6)&(xw.Range(bill_cell).color ==None)):
							xw.Range(bill_cell).color = (100, 255, 0)

				#Compare money cell
				if((xw.Range(bill_cell).color ==None)&(len(line_split[j])>0.8*len(money))&(len(line_split[j])<1.5*len(money))):
					money_compare = line_split[j].replace('.','')
					money = str(money)
					money.replace('.',',')
					if(similar(money_compare,money)>0.8):
						xw.Range(bill_cell).color = (0, 255, 0)
					elif((similar(money_compare,money)>0.6)&(xw.Range(money_cell).color ==None)):
						xw.Range(money_cell).color = (100, 255, 0)
				
				#Stop if finding all
				if((xw.Range(money_cell).color !=None)&(xw.Range(bill_cell).color !=None)&(xw.Range(cnpj_cell).color !=None)):
					break
#Rotate Image
def rotate_90_image():
	os.chdir(path)
	jpgs = []
	for file in glob.glob("*.jpg"):
		print(file)
		jpgs.append(file)
	for name in jpgs:
		colorImage  = Image.open(name)
		rotated     = colorImage.rotate(90)
		rotated.save(name)

#Check if represents a int
def RepresentsInt(s):
	try: 
		int(s)
		return True
	except ValueError:
		return False

#Read PDF text
def read_to_text(pdf_name,complete_adress):
	text=str(parser.from_file(pdf_name))
	outfile = os.path.splitext(pdf_name)[0] + "_read.txt"
	f = open(outfile, "a") 
	text = text.replace('\\n', ' ')
	f.write(text)
	f.close()

#Transform into image
def transform_image(pdf_name,complete_adress):
	# Path of the pdf 
	PDF_file = pdf_name

	''' 
	Part #1 : Converting PDF to images 
	'''

	# Store all the pages of the PDF in a variable 
	pages = convert_from_path(PDF_file, 500) 

	# Counter to store images of each page of PDF to image 
	image_counter = 1

	# Iterate through all the pages stored above 
	for page in pages: 

		# Declaring filename for each page of PDF as JPG 
		# For each page, filename will be: 
		# PDF page 1 -> page_1.jpg 
		# PDF page 2 -> page_2.jpg 
		# PDF page 3 -> page_3.jpg 
		# .... 
		# PDF page n -> page_n.jpg 
		filename = complete_adress+"page_"+str(image_counter)+".jpg"
		
		# Save the image of the page in system 
		page.save(filename, 'JPEG') 

		# Increment the counter to update filename 
		image_counter = image_counter + 1
	return image_counter

#Convert PDF to text
def convert_to_text(pdf_name,complete_adress,image_counter):
	''' 
	Part #2 - Recognizing text from the images using OCR 
	'''
	# Path of the pdf 
	PDF_file = pdf_name
	# Variable to get count of total number of pages 
	filelimit = image_counter-1

	# Creating a text file to write the output 
	outfile = os.path.splitext(PDF_file)[0] + "_read.txt"

	# Open the file in append mode so that 
	# All contents of all images are added to the same file 
	f = open(outfile, "a") 

	# Iterate from 1 to total number of pages 
	for i in range(1, filelimit + 1): 

		# Set filename to recognize text from 
		# Again, these files will be: 
		# page_1.jpg 
		# page_2.jpg 
		# .... 
		# page_n.jpg 
		filename = complete_adress+"page_"+str(i)+".jpg"
			
		# Recognize the text as string in image using pytesserct 
		text = str(((pytesseract.image_to_string(Image.open(filename))))) 

		# The recognized text is stored in variable text 
		# Any string processing may be applied on text 
		# Here, basic formatting has been done: 
		# In many PDFs, at line ending, if a word can't 
		# be written fully, a 'hyphen' is added. 
		# The rest of the word is written in the next line 
		# Eg: This is a sample text this word here GeeksF- 
		# orGeeks is half on first line, remaining on next. 
		# To remove this, we replace every '-\n' to ''. 
		text = text.replace('-\n', '')	 

		# Finally, write the processed text to the file. 
		f.write(text) 

	# Close the file after writing all the text. 
	f.close() 


#Reading Worksheet
def hello_xlwings():
	#Conecting to worksheet
    wb = xw.Book.caller()
    i = wb.sheets[0].range("H2").value
    i = int(i) + 1

    #Locating cells and adresses
    file_name_cell = "E" + str(i)
    cnpj_cell = "B" + str(i)
    money_cell = "C" + str(i)
    bill_cell = "D" + str(i)
    pdf_name = wb.sheets[0].range(file_name_cell).value
    complete_adress = os.path.abspath(__file__)
    complete_adress = complete_adress.replace("myproject.py","")
    pdf_name = complete_adress +pdf_name

    #Geting cell values
    cnpj = wb.sheets[0].range(cnpj_cell).value
    money = wb.sheets[0].range(money_cell).value
    bill = wb.sheets[0].range(bill_cell).value

    #Rotation atempts
    rotate_try = 0

    while(pdf_name):
    	#Trying to read in text mode
    	wb.sheets[0].range("I2").value = "OK"
    	read_to_text(pdf_name,complete_adress)
    	wb.sheets[0].range("I3").value = "OK"
    	read_name  = os.path.splitext(pdf_name)[0] + "_read.txt"
    	information = compare_info(read_name,cnpj,money,bill,str(i))
    	wb.sheets[0].range("I4").value = "OK"
    	#Reading from images
    	if(xw.Range(cnpj_cell).color == None):
    		image_counter = transform_image(pdf_name,complete_adress)
    		convert_to_text(pdf_name,complete_adress, image_counter)
    		wb.sheets[0].range("I5").value = "OK"
    		read_name = s.path.splitext(pdf_name)[0] + "_read.txt"
    		information = compare_info(read_name,cnpj,money,bill,str(i))
    		while(xw.Range(cnpj_cell).color == None):
    			wb.sheets[0].range("I6").value = "OK"
    			rotate_90_image()
    			convert_to_text(pdf_name,complete_adress, image_counter)
    			read_name = s.path.splitext(pdf_name)[0] + "_read.txt"
    			information = compare_info(read_name,cnpj,money,bill,str(i))
    			rotate_try = rotate_try + 1
    			if(rotate_try==3):
    				rotate_try = 0
    				break
    	#Updating cell
    	i = i+1
    	wb.sheets[0].range("H2").value = i
    	file_name_cell = "E" + str(i)
    	cnpj_cell = "B" + str(i)
    	money_cell = "C" + str(i)
    	bill_cell = "D" + str(i)
    	pdf_name = wb.sheets[0].range(file_name_cell).value
    	if(pdf_name==""):
    		break
    	pdf_name = complete_adress +pdf_name
    	cnpj = wb.sheets[0].range(cnpj_cell).value
    	money = wb.sheets[0].range(money_cell).value
    	bill = wb.sheets[0].range(bill_cell).value

