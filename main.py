"""
Get a document, from source directory, with all paycheck of every workers and split it in many PDF one for every worker, the pdf can has one or more pages

# The get_text("word") function of the library PyMuPDF extracts the words as show below
# Example of items that contain the words for month and year
(341.81988525390625, 109.61735534667969, 386.1297912597656, 122.79430389404297, 'Ottobre', 22, 7, 3)
(405.1197509765625, 109.61735534667969, 430.439697265625, 122.79430389404297, '2021', 22, 7, 4)
(449.4296569824219, 109.61735534667969, 468.41961669921875, 122.79430389404297, 'Del', 22, 7, 5)

# Example of items that contain the words for first and last names
(63.300010681152344, 150.3373260498047, 88.6200180053711, 163.51426696777344, 'ROSSI', 22, 11, 1)
(94.95001983642578, 150.3373260498047, 132.93002319335938, 163.51426696777344, 'MARIO', 22, 11, 2)
(316.49993896484375, 150.3373260498047, 417.77972412109375, 163.51426696777344, 'RSSMRA14D79R145S', 22, 11, 3)

(63.300010681152344, 150.3373260498047, 88.6200180053711, 163.51426696777344, 'ROSSI', 22, 11, 1)
(94.95001983642578, 150.3373260498047, 132.93002319335938, 163.51426696777344, 'DAVIDE', 22, 11, 2)
(139.26002502441406, 150.3373260498047, 189.90003967285156, 163.51426696777344, 'LUCA', 22, 11, 3)
(316.49993896484375, 150.3373260498047, 417.77972412109375, 163.51426696777344, 'RSSMRL14D79R145S', 22, 11, 4)
"""

import fitz
import sys
from os import path, makedirs

source_filename = 'Source/CEDOLINI.pdf' # PDF with all paycheck

# Check if source file exists or abort
if(not path.exists(source_filename)):
  sys.exit("[ERROR] - File is not found !")

# Points position for names and sourname
name_r1 = 150.3373260498047   # That is the upper position of row's [1] == r1
name_r2 = 163.51426696777344  # That is the bottom position of row's [3] == r2
name_c1 = 63.300010681152344  # That is the start position of column's [0] >= c1
name_c2 = 316.49993896484375  # That is the end position of column's[0] < c2
# Points position for month and year
data_r1 = 109.61735534667969  # That is the upper position of row's
data_r2 = 122.79430389404297  # That is the bottom position of row's
data_c1 = 341.81988525390625  # That is the start position of column's
data_c2 = 405.1197509765625   # That is the middle position of column's
data_c3 = 449.4296569824219   # That is the end position of column's

# Translation dictionary from month to number 
month_to_int = {
  "Gennaio": "1",
  "Febbraio": "2",
  "Marzo": "3",
  "Aprile": "4",
  "Maggio": "5",
  "Giugno": "6",
  "Luglio": "7",
  "Agosto": "8",
  "Settembre": "9",
  "Ottobre": "10",
  "Novembre": "11",
  "Dicembre": "12"}

doc = fitz.open(source_filename) # Open file cedolini with PyMuPDF

# Iterate all pages in PDF
for page in doc.pages(0, 20, 1): # If you want get only a range use doc.pages(START, STOP, STEP)
  words = page.get_text("words")  # Create a list of items that contain the word of a signle page
  # Variables for name, month, year and output filename
  name = ''
  year = ''
  month = ''
  output_filename = ''
  
  # If the page is empty, skip it
  if(len(words) <= 1):
    continue

  # Iterate all words of the current page to get first and last names, month and year 
  for i in words:
    # Get first and last names
    if((i[1] == name_r1 and i[3] == name_r2) and (i[0] >= name_c1 and i[0] < name_c2)):
      name += i[4].replace("'","")
    
    # Get month and year
    if((i[1] == data_r1 and i[3] == data_r2) and (i[0] >= data_c1 and i[0] < data_c2)):
      month = i[4]
    if((i[1] == data_r1 and i[3] == data_r2) and (i[0] >= data_c2 and i[0] < data_c3)):
      year = i[4]
  
  # Combine the name, month and year to create the destination folder and output filename
  destination_folder = 'Destination/' + name + '/'
  filename = name + '_' + month_to_int[month] + '_' + year + '.pdf'
  output_filename = destination_folder + filename
  
  # If destination folder don't exists make it
  if(not path.exists(destination_folder)):
    makedirs(destination_folder)
  
  # If file already exists merge it with current page
  if(path.exists(output_filename)):
    output_doc = fitz.open(output_filename) # Open a file called output_filename
    output_doc.insert_pdf(doc, from_page=page.number, to_page=page.number) # Append current page of main document
    output_doc.save(output_filename, incremental=True, encryption=0) # Save file
  else:
    output_doc = fitz.open() # Open an empty file
    output_doc.insert_pdf(doc, from_page=page.number, to_page=page.number) # Append current page of main document
    output_doc.save(output_filename) # Save file

  output_doc.close() # Close the output file

doc.close() # Close the main file