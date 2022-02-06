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

# Dictionary for translate month to number 
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
  "Dicembre": "12"
  }

def get_words(words, row, start_point, end_point):
  string = ''
  
  # Iterate all words and return the word specified by row, start and end points 
  for i in words:
    # Get first and last names
    if(i[3] == row and (i[0] >= start_point and i[0] < end_point)):
      string += i[4].replace("'","")
  
  return string

def save_file(file, doc, page_number):
  # If file exsists than it append current page
  if(path.exists(file)):
    output_doc = fitz.open(file) # Open a file
    output_doc.insert_pdf(doc, from_page=page_number, to_page=page_number) # Append current page of main document
    output_doc.save(file, incremental=True, encryption=0) # Save file
  # otherwise create a new file with current page
  else:
    output_doc = fitz.open() # Open an empty file
    output_doc.insert_pdf(doc, from_page=page_number, to_page=page_number) # Append current page of main document
    output_doc.save(file) # Save file
  
  output_doc.close() # Close the file

# TODO: Make a generic function for different documents, for my case cedolini and staced
def make_cedolini():
  # This row and poit is pecific for my document
  # Points position for names and sourname
  name_row = 163.51426696777344  # That is the bottom position of row's [3] == r2
  name_p1 = 63.300010681152344  # That is the start position of column's [0] >= c1
  name_p2 = 316.49993896484375  # That is the end position of column's[0] < c2
  # Points position for month and year
  data_row = 122.79430389404297  # That is the bottom position of row's
  data_p1 = 341.81988525390625  # That is the start position of column's
  data_p2 = 405.1197509765625   # That is the middle position of column's
  data_p3 = 449.4296569824219   # That is the end position of column's
  
  source_filename = 'Source/CEDOLINI.pdf' # PDF with all paycheck

  # Check if the source file exists and open or abort programm
  if(not path.exists(source_filename)):
    sys.exit("[ERROR] - File is not found !")
  doc = fitz.open(source_filename) # Open file

  # Iterate all pages in PDF
  for page in doc.pages(0, doc.page_count-1, 1): # Range of pages doc.pages(START, STOP, STEP)
    words_of_page = page.get_text("words")  # Create a list of items that contain the words of the current page

    # If the page is empty, skip it
    if(len(words_of_page) <= 1):
      continue

    # Iterate all words of the current page to get first and last names, month and year 
    for i in words_of_page:
      name = get_words(words_of_page, name_row, name_p1, name_p2) # Name of worker
      month = month_to_int[get_words(words_of_page, data_row, data_p1, data_p2)] # Month of the year
      year = get_words(words_of_page, data_row, data_p2, data_p3) # Year of the paycheck
    
    # Combine the name, month and year to create the destination folder and output filename
    destination_folder = 'Destination/' + name + '/'
    output_filename = destination_folder + name + '_' + month + '_' + year + '.pdf'
    
    # If destination folder doesn't exists make it
    if(not path.exists(destination_folder)):
      makedirs(destination_folder)
    save_file(output_filename, doc, page.number) # Save or append current page in the pdf

  doc.close() # Close the main file

if __name__ == "__main__":
  make_cedolini()