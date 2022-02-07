"""
Get a document, from source directory, with all paycheck of every workers and split it in many PDF one for every worker, the pdf can has one or more pages

"""

import fitz
import sys
from os import path, makedirs

# Dictionary for translate month to number 
month_to_int = {
  "gennaio": "01",
  "febbraio": "02",
  "marzo": "03",
  "aprile": "04",
  "maggio": "05",
  "giugno": "06",
  "luglio": "07",
  "agosto": "08",
  "settembre": "09",
  "ottobre": "10",
  "novembre": "11",
  "dicembre": "12"
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

#### TODO: Make a generic function for different documents, for my case cedolini and staced
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
  for page in doc.pages(0, doc.page_count): # Range of pages doc.pages(START, STOP, STEP) doc.page_count-1 
    words_of_page = page.get_text("words")  # Create a list of items that contain the words of the current page

    # If the page is empty, skip it
    if(len(words_of_page) <= 1):
      continue

    # Iterate all words of the current page to get first and last names, month and year 
    for i in words_of_page:
      name = get_words(words_of_page, name_row, name_p1, name_p2) # Name of worker
      month = month_to_int[get_words(words_of_page, data_row, data_p1, data_p2).lower()] # Month of the year
      year = get_words(words_of_page, data_row, data_p2, data_p3) # Year of the paycheck
    
    # Combine the name, month and year to create the destination folder and output filename
    destination_folder = 'Destination/' + name + '/'
    output_filename = destination_folder + name + '_' + month + '_' + year + '.pdf'
    
    # If destination folder doesn't exists make it
    if(not path.exists(destination_folder)):
      makedirs(destination_folder)
    save_file(output_filename, doc, page.number) # Save or append current page in the pdf

  doc.close() # Close the main file

def make_staced():
  """
  (272.998046875, 52.28958511352539, 301.7962951660156, 63.390953063964844, 'ZOLI', 0, 3, 3)
  (308.99798583984375, 52.28958511352539, 352.1961975097656, 63.390953063964844, 'ANDREA', 0, 3, 4)
  (560.9973754882812, 76.28958129882812, 582.5955200195312, 87.39095306396484, ',00', 0, 5, 9)

  (272.998046875, 52.28958511352539, 323.396240234375, 63.390953063964844, "VERITA'", 0, 3, 3)
  (330.59796142578125, 52.28958511352539, 380.99615478515625, 63.390953063964844, 'MASSIMO', 0, 3, 4)
  (560.9973754882812, 76.28958129882812, 582.5955200195312, 87.39095306396484, ',00', 0, 5, 9)

  (258.59808349609375, 40.28958511352539, 301.7962951660156, 51.390953063964844, 'MAGGIO', 0, 2, 0)
  (323.3979797363281, 40.28958511352539, 352.1961975097656, 51.390953063964844, '2017', 0, 2, 1)
  (28.19898796081543, 52.28958511352539, 56.997161865234375, 63.390953063964844, 'Cod.', 0, 3, 0)
  """
  name_row = 63.390953063964844
  name_p1 = 272.998046875
  name_p2 = 582.5955200195312
  data_row = 51.390953063964844
  data_p1 = 20.99901580810547 # Page's margin left
  data_p2 = 323.3979797363281 # Middle poit
  data_p3 = 582.5955200195312 # Pages's margin right
  source_filename = "Source/STACED.pdf"

  # Check if the source file exists and open or abort programm
  if(not path.exists(source_filename)):
    sys.exit("[ERROR] - File is not found !")
  doc = fitz.open(source_filename) # Open file

  # Iterate all pages in PDF
  for page in doc.pages(0, doc.page_count): # Range of pages doc.pages(START, STOP, STEP) doc.page_count-1 
    words_of_page = page.get_text("words")  # Create a list of items that contain the words of the current page

    # If the page is empty, skip it
    if(len(words_of_page) <= 1):
      continue

    # Iterate all words of the current page to get first and last names, month and year 
    for i in words_of_page:
      name = get_words(words_of_page, name_row, name_p1, name_p2) # Name of worker
      month = month_to_int[get_words(words_of_page, data_row, data_p1, data_p2).lower()] # Month of the year
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
  make_staced()
  print("finish!")