"""
Get a document, from source directory, with all paycheck of every workers and split it in many PDF one for every worker, the pdf can has one or more pages

"""
### LIBRARY
import fitz
import sys
import time
from os import path, makedirs, getpid
from multiprocessing import Process


### VARIABLES AND CONSTATNS
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


### FUNCTION
def get_words(words, row, start_point, end_point):
  string = ''
  
  # Iterate all words and return the word specified by row, start and end points 
  for i in words:
    # Get first and last names
    if(i[3] == row and (i[0] >= start_point and i[0] < end_point)):
      string += i[4].replace("'","")
  
  return string


# Create or merge a pdf with a desired page numeber from the original doc
def save_file(file, doc, page_number):
  # If file exsists than it append current page
  if(path.exists(file)):
    with fitz.open(file) as output_doc: # Open a file
      output_doc.insert_pdf(doc, from_page=page_number, to_page=page_number) # Append current page of main document
      output_doc.save(file, incremental=True, encryption=0) # Save file
  # otherwise create a new file with current page
  else:
    with fitz.open() as output_doc: # Open an empty file
      output_doc.insert_pdf(doc, from_page=page_number, to_page=page_number) # Append current page of main document
      output_doc.save(file) # Save file


def make_document(source_filename, name_row, name_p1, name_p2, data_row, data_p1, data_p2, data_p3):
  print('Start process id:', getpid())
  start_time = time.time()

  # Check if the source file exists and open or abort programm
  if(not path.exists(source_filename)):
    sys.exit("[ERROR] - File is not found !")
  
  with fitz.open(source_filename) as doc: # Open file

    # Iterate all pages in PDF
    for page in doc.pages(0, doc.page_count): # Range of pages doc.pages(START, STOP, STEP) | doc.page_count 
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
  print('Finish process id:', getpid())
  print("--- %s seconds ---" % (time.time() - start_time))


### MAIN
if __name__ == "__main__":

  # CEDOLINI
  source_filename = 'Source/CEDOLINI_SOCI.pdf'
  name_row = 163.51426696777344 # line_no 11
  name_p1 = 63.300010681152344
  name_p2 = 316.49993896484375
  data_row = 122.79430389404297 # line no 7  # That is the bottom position of row's
  data_p1 = 341.81988525390625
  data_p2 = 405.1197509765625
  data_p3 = 449.4296569824219
  cedolini = Process(target=make_document, args=(source_filename, name_row, name_p1, name_p2, data_row, data_p1, data_p2, data_p3,))
  #make_document(source_filename, name_row, name_p1, name_p2, data_row, data_p1, data_p2, data_p3)

  # STACED
  source_filename = 'Source/STACED_SOCI.pdf'
  name_row = 63.390953063964844 # line_no 3
  name_p1 = 272.998046875
  name_p2 = 582.5955200195312
  data_row = 51.390953063964844 # line_no 2
  data_p1 = 20.99901580810547
  data_p2 = 323.3979797363281
  data_p3 = 582.5955200195312
  staced = Process(target=make_document, args=(source_filename, name_row, name_p1, name_p2, data_row, data_p1, data_p2, data_p3,))
  
  cedolini.start()
  staced.start()

  cedolini.join()
  staced.join()
  #make_document(source_filename, name_row, name_p1, name_p2, data_row, data_p1, data_p2, data_p3)
  print("finish!")