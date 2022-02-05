# Created by Luca Canali (c)2022

import fitz
from os import path

# The words is between this points
# Items for the month and year
#(341.81988525390625, 109.61735534667969, 386.1297912597656, 122.79430389404297, 'Ottobre', 22, 7, 3)
#(405.1197509765625, 109.61735534667969, 430.439697265625, 122.79430389404297, '2021', 22, 7, 4)
#(449.4296569824219, 109.61735534667969, 468.41961669921875, 122.79430389404297, 'Del', 22, 7, 5)

# Items for the names and sourname
#(63.300010681152344, 150.3373260498047, 88.6200180053711, 163.51426696777344, 'ROSSI', 22, 11, 1)
#(94.95001983642578, 150.3373260498047, 132.93002319335938, 163.51426696777344, 'MARIO', 22, 11, 2)
#(316.49993896484375, 150.3373260498047, 417.77972412109375, 163.51426696777344, 'RSSMRA14D79R145S', 22, 11, 3)

#(63.300010681152344, 150.3373260498047, 88.6200180053711, 163.51426696777344, 'ROSSI', 22, 11, 1)
#(94.95001983642578, 150.3373260498047, 132.93002319335938, 163.51426696777344, 'DAVIDE', 22, 11, 2)
#(139.26002502441406, 150.3373260498047, 189.90003967285156, 163.51426696777344, 'LUCA', 22, 11, 3)
#(316.49993896484375, 150.3373260498047, 417.77972412109375, 163.51426696777344, 'RSSMRL14D79R145S', 22, 11, 4)

# Points for names and sourname
name_r1 = 150.3373260498047  # [1] == r1
name_r2 = 163.51426696777344 # [3] == r2
name_c1 = 63.300010681152344 # [0] >= c1
name_c2 = 316.49993896484375 # [0] < c2
# Point for month and date
data_r1 = 109.61735534667969
data_r2 = 122.79430389404297
data_c1 = 341.81988525390625  # column start month
data_c2 = 405.1197509765625   # column end month and start year
data_c3 = 449.4296569824219   # column end year

cedolini = 'CEDOLINI.pdf' # PDF with all paycheck

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

doc = fitz.open(cedolini) # Open file cedolini with PyMuPDF

# Iterate all pages in PDF
for page in doc: # If you want get only a range use doc.pages(START, STOP, STEP)
  words = page.get_text("words")  # Make a list of words, from the document, with points that delimit it
  # Variables for name, month, year and filename of output
  name = ''
  year = ''
  month = ''
  output_filename = ''
  
  # If the page is empty, skip it
  if(len(words) <= 1):
    continue

  # Iterate all pages of document
  for i in words:
    # Get names and sourname
    if((i[1] == name_r1 and i[3] == name_r2) and (i[0] >= name_c1 and i[0] < name_c2)):
      name += i[4].replace("'","")
    
    # Get month and year
    if((i[1] == data_r1 and i[3] == data_r2) and (i[0] >= data_c1 and i[0] < data_c2)):
      month = i[4]
    if((i[1] == data_r1 and i[3] == data_r2) and (i[0] >= data_c2 and i[0] < data_c3)):
      year = i[4]
  
  # Make a output filename with name, month, year and extension
  output_filename = name + '_' + month_to_int[month] + '_' + year + '.pdf'
  
  #TODO: If file already exists merge it with new page
  if(path.exists(output_filename)):
    # Open a file called output_filename
    output_doc = fitz.open(output_filename)
    output_doc.insert_pdf(doc, from_page=page.number, to_page=page.number)
    output_doc.save(output_filename, incremental=True, encryption=0)
  else:
    # Open an empty file
    output_doc = fitz.open()
    output_doc.insert_pdf(doc, from_page=page.number, to_page=page.number)
    output_doc.save(output_filename)
  
  # Close the output file
  output_doc.close()

# Close the main file
doc.close()