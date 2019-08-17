import xlrd
from time import sleep
from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Font
from openpyxl.styles import Alignment


book1_first_row = []
book1_data = []

"""
This function creates a dictionary with the first row values being keys and everything else in the column being paired with the key as a value
"""

def parse_data_from_book1():
    """PARSES DATA FROM EXCEL TO BE USED AND CONVERTED"""
    workbook = xlrd.open_workbook('Path to excel spreadsheet',
                                  on_demand=True)
    worksheet = workbook.sheet_by_index(0)  # THE FIRST WORKSHEET IN THE SPREADSHEET

    # THIS FIRST PORTION TAKES COLUMN NAMES THAT WILL BE THE DICTIONARY KEYS
    for col in range(worksheet.ncols):
        #  THE ZERO BELOW REFLECTS THE ROW WITH THE COLUMN HEADS AND MUST BE ACCURATE TO GET PROPER DICTIONARY PAIRINGS
        book1_first_row.append(worksheet.cell_value(0, col))

    # THIS AREA PAIRS DATA FROM EACH COLUMN WITH ITS COLUMN LABEL IN THE DICTIONARY EX. NAME:CHRIS
    for row in range(1, worksheet.nrows):
        dictionary_values = {}
        for col in range(worksheet.ncols):  # ITERATES OVER THE COLUMNS
            dictionary_values[book1_first_row[col]] = worksheet.cell_value(row, col)
            # PAIRS THE COLUMN HEADS WITH THE ROW VALUE
        book1_data.append(dictionary_values)