import pyodbc
from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Font

data = []  # HOLDS LISTS INSTEAD OF pyobdc.rows DATA TYPE.

#  BELOW SPECIFIES LOCATION OF TARGET DB, DRIVER, AND PASSWORD IF NECESSARY
data_target = 'path to .mdb'; DRV = '{Microsoft Access Driver (*.mdb)}'; PWD = ''

# LIKE SQLLITE THIS CONNECTS UP TO THE DATABASE
conn = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV, data_target, PWD))

cur = conn.cursor()

for row in cur.tables():  # QUERIES ALL TABLE NAMES
    print('AVAILABLE TABLE NAME: ' + str(row.table_name))

for row in cur.columns(table='Target table'):  # GETS COLUMN NAMES FROM SPECIFIC TABLE
    print('COLUMN HEADER: ' + str(row.column_name))

sql = 'SELECT * FROM Target table'  # SELECTS ALL DATA FROM THIS TABLE

rows = cur.execute(sql).fetchall()
for item in rows:
    data.append(list(item))  # CONVERTS pyodbc.rows FORMAT INTO A LIST AND SENDS IT TO THE DATA LIST

cur.close()
conn.close()