print("Hello")

import pyodbc

list = pyodbc.drivers()

print(list)

# "i-model ODBC Driver for Windows"

# Replace with your actual path to the Access file
# db_path = r"D:\OneDrive - MPS\BMC\OPPlant\Test.accdb"
#
# # Connect to the database
# conn = pyodbc.connect(
#     r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % db_path)
#
# # Create a cursor object
# cursor = conn.cursor()
#
# # Get Linked Table Name from Access file.
# cursor.execute("SELECT * FROM _LinkedTableList")
#
# linkedTable = []
# for row in cursor.fetchall():
#     linkedTable.append(row[0])
#
#
# cursor.execute("Select * from OpenPlant_3D_01_08_Ball_Valve")
# for row in cursor.fetchall():
#     print(row)
#
#
# # Close connection
# conn.close()
#
# # print(linkedTable)
