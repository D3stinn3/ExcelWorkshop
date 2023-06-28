import openpyxl

# Obtain a single sheet from Workbook
wb = openpyxl.load_workbook('Data.xlsx')
sheet = wb['Sheet1']

# Get the cell value from sheet
cell_A1 = sheet['A1']
A1_value = cell_A1.value
print(A1_value)

cell_B2 = sheet['B2']
B2_value = cell_B2.value
print(B2_value)