import openpyxl

# opening the workbook using the openpyxl module
wb = openpyxl.load_workbook('Data.xlsx')

# identifying the workbook
print(type(wb))

# identifying sheetnames embedded in the workbook
sheetnames_in_wb = wb.sheetnames
print(sheetnames_in_wb)