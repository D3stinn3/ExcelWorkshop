import openpyxl

wb = openpyxl.load_workbook('Data.xlsx')
sheet = wb['Sheet1']

class Iterator:
    
    # Add the variable classes to the Iterator class
    def __init__(self):
        self.workbook = wb
        self.sheet = sheet
        
    def loops(self):
        # Initialize loop here and specify the row range using a two step parameter
        # Also specify the seed cell
        for row_ in range(1, 20, 2):
            print(row_, self.sheet.cell(row=row_, column=7).value)
        print("\n")
        
        # Initialize loop here and specify the column range using a two step parameter
        # Also specify the seed cell  
        for column_ in range(1, 8, 2):
            print(column_, self.sheet.cell(row=19, column=column_).value)
        print("\n")
        
        
        # Lets base the loop off limits
        row_lower_limit = self.sheet.min_row
        row_upper_limit = self.sheet.max_row
        column_lower_limit = self.sheet.min_column
        column_upper_limit = self.sheet.max_column
        
        # Iterate using the row limits
        for row_ in range(row_lower_limit, row_upper_limit):
            print(row_, self.sheet.cell(row=row_, column=7).value)
        print("\n")
        
        # Iterate using the column limits  
        for column_ in range(column_lower_limit, column_upper_limit):
            print(column_, self.sheet.cell(row=19, column=column_).value)
        print("\n")
        

            
            
            
if __name__ == "__main__":
    Iterator().loops()