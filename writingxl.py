import openpyxl

# Ukitaka kuandika workbook mpya
new_wb = openpyxl.Workbook('New.xlsx')
print(f"Workbook created as: {new_wb}")

# Load workbook to write on
wb = openpyxl.load_workbook('Data.xlsx')
print(f"Workbook loaded as: {wb}")
print("\n")

# Tutengeneze dunia ya worksheet yetu hivi
class Creation:
    
    def __init__(self) -> None:
        self.newworkbook = new_wb
        self.workbook = wb
    
    # Hii ni repesentror tu    
    def __repr__(self) -> str:
        return str(self.workbook)
    
    def create_wb_sheet(self):
        # Wacha tutengeneze sheet yetu
        sheet_creator = self.workbook.create_sheet(title='CodeCreatedSheet')
        print(f"Sheet created as: {sheet_creator}")
        print("\n")
        
        # Verify hapa Kwanza kwa hii instance
        while sheet_creator:
            sheet_name = self.workbook.sheetnames[1]
            print("Sheet exists as: {0}".format(sheet_name))
            print("\n")
            break
        else:
            print("Null")
            
    
    # Lets verify if the sheet was created in another instance
    def verify(self):
        print(self.workbook.sheetnames)
        print("\n")


if __name__ == "__main__":
    Creation().create_wb_sheet()
    Creation().verify()