'''
All Rights Reversed.
2020/12/09
My First Python Script
Developed With LOVE By Ashkan Ebtekari
'''

'''
Install pypiwin32
pip install pypiwin32
'''

# imports

import openpyxl
from openpyxl.styles import Alignment
import time

# Main Local Variables

Introduction = "\n \n \n Welcome To ExcelGuy Script " \
               "\n You Can Easily Handle Your Excel WorkSheets " \
               "\n in Diffrent Cool Ways ... \n \n \n"

Menu = "1- Excel Merger " \
       "\n 2- Excel Scraper" \
       "\n 3- Excel Converter" \
       "\n \n \n"

# Print Instructions

print(Introduction)
print(Menu)

# User Input

UserInput = input("Select From Menu: [1 TO 3] :  ")

# Menu Checker

Menu_Merger = UserInput == "1"
Menu_Scraper = UserInput == "2"
Menu_Convertor = UserInput == "3"


# Sleep Method

def Sleeper():
    time.sleep(1.7)


# Statements

if Menu_Merger:

    # Introduction
    print("Selected Excel Merger")
    Sleeper()
    Unmerged_File_Input = input("Unmerged File Name : ")
    WorkBook = openpyxl.load_workbook(Unmerged_File_Input)
    Sleeper()
    Merge_File_Input = input("Merge File Name : ")
    Sheet = WorkBook[Merge_File_Input]
    Data = Sheet['B4'].value
    Sheet.merge_cells('B4:E4')
    Sheet['B4'] = Data
    Sheet['B4'].alignment = Alignment.horizontal
    WorkBook.save(Unmerged_File_Input)
    print("Operation Done Successfully ...")
    Sleeper()
    exit()





# Excel Scraper

elif Menu_Scraper:

    # Instruction
    print("Selected Excel Scraper")
    Sleeper()



# Excel Converter

elif Menu_Convertor:

    # Instruction
    print("Selected Excel Converter")
    Sleeper()

    # Excel File Collector
    Excel_Source_Input = input("Give Me The Excel Source : ")
    Excel_Source = Excel_Source_Input
    Sleeper()
    Excel_Name_Input = input("Excel Name : ")
    Excel_Name = Excel_Name_Input
    Sleeper()
    Excel_Final_Output = r'{}{}'.format(Excel_Source, Excel_Name)
    print("Excel Source Collected Successfully ... \n \n")
    Sleeper()

    # PDF File Collector
    PDF_Source_Input = input("PDF OutPut File Directory : ")
    PDF_Source = PDF_Source_Input
    Sleeper()
    PDF_Name_Input = input("PDF Name : ")
    PDF_Name = PDF_Name_Input
    Sleeper()
    PDF_Final_Output = r'{}{}'.format(PDF_Source, PDF_Name)
    print("PDF Info Collected Successfully ... \n \n")
    Sleeper()

    # Excel Application Warmups
    Application = client.DispatchEx("Excel.Application")
    Application.Interactive = False
    Application.Visible = False
    Student = Application.Workbooks.Open(Excel_Final_Output)

    try:

        Student.ActiveSheet.ExportAsFixedFormat(0, PDF_Final_Output)

    except Exception as exc:

        print("Convertation Process Failed")
        print(exc)

    finally:

        Student.Close()
        Application.Exit()
        print("PDF FILE GENERATED SUCCESSFULLY !!!")



else:

    Sleeper
    print("Wrong Credentials !!!")  # Wrong Menu Selection
    Sleeper()