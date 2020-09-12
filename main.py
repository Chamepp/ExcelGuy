'''
All Rights Reversed.
2020/12/09
My First Python Script
Developed With LOVE By Ashkan Ebtekari
'''


#imports

from win32com import client
import win32api



# Main Local Variables

Introduction = "Welcome To Ashkans Script " \
               "\n You Can Easily Handle Your Excel WorkSheets " \
               "\n in Diffrent Cool Ways ..."

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
Menu_Checker = UserInput == "3"



# Excel Converter

# Statements

if Menu_Merger:

    print("Selected Excel Merger")  # Excel Merger

elif Menu_Scraper:

    print("Selected Excel Scraper")  # Excel Scraper

elif Menu_Checker:

    # Instruction
    print("Selected Excel Converter")

    # Excel File Collector
    Excel_Source_Input = input("Give Me The Excel Source : ")
    Excel_Source = Excel_Source_Input
    Excel_Name_Input = input("Excel Name : ")
    Excel_Name = Excel_Name_Input
    Excel_Final_Output = r'{}{}'.format(Excel_Source , Excel_Name)
    print("Excel Source Collected Successfully ... \n \n")

    # PDF File Collector
    PDF_Source_Input = input("PDF OutPut File Directory : ")
    PDF_Source = PDF_Source_Input
    PDF_Name_Input = input("PDF Name : ")
    PDF_Name = PDF_Name_Input
    PDF_Final_Output = r'{}{}'.format(PDF_Source, PDF_Name)
    print("PDF Info Collected Successfully ... \n \n")


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

    print("Wrong Credentials !!!")  # Wrong Menu Selection
