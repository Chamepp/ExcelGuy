'''
All Rights Reversed.
2020/12/09
My First Python Script
Developed With LOVE By Ashkan Ebtekari
'''

# Main Local Variables


Introduction = "Welcome To Ashkans Script " \
               "\n You Can Easily Handle Your Excel WorkSheets " \
               "\n in Diffrent Cool Ways ..."

Menu = "1- Excel Merger " \
       "\n 2- Excel Scraper" \
       "\n 3- Excel Checker" \
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

# Statements

if Menu_Merger:

    print("Selected Excel Merger")  # Excel Merger

elif Menu_Scraper:

    print("Selected Excel Scraper")  # Excel Scraper

elif Menu_Checker:

    print("Selected Excel Checker")  # Excel Checker


else:

    print("Wrong Credentials !!!")  # Wrong Menu Selection
