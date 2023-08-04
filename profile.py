#Importing packages 
from openpyxl import Workbook
from openpyxl.styles import Font

from openpyxl import load_workbook

# Load the existing workbook
workbook = load_workbook("profiles.xlsx")

# Select the sheet from the workbook
sheet = workbook["Sheet"]
number_column = 1  # Column number for candidate numbers
link_column = 2   # Column number for candidate links

# Loop to collect candidate names and links
while True:
      # Ask the user for candidate name (exit loop if "exit" is entered)
    user_name = input("Please enter candidate name (enter 'exit' to quit): ")
    if user_name.lower() == "exit":
        break
       # Ask the user for candidate link
    user_link = input("Please enter a candidate link: ")
    # Find the last used row with a hyperlink in the specified column
    last_used_row = None
    for row in range(sheet.max_row, 0, -1):
        if sheet.cell(row=row, column=link_column).hyperlink:
            last_used_row = row
            break
     # Determine the row for the new entry    
    current_row = last_used_row + 1 if last_used_row else 2
       # Calculate the candidate number for the new entry
    start_number = int(sheet.cell(row=current_row - 1, column=number_column).value) + 1 if current_row > 2 else 1

    # Write the candidate number to the sheet
    cell_number = sheet.cell(row=current_row, column=number_column)
    cell_number.value = start_number
      # Write the candidate name and link to the sheet
    cell_link = sheet.cell(row=current_row, column=link_column)
    cell_link.value = user_name
    cell_link.hyperlink = user_link

     # Save the workbook after each entry
    workbook.save("profiles.xlsx")






