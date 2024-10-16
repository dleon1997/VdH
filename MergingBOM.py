# -*- coding: utf-8 -*-
"""
Created on Fri Oct  4 11:46:21 2024

@author: diegobecerra
"""
import pandas as pd
from tkinter import Tk, filedialog
import numpy as np

# Function to open file dialog and select files
def upload_file(label):
    # Create a hidden Tkinter window (for file dialog)
    root = Tk()
    root.withdraw()  # Hide the root window
    root.call('wm', 'attributes', '.', '-topmost', True)  # Bring file dialog to the front
    
    # Open file dialog
    file_path = filedialog.askopenfilename(
        title=f'Select {label} file',
        filetypes=[('Excel files', '*.xlsx')]
    )
    
    # Check if a file was selected
    if file_path:
        print(f'{label} file selected: {file_path}')
        return file_path
    else:
        print(f'No {label} file selected.')
        return None

# Manually input file paths for BOM A and BOM B using file dialog
print('Please select BoM file coming from Tacton:')
bom_a_file = upload_file('BOM A')

print('Please select BoM file coming from 3D:')
bom_b_file = upload_file('BOM B')

bom_a=pd.read_excel(bom_a_file,header=None) #We read the file eliminating blank spaces
bom_a_headers= bom_a.apply(lambda x: x.notna().all(), axis=1).idxmax() #obtaing the first non empty row in order to get the headers
bom_a=pd.read_excel(bom_a_file, header=bom_a_headers).dropna(how='all') #Obtaing the files with the headers as we extracted before

#Same procedure done for bom a now for bom b
bom_b=pd.read_excel(bom_b_file,header=None)
bom_b_headers= bom_b.apply(lambda x: x.notna().all(), axis=1).idxmax() 
bom_b=pd.read_excel(bom_b_file, header=bom_b_headers).dropna(how='all') 


#Optional to ask information to ensure correct proccessing
headers_bom_a=list(bom_a.columns)
headers_bom_b=list(bom_b.columns)

article_code_col_a = headers_bom_a[0] #input('Enter the column name for Article Code in BOM A: ') #This is just in ase of wanting manual input 
quantity_col_a = 'Qty' #input('Enter the column name for Quantity in BOM A: ')
article_code_col_b = headers_bom_b[0] #input('Enter the column name for Article Code in BOM B: ')
quantity_col_b = 'Qty' #input('Enter the column name for Quantity in BOM B: ')

# Setting the indices for comparison
bom_a.set_index(article_code_col_a, inplace=True)
bom_b.set_index(article_code_col_b, inplace=True)

if bom_a.index.name != bom_b.index.name :
    bom_a.index.name = 'Article Code'
    bom_b.index.name = 'Article Code'
    
bom_a=bom_a.sort_index(ascending=True)
bom_b=bom_b.sort_index(ascending=True)

# Comparison
missing_in_b = bom_a[~bom_a.index.isin(bom_b.index)]
missing_in_a = bom_b[~bom_b.index.isin(bom_a.index)]
quantity_diff = bom_a.merge(bom_b,left_index=True,right_index=True,suffixes=('_A', '_B'))
quantity_diff['Quantity Difference'] = quantity_diff['Qty_A'] - quantity_diff['Qty_B']
# Using numpy.where to create the 'Notes' column based on 'Quantity Difference'
quantity_diff['Status'] = np.where(
    quantity_diff['Quantity Difference'] < 0,
    "There are " + quantity_diff['Quantity Difference'].abs().astype(str) + " pieces missing in Tacton",
    np.where(
        quantity_diff['Quantity Difference'] == 0,
        "Tacton and 3D matching",
        "There are " + quantity_diff['Quantity Difference'].abs().astype(str) + " pieces missing in 3D"
    )
)
print('Comparative View of the two Excels generated')

# Exporting to an Excel File
writer = pd.ExcelWriter('Comparison.xlsx', engine='xlsxwriter')

# Convert the dataframes to an XlsxWriter Excel object.
# Set index=True to include the Article Code index
quantity_diff.to_excel(writer, sheet_name='Quantities', index=True)
missing_in_b.to_excel(writer, sheet_name='MissinginSP', index=True)
missing_in_a.to_excel(writer, sheet_name='MissinginTacton', index=True)

# Get the workbook and the worksheets
workbook = writer.book
worksheet_quantities = writer.sheets['Quantities']
worksheet_missinginSP = writer.sheets['MissinginSP']
worksheet_missinginTacton = writer.sheets['MissinginTacton']

# Define a format for bordered cells
border_format = workbook.add_format({
    'border': 1  # Adds a border to all sides of the cell
})

# Set column width and apply borders for each dataframe
for worksheet, dataframe in zip([worksheet_quantities, worksheet_missinginSP, worksheet_missinginTacton],[quantity_diff, missing_in_b, missing_in_a]):
    # Apply border format to the entire dataframe
    last_row = len(dataframe)  # total number of data rows
    last_col = len(dataframe.columns)+1  # zero-based index for the last column
    worksheet.conditional_format(0, 0, last_row, last_col, 
                                 {'type': 'no_blanks', 'format': border_format})

# To highlight cells with non-zero values in the "Quantity Difference" column

# Define a format for cells that should be highlighted (non-zero values)
highlight_format = workbook.add_format({
    'bg_color': '#FF9999',  # Light red background
    'font_color': '#9C0006'  # Dark red text
})

# Get the row count for the quantity_diff dataframe
quantity_diff_rows = len(quantity_diff)  # This is the number of data rows
# If you want to include the header row in the color formatting, use:
# quantity_diff_rows += 1

# Locate the "Quantity Difference" column index
quantity_diff_col_index = quantity_diff.columns.get_loc('Quantity Difference')+1

# Apply conditional formatting to highlight non-zero values in "Quantity Difference" column
# Adjust the start row to 1 to skip the header
worksheet_quantities.conditional_format(1, quantity_diff_col_index, quantity_diff_rows, quantity_diff_col_index,
    {'type': 'cell',
     'criteria': '!=',
     'value': 0,
     'format': highlight_format})

# Save the workbook
writer.close()


 