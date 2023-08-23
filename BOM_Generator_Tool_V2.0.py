# -*- coding: utf-8 -*-
"""
Created on Sat Aug 19 23:24:18 2023

@author: RAJA REDDY
"""

import tkinter as tk
import os
from tkinter import filedialog
#from PIL import Image, ImageTk  # Import PIL modules
import pandas as pd
import re
from openpyxl import load_workbook
#import openpyxl
#import time
#import odf
#from odf import text  # Importing the 'text' module from odfpy
#from odf import text
#from odf import odf

def alphanumeric_key(s):
    parts = re.split(r'(\d+)', str(s))
    for i in range(len(parts)):
        if parts[i].isdigit():
            parts[i] = int(parts[i])
    parts = [part for part in parts if part]
    return parts

def sort_alphanumeric(dataframe, column_number):
    # Get the column name based on the column number
    column_name = dataframe.columns[column_number]
    
    # Convert the values in the specified column to strings
    dataframe[column_name] = dataframe[column_name].astype(str)
    
    # Sort the dataframe based on the alphanumeric_key function
    sorted_df = dataframe.sort_values(by=column_name, key=lambda x: x.map(alphanumeric_key))
    return sorted_df

def sort_and_copy():
    global input_file_path
    input_file_path = filedialog.askopenfilename(title="Select Input Excel File")
    if not input_file_path:
        return

    try:
        sheet_name1 = 'Sheet1'
        data = pd.read_excel(input_file_path, sheet_name=sheet_name1)
        print(data)
        column_to_sort = 0 
       # Sort the data alphanumericly based on the specified column
        sorted_data = sort_alphanumeric(data, column_to_sort)
        print(sorted_data)
        input_directory = os.path.dirname(input_file_path)
        output_file_name = "AlphaNumerically_Sorted.xlsx"   
        output_file_path = os.path.join(input_directory, output_file_name)
        #column_name = data.columns[1]
        sorted_data.to_excel(output_file_path, index=False)
        status_label.config(text="BOM generated successfully! Please close application to generate final BOM", fg="green")
    except Exception as e:
        status_label.config(text=f"Error: {str(e)}", fg="red")

# GUI Setup
root = tk.Tk()
root.title("Medha BOM Generation Tool")
root.geometry("850x450")

# Widgets
title_label = tk.Label(root, text="MEDHA BOM Generation Tool", font=("Helvetica", 20, "bold"),fg="blue")
title_label.pack(pady=20)

sort_button = tk.Button(root, text="Provide input file", command=sort_and_copy)
sort_button.pack(pady=20)

#instructions_label = tk.Label(root, text="Instructions:\n1. Enter the path of the input Excel file in the entry box below or use the 'Browse' button to select the file.\n2. First Column must be lables of the components like R1,C2,...\n3. Second  Column must be SAP CODE.\n4. Third  Column must be Description. \n5. Fourth column must be Package .\n6. Fifth column must be 'Top' or 'Bottom'sides  of PCB  case sensitive Use Only 'Top' and 'Bottom'. \n7. Code is written python and can be edited any time. \n8. All the files will be generated after closing application \n9. Currently supports xlsx format only. Please convert to *.xlsx format from *.ods using 'save as using 'File' menu ", font=("Arial", 10))
instructions_label = tk.Label(root, text="""Instructions:
1. Enter the path of the input Excel file in the entry box below or use the 'Browse' button to select the file.
2. First Column must be labels of the components like R1, C2,...
3. Second Column must be SAP CODE.
4. Third Column must be Description.
5. Fourth column must be Package.
6. Fifth column must be 'Top' or 'Bottom' sides of PCB case sensitive. Use Only 'Top' and 'Bottom'.
7. Code is written in Python and can be edited any time.
8. All the files will be generated after closing the application.
9. Currently supports xlsx format only. Please convert to *.xlsx format from *.ods using 'Save As' in the 'File' menu.""",
font=("Arial", 12))  # Set font size to 14 and bold
instructions_label.pack(pady=5)

status_label = tk.Label(root, text="", fg="green")
status_label.pack()

root.mainloop()



def highlight_duplicates(s):
    # Create a boolean mask for duplicate values
    is_duplicate = s.duplicated(keep=False)
    
    # Create a style function to highlight duplicate rows
    def highlight(val):
        return ['background-color: yellow' if v else '' for v in is_duplicate]
    
    return s.apply(highlight)

# Load Excel file
#excel_file = "D:/BOM Generator Tool/BOM1.xlsx"
input_directory = os.path.dirname(input_file_path)
excel_file_path = os.path.join(input_directory, "AlphaNumerically_Sorted.xlsx")
#excel_file_path = 'C:/Users/br102314/Desktop/BOM1.xlsx'
sheet_name='Sheet1'

df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# print("Original DataFrame:")
# print(df)

# Choose the column you want to highlight duplicates in
column_to_check = df.columns[0] #'Designator'

# Apply the highlighting function to the selected column
highlighted_df = df.style.apply(highlight_duplicates, subset=[column_to_check], axis=None)

# Write the highlighted DataFrame to another Excel file
output_file_name="AlphaNumerically_Sorted.xlsx"
output_file_path = os.path.join(input_directory, output_file_name)
highlighted_df.to_excel(output_file_path, index=False)




# Define the find_missing_items function
def find_missing_items(sorted_list, prefix, max_value):
    missing_items = []
    expected_items = [f"{prefix}{i}" for i in range(1, max_value + 1)]
    
    sorted_set = set(sorted_list)
    expected_set = set(expected_items)
    
    missing_set = expected_set - sorted_set
    
    for missing_item in sorted(missing_set):
        missing_items.append(missing_item)
    return missing_items

# Specify input and output file paths
input_directory = os.path.dirname(input_file_path)
excel_file_path = os.path.join(input_directory, "AlphaNumerically_Sorted.xlsx")
output_file_name = "Missing_Items.xlsx"
output_file_path = os.path.join(input_directory, output_file_name)

# Read data from the input Excel file
df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
sorted_column_data = df.iloc[:, 0]  # Get data from the first column

# Extract prefixes and max values
prefixes_and_max_values = {item[0]: int(item[1:]) for item in sorted_column_data}
data_prefixes = list(prefixes_and_max_values.keys())

# Create an empty DataFrame to store all missing items
all_missing_items_df = pd.DataFrame(columns=['Missing_Items', 'SAP code', 'Description', 'Package', 'Position'])

# Loop through each prefix and append missing items
for prefix in data_prefixes:
    max_value = prefixes_and_max_values[prefix]
    prefix_sorted = [item for item in sorted_column_data if item.startswith(prefix)]
    missing_items = find_missing_items(prefix_sorted, prefix, max_value)

    # Append the missing items to the consolidated DataFrame
    if missing_items:
        missing_items_df = pd.DataFrame({'Missing_Items': missing_items, 'SAP code': " NA ", 'Description': " NA ", 'Package': " NA ", 'Position': 'Top'})
        all_missing_items_df = pd.concat([all_missing_items_df, missing_items_df], ignore_index=True)

# Write the consolidated DataFrame to the output Excel file
all_missing_items_df.to_excel(output_file_path, index=False)



# Define file paths
input_directory = os.path.dirname(input_file_path)
source_file_path = os.path.join(input_directory, "Missing_Items.xlsx")   # Update with the actual path
source_sheet_name = 'Sheet1'  # Update with the actual sheet name

target_file_path = os.path.join(input_directory, "AlphaNumerically_Sorted.xlsx")  # Update with the actual path
target_sheet_name = 'Sheet1'  # Update with the actual sheet name

# Load data from the source Excel file
source_data = pd.read_excel(source_file_path, sheet_name=source_sheet_name, skiprows=1)

# Load the target Excel file
target_workbook = load_workbook(target_file_path)
target_sheet = target_workbook[target_sheet_name]

# Append source data to the target sheet
for r_idx, row in source_data.iterrows():
    target_sheet.append(row.tolist())

# Save the updated target Excel file
target_workbook.save(target_file_path)


#############################################
#Top Side Components Sorting ans saving to Top_Side_Components
#input_directory = os.getcwd()
input_directory = os.path.dirname(input_file_path)
excel_file_path = os.path.join(input_directory, "AlphaNumerically_Sorted.xlsx")
#excel_file_path = 'C:/Users/br102314/Desktop/BOM1.xlsx'
sheet_name='Sheet1'
data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
column_to_sort1 =  data.columns[4] #Position
Top_Side_Components = data[data[column_to_sort1] == ('Top' or  'top')] #to sort top and bottom side
output_file_name="Top_Side_Components.xlsx"
output_file_path = os.path.join(input_directory, output_file_name)
Top_Side_Components.to_excel(output_file_path, index=False)
######################################################

#Top Side Components Sorting ans saving to Top_Side_Components
#current_directory = os.getcwd()
input_directory = os.path.dirname(input_file_path)
excel_file_path = os.path.join(input_directory, "AlphaNumerically_Sorted.xlsx")
#excel_file_path = 'C:/Users/br102314/Desktop/BOM1.xlsx'
sheet_name='Sheet1'
data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
column_to_sort1 = data.columns[4] #Position
Top_Side_Components = data[data[column_to_sort1] == ('Bottom' or 'bottom')] #to sort top and bottom side
output_file_name="Bottom_Side_Components.xlsx"
output_file_path = os.path.join(input_directory, output_file_name)
Top_Side_Components.to_excel(output_file_path, index=False)
######################################################


# ################ Sorting Top_Side_Components as per SAP Code####################
# #current_directory = os.getcwd()
# input_directory = os.path.dirname(input_file_path)
# excel_file_path = os.path.join(input_directory, "Top_Side_Components.xlsx")
# sheet_name='Sheet1'
# data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
# column_to_sort1 = data.columns[1] #'SAP Code'
# df_sorted = data.sort_values(by=column_to_sort1) #to sort top and bottom side
# print(df_sorted)
# output_file_name="Top_side_SAP_Sort.xlsx"
# output_file_path = os.path.join(input_directory, output_file_name)
# df_sorted.to_excel(output_file_path, index=False)
# #########################################################
# #Okay Till Now

# ################ Sorting Top_Side_Components as per SAP Code####################
# current_directory = os.path.dirname(input_file_path) #os.getcwd()
# excel_file_path = os.path.join(current_directory, "Bottom_Side_Components.xlsx")
# sheet_name='Sheet1'
# data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
# column_to_sort1 = 'SAP Code'
# df_sorted = data.sort_values(by=column_to_sort1) #to sort top and bottom side
# print(df_sorted)
# output_file_name="Bottom_side_SAP_Sort.xlsx"
# output_file_path = os.path.join(current_directory, output_file_name)
# df_sorted.to_excel(output_file_path, index=False)
# #########################################################

################### Generate Final BOM (Top Side) #######################################

def combine_strings_if_second_column_same(file_path):
    try:
        # Read the Excel file into a pandas DataFrame
        #current_directory = os.getcwd()
        current_directory = os.path.dirname(input_file_path)
        excel_file_path = os.path.join(current_directory, "Top_Side_Components.xlsx")
        df = pd.read_excel(excel_file_path)
        grouped = df.groupby(df.columns[1])[df.columns[0]].apply(','.join).reset_index()
        grouped = df.groupby(df.columns[1]).agg(lambda x: ','.join(x.astype(str))).reset_index()
        # Define the aggregation dictionary to apply different aggregation functions to each column
        aggregation_dict = {
        df.columns[0]: ','.join,  # Combine 'Column2' values with commas
        df.columns[2]: 'first',   # Keep the first value of 'Column3'
        df.columns[3]: 'first',     # Keep the first non-False value of 'Column4'
        df.columns[4]: 'count',     # Keep the first non-False value of 'Column4'
        #df.columns[4]: range(1, 1 + len(grouped)),
        }

# Perform the groupby operation and apply the aggregation dictionary
        grouped = df.groupby(df.columns[1]).agg(aggregation_dict).reset_index()
        #grouped = grouped[[df.columns[0], df.columns[1],df.columns[4],  df.columns[2], df.columns[3]]]
        grouped = grouped[[df.columns[1], df.columns[2],df.columns[4],  df.columns[0]]]
        grouped = grouped.rename(columns={df.columns[4]: 'Quantity'})
        #grouped = grouped.insert(0, 'S.No.', range(1, 1 + len(grouped)))
        # grouped = grouped.insert(0, 'S.No.', range(1, 1 + len(grouped)))
        grouped.insert(0, 'S.No.', range(1, 1 + len(grouped))) 
        return grouped

    except Exception as e:
        print("Error:", e)
        return None

if __name__ == "__main__":
    # Replace 'your_file_path.xlsx' with the actual path to your Excel file
    #file_path = 'file_path.xlsx'
    output_file_name="Final_Top_BOM.xlsx"
    current_directory = os.path.dirname(input_file_path)
    output_file_path = os.path.join(current_directory, output_file_name)
#    file_path = 'C:/Users/br102314/Desktop/BOM7.xlsx'
    result = combine_strings_if_second_column_same(output_file_path)
    if result is not None:
        print("Combined strings if the data in the second column is the same:")
        print(result)
    else:
        print("Failed to process the file.")
# output_file= 'C:/Users/br102314/Desktop/Top_BOM.xlsx'
result.to_excel(output_file_path, index=False)

##############################################

################### Generate Final BOM (Bottom Side) #######################################

def combine_strings_if_second_column_same(file_path):
    try:
        # Read the Excel file into a pandas DataFrame
        #current_directory = os.getcwd()
        current_directory = os.path.dirname(input_file_path)
        excel_file_path = os.path.join(current_directory, "Bottom_Side_Components.xlsx")
        df = pd.read_excel(excel_file_path)
        grouped = df.groupby(df.columns[1])[df.columns[0]].apply(','.join).reset_index()
        grouped = df.groupby(df.columns[1]).agg(lambda x: ','.join(x.astype(str))).reset_index()
        # Define the aggregation dictionary to apply different aggregation functions to each column
        aggregation_dict = {
        df.columns[0]: ','.join,  # Combine 'Column2' values with commas
        df.columns[2]: 'first',   # Keep the first value of 'Column3'
        df.columns[3]: 'first',     # Keep the first non-False value of 'Column4'
        df.columns[4]: 'count',     # Keep the first non-False value of 'Column4'
        }
     
# Perform the groupby operation and apply the aggregation dictionary
        grouped = df.groupby(df.columns[1]).agg(aggregation_dict).reset_index()
        grouped = grouped[[df.columns[1], df.columns[2],df.columns[4], df.columns[0]  ]]
        grouped = grouped.rename(columns={df.columns[4]: 'Quantity'})
        grouped.insert(0, 'S.No.', range(1, 1 + len(grouped)))
        return grouped

    except Exception as e:
        print("Error:", e)
        return None

if __name__ == "__main__":
    # Replace 'your_file_path.xlsx' with the actual path to your Excel file
    #file_path = 'file_path.xlsx'
    output_file_name="Final_Bottom_BOM.xlsx"
    current_directory = os.path.dirname(input_file_path)
    output_file_path = os.path.join(current_directory, output_file_name)
#    file_path = 'C:/Users/br102314/Desktop/BOM7.xlsx'
    result = combine_strings_if_second_column_same(output_file_path)
    if result is not None:
        print("Combined strings if the data in the second column is the same:")
        print(result)
    else:
        print("Failed to process the file.")
# output_file= 'C:/Users/br102314/Desktop/Top_BOM.xlsx'
result.to_excel(output_file_path, index=False)

###############################################

input_directory = os.path.dirname(input_file_path)
# file_path = os.path.join(input_directory, "Top_Bottom_Sorted.xlsx")
# os.remove(file_path)
# file_path = os.path.join(input_directory, "Top_Side_Components.xlsx")
# os.remove(file_path)
# file_path = os.path.join(input_directory, "Bottom_Side_Components.xlsx")
# os.remove(file_path)
# file_path = os.path.join(input_directory, "Top_side_SAP_Sort.xlsx")
# os.remove(file_path)
# file_path = os.path.join(input_directory, "Bottom_side_SAP_Sort.xlsx")
# os.remove(file_path)

# time(100)
