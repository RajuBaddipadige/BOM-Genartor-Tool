import tkinter as tk
import os
from tkinter import filedialog
from PIL import Image, ImageTk  # Import PIL modules
import pandas as pd
import re
import openpyxl
import time
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
root.geometry("800x350")

# Widgets
title_label = tk.Label(root, text="MEDHA BOM Generation Tool", font=("Helvetica", 20),fg="blue")
title_label.pack(pady=20)

sort_button = tk.Button(root, text="Provide input file", command=sort_and_copy)
sort_button.pack(pady=20)

instructions_label = tk.Label(root, text="Instructions:\n1. Enter the path of the input Excel file in the entry box below or use the 'Browse' button to select the file.\n2. First Column must be lables of the components like R1,C2.\n3. Second  Column must be SAP CODE.\n4. Third  Column must be Description. \n5. Fourth column must be Package .\n6. Fifth column must be 'Top' or 'Bottom'sides  of PCB  case sensitive Use Only 'Top' and 'Bottom'. \n7. Code is written python and can be edited any time. \n8. All the files will be generated after closing application")
instructions_label.pack(pady=5)

status_label = tk.Label(root, text="", fg="green")
status_label.pack()

root.mainloop()

time.sleep(0)
# Top and Bottom side compeonents  sorting and saving to Top_Bottom_Sorted
#print('Sort top and Bottom Components')
current_directory = os.getcwd()
excel_file_path = os.path.join(current_directory, "AlphaNumerically_Sorted.xlsx")
sheet_name='Sheet1'
data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
column_to_sort1 = data.columns[1]
df_sorted = data.sort_values(by=column_to_sort1) #to sort top and bottom side
print(df_sorted)
output_file_name="Top_Bottom_Sorted.xlsx"
output_file_path = os.path.join(current_directory, output_file_name)
df_sorted.to_excel(output_file_path, index=False)

#############################################################

#Top Side Components Sorting ans saving to Top_Side_Components
current_directory = os.getcwd()
excel_file_path = os.path.join(current_directory, "AlphaNumerically_Sorted.xlsx")
#excel_file_path = 'C:/Users/br102314/Desktop/BOM1.xlsx'
sheet_name='Sheet1'
data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
column_to_sort1 =  data.columns[4] #Position
Top_Side_Components = data[data[column_to_sort1] == ('Top' or  'top')] #to sort top and bottom side
output_file_name="Top_Side_Components.xlsx"
output_file_path = os.path.join(current_directory, output_file_name)
Top_Side_Components.to_excel(output_file_path, index=False)
######################################################

#Top Side Components Sorting ans saving to Top_Side_Components
current_directory = os.getcwd()
excel_file_path = os.path.join(current_directory, "AlphaNumerically_Sorted.xlsx")
#excel_file_path = 'C:/Users/br102314/Desktop/BOM1.xlsx'
sheet_name='Sheet1'
data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
column_to_sort1 = data.columns[4] #Position
Top_Side_Components = data[data[column_to_sort1] == ('Bottom' or 'bottom')] #to sort top and bottom side
output_file_name="Bottom_Side_Components.xlsx"
output_file_path = os.path.join(current_directory, output_file_name)
Top_Side_Components.to_excel(output_file_path, index=False)
######################################################


################ Sorting Top_Side_Components as per SAP Code####################
current_directory = os.getcwd()
excel_file_path = os.path.join(current_directory, "Top_Side_Components.xlsx")
sheet_name='Sheet1'
data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
column_to_sort1 = data.columns[1] #'SAP Code'
df_sorted = data.sort_values(by=column_to_sort1) #to sort top and bottom side
print(df_sorted)
output_file_name="Top_side_SAP_Sort.xlsx"
output_file_path = os.path.join(current_directory, output_file_name)
df_sorted.to_excel(output_file_path, index=False)
#########################################################

################ Sorting Top_Side_Components as per SAP Code####################
current_directory = os.getcwd()
excel_file_path = os.path.join(current_directory, "Bottom_Side_Components.xlsx")
sheet_name='Sheet1'
data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
column_to_sort1 = 'SAP Code'
df_sorted = data.sort_values(by=column_to_sort1) #to sort top and bottom side
print(df_sorted)
output_file_name="Bottom_side_SAP_Sort.xlsx"
output_file_path = os.path.join(current_directory, output_file_name)
df_sorted.to_excel(output_file_path, index=False)
#########################################################

################### Generate Final BOM (Top Side) #######################################

def combine_strings_if_second_column_same(file_path):
    try:
        # Read the Excel file into a pandas DataFrame
        current_directory = os.getcwd()
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
        grouped = grouped[[df.columns[0], df.columns[1],df.columns[4],  df.columns[2]]]
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

################### Generate Final BOM (Bottom Side) #######################################

def combine_strings_if_second_column_same(file_path):
    try:
        # Read the Excel file into a pandas DataFrame
        current_directory = os.getcwd()
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
        grouped = grouped[[df.columns[0], df.columns[1],df.columns[4],  df.columns[2]]]
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

os.remove("Top_Bottom_Sorted.xlsx")
os.remove("Top_Side_Components.xlsx")
os.remove("Bottom_Side_Components.xlsx") 
os.remove("Top_side_SAP_Sort.xlsx") 
os.remove("Bottom_side_SAP_Sort.xlsx") 

