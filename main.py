import pandas as pd
import numpy as np
from datetime import datetime
from tkinter import *
from tkinter import filedialog, messagebox
import tkinter as tk

# Function to read until null in Excel
def read_until_null_excel(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    null_row_index = df.isnull().all(axis=1).idxmax() if df.isnull().all(axis=1).any() else None
    if null_row_index is not None:
        df = df.iloc[:null_row_index]
    return df

# Function to check if a group is in an item
def is_group_in_item(item, group):
    return str(group) in str(item)

# Function to map group numbers [ này chat gpt hỗ trợ - chứ t bí ý tưởng :))) ] => cụ thể là match được item -- group xong match vendor thì nó mới trả group name
def map_group_numbers(df1, df2, item_column, vendor_column, groups_column):
    result = []
    for index1, row in df1.iterrows():
        item = row[item_column]
        vendor = str(row[vendor_column])
        matched_group = ''
        for index2, group_row in df2.iterrows():
            group_vendor = str(group_row[vendor_column])
            if group_vendor == vendor:
                group = group_row[groups_column]
                if is_group_in_item(item, group):
                    matched_group = group
                    print(f"Match found: Item '{item}' (Vendor: '{vendor}') matches Group '{group}'")
                    break
        result.append(matched_group)
        if not matched_group:
            print(f"No match for Item '{item}' (Vendor: '{vendor}')")
    return result

# Function to find the first non-null column name
def first_non_null_column_name(df, start_col, end_col, new_col_name):
    columns_to_check = df.loc[:, start_col:end_col].columns
    df[new_col_name] = df.apply(lambda row: next((col for col in columns_to_check if pd.notnull(row[col]) and row[col] != 0), None), axis=1)
    return df

def sum_pairs(df):
    # Initialize an empty list to store the results
    result_list = []
    
    # Separate numeric and non-numeric columns
    numeric_cols = df.select_dtypes(include='number').columns
    non_numeric_cols = df.select_dtypes(exclude='number').columns
    
    # Iterate through the DataFrame in steps of 2
    for i in range(0, len(df), 2):
        # Check if the next row exists
        if i + 1 < len(df):
            # Sum the current row and the next row for numeric columns
            summed_row = df.iloc[i:i+2][numeric_cols].sum()
            # Keep the non-numeric column from the first row
            non_numeric = df.iloc[i][non_numeric_cols]
            # Combine numeric and non-numeric results
            combined = pd.concat([summed_row, non_numeric])
            # Append the result to the list
            result_list.append(combined)
    
    # Convert the list of results to a DataFrame
    result = pd.DataFrame(result_list)
    
    return result

def parse_dates(date_str):
    try:
        return pd.to_datetime(date_str, format='%Y-%m-%d')
    except ValueError:
        return pd.NA
    
# get the first day between Ship and Firm when start to validation
def returnS_F(df_1):
    result_list = []
    for row in range(0, len(df_1), 2):
        found = False
        for col in range(0, len(df_1.columns),1):
            if df_1.iloc[row, col] > 0:
                result_list.append(df_1.columns[col])
                found = True
                break
        if not found and row + 1 < len(df_1):
            for col in range(0, len(df_1.columns),1):
                if df_1.iloc[row + 1, col] > 0:
                    result_list.append(df_1.columns[col])
                    found = True
                    break
        if not found:
            result_list.append("")
    result = pd.DataFrame(result_list, columns=['New-Outcome'])
    return result

def remove_extension(filename,extentions):
    for ext in extentions:
        filename = str(filename)
        filename = filename.replace(ext,'')
    return filename

def highlight_condition_Arca_EC(df, s1, s2,date_columns_1):
    # Define a function to apply to each row
    def highlight(row):
        # Initialize a list of empty styles for each column
        styles = [''] * len(row)
        
        # Apply red background if the condition is met
        if row[s1] < row[s2]:
            styles[df.columns.get_loc(s1)] = 'background-color: red'
            styles[df.columns.get_loc(s2)] = 'background-color: red'
        
        return styles
    df[date_columns_1] = df[date_columns_1].astype(str)
    df[date_columns_1] = df[date_columns_1].replace('1999-01-01', '0')
    df['Vendor #'] = df['Vendor #'].astype(int)
    # Apply the function row-wise
    return df.style.apply(highlight, axis=1)

# Processing function - sheet 1
def processing(smartsheet_file, system_file, sheet_used, num_start, num_end):
    print(f"Processing {smartsheet_file}, {system_file}, {sheet_used}, {num_start}, {num_end}")
    df_2 = pd.read_csv(smartsheet_file, skiprows=6)  # Skip the system docs info lines
    df_2['Vendor #'] = df_2['Vendor #'].astype(float)
    df_1 = read_until_null_excel(system_file, sheet_used)
    
    vendor_take = df_1['Vendor #'].unique()
    df_2 = df_2[df_2['Vendor #'].isin(vendor_take)]
    
    split_columns = df_1['Group Number'].str.split(' ', expand=True)
    
    if split_columns.shape == 2:
        split_columns.columns = ['Group Number Split','Additional Component']
    else:
        split_columns = split_columns.rename(columns={0: 'Group Number Split'})
        split_columns['Additional Component'] = None
    
    df_1 = pd.concat([df_1, split_columns], axis=1)
    
    df_1.drop(columns=['Group Number'], inplace=True)
    df_1 = df_1.rename(columns={'Group Number Split': 'Group Number'})
    
    df_2['Item #'] = df_2['Item #'].str.replace('="', '').str.replace('"', '')
    
    df_2_sample = df_2.loc[(df_2['S/F/P'] == 'F') | (df_2['S/F/P'] == 'S')]
    df_2_sample = df_2_sample.loc[:,df_2_sample.columns[num_start]:df_2_sample.columns[num_end]]
    
    df_2_sample  = df_2_sample.astype(float)
    df = pd.DataFrame(df_2_sample)
    result = returnS_F(df)    
    
    # Day First Valid => Ship/Confirmation
    df_2 = df_2.loc[df_2['S/F/P'] == 'F'].reset_index(drop=True)
    df_2['Date_First_Value'] = result
    
    # mapping section
    item_column = 'Item #'
    vendor_column = 'Vendor #'
    groups_column = 'Group Number'
    df_2['Group Number'] = map_group_numbers(df_2, df_1, item_column, vendor_column, groups_column)
    
    df_check = df_2.loc[df_2['Group Number'] != '']
    
    df_check['Arcadia ETD System'] = np.where(df_check['Whse'].isin({'1'}), 1, 0)
    df_check['EC ETD System'] = np.where(df_check['Whse'].isin({'15', '17', 'ECR'}), 1, 0)
    df_check['WC ETD System'] = np.where(df_check['Whse'].isin({'42', '28', '5'}), 1, 0)
    
    df_check_v2 = pd.merge(df_check, df_1[['Vendor #', 'Group Number', 'Additional Component', 'Arcadia ETD', 'EC ETD', 'WC ETD','Categories']], on=['Group Number','Vendor #'], how='left')
    
    df_check_v3 = df_check_v2.copy()
    
    current_year = datetime.now().strftime('%Y')

    #pretend solution - not the official - need medicate here !!!!
    df_check_v3['Date_First_Value'] = df_check_v3['Date_First_Value'].str.strip() + '/' + current_year
    df_check_v3['Date_First_Value'] = df_check_v3['Date_First_Value'].replace('/2024','01/01/2027')
    df_check_v3['Date_First_Value'] = pd.to_datetime(df_check_v3['Date_First_Value'], format='%m/%d/%Y')
    df_check_v3['Date_First_Value'] = df_check_v3['Date_First_Value'].dt.strftime('%Y-%m-%d')
    df_check_v3['Date_First_Value'] = df_check_v3['Date_First_Value'].replace('2027-01-01','0')
    
    df_check_v3['Arcadia ETD System Final'] = np.where(df_check_v3['Arcadia ETD System'] == 1, df_check_v3['Date_First_Value'], 0)
    df_check_v3['EC ETD System Final'] = np.where(df_check_v3['EC ETD System'] == 1, df_check_v3['Date_First_Value'], 0)
    df_check_v3['WC ETD System Final'] = np.where(df_check_v3['WC ETD System'] == 1, df_check_v3['Date_First_Value'], 0)

    for col in ['Arcadia ETD', 'EC ETD', 'WC ETD']:
        df_check_v3[col] = df_check_v3[col].dt.strftime('%Y-%m-%d')
    
    df_check_v3['Arcadia ETD Smartsheet'] = np.where(df_check_v3['Arcadia ETD System'] == 1, df_check_v3['Arcadia ETD'], 0)
    df_check_v3['EC ETD Smartsheet'] = np.where(df_check_v3['EC ETD System'] == 1, df_check_v3['EC ETD'], 0)
    df_check_v3['WC ETD Smartsheet'] = np.where(df_check_v3['WC ETD System'] == 1, df_check_v3['WC ETD'], 0)
    
    df_filtered_next = df_check_v3[['Categories','Item #', 'Whse', 'Group Number', 'Additional Component', 'Vendor #',
                                    'Arcadia ETD System Final', 'EC ETD System Final', 'WC ETD System Final',
                                    'Arcadia ETD Smartsheet', 'EC ETD Smartsheet', 'WC ETD Smartsheet']]
    
    df_filtered_next = df_filtered_next.loc[df_filtered_next['Whse'].isin({'1','15','17','ECR','42','28','5'})]
    df_filtered_next['Additional Component'] = df_filtered_next['Additional Component'].fillna('None')
    
    # Replace non-date strings with NaN
    date_columns = ['Arcadia ETD System Final','EC ETD System Final','WC ETD System Final','Arcadia ETD Smartsheet','EC ETD Smartsheet','WC ETD Smartsheet']
    df_filtered_next[date_columns] = df_filtered_next[date_columns].astype(str)
    df_filtered_next[date_columns] = df_filtered_next[date_columns].replace('0', np.nan)

    for col in date_columns:
        df_filtered_next[col] = df_filtered_next[col].apply(parse_dates)
        
    # Fill NaN values with a placeholder date for aggregation purposes
    placeholder_date = '2027-12-31'
    df_filtered_next[date_columns] = df_filtered_next[date_columns].fillna(placeholder_date)
        
    # Convert date columns to datetime
    df_filtered_next[date_columns] = df_filtered_next[date_columns].apply(pd.to_datetime, format='%Y-%m-%d')
        
    df_new = df_filtered_next.groupby(['Categories','Item #', 'Group Number','Additional Component','Vendor #','Arcadia ETD Smartsheet','EC ETD Smartsheet','WC ETD Smartsheet']).agg({
        'Whse': lambda x: set(x),
        'Arcadia ETD System Final': 'min',
        'EC ETD System Final': 'min',
        'WC ETD System Final': 'min'
    }).reset_index()
    
    df_new['Check True/False'] = np.where(
        (df_new['Arcadia ETD System Final'] == df_new['Arcadia ETD Smartsheet']) &
        (df_new['EC ETD System Final'] == df_new['EC ETD Smartsheet']) &
        (df_new['WC ETD System Final'] == df_new['WC ETD Smartsheet']),
        'True', 'False'
    )
    
    # turn from datetime to string 
    df_new = df_new[['Categories','Item #','Whse','Vendor #','Group Number','Additional Component','Arcadia ETD System Final','EC ETD System Final','WC ETD System Final','Arcadia ETD Smartsheet','EC ETD Smartsheet','WC ETD Smartsheet','Check True/False']]

    for col in date_columns:
        df_new[col] = df_new[col].dt.date

    existing_columns = [col for col in date_columns if col in df_new.columns]

    if existing_columns:
        df_new[existing_columns] = df_new[existing_columns].astype(str)
    else:
        print("No valid columns found for conversion.")

    df_new = df_new.replace(placeholder_date, '0', inplace=False)

    return df_new

# Processing conditional formatting - sheet 2
def processing_2(df_new):
    df_new_1 = df_new.copy()

    date_columns_1 = ['Arcadia ETD System Final','EC ETD System Final','WC ETD System Final','Arcadia ETD Smartsheet','EC ETD Smartsheet','WC ETD Smartsheet']
    
    df_new_1[date_columns_1] = df_new_1[date_columns_1].astype(str)
    df_new_1[date_columns_1] = df_new_1[date_columns_1].replace('0',np.nan)
    place_nan = '1999-01-01'
    df_new_1[date_columns_1] = df_new_1[date_columns_1].fillna(place_nan)
    df_new_1[date_columns_1] = df_new_1[date_columns_1].apply(pd.to_datetime, format='%Y-%m-%d')
    
    df_new_1 = df_new_1[['Categories','Item #','Vendor #','Group Number','Additional Component','Arcadia ETD System Final','EC ETD System Final','WC ETD System Final','Arcadia ETD Smartsheet','EC ETD Smartsheet','WC ETD Smartsheet']]

    df_new_1 = df_new_1.groupby(['Categories', 'Item #','Vendor #','Group Number','Additional Component'])[date_columns_1].max().reset_index()

    styled_df = highlight_condition_Arca_EC(df_new_1, 'Arcadia ETD System Final', 'EC ETD System Final',date_columns_1)

    return styled_df

# Concatenate the results for the new filename
def export_file(df1, df2, system_file, smartsheet_file):
    filename = system_file.split('/')[-1]
    smartsheetname = smartsheet_file.split('/')[-1]
    extension_to_remove = ['.xlsx', '.csv']
    
    clean_system_file = remove_extension(filename, extension_to_remove)
    clean_smartsheet_file = remove_extension(smartsheetname, extension_to_remove)
    
    new_filename = f'File_check [{clean_system_file} - {clean_smartsheet_file}].xlsx'
    
    # Create an Excel writer object
    with pd.ExcelWriter(new_filename, engine='xlsxwriter') as writer:
        df1.to_excel(writer, index=False, sheet_name='Check True False')
        df2.to_excel(writer, index=False, sheet_name='Condition Formatting')
    
    print("Processing complete - Check the file in the repository.")

def open_smartsheet_file():
    global smartsheet_file
    smartsheet_file = filedialog.askopenfilename(title="Select Smartsheet File", filetypes=[("CSV Files", "*.csv")])
    smartsheet_label.config(text=smartsheet_file.split('/')[-1])

def open_system_file():
    global system_file, system_sheets_name
    system_file = filedialog.askopenfilename(title="Select System File", filetypes=[("Excel Files", "*.xlsx")])
    system_sheets_name = pd.ExcelFile(system_file).sheet_names
    system_label.config(text=system_file.split('/')[-1])
    
    # Update dropdown with sheet names
    clicked.set("Choose the sheet")
    drop['menu'].delete(0, 'end')
    for sheet in system_sheets_name:
        drop['menu'].add_command(label=sheet, command=lambda value=sheet: clicked.set(value))

# Function to update start and end column dropdown based on selected sheet
def update_column_options(*args):
    df = pd.read_csv(smartsheet_file, skiprows = 6)
    
    # Update start options
    start_options = list(range(4, len(df.columns)-1))  # Start from 4 to len(columns) - 1
    num_start_var.set(start_options[0])  # Set default value
    start_menu['menu'].delete(0, 'end')  # Clear existing options
    for option in start_options:
        start_menu['menu'].add_command(label=option, command=tk._setit(num_start_var, option))

    # Update end options based on the start column selected
    update_end_options()

def update_end_options(*args):
    df = pd.read_csv(smartsheet_file, skiprows = 6)
    start = num_start_var.get()
    end_options = list(range(start + 1, len(df.columns)))  # End options start from num_start + 1
    num_end_var.set(end_options[0])  # Set default value for num_end_var
    end_menu['menu'].delete(0, 'end')  # Clear existing options
    for option in end_options:
        end_menu['menu'].add_command(label=option, command=tk._setit(num_end_var, option))

# Retrieve start and end column numbers
def retrieve_values():
    num_start = num_start_var.get()
    num_end = num_end_var.get()
    return num_start, num_end

# Show function for GUI
def show():
    sheet_used = clicked.get()
    if sheet_used == "Choose the sheet":
        messagebox.showerror("Error", "Please select a valid sheet.")
    else:
        num_start, num_end = retrieve_values()
        df_new = processing(smartsheet_file, system_file, sheet_used, num_start, num_end)
        df_styled = processing_2(df_new)
        export_file(df_new, df_styled, system_file, smartsheet_file)
        messagebox.showinfo("Success", "Processing completed successfully!")

# Create GUI
root = tk.Tk()
root.geometry("600x300")
root.title("Sheet Selection")

# Create frames for better organization
frame1 = tk.Frame(root)
frame1.pack(pady=10, padx=10, fill='x')

frame2 = tk.Frame(root)
frame2.pack(pady=10, padx=10, fill='x')

frame3 = tk.Frame(root)
frame3.pack(pady=10, padx=10, fill='x')

frame5 = tk.Frame(root)
frame5.pack(pady=10, padx=10, fill='x')

# Smartsheet file selection
Label(frame1, text="Smartsheet File: ").pack(side=LEFT)
smartsheet_button = Button(frame1, text="Select File", command=open_smartsheet_file)
smartsheet_button.pack(side=LEFT)
smartsheet_label = Label(frame1, text="No file selected")
smartsheet_label.pack(side=LEFT)

# System file selection
Label(frame2, text="System File: ").pack(side=LEFT)
system_button = Button(frame2, text="Select File", command=open_system_file)
system_button.pack(side=LEFT)
system_label = Label(frame2, text="No file selected")
system_label.pack(side=LEFT)

# Sheet selection dropdown
Label(frame3, text="Smartsheet Sheet Selection: ").pack(side=tk.LEFT)
clicked = StringVar()
clicked.set("Choose the sheet")
drop = OptionMenu(frame3, clicked, "Sheet 1", "Sheet 2", "Sheet 3")  # Dummy options; updated by open_system_file()
drop.pack(side=tk.LEFT)
clicked.trace("w", update_column_options)

Label(frame5, text="Column Start: ").pack(side=tk.LEFT)
num_start_var = IntVar()
num_start_var.set(4)  # Default value
start_menu = OptionMenu(frame5, num_start_var, [])  # Empty list, updated later
start_menu.pack(side=tk.LEFT)
num_start_var.trace("w", update_end_options)  # Update end options when start changes

# Column end dropdown
Label(frame5, text="Column End: ").pack(side=tk.LEFT)
num_end_var = IntVar()
end_menu = OptionMenu(frame5, num_end_var, [])  # Empty list, updated later
end_menu.pack(side=tk.LEFT)

# Process button
button = Button(root, text="Process", command=show)
button.pack(pady=10)

# Instructions label
label = Label(root, text="Select the files and sheet, then click 'Process'")
label.pack(pady=10)

root.mainloop()