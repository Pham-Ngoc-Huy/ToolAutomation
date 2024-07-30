import pandas as pd
import numpy as np
from datetime import datetime
from tkinter import *

# Function Support
def read_until_null_excel(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    null_row_index = df.isnull().all(axis=1).idxmax() if df.isnull().all(axis=1).any() else None

    if null_row_index is not None:
        df = df.iloc[:null_row_index]
    return df

def is_group_in_item(item, group):
    item_str = str(item)
    group_str = str(group)
    return group_str in item_str
def map_group_numbers(item, groups):
    for group in groups:
        if is_group_in_item(item, group):
            return group
    return ''

def first_non_null_column_name(df, start_col, end_col, new_col_name):
    columns_to_check = df.loc[:, start_col:end_col].columns

    def find_first_non_null(row):
        for col in columns_to_check:
            if pd.notnull(row[col]) and row[col] != 0:
                return col
        return None

    df[new_col_name] = df.apply(find_first_non_null, axis=1)
    return df

# Processing file 
def processing(smartsheet_file, system_file, sheet_used, num_start, num_end):
    df_2 = pd.read_csv(smartsheet_file, skiprows = 6) #6 here are the system docs info lines
    df_2['Vendor #'] = df_2['Vendor #'].astype(float)
    df_1 = read_until_null_excel(system_file, sheet_used)

    vendor_take = df_2['Vendor #'].unique()
    df_1 = df_1[df_1['Vendor #'].isin(vendor_take)]

    split_columns = df_1['Group Number'].str.split(' ', expand=True)
    split_columns.columns = ['Group Number Split','Additional Component']

    df_1 = pd.concat([df_1, split_columns], axis=1)
    group_number_list = df_1['Group Number Split'].unique().tolist()
    df_1.drop(columns=['Group Number'], inplace=True)
    df_1 = df_1.rename(columns={'Group Number Split': 'Group Number'})

    df_2['Item #'] = df_2['Item #'].str.replace('="', '').str.replace('"', '')
    df_2 = df_2.loc[df_2['S/F/P'] == 'F']
    df_2['Group Number'] = df_2['Item #'].apply(lambda x: map_group_numbers(x, group_number_list))

    df_check = df_2.loc[df_2['Group Number'] != '']
    df_check['Arcadia ETD System'] = np.where(df_check['Whse'].isin({'1'}), 1, 0)
    df_check['EC ETD System'] = np.where(df_check['Whse'].isin({'15', '17', 'ECR'}), 1, 0)
    df_check['WC ETD System'] = np.where(df_check['Whse'].isin({'42', '28', '5'}), 1, 0)

    df_check_v2  = pd.merge(df_check, df_1[['Group Number','Additional Component','Arcadia ETD','EC ETD', 'WC ETD']], on ='Group Number', how = 'left')

    start_column = df_check_v2.columns[num_start]
    end_column = df_check_v2.columns[num_end]

    df_check_v3 = first_non_null_column_name(df_check_v2, start_column, end_column, 'Date_First_Value')
    current_year = datetime.now().strftime('%Y')
    df_check_v3['Date_First_Value'] = df_check_v3['Date_First_Value'].str.strip()
    df_check_v3['Date_First_Value']= df_check_v3['Date_First_Value'] + '/' + current_year
    df_check_v3['Date_First_Value'] = pd.to_datetime(df_check_v3['Date_First_Value'], format='%m/%d/%Y', errors='coerce')
    df_check_v3['Date_First_Value'] = df_check_v3['Date_First_Value'].dt.strftime('%Y-%m-%d')

    df_check_v3['Arcadia ETD System Final'] = np.where(df_check_v3['Arcadia ETD System'] == 1, df_check_v3['Date_First_Value'],0)
    df_check_v3['EC ETD System Final'] = np.where(df_check_v3['EC ETD System'] == 1, df_check_v3['Date_First_Value'],0)
    df_check_v3['WC ETD System Final'] = np.where(df_check_v3['WC ETD System'] == 1, df_check_v3['Date_First_Value'],0)

    df_check_v3['Arcadia ETD'] = df_check_v3['Arcadia ETD'].dt.strftime('%Y-%m-%d')
    df_check_v3['EC ETD'] = df_check_v3['EC ETD'].dt.strftime('%Y-%m-%d')
    df_check_v3['WC ETD'] = df_check_v3['WC ETD'].dt.strftime('%Y-%m-%d')

    df_check_v3['Arcadia ETD Smartsheet'] = np.where(df_check_v3['Arcadia ETD System'] == 1, df_check_v3['Arcadia ETD'], 0)
    df_check_v3['EC ETD Smartsheet'] = np.where(df_check_v3['EC ETD System'] == 1, df_check_v3['EC ETD'], 0)
    df_check_v3['WC ETD Smartsheet'] = np.where(df_check_v3['WC ETD System'] == 1, df_check_v3['WC ETD'], 0)

    df_filtered_next = df_check_v3[['Item #','Whse','Group Number','Additional Component','Vendor #','Arcadia ETD System Final','EC ETD System Final','WC ETD System Final','Arcadia ETD Smartsheet','EC ETD Smartsheet','WC ETD Smartsheet']]
    df_filtered_next['Check True/False'] = np.where(
    (df_filtered_next['Arcadia ETD System Final'] == df_filtered_next['Arcadia ETD Smartsheet']) &
    (df_filtered_next['EC ETD System Final'] == df_filtered_next['EC ETD Smartsheet']) &
    (df_filtered_next['WC ETD System Final'] == df_filtered_next['WC ETD Smartsheet']),
    'True', 'False'
    )
    return df_filtered_next
def show(): 
    label.config( text = clicked.get() ) 

# Build GUI App




# main.py
# input smartsheet file
smartsheet_file = 'Production ScheduleDoLuong44218.csv'

# input system file and sheet name
system_file = 'Jan LV 2024 Price Confirmation Roll-Up 7.20.xlsx'
system_file_object = pd.ExcelFile(system_file)
system_sheets_name = system_file_object.sheet_names
# create drop down
root = Tk()
root.geometry( "200x200" ) 
options = [system_sheets_name]

clicked = StringVar()
clicked.set( "Choose the sheet of smartsheet" )
drop = OptionMenu( root , clicked , *options ) 
drop.pack() 
button = Button( root , text = "click Me" , command = show ).pack() 
label = Label( root , text = " " ) 
label.pack() 
root.mainloop() 

#sheet_used = 'Jan LV 2024 Price Confirmation '

# input start - end column that will detect the date to compare between smartsheet and system - count from zero
num_start = 4
num_end = 22

