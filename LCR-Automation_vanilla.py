#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import os
import re
import csv
import glob
import numpy as np
import xlwings as xw
from datetime import datetime, timedelta
import time

# Record the start time
start_time = time.time()

def rename_and_save(file_path, new_name):
    try:
        # Remove the old extension and add .txt to the new name
        new_file_path = os.path.join(parent_folder, new_name + '.txt')
        os.rename(file_path, new_file_path)
        print(f"Renamed and saved: {new_name}")
    except Exception as e:
        print(f"Error renaming and saving {file_path}: {e}")

try:
    # Get user input for the parent folder path
    parent_folder = input("Enter the path of parent folder consisting .txt files (Lombard Files): ")

    # Check if the directory exists
    if not os.path.isdir(parent_folder):
        raise FileNotFoundError(f"The directory '{parent_folder}' does not exist.")


    # Get a list of all files in the parent folder
    all_files = os.listdir(parent_folder)

    # Filter files that end with .txt
    txt_files = [file for file in all_files if file.endswith('.txt')]

    # Rename and save the .txt files
    for txt_file in txt_files:
        # Remove 'Lombard' and 'Report', and remove special characters and numbers
        new_name = re.sub(r'[^a-zA-Z]+', '', txt_file.replace('Lombard', '').replace('Report', ''))

        # Remove spaces (without converting to lowercase)
        new_name = new_name.replace(' ', '')

        # Remove the 'txt' extension
        new_name = new_name.replace('txt', '')

        file_path = os.path.join(parent_folder, txt_file)
        rename_and_save(file_path, new_name)

    # Verify the renamed files
    renamed_files = os.listdir(parent_folder)
    print(f"Renamed files in '{parent_folder}': {renamed_files}")

except Exception as ex:
    print(f"An error occurred: {ex}")

# Get user input for the parent folder
#parent_folder = input("Enter the parent folder path: ")

# Get a list of all files in the parent folder
all_files = os.listdir(parent_folder)

#all_files

def read_excel_file(parent_folder, file_name):
    """
    Check if the specified Excel file exists in the given folder and read it into a DataFrame.

    Args:
    parent_folder (str): Path of the parent folder.
    file_name (str): Name of the Excel file.

    Returns:
    pd.DataFrame: DataFrame containing the data from the Excel file, or None if the file does not exist.
    """
    file_path = os.path.join(parent_folder, file_name)

    # Check if the file exists before reading it
    if os.path.exists(file_path):
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)
        print(f"File '{file_name}' loaded successfully.")
        return df
    else:
        print(f"The file '{file_name}' does not exist in the specified folder.")
        return None

# Usage
#parent_folder = 'path_to_your_parent_folder'
file_name = 'Logic_for_Deposit_Outflows.xlsx'
Logic_for_Deposit_Outflows_df = read_excel_file(parent_folder, file_name)

'''file_name = 'Logic_for_Deposit_Outflows.xlsx'
file_path = os.path.join(parent_folder, file_name)

# Check if the file exists before reading it
if os.path.exists(file_path):
    # Read the Excel file into a DataFrame
    Logic_for_Deposit_Outflows_df = pd.read_excel(file_path)

else: 
    print(f"The file '{file_name}' does not exist in the specified folder.")
'''

#Logic_for_Deposit_Outflows_df

# Filter files that end with .txt
txt_files = [file for file in all_files if file.endswith('.txt')]

# Loop through the txt_files and read "FSCSSUB" file into a dataframe
for txt_file in txt_files:
    if "FSCSSUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_FSCSSUB = pd.read_csv(file_path)
        
    if "AccountBalanceSUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_accbal = pd.read_csv(file_path)
        
    if "AMMSUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_amm = pd.read_csv(file_path)
    
    if "CustSUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_cust = pd.read_csv(file_path)
    
    if "DerivativeSUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_derivative = pd.read_csv(file_path)
        
    if "ExchangeRateSUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_exch = pd.read_csv(file_path)
        
    if "GLBalanceSUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_glbal = pd.read_csv(file_path)
        
    if "LimitSUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_limit = pd.read_csv(file_path)
    
    if "RepaySUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_repay = pd.read_csv(file_path)
        
    if "LoanDepSUB" in txt_file:
        file_path = os.path.join(parent_folder,txt_file)
        df_loandep = pd.read_csv(file_path)
        
    #if "CollateralSUB" in txt_file:
    #    file_path = os.path.join(parent_folder,txt_file)
    #    df_collat = pd.read_csv(file_path)
        
    if "SecuritySUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_security = pd.read_csv(file_path)
        
    if "SECURITYPORTFOLIOREPORTSUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_securityportfolio = pd.read_csv(file_path)
    
    if "FXDealsSUB" in txt_file:
        file_path = os.path.join(parent_folder, txt_file)
        df_fxdeals = pd.read_csv(file_path)


'''cuba holding :- 800002322
Nijjar Holding:- 800002914
N And S:- 610963269
Lamba Pension Fund:- 800002588
Change the sector code for above to be 3410
And Change Industry code to 1001'''

#xcvbn = df_cust[df_cust['Customer Identifier'] == '610963269']
#xcvbn

# Define the list of customer identifiers

customer_identifiers = [800002322, 800002914, '610963269', 800002588]

# Define the new sector code and industry code
new_sector_code = 3410
new_industry_code = 1001

# Update the values based on conditions
mask = df_cust['Customer Identifier'].isin(customer_identifiers)
df_cust.loc[mask, 'Sector Code'] = new_sector_code
df_cust.loc[mask, 'Industry Code'] = new_industry_code

columns_to_drop = ['Currency Name', 'Material Currency', 'Significant Currency', 'Multiplier/Divisor']
df_exch = df_exch.drop(columns=columns_to_drop)

#df_exch

# Filter rows where Currency is "USD"
exch_usd = df_exch[df_exch['Currency Code'] == 'USD'].reset_index(drop=True)
exch_GBP = df_exch[df_exch['Currency Code'] == 'GBP'].reset_index(drop=True)
exch_EUR = df_exch[df_exch['Currency Code'] == 'EUR'].reset_index(drop=True)
exch_CAD = df_exch[df_exch['Currency Code'] == 'CAD'].reset_index(drop=True)
exch_INR = df_exch[df_exch['Currency Code'] == 'INR'].reset_index(drop=True)

# # Dates ( future_date; notice_acc_dt; casa_excl_notice_acc_dt; one_yr_old_dt)

# Get user input for the date in dd/mm/yyyy format
date_input = input("Enter the date (dd/mm/yyyy): ")

try:
    # Parse the user input into a datetime object
    date_obj = datetime.strptime(date_input, '%d/%m/%Y')
    
    # Calculate the date 30 days after the input date
    future_date = (date_obj + timedelta(days=30)).strftime('%d/%m/%Y')
    
    # Calculate the notice Account date ( 35 days from actual date)
    notice_acc_dt = (date_obj + timedelta(days=35)).strftime('%d/%m/%Y')
    
    # Calculate the CASA excluding Notice Account date ( 1  day from future date)
    casa_excl_notice_acc_dt = (date_obj + timedelta(days=1)).strftime('%d/%m/%Y')
    
    # Calculate the 12-month old date ( 365 days or 1 year from date_input)
    one_yr_old_dt = (date_obj - timedelta(days=365)).strftime('%d/%m/%Y')
    
except ValueError:
    print("Invalid date format. Please enter the date in dd/mm/yyyy format.")

'''try:
    # Parse the user input into a datetime object
    date_obj = datetime.strptime(date_input, '%d/%m/%Y')
    
    # Calculate the date 30 days after the input date
    future_date = (date_obj + timedelta(days=30))
    #future_date = (date_obj + timedelta(days=30)).strftime('%d/%m/%Y')
    
    # Calculate the notice Account date ( 35 days from actual date)
    notice_acc_dt = date_obj + timedelta(days=35)
    #notice_acc_dt = date_obj + timedelta(days=35).strftime('%d/%m/%Y')
    
    # Calculate the CASA excluding Notice Account date ( 1  day from future date)
    casa_excl_notice_acc_dt = date_obj + timedelta(days=1)
    #casa_excl_notice_acc_dt = date_obj + timedelta(days=1).strftime('%d/%m/%Y')
    
    # Calculate the 12month old date ( 365 days or 1 year from date_input)
    one_yr_old_dt = date_obj - timedelta(days=365)
    #one_yr_old_dt = date_obj - timedelta(days=365).strftime('%d/%m/%Y')
    
    # Format the dates for display
    #current_date = date_obj.strftime('%d/%m/%Y')
    #future_date = future_date.strftime('%d/%m/%Y')
    #notice_acc_dt = notice_acc_dt.strftime('%d/%m/%Y')
    #casa_excl_notice_acc_dt = casa_excl_notice_acc_dt.strftime('%d/%m/%Y')
    #one_yr_old_dt = one_yr_old_dt.strftime('%d/%m/%Y')
    
except ValueError:
    print("Invalid date format. Please enter a date in dd/mm/yyyy format.")'''

print('Future Date: ', future_date, '\n')
print('Notice account Date: ', notice_acc_dt, '\n')
print('CASA Date: ', casa_excl_notice_acc_dt, '\n')
print('1Year Old Date: ', one_yr_old_dt, '\n')

#future_date
#exch_usd.info()

# # Security File Start

#df_security.info()
#df_security.isnull().sum()
#df_security

# Update Investment Type for BPI France And KFW and KONOCKR
def update_security_data(df_security):
    # Define the conditions for BPIFSEC
    condition_1_bpi = (df_security['Issuer Id'] == 'BPIFSEC') & (df_security['Product Type'] == 70013016)
    condition_2_bpi = (df_security['Investment Type'] == 'GOVERNMENT SECURITIES') & (df_security['HQ Indicator'] == 'I')

    # Apply the conditions for BPIFSEC
    df_security.loc[condition_1_bpi, ['Investment Type', 'HQ Indicator']] = df_security.loc[condition_1_bpi, ['Investment Type', 'HQ Indicator']].fillna('')  # Fill NaN values with empty strings

    df_security.loc[condition_1_bpi & ~condition_2_bpi, 'Investment Type'] = 'GOVERNMENT SECURITIES'
    df_security.loc[condition_1_bpi & ~condition_2_bpi, 'HQ Indicator'] = 'I'

    # Define the conditions for KFW
    condition_1_kfw = (df_security['Issuer Id'] == 'KFW') & (df_security['Product Type'] == 70013018)
    condition_2_kfw = (df_security['Investment Type'] == 'GOVERNMENT SECURITIES') & (df_security['HQ Indicator'] == 'I')

    # Apply the conditions for KFW
    df_security.loc[condition_1_kfw, ['Investment Type', 'HQ Indicator']] = df_security.loc[condition_1_kfw, ['Investment Type', 'HQ Indicator']].fillna('')  # Fill NaN values with empty strings

    df_security.loc[condition_1_kfw & ~condition_2_kfw, 'Investment Type'] = 'GOVERNMENT SECURITIES'
    df_security.loc[condition_1_kfw & ~condition_2_kfw, 'HQ Indicator'] = 'I'
    
    # Define the conditions for KONOCKR
    condition_1_knockr = (df_security['Issuer Id'] == 'KONOCKR') & (df_security['Product Type'] == 70013020)
    condition_2_knockr = (df_security['Investment Type'] == 'DEBENTURES & BONDS') & (df_security['HQ Indicator'] == 'II A')

    # Apply the conditions for KONOCKR
    df_security.loc[condition_1_knockr, ['Investment Type', 'HQ Indicator']] = df_security.loc[condition_1_knockr, ['Investment Type', 'HQ Indicator']].fillna('')  # Fill NaN values with empty strings

    df_security.loc[condition_1_knockr & ~condition_2_knockr, 'Investment Type'] = 'DEBENTURES & BONDS'
    df_security.loc[condition_1_knockr & ~condition_2_knockr, 'HQ Indicator'] = 'II A'

    return df_security

# Call the function to update the DataFrame
df_security = update_security_data(df_security)

# Display the updated DataFrame
#df_security

# Merge the DataFrames on 'Currency Code'
df_security = df_security.merge(df_exch, left_on='Currency', right_on='Currency Code', how='left')

# Save a copy of the data frame df_securoty
#df_security.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/August ReDIS Files/03.10.2023/security_test.csv', index = False)

# Drop the column Currency code
columns_to_drop = ['Currency Code']
df_security = df_security.drop(columns=columns_to_drop)

#df_security.columns

# Check the ['Maturity Date'] column 
#df_security['Maturity Date']

# Change the dtype of the column Maturity Date to datentime format
#df_security['Maturity Date'] = pd.to_datetime(df_security['Maturity Date'])

# Change the dtype of the column 'Maturity Date' to datetime format with explicit date format
df_security['Maturity Date'] = pd.to_datetime(df_security['Maturity Date'], format='%d-%m-%Y')

# Column Maturity Date after converting the format to date and time
#df_security['Maturity Date']

# Format the dates in 'Maturity Date' column as 'dd/mm/yyyy'
df_security['Maturity Date'] = df_security['Maturity Date'].dt.strftime('%d/%m/%Y')

#print(type(df_security['Maturity Date']), '\n')
#print(df_security['Maturity Date'].dtype, '\n')
#print(df_security['Maturity Date'],'\n')
#print(df_security.columns)

# Column MTM GBP
df_security['MTM GBP'] = (df_security['Market Value CCY']/ df_security['Exchange Rate']).round(2)

df_security['MTM GBP'].sum().round(2)

# Column Principle GBP
df_security['Principle GBP'] = (df_security['Nominal Amount CCY']/df_security['Exchange Rate']).round(2)
df_security['Principle GBP'].sum()

''' # Logic for the column MTM GBP
def mtm_gbp(row):
    if row['Currency'] == 'USD':
        return round (row['Market Value CCY'] / exch_usd['Exchange Rate'], 2)
    else:
        return round (row['Market Value CCY']/exch_GBP['Exchange Rate'], 2)
        
# Apply the conversion function (mtm_gbp) to create a new column
df_security['MTM GBP'] = df_security.apply(mtm_gbp, axis=1)

# Logic for the column Principle GBP
    def principle_gbp (row):
    if row['Currency'] == 'USD':
        return round(row['Principal Amount CCY']/ exch_usd['Exchange Rate'], 2)
    elif row['Currency'] == 'EUR':
        return round(row['Principal Amount CCY']/ exch_EUR['Exchange Rate'], 2)
    else:
        return round( row['Principal Amount CCY']/ exch_GBP['Exchange Rate'], 2)
# Apply the conversion function (principle_gbp) to create a new column
df_security['Principle GBP'] = df_security.apply(principle_gbp, axis=1)

# Columns to convert to datetime
date_columns = ['Start Date', 'Maturity Date', 'Settlement Date']
# Convert columns to datetime format
for column in date_columns:
    df_security[column] = pd.to_datetime(df_security[column])'''

#type(future_date)

# Column LCR row

def lcr_row_logic(row):
        if row['Investment Type'] == 'GOVERNMENT SECURITIES' and row['HQ Indicator'] == 'I':
            return "C72070"
        elif row['MDB Flag'] == 'Y' and row['HQ Indicator'] == 'I':
            return "C72120"
        elif row['HQ Indicator'] == 'II A':
            return "C72250"
        elif row['HQ Indicator'] == 'II B' and row['Maturity Date'] > future_date:
            return "C72360"
        elif row['HQ Indicator'] == 'II B' and row['Maturity Date'] < future_date:
            return "C74190"
        elif pd.isnull(row['HQ Indicator']) and row['Maturity Date'] < future_date:
            return "C74190"
        else:
            return "0"
# Apply the logic to create the 'LCR row' column
df_security['LCR row'] = df_security.apply(lcr_row_logic, axis=1)

#df_security['LCR row'].value_counts()

a = df_security[df_security['LCR row']== '0']
#a

# Column MDB Flag
#Get the values from column MDB Flag to MDB Flag2
df_security['MDB FLAG2'] = df_security['MDB Flag']

#df_security['MDB FLAG2'].value_counts()
#df_security['Exchange Rate']

'''# Column Check
# Create a new column 'Check' based on conditions
df_security['Check'] = np.where(df_security['Currency'] == 'USD', df_security['Market Value CCY'], 
                                  df_security['Market Value CCY'] * df_security['Exchange Rate'])

# Round the values in the 'MTM USD' column to 2 decimals
df_security['Check'] = df_security['Check'].round(2)'''

# Column Check
def check_logic(row):
    if row['Currency'] == 'USD':
        return round (row['Market Value CCY'], 2)
    else:
        return (row['Market Value CCY'] * exch_usd['Exchange Rate']).round(2)
# Apply the conversion function (mtm_gbp) to create a new column
df_security['Check'] = df_security.apply(check_logic, axis=1)
#df_security['Check'].sum()

# Column Difference
df_security['Difference'] = df_security['MTM'] - df_security['Check']

# Round down the column upto 2 decimal points
df_security['Difference'] = df_security['Difference'].round(2)

#df_security['Difference'].sum()

# Column Book Vlaue GBP
def book_value_gbp(row):
    if row['Trading Indicator'] == 'AFS':
        return row ['MTM GBP']
    elif row['Currency'] == 'USD':
        return round ((row['Principal Amount CCY'] / row['Exchange Rate']), 2)
    elif row['Currency'] == 'EUR':
        return round ((row['Principal Amount CCY']/row['Exchange Rate']), 2)
    elif row['Currency'] == 'GBP':
        return round(row['Principal Amount CCY'], 2)

# Apply the conversion function (book_value_gbp) to create a new column
df_security['Book Value GBP'] = df_security.apply(book_value_gbp, axis=1)
#df_security['Book Value GBP'].sum()

# Create a dictionary mapping Customer Identifier to Residency Code
cust_residency_mapping = df_cust.set_index('Customer Identifier')['Residency Code'].to_dict()

# Create a new column 'Residence Code' in df_security and fill it based on the mapping
df_security['Residence Code'] = df_security['Issuer Id'].map(cust_residency_mapping)

# If there's no match, fill the 'Residence Code' column with a default value (e.g., 'N/A')
df_security['Residence Code'].fillna('N/A', inplace=True)
#df_security['Residence Code'].value_counts()

# Column Nominal GBP
df_security['Nominal GBP'] = df_security['Nominal Amount CCY']/df_security['Exchange Rate']

df_security['Nominal GBP'] = df_security['Nominal GBP'].round(2)
#df_security['Nominal GBP'].sum()

# Create empty columns Days to Maturity, PRA Time Bucket,
df_security['Days to Maturity'] = 'COLUMN NOT REQUIRED'
df_security['PRA Time Bucket'] = 'COLUMN NOT REQUIRED'

# Create a dictionary to map 'Issuer Id' to 'LEI CODE' from df_cust
lei_mapping = df_cust.set_index('Customer Identifier')['LEI CODE'].to_dict()

# Create a new column 'dfgh' in df_security by mapping 'Issuer Id' to 'LEI CODE'
df_security['LEI'] = df_security['Issuer Id'].map(lei_mapping)

# Fill any NaN values in the 'dfgh' column with empty cells
df_security['LEI'].fillna('', inplace=True)
#df_security['LEI'].value_counts()

# Create another empty column CB Eligibility
df_security['CB Eligibility'] = 'COLUMN NOT REQUIRED'

col_to_drop = ['Exchange Rate']
df_security = df_security.drop(columns=col_to_drop)
#df_security.columns
#df_security


# # FSCS File Start

#print('Columns in FSCS file: ''\n',df_FSCSSUB.columns)
#print('\n \n \n \n Columns in Cust file: ''\n',df_cust.columns)

# Create a copy of the cust file
df_cust_copy = df_cust.copy()

#Drop all the irrelevant columns from copy of the cust file 
cols_to_drop = ['Branch id', 'Customer Name','Risk Code', 'Incorporation Code','Connected Code', 'Deposit Guarantee Scheme',
                'Development Bank Indicator', 'Parent Code','Ultimate Parent Code','Short Term Credit Rating Code - Moody',
                'Long Term Credit Rating Code - Moody','Short Term Credit Rating Code - Fitch',
                'Long Term Credit Rating Code - Fitch','Short Term Credit Rating Code - S and P',
                'Long Term Credit Rating Code - S and P','Sovereign Risk Transfer Indicator', 'LECustGroup', 'LECustomerCode',
                'LEGroupCode', 'LEI CODE' ]

df_cust_copy = df_cust_copy.drop(columns=cols_to_drop)
#df_cust_copy.info()

# Convert the column Customer Identifier of data customer data frame to string format
df_cust_copy['Customer Identifier'] = df_cust_copy['Customer Identifier'].astype(str)
df_cust_copy.sort_values(by='Customer Identifier', ascending=True, inplace=True)

#print(df_FSCSSUB.info())
#print(df_FSCSSUB.columns)

# ADD THE FEATURE ENGINEERING COLUMNS TO THE DATA FRAME 
# Add all the necessary column to be derived in the FSCS file
# List of columns to add as empty columns
empty_columns = [
    'Balance BILR','Maturity Date LCR','RETAIL FLAG','SME/NON-SME FLAG','RESIDENT CODE','BALANCE>£880,000',
    'Uncovered Amount GBP','Sector Code','Notice Account Y/N','fiancial customer Y/N','LCR PERIOD','INDIVIDUAL + SME',
    'TOTAL CUSTOMER Balance LESS than £440K','NON-UK Resident','NON-GBP Denominated','FSCS COVERED','OLDER than 12 Months ACTIVE Relationship','VLOOKUP TOTAL','CAT_1','LIAB_OUTF_CAT','NON-RETAIL DEPOSIT ROW250','Covered Amt CCY',
    'Uncovered amt CCY','Stable Deposit covered Balance','Stable Deposit Balance CCY','Uncovered amt CCY2','CAA+SBA+TDA Split',
    'NSFR CATEGORY']

# Add empty columns to the DataFrame
for column in empty_columns:
    df_FSCSSUB[column] = pd.Series(dtype='object')

#print('FSCS columns: ', df_FSCSSUB.columns)
#print('FSCS information: \n', df_FSCSSUB.info())

#Check for the null values in all the columns in data set
#df_FSCSSUB.isnull().sum()
#df_FSCSSUB.info()
#df_FSCSSUB['Current Date']

# Add the addtional date columns from the date section above

# Add Column for the Future Date
df_FSCSSUB['Current Date'] = date_obj

# Add Column for the Future Date
df_FSCSSUB['Future Date'] = future_date

# Add Column for the Notice Account Date
df_FSCSSUB['Notice Account Date'] = notice_acc_dt

# Add Column for the CASA Date
df_FSCSSUB['CASA Date'] = casa_excl_notice_acc_dt

# Add Column for 1Yr Old Date
df_FSCSSUB['1YR OLD Date'] = one_yr_old_dt

#df_FSCSSUB.head(15)
#df_FSCSSUB['START_DATE'].head(2)
#df_FSCSSUB['1YR OLD Date'].head(3)


# ## Handling of Date Columns in FSCS File

# Convert column START_DATE to datetime keeping to its original format
df_FSCSSUB['START_DATE'] = pd.to_datetime(df_FSCSSUB['START_DATE'], format='%d-%m-%Y')

# Convert column MATURITY_DATE to datetime keeping to its original format
df_FSCSSUB['MATURITY_DATE'] = pd.to_datetime(df_FSCSSUB['MATURITY_DATE'], format='%d-%m-%Y')

# Convert column Maturity Date LCR to datetime keeping to its original format
df_FSCSSUB['Maturity Date LCR'] = pd.to_datetime(df_FSCSSUB['Maturity Date LCR'], format='%d-%m-%Y')

# Convert column Future Date  to datetime keeping to its original format
df_FSCSSUB['Future Date'] = pd.to_datetime(df_FSCSSUB['Future Date'], format='%d/%m/%Y')

# Convert column Notice Account Date to datetime keeping to its original format
df_FSCSSUB['Notice Account Date'] = pd.to_datetime(df_FSCSSUB['Notice Account Date'], format='%d/%m/%Y')

# Convert column CASA Date to datetime keeping to its original format
df_FSCSSUB['CASA Date'] = pd.to_datetime(df_FSCSSUB['CASA Date'], format='%d/%m/%Y')

# Convert column 1YR OLD Date to datetime keeping to its original format
df_FSCSSUB['1YR OLD Date'] = pd.to_datetime(df_FSCSSUB['1YR OLD Date'], format='%d/%m/%Y')

# Convert column Current Date to datetime keeping to its original format
df_FSCSSUB['Current Date'] = pd.to_datetime(df_FSCSSUB['Current Date'], format='%d/%m/%Y')

'''# Columns to convert to datetime format

date_columns = ['START_DATE', 'MATURITY_DATE','Maturity Date LCR']
# Convert columns to datetime format
for column in date_columns:
    df_FSCSSUB[column] = pd.to_datetime(df_FSCSSUB[column], format='%d-%m-%Y')

# If the date columns are not already in the dd/mm/yyyy format, you can use the strftime method
#df_FSCSSUB['START_DATE'] = df_FSCSSUB['START_DATE'].dt.strftime('%d/%m/%Y')
#df_FSCSSUB['MATURITY_DATE'] = df_FSCSSUB['MATURITY_DATE'].dt.strftime('%d/%m/%Y')'''

# Column START_DATE after converting to date and time format
df_FSCSSUB[['START_DATE', 'MATURITY_DATE', 'Maturity Date LCR', 'Future Date','Notice Account Date','CASA Date','1YR OLD Date','Current Date']].head()

'''# Format the dates in 'START_DATE','MATURITY_DATE','Maturity Date LCR'  column as 'dd/mm/yyyy'
df_FSCSSUB['START_DATE'] = df_FSCSSUB['START_DATE'].dt.strftime('%d/%m/%Y')
df_FSCSSUB['MATURITY_DATE'] = df_FSCSSUB['MATURITY_DATE'].dt.strftime('%d/%m/%Y')
df_FSCSSUB['Maturity Date LCR'] = df_FSCSSUB['Maturity Date LCR'].dt.strftime('%d/%m/%Y')
# Column START_DATE after converting to dd/mm/yyyy format
df_FSCSSUB['START_DATE']'''

#df_FSCSSUB['Maturity Date LCR'].info()
#df_FSCSSUB['MATURITY_DATE'].head(15)
#df_FSCSSUB.info()
#df_cust_copy.isnull().sum()

df_FSCSSUB = df_FSCSSUB[df_FSCSSUB['ORIG_BAL_BEFORE_INT'] > 0]

columns_to_drop = ['ACC_BALANCE_IN_ACCT_CRNCY']
df_FSCSSUB = df_FSCSSUB.drop(columns=columns_to_drop)

# Drop the 'ORIG_BAL_BEFORE_INT' column
# The pop () method is used to remove the 'ORIG_BAL_BEFORE_INT' column from the DataFrame and save its values.
orig_bal_before_int = df_FSCSSUB.pop('ORIG_BAL_BEFORE_INT')

# Insert the 'ORIG_BAL_BEFORE_INT' column after the 'ACC_CRNCY' column
df_FSCSSUB.insert(df_FSCSSUB.columns.get_loc('ACC_CRNCY') + 1, 'ORIG_BAL_BEFORE_INT', orig_bal_before_int)

# Rename the column name from ORIG_BAL_BEFORE_INT to Amount in Acct Crncy
df_FSCSSUB.rename(columns={'ORIG_BAL_BEFORE_INT': 'Amount in Acct Crncy', 'ACC_CRNCY': 'Currency', 'TRANSFERABLE_ELIG_DEPOSIT': 'Covered Amount GBP' }, inplace=True)

# Merge the DataFrames on 'Currency Code'
df_FSCSSUB = df_FSCSSUB.merge(df_exch, left_on='Currency', right_on='Currency Code', how='left')

columns_to_drop = ['Currency Code']
df_FSCSSUB = df_FSCSSUB.drop(columns=columns_to_drop)
#df_FSCSSUB.info()

df_FSCSSUB['Balance BILR'] = df_FSCSSUB['Amount in Acct Crncy']/df_FSCSSUB['Exchange Rate']
df_FSCSSUB['Balance BILR'] = df_FSCSSUB['Balance BILR'].round(2)
#df_FSCSSUB['Balance BILR'].sum()

''' # Logic function for col Balance BILR
def bal_bilr_logic(row):
    if row['Currency'] == 'GBP':
        return round(row['Amount in Acct Crncy'], 2)
    elif row['Currency'] == 'EUR':
        return round(row['Amount in Acct Crncy']/ exch_EUR['Exchange Rate'], 2).values[0]
    elif row['Currency'] == 'USD':
        return round(row['Amount in Acct Crncy']/ exch_usd['Exchange Rate'], 2).values[0]
    elif row['Currency'] == 'CAD':
        return round(row['Amount in Acct Crncy']/ exch_CAD['Exchange Rate'], 2).values[0]
        
# Assign the logic (bal_bilr_logic) to the column Balance BILR
df_FSCSSUB['Balance BILR'] = df_FSCSSUB.apply(bal_bilr_logic, axis=1)    
'''

# Check the uniquevalues in column Product_type
#df_FSCSSUB['PRODUCT_TYPE'].unique()

# Column Maturity Date LCR
conditions = [
    (
        (df_FSCSSUB['PRODUCT_TYPE'].isin(['FD1', 'FD2', 'FD4', 'FP4P', 'ISA'])) &
        (~df_FSCSSUB['SCHM_CODE'].isin(['SBF0M', 'SBF0Y', 'SB163'])) &
        (df_FSCSSUB['MATURITY_DATE'] > df_FSCSSUB['Current Date'])
    ),
    df_FSCSSUB['SCHM_CODE'].isin(['NSB15', 'NSB16', 'NSB17'])
]

choices = [
    df_FSCSSUB['MATURITY_DATE'],
    df_FSCSSUB['Notice Account Date']
]

# Use numpy.select to apply conditions and choices
df_FSCSSUB['Maturity Date LCR'] = np.select(conditions, choices, default=df_FSCSSUB['CASA Date'])
#df_FSCSSUB['Maturity Date LCR'].isnull().sum()

'''conditions = [
    df_FSCSSUB['PRODUCT_TYPE'].isin(['FD1', 'FD2', 'FD4', 'FP4P', 'ISA']) &
    ~df_FSCSSUB['SCHM_CODE'].isin(['SBF0M', 'SBF0Y', 'SB163']),
    df_FSCSSUB['SCHM_CODE'].isin(['NSB15', 'NSB16', 'NSB17'])
]

choices = [
    df_FSCSSUB['MATURITY_DATE'],
    df_FSCSSUB['Notice Account Date']
    
]

default_selection = [df_FSCSSUB['CASA Date']]

# Use numpy.select to apply conditions and choices
df_FSCSSUB['Maturity Date LCR'] = np.select(conditions, choices, default=casa_excl_notice_acc_dt)'''
#df_FSCSSUB['MATURITY_DATE'].head(15)
#df_FSCSSUB['Maturity Date LCR'].head(15)
#df_cust_copy.head()

# Save the fscs file to csv format check the progress 

#df_FSCSSUB.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/August ReDIS Files/03.10.2023/fscs_test.csv', index = False)

# Print the columns of the cust file

#df_cust.columns
#df_cust_copy.info()

# Change the dtype of the customer Id columns in both dataframes

df_cust_copy['Customer Identifier'] = df_cust_copy['Customer Identifier'].astype(str)
df_FSCSSUB['CUSTOMER_ID'] = df_FSCSSUB['CUSTOMER_ID'].astype(str)

#print(df_FSCSSUB.shape)
#print(df_cust_copy.shape)

# Merge the df_cust_copy and df_FSCSSUB file based on the customer ID column

df_FSCSSUB = df_FSCSSUB.merge(df_cust_copy, left_on='CUSTOMER_ID', right_on='Customer Identifier', how='left')

# drop the column Customer Identifier as it is no loger needed
df_FSCSSUB = df_FSCSSUB.drop(['Customer Identifier'], axis =1)

# check for the null values after the merging of 2 dataframes
#df_FSCSSUB.isnull().sum()

# Replace missing values in column SME Flag with 0
df_FSCSSUB['SME Flag'].fillna(0, inplace=True)

#df_FSCSSUB.columns

# Column RETAIL FLAG
df_FSCSSUB['RETAIL FLAG'] = df_FSCSSUB['Retail Flag']
#df_FSCSSUB['RETAIL FLAG'].value_counts()

# Column SME/NON-SME FLAG
df_FSCSSUB['SME/NON-SME FLAG'] = df_FSCSSUB['SME Flag']
#df_FSCSSUB['SME/NON-SME FLAG'].value_counts()

'''# logic for RETAIL FLAG column
df_FSCSSUB['RETAIL FLAG'] = df_FSCSSUB['CUSTOMER_ID'].map(df_cust_copy.set_index('Customer Identifier')['Residency Code'])
df_FSCSSUB['RETAIL FLAG'].fillna(0, inplace=True)'''

''' # Logic for SME/NON-SME FLAG
df_FSCSSUB['SME/NON-SME FLAG'] = df_FSCSSUB['CUSTOMER_ID'].map(df_cust.set_index('Customer Identifier')['SME Flag'])
df_FSCSSUB['SME/NON-SME FLAG'].fillna(0, inplace=True)'''

# Column RESIDENT CODE

df_FSCSSUB['RESIDENT CODE'] = df_FSCSSUB['Residency Code']

# Replace missing values in 'SME/NON-SME FLAG' with 0
#df_FSCSSUB['RESIDENT CODE'].fillna('NA', inplace=True)

'''df_FSCSSUB['RESIDENT CODE'] = df_FSCSSUB['CUSTOMER_ID'].map(df_cust_copy.set_index('Customer Identifier')['Residency Code'])
#df_FSCSSUB['RESIDENT CODE'].fillna(0, inplace=True)'''

'''# Create a dictionary mapping Customer Identifier to Residency Code
cust_identifier_mapping = df_cust.set_index('Customer Identifier')['Residency Code'].to_dict()

# Create a new column 'cvbn' in df_FSCSSUB and fill it based on the mapping
df_FSCSSUB['RESIDENT CODE'] = df_FSCSSUB['CUSTOMER_ID'].map(cust_identifier_mapping)

# If there's no match, fill the 'cvbn' column with a default value (e.g., 'N/A')
df_FSCSSUB['RESIDENT CODE'].fillna('N/A', inplace=True)'''

'''# Merge df_FSCSSUB and df_cust on 'CUSTOMER_ID' and 'Customer Identifier' columns
merged = df_FSCSSUB.merge(df_cust[['Customer Identifier', 'Residency Code']], left_on='CUSTOMER_ID', right_on='Customer Identifier', how='left')

# Rename the 'Residency Code' column to 'cvbn'
merged.rename(columns={'Residency Code': 'cvbn'}, inplace=True)'''

'''
# Extract the 'RESIDENT CODE' column from the merged DataFrame
df_FSCSSUB['RESIDENT CODE'] = merged_df['Residency Code']
# Replace missing values in 'RESIDENT CODE' with 0
#df_FSCSSUB['RESIDENT CODE'].fillna(0, inplace=True)
'''

# ## Creation of DepositBalChk

# Group the DataFrame by 'CUSTOMER_ID' and calculate the sum of 'Balance BILR'
DepositBalChk = df_FSCSSUB.groupby('CUSTOMER_ID')['Balance BILR'].sum().reset_index()

#DepositBalChk
#df_cust_copy.columns

# Merge the df_FSCSSUB and DepositBalChk file based on the customer ID column

df_FSCSSUB = df_FSCSSUB.merge(DepositBalChk, left_on='CUSTOMER_ID', right_on='CUSTOMER_ID', how='left')
#df_FSCSSUB.columns

# Rename the columns in the FSCSSUB file to the actual names of the columns
df_FSCSSUB.rename(columns={'Balance BILR_x': 'Balance BILR',
                           'Balance BILR_y': 'Deposit_BAL_Balance_BILR', 
                           'Sector Code_x': 'Sector Code' }, inplace=True)

'''# Merge the two DataFrames on 'CUSTOMER_ID'
deposit_bal_merged_df = df_FSCSSUB.merge(DepositBalChk, on='CUSTOMER_ID', how='left')
deposit_bal_merged_df.columns'''

# Column BALANCE>£880,000

condition = (df_FSCSSUB['Deposit_BAL_Balance_BILR'] > 880000)
df_FSCSSUB['BALANCE>£880,000'] = np.where(condition, 'YES', 'NO')

'''# Define the condition and create a new column 'Result'
condition = (deposit_bal_merged_df['Balance BILR_y'] > 880000)
df_FSCSSUB['BALANCE>£880,000'] = np.where(condition, 'YES', 'NO')'''

#df_FSCSSUB['BALANCE>£880,000'].value_counts()
#df_FSCSSUB.info()

# Column Uncovered Amount GBP

df_FSCSSUB['Uncovered Amount GBP'] = np.where(df_FSCSSUB['Balance BILR'] > df_FSCSSUB['Covered Amount GBP'], df_FSCSSUB['Balance BILR'] - df_FSCSSUB['Covered Amount GBP'], 0)

#Round down the column Uncovered Amount GBP pto 2 decimal points
df_FSCSSUB['Uncovered Amount GBP'] = df_FSCSSUB['Uncovered Amount GBP'].round(2)

#df_FSCSSUB['Uncovered Amount GBP'].sum()

'''# Column Uncovered Amount GBP
def uncovered_amt_gbp(row):
    if row ['Balance BILR'] > row ['Covered Amount GBP']:
        return round ( row ['Balance BILR'] - row['Covered Amount GBP'], 2)
    else:
        return "0"

# Assign the logic (uncovered_amt_gbp) to the column Balance BILR
df_FSCSSUB['Uncovered Amount GBP'] = df_FSCSSUB.apply(uncovered_amt_gbp, axis=1)'''

# Column Sector Code
df_FSCSSUB['Sector Code'] = df_FSCSSUB['Sector Code_y']

# Replace missing values in 'SME/NON-SME FLAG' with 0
df_FSCSSUB['Sector Code'].fillna(0, inplace=True)

# Convert 'Sector Code' column to integer type
df_FSCSSUB['Sector Code'] = df_FSCSSUB['Sector Code'].astype(int)
#df_FSCSSUB['Sector Code'].value_counts()

# condition for notice_Account check
notice_Account_condition = df_FSCSSUB['SCHM_CODE'].isin(['NSB15', 'NSB16', 'NSB17'])

# Create a new column 'Result' based on the condition
df_FSCSSUB['Notice Account Y/N'] = np.where(notice_Account_condition, 'YES', 'NO')

#df_FSCSSUB['Notice Account Y/N'].value_counts()

# Column fiancial customer Y/N

# condition for fiancial_customer check
fiancial_customer_condition = (
    (df_FSCSSUB['Sector Code'] == 3350)|
    (df_FSCSSUB['Sector Code'] == 3370)|
    (df_FSCSSUB['Sector Code'] == 4539)
    
)

# Create a new column 'Result' based on the condition
df_FSCSSUB['fiancial customer Y/N'] = np.where(fiancial_customer_condition, 'YES', 'NO')
#df_FSCSSUB['fiancial customer Y/N'].value_counts()

'''
# Define your condition
condition = df_FSCSSUB['Sector Code'].isin(['3350', '3370', '4539'])

# Create a new column 'Result' based on the condition
df_FSCSSUB['fiancial customer Y/N'] = np.where(condition, 'YES', 'NO')
'''
#df_FSCSSUB['Maturity Date LCR'].info()
#df_FSCSSUB['Future Date'].info()

'''# Convert 'Maturity Date LCR' column to datetime format
df_FSCSSUB['Maturity Date LCR'] = pd.to_datetime(df_FSCSSUB['Maturity Date LCR'])'''
#df_FSCSSUB['Maturity Date LCR']

# Column LCR PERIOD
def lcr_period_logic(row):
    if row ['Maturity Date LCR'] <= row['Future Date']:
        return "YES"
    else:
        return "NO"

# Assign the logic (lcr_period_logic) to the column Balance BILR
df_FSCSSUB['LCR PERIOD'] = df_FSCSSUB.apply(lcr_period_logic, axis=1)
#df_FSCSSUB['LCR PERIOD'].value_counts()

# Column INDIVIDUAL + SME

df_FSCSSUB['INDIVIDUAL + SME'] = np.where(
    (
        (df_FSCSSUB['SME/NON-SME FLAG'] == "SME") &
        (df_FSCSSUB['BALANCE>£880,000'] == "YES") &
        (~df_FSCSSUB['CUSTOMER_ID'].isin(['800002460', '800003329', '800003440']))  # the numbers are put into '' as the column is not in a integer format
    ) |
    (df_FSCSSUB['SME/NON-SME FLAG'] == "NON-SME"),
    "NO",
    "YES"
)

#df_FSCSSUB['INDIVIDUAL + SME'].value_counts()
#df_FSCSSUB['INDIVIDUAL + SME'].value_counts()

'''# Column INDIVIDUAL + SME
def individualor_SME_flag(row):
    if (
        (row['SME/NON-SME FLAG'] == 'SME' and
        row['BALANCE>£880,000'] not in [800002460, 800003329, 800003440]) or
        (row['SME/NON-SME FLAG'] == 'NON-SME')
    ):
        return "NO"
    else:
        return "YES"

# Apply the logic to create a new column 'LCR Flag'
df_FSCSSUB['INDIVIDUAL + SME'] = df_FSCSSUB.apply(individualor_SME_flag, axis=1)'''
#df_FSCSSUB.info()

# Column TOTAL CUSTOMER Balance LESS than £440K

def check_conditions(row):
    if row['fiancial customer Y/N'] == 'NO' and row['INDIVIDUAL + SME'] == 'YES':
        if row['Deposit_BAL_Balance_BILR'] < 440000:
            return 'YES'
        else:
            return 'NO'
    return ''

# Apply the function to create the 'bv' column
df_FSCSSUB['TOTAL CUSTOMER Balance LESS than £440K'] = df_FSCSSUB.apply(check_conditions, axis=1)

#df_FSCSSUB['TOTAL CUSTOMER Balance LESS than £440K'].value_counts()

# Column NON-UK Resident

def non_uk_resident(row):
    if row['fiancial customer Y/N'] == 'NO' and row['INDIVIDUAL + SME'] == 'YES':
        if row['RESIDENT CODE'] == 'GB':
            return "NO"  # Apply Logic1 when conditions are met
        else:
            return "YES"  # Apply Logic1 when conditions are met
    
    return ""  # Apply Logic2 when conditions are not met

# Apply the logic to create a new column 
df_FSCSSUB['NON-UK Resident'] = df_FSCSSUB.apply(non_uk_resident, axis=1)
#df_FSCSSUB['NON-UK Resident'].value_counts()

# Column NON-GBP Denominated
def non_gbp_denomination (row):
    if row['fiancial customer Y/N'] == 'NO' and row['INDIVIDUAL + SME'] == 'YES':
        if row['Currency'] == 'GBP':
            return "NO"
        else:
            return "YES"
    return ""

# Apply the logic to create a new column NON-GBP Denominated
df_FSCSSUB['NON-GBP Denominated'] = df_FSCSSUB.apply(non_gbp_denomination, axis=1)
#df_FSCSSUB['NON-GBP Denominated'].value_counts()
#df_FSCSSUB.columns

# Column FSCS COVERED
def fscs_cover(row):
    if row['fiancial customer Y/N'] == 'NO' and row['TOTAL CUSTOMER Balance LESS than £440K'] == 'YES':
        if row['COVERED_IND'] == 'YES':
            return "YES"
        else:
            return "NO"
    return ""

# Apply the logic to create a new column NON-GBP Denominated
df_FSCSSUB['FSCS COVERED'] = df_FSCSSUB.apply(fscs_cover, axis=1)

#df_FSCSSUB['FSCS COVERED'].value_counts()

#check from below for the actiev relationship column
#df_FSCSSUB.columns
#df_FSCSSUB['1YR OLD Date'].head(2)
#df_FSCSSUB['START_DATE'].head(2)

# Create a Copy of FSCS data frame 
df1 = df_FSCSSUB.copy()
#df1.columns

df2 = df1[['CUSTOMER_ID','START_DATE', '1YR OLD Date', 'fiancial customer Y/N', 'TOTAL CUSTOMER Balance LESS than £440K']]
#df2.head(5)
#df2.info()

def And_condition(row):
    if  row['fiancial customer Y/N']== 'NO' and row['TOTAL CUSTOMER Balance LESS than £440K']== 'YES':
        return "1"
    else:
        return ""
    
    
# Apply the logic to create a new column OLDER than 12 Months
df2['AND_Condition'] = df2.apply(And_condition, axis=1)
#df2['AND_Condition'].value_counts()

def Date_check_condition(row):
    if  row['START_DATE'] < row['1YR OLD Date']:
        return "1"
    else:
        return "0"
    
    
# Apply the logic to create a new column OLDER than 12 Months
df2['Date_Condition'] = df2.apply(Date_check_condition, axis=1)

#df2['Date_Condition'].value_counts()

def final_condition(row):
    if  row['AND_Condition']== '1':
        return row['Date_Condition']
    else:
        return ""
    
    
# Apply the logic to create a new column OLDER than 12 Months
df2['final_Condition'] = df2.apply(final_condition, axis=1)
#df2['final_Condition'].value_counts()

#df2.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/August ReDIS Files/03.10.2023/df2.csv', index = False)

# Column OLDER than 12 Months ACTIVE Relationship

df_FSCSSUB['OLDER than 12 Months ACTIVE Relationship'] = df2['final_Condition']

# Replace 1 with 'YES' and 0 with 'NO' in the specified column
df_FSCSSUB['OLDER than 12 Months ACTIVE Relationship'] = df_FSCSSUB['OLDER than 12 Months ACTIVE Relationship'].replace({'1': "YES", '0': "NO"})

#type(one_yr_old_dt)

# Check if a column is of string type
column_name = 'START_DATE'
is_string_column = isinstance(df_FSCSSUB[column_name].iloc[0], str)

if is_string_column:
    print(f"{column_name} is a string column.")
else:
    print(f"{column_name} is not a string column.")

#print(df_FSCSSUB['START_DATE'].dtype)
#print(df_FSCSSUB['START_DATE'].isnull().sum())

# Filling the missing values in the column START_DATE with the user input date
df_FSCSSUB['START_DATE'].fillna(date_obj.strftime('%d/%m/%Y'), inplace=True)

'''
# Use this logic if the missing values are to be kept as it is 

# Column OLDER than 12 Months ACTIVE Relationship

# Since we have an null value we go ahead with following code
def active_relationship(row):
    if (
        row['fiancial customer Y/N'] == 'NO' 
        and row['TOTAL CUSTOMER Balance LESS than £440K'] == 'YES'
    ):
        start_date = row['START_DATE']
        if pd.notnull(start_date) and start_date < one_yr_old_dt:
            return "YES"
        else:
            return "NO"
    return ""

# Apply the logic to create a new column NON-GBP Denominated
df_FSCSSUB['OLDER than 12 Months ACTIVE Relationship'] = df_FSCSSUB.apply(active_relationship, axis=1)
# Check for the total counts for all the values in the column
df_FSCSSUB['OLDER than 12 Months ACTIVE Relationship'].value_counts()'''

'''# Column OLDER than 12 Months ACTIVE Relationship
def active_relationship(row):
    if row['fiancial customer Y/N'] == 'NO' and row['TOTAL CUSTOMER Balance LESS than £440K'] == 'YES':
        if row['START_DATE'] < one_yr_old_dt:
            return "YES"
        else:
            return "NO"
    return ""

# Apply the logic to create a new column NON-GBP Denominated
df_FSCSSUB['OLDER than 12 Months ACTIVE Relationship'] = df_FSCSSUB.apply(active_relationship, axis=1)'''

#df_FSCSSUB['OLDER than 12 Months ACTIVE Relationship'].info()
#df_FSCSSUB.columns

# Column VLOOKUP TOTAL
def vlookup_total(row):
    return f"{row['fiancial customer Y/N']}|{row['INDIVIDUAL + SME']}|{row['TOTAL CUSTOMER Balance LESS than £440K']}|{row['NON-UK Resident']}|{row['NON-GBP Denominated']}|{row['FSCS COVERED']}|{row['OLDER than 12 Months ACTIVE Relationship']}"

# Apply logic (combine_columns) to create the column LCR Sheet Lookup
df_FSCSSUB['VLOOKUP TOTAL'] = df_FSCSSUB.apply(vlookup_total, axis=1)
#df_FSCSSUB['VLOOKUP TOTAL'].value_counts()

#Logic_for_Deposit_Outflows_df.head()
#Logic_for_Deposit_Outflows_df.columns
#Logic_for_Deposit_Outflows_df['Category'].value_counts()


# Create a copy of Logic_for_Deposit_Outflows_df 
deposit_outflow_copy = Logic_for_Deposit_Outflows_df.copy()

# Consider only 2 columns 'Vlookup Total', 'Category'
deposit_outflow_copy = deposit_outflow_copy[['Vlookup Total', 'Category']]
# Rename the 'Category' column to 'CAT_1'
deposit_outflow_copy.rename(columns={'Vlookup Total': 'vlookup_total'}, inplace=True)
#deposit_outflow_copy.shape

# check for the duplications in columnn Catefory
#Logic_for_Deposit_Outflows_df['Vlookup Total'].duplicated()

'''# Find and display the exact duplicate values in the 'Vlookup Total' column
duplicate_values = Logic_for_Deposit_Outflows_df[Logic_for_Deposit_Outflows_df['Vlookup Total'].duplicated(keep=False)]
duplicate_values'''

#deposit_outflow_copy['Category'].value_counts()

# Remove all the occurances of the 'NOTICE ACCT', 'HIGHER OUTFLOW CAT 2' from the column CAT_1
deposit_outflow_copy = deposit_outflow_copy[~deposit_outflow_copy['Category'].isin(['NOTICE ACCT'])]

#deposit_outflow_copy.shape
#df_FSCSSUB.columns

# Merge the DataFrames on column 'VLOOKUP TOTAL' and 'vlookup_total'
df_FSCSSUB = df_FSCSSUB.merge(deposit_outflow_copy, left_on='VLOOKUP TOTAL', right_on='vlookup_total', how='left')

# Column CAT_1
df_FSCSSUB['CAT_1'] = df_FSCSSUB['Category']
#df_FSCSSUB['CAT_1'].value_counts()
#df_FSCSSUB['CAT_1'].isnull().sum()

# Column LIAB_OUTF_CAT
def categorize_accounts(row):
    if row['SCHM_CODE'] in ["NSB15", "NSB16", "NSB17"]:
        return "Notice Accounts"
    elif row['CUSTOMER_ID'] in ['800002460', '800003329', '800003440'] and row['SCHM_CODE'] not in ["NSB15", "NSB16", "NSB17"]:
        return "HIGHER OUTFLOW CAT 2"
    elif row['fiancial customer Y/N'] == "NO" and row['INDIVIDUAL + SME'] == "NO":
        return "NON RETAIL Deposit"
    else:
        return row['CAT_1']

# Apply the categorization logic to create a new column 'Category'
df_FSCSSUB['LIAB_OUTF_CAT'] = df_FSCSSUB.apply(categorize_accounts, axis=1)

#df_FSCSSUB['LIAB_OUTF_CAT'].nunique()
#df_FSCSSUB['LIAB_OUTF_CAT'].value_counts()
#df_FSCSSUB.columns

#deposit_bal_merged_df.head()

# Define the condition and create a new column 'Result'
lookup_condition = df_FSCSSUB['Deposit_BAL_Balance_BILR'] <= 85000

#lookup_condition

# Column NON-RETAIL DEPOSIT ROW250
def categorize_rtyu(row):
    if (
        row['LIAB_OUTF_CAT'] == 'NON RETAIL Deposit' and
        row['COVERED_IND'] == 'YES' and
        #(df_FSCSSUB['Deposit_BAL_Balance_BILR'] <= 85000).any()  # Use .any() to check if any row satisfies the condition
        row['Deposit_BAL_Balance_BILR']<= 85000
    ):
        return 'YES'
    else:
        return 'NO'

# Apply the categorization logic to create a new column 'NON-RETAIL DEPOSIT ROW250'
df_FSCSSUB['NON-RETAIL DEPOSIT ROW250'] = df_FSCSSUB.apply(categorize_rtyu, axis=1)
#df_FSCSSUB['NON-RETAIL DEPOSIT ROW250'].value_counts()
#df_FSCSSUB.columns

# Column Covered Amt CCY

df_FSCSSUB['Covered Amt CCY'] = df_FSCSSUB['Covered Amount GBP']*df_FSCSSUB['Exchange Rate']

df_FSCSSUB['Covered Amt CCY'] = df_FSCSSUB['Covered Amt CCY'].round(2)
#df_FSCSSUB['Covered Amt CCY'].sum().round(2)

'''def covered_amt_ccy(row):
    if row['Currency'] == 'USD':
        return round ( row ['Covered Amount GBP'] * exch_usd['Exchange Rate'], 2).values[0]
    elif row['Currency'] == 'EUR':
        return round ( row ['Covered Amount GBP'] * exch_EUR['Exchange Rate'], 2).values[0]
    elif row['Currency'] == 'CAD':
        return round ( row ['Covered Amount GBP'] * exch_CAD['Exchange Rate'], 2).values[0]
    elif row['Currency'] == 'GBP':
        return round ( row['Covered Amount GBP'], 2)
# Apply the covered_amt_ccy logic to create a new column 'Covered Amt CCY'
df_FSCSSUB['Covered Amt CCY'] = df_FSCSSUB.apply(covered_amt_ccy, axis=1)   '''   

# Column Uncovered amt CCY
def uncovered_amt_ccy(row):
    if row['Amount in Acct Crncy'] > row ['Covered Amt CCY']:
        return round ( row['Amount in Acct Crncy'] - row ['Covered Amt CCY'],2)
    else:
        return "0"
# Apply the uncovered_amt_ccy logic to create a new column 'Uncovered amt CCY'
df_FSCSSUB['Uncovered amt CCY'] = df_FSCSSUB.apply(uncovered_amt_ccy, axis=1)   
#df_FSCSSUB['Uncovered amt CCY'].info()

#df_FSCSSUB.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/August ReDIS Files/03.10.2023/trial/a1.csv', index = False)

# Round down to 2 decimal places
df_FSCSSUB['Balance BILR'] = df_FSCSSUB['Balance BILR'].round(2)

#print('Sum of column Balance BILR: ', df_FSCSSUB['Balance BILR'].sum().round(2))
#print('Sum of column Covered Amount GBP: ', df_FSCSSUB['Covered Amount GBP'].sum())

# Column Stable Deposit covered Balance
def stable_deposit_bal_covered(row):
    if row['Balance BILR'] <= row['Covered Amount GBP']:
        return row['Balance BILR']
    else:
        return row['Covered Amount GBP']
    
# Apply the stable_deposit_bal_covered logic to create a new column 'Stable Deposit covered Balance'
df_FSCSSUB['Stable Deposit covered Balance'] = df_FSCSSUB.apply(stable_deposit_bal_covered, axis=1)   

df_FSCSSUB['Stable Deposit covered Balance'].sum().round(2)
#df_FSCSSUB['Currency'].value_counts()


# Column Stable Deposit Balance CCY

df_FSCSSUB['Stable Deposit Balance CCY'] = df_FSCSSUB['Stable Deposit covered Balance'] * df_FSCSSUB['Exchange Rate']

df_FSCSSUB['Stable Deposit Balance CCY'] = df_FSCSSUB['Stable Deposit Balance CCY'].round(2)

#df_FSCSSUB['Stable Deposit Balance CCY'].sum()

'''
# Column Stable Deposit Balance CCY

def stable_deposit_ccy_bal(row):
    if row['Currency'] == 'USD':
        return round( row['Stable Deposit covered Balance'] * exch_usd['Exchange Rate'], 2).values[0]
    elif row['Currency'] == 'EUR':
        return round( row['Stable Deposit covered Balance'] * exch_EUR['Exchange Rate'], 2).values[0]
    elif row['Currency'] == 'CAD':
        return round( row['Stable Deposit covered Balance'] * exch_CAD['Exchange Rate'], 2).values[0]
    elif row['Currency'] == 'GBP':
        return row['Stable Deposit covered Balance']
        
# Apply the stable_deposit_bal_covered logic to create a new column 'Stable Deposit Balance CCY'
df_FSCSSUB['Stable Deposit Balance CCY'] = df_FSCSSUB.apply(stable_deposit_ccy_bal, axis=1)    ''' 

# Column Uncovered amt CCY2

df_FSCSSUB['Uncovered amt CCY2'] = df_FSCSSUB['Amount in Acct Crncy'] - df_FSCSSUB['Stable Deposit Balance CCY']

df_FSCSSUB['Uncovered amt CCY2'] = df_FSCSSUB['Uncovered amt CCY2'].round(2)  # Round the column to be upto 2 decimal points

#df_FSCSSUB['Uncovered amt CCY2'].sum()

# Column CAA+SBA+TDA Split

def categorize_schm_code(code):
    if code.startswith(("CA", "OD")):
        return "CAA"
    elif code.startswith(("SB", "RS", "LA")):
        return "SBA"
    elif code.startswith("TD"):
        return "TDA"
    else:
        return "NSB"

# Apply the categorization function to create a new 'Result' column
df_FSCSSUB['CAA+SBA+TDA Split'] = df_FSCSSUB['SCHM_CODE'].apply(lambda x: categorize_schm_code(x))
#df_FSCSSUB['CAA+SBA+TDA Split'].value_counts()

# Column NSFR CATEGORY

def nsfr_categ(row):
    if row ['LIAB_OUTF_CAT'] == 'STABLE Deposit':
        return "STABLE Deposit"
    elif row ['LIAB_OUTF_CAT'] in ["HIGHER OUTFLOW CAT 1", "HIGHER OUTFLOW CAT 2", "OTHER RETAIL Deposit"]:
        return "OTHER RETAIL Deposit"
    elif row['LIAB_OUTF_CAT'] == "Notice Accounts" and row ['INDIVIDUAL + SME'] == 'YES':
        return "OTHER RETAIL Deposit"
    elif row ['LIAB_OUTF_CAT'] == "FINANCIAL CUSTOMER":
        return "FINANCIAL CUSTOMER"
    elif [row ['LIAB_OUTF_CAT'] == "Notice Accounts" and row ['INDIVIDUAL + SME'] == 'NO'] or row ['LIAB_OUTF_CAT'] == "NON RETAIL Deposit":
        return "NON RETAIL Deposit"


# Apply the categorization function to create a new 'Result' column
df_FSCSSUB['NSFR CATEGORY'] = df_FSCSSUB.apply(nsfr_categ, axis = 1)
#df_FSCSSUB['NSFR CATEGORY'].value_counts()

# Drop the column Exchange Rate as we no longer require the column
col_to_drop = ['Current Date','Future Date', 'Notice Account Date', 'CASA Date',
       '1YR OLD Date', 'Exchange Rate', 'Residency Code', 'Sector Code_y',
       'Industry Code', 'Retail Flag', 'SME Flag', 'Deposit_BAL_Balance_BILR',
       'vlookup_total', 'Category']
df_FSCSSUB = df_FSCSSUB.drop(columns=col_to_drop)

#df_FSCSSUB.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/December Validation/December-23/fscs_test_29thFeb1.csv', index = False)


# # FX Deals

# Add an empty column without any header to the data set
df_fxdeals[''] = ''

# Merge the DataFrames on column 'FX Sold Currency'
df_fxdeals = df_fxdeals.merge(df_exch, left_on='FX Sold Currency', right_on='Currency Code', how='left')

# Drop the unwanted column 'Currency Code'
df_fxdeals = df_fxdeals.drop(['Currency Code'], axis =1 )

# Rename the column name from ORIG_BAL_BEFORE_INT to Amount in Acct Crncy
df_fxdeals.rename(columns={'Exchange Rate': 'FX Sold Exchange Rate'}, inplace=True)

# Merge the DataFrames on column 'FX Bought Currency'
df_fxdeals = df_fxdeals.merge(df_exch, left_on='FX Bought Currency', right_on='Currency Code', how='left')

# Drop the unwanted column 'Currency Code'
df_fxdeals = df_fxdeals.drop(['Currency Code'], axis =1 )

# Rename the column name from ORIG_BAL_BEFORE_INT to Amount in Acct Crncy
df_fxdeals.rename(columns={'Exchange Rate': 'FX Bought Exchange Rate'}, inplace=True)

# Column MTM CCY
df_fxdeals['MTM CCY'] = df_fxdeals['FX MTM Value'] * df_fxdeals['FX Sold Exchange Rate']

df_fxdeals['MTM CCY'] = df_fxdeals['MTM CCY'].round(2)  # round down the column to 2 decimal points
#df_fxdeals['MTM CCY'].sum()

'''# Column MTM CCY
def mtm_ccy (row):
    if row['FX Sold Currency'] == 'USD':
        return round( row ['FX MTM Value'] * exch_usd['Exchange Rate'], 2).values[0]
    elif row['FX Sold Currency'] == 'EUR':
        return round( row ['FX MTM Value'] * exch_EUR['Exchange Rate'], 2).values[0]
    elif row['FX Sold Currency'] == 'GBP':
        return round (row['FX MTM Value'], 2)
    
# Apply the function (mtm_ccy) to create a new column
df_fxdeals['MTM CCY'] = df_fxdeals.apply(mtm_ccy, axis=1)'''

# Column LCR ROW
def lcr_row_fxdeals(row):
    if row ['FX MTM Value'] > 0:
        return "C74240"
    else:
        return "C73310"
    
# Apply the function (lcr_row) to create a new column
df_fxdeals['LCR ROW'] = df_fxdeals.apply(lcr_row_fxdeals, axis=1)
#df_fxdeals['LCR ROW'].value_counts()

# Column SWAP PAYBLE GBP
df_fxdeals['SWAP PAYBLE GBP'] = df_fxdeals['FX Sold Amount'] / df_fxdeals['FX Sold Exchange Rate']

# Round down the column to 2 decimal points
df_fxdeals['SWAP PAYBLE GBP'] = df_fxdeals['SWAP PAYBLE GBP'].round(2) 
#df_fxdeals['SWAP PAYBLE GBP'].sum()

# Column SWAP RECV GBP
df_fxdeals['SWAP RECV GBP'] = df_fxdeals['FX Bought Amount']/ df_fxdeals['FX Bought Exchange Rate']

# Round down the column SWAP RECV GBP to 2 decimal points
df_fxdeals['SWAP RECV GBP'] = df_fxdeals['SWAP RECV GBP'].round(2)
#df_fxdeals['SWAP RECV GBP'].sum()

'''def swp_payable_gbp (row):
    if row ['FX Sold Currency'] == 'USD':
        return round (row ['FX Sold Amount'] / exch_usd['Exchange Rate'], 2)
    elif row ['FX Sold Currency'] == 'EUR':
        return round (row ['FX Sold Amount'] / exch_EUR['Exchange Rate'], 2)
    elif row ['FX Sold Currency'] == 'GBP':
        return row ['FX Sold Amount']
    
# Apply the function (swp_payable_gbp) to create a new column
df_fxdeals['SWAP PAYBLE GBP'] = df_fxdeals.apply(swp_payable_gbp, axis=1)


def swp_recv_gbp (row):
    if row ['FX Bought Currency'] == 'USD':
        return round (row ['FX Bought Amount'] / exch_usd['Exchange Rate'], 2)
    elif row ['FX Bought Currency'] == 'EUR':
        return round (row ['FX Bought Amount'] / exch_EUR['Exchange Rate'], 2)
    elif row ['FX Bought Currency'] == 'GBP':
        return row ['FX Bought Amount']
    
# Apply the function (swp_payable_gbp) to create a new column
df_fxdeals['SWAP RECV GBP'] = df_fxdeals.apply(swp_recv_gbp, axis=1)


'''

# Drop the columns 'FX Sold Exchange Rate', 'FX Bought Exchange Rate' as they are no longer required

cols_to_drop = ['FX Sold Exchange Rate', 'FX Bought Exchange Rate']

df_fxdeals = df_fxdeals.drop(columns=cols_to_drop)

#merged_fxdeals_exch.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/August ReDIS Files/28.09.2023/fxdeasls1.csv', index = False)

# # Derivatives File

# Merge the DataFrames on column 'Currency Code1'
df_derivative = df_derivative.merge(df_exch, left_on='Currency Code1', right_on='Currency Code', how='left')

# Drop the unwanted column 'Currency Code'
df_derivative = df_derivative.drop(['Currency Code'], axis =1 )

# Rename the column name from ORIG_BAL_BEFORE_INT to Amount in Acct Crncy
df_derivative.rename(columns={'Exchange Rate': 'Currency Code1 Exch Rate'}, inplace=True)


#Merge the DataFrames on column 'Currency Code2'
df_derivative = df_derivative.merge(df_exch, left_on='Currency Code2', right_on='Currency Code', how='left')

# Drop the unwanted column 'Currency Code'
df_derivative = df_derivative.drop(['Currency Code'], axis =1 )

# Rename the column name from ORIG_BAL_BEFORE_INT to Amount in Acct Crncy
df_derivative.rename(columns={'Exchange Rate': 'Currency Code2 Exch Rate'}, inplace=True)

# Column MTM CCY
df_derivative['MTM CCY'] = df_derivative['MTM TOTAL GBP'] * df_derivative['Currency Code2 Exch Rate']

# Round down the column to 2 decimal points
df_derivative['MTM CCY']= df_derivative['MTM CCY'].round(2)

#df_derivative['MTM CCY'].sum()

'''
# Column MTM CCY
def derivatives_mtm_ccy (row):
    if row['Currency Code2'] == 'USD':
        return round (row['MTM TOTAL GBP'] * exch_usd['Exchange Rate'], 2).values[0]
    elif row ['Currency Code2'] == 'EUR':
        return round (row['MTM TOTAL GBP'] * exch_EUR['Exchange Rate'], 2).values[0]
    elif row ['Currency Code2'] == 'GBP':
        return row ['MTM TOTAL GBP']
    else:
        return "0"
    
# Apply the function (derivatives_mtm_ccy) to create a new column
df_derivative['MTM CCY'] = df_derivative.apply(derivatives_mtm_ccy, axis=1)'''


# Column LCR ROW

def lcr_row_derivaties(row):
    if row ['MTM TOTAL GBP'] < 0:
        return "C74240"
    else:
        return "C73310"
    
# Apply the function (lcr_row_derivaties) to create a new column
df_derivative['LCR ROW'] = df_derivative.apply(lcr_row_derivaties, axis=1)
#df_derivative['LCR ROW'].value_counts()

# Column FUTURE PAYABLE GBP

df_derivative['FUTURE PAYABLE GBP'] = -(df_derivative['Notional Amount2 CCY']/ df_derivative['Currency Code2 Exch Rate'])

# ound down the column upto 2 decimal points
df_derivative['FUTURE PAYABLE GBP'] =df_derivative['FUTURE PAYABLE GBP'].round(2)
#df_derivative['FUTURE PAYABLE GBP'].sum()

# Column FUTURE RECEIVABLE GBP

df_derivative['FUTURE RECEIVABLE GBP'] = df_derivative['Notional Amount1 GBP'] / df_derivative['Currency Code1 Exch Rate']

# Round down the column upto 2 decimal points
df_derivative['FUTURE RECEIVABLE GBP'] = df_derivative['FUTURE RECEIVABLE GBP'].round(2)
#df_derivative['FUTURE RECEIVABLE GBP'].sum()

# Drop the columns Currency Code2 Exch Rate and Currency Code1 Exch Rate

cols_to_drop = ['Currency Code2 Exch Rate', 'Currency Code1 Exch Rate']

df_derivative =df_derivative.drop(columns=cols_to_drop)

'''
# Column FUTURE PAYABLE GBP
def future_pyble_gbp(row):
    if row['Currency Code2'] == 'USD':
        return round (-(row['Notional Amount2 CCY'] / exch_usd['Exchange Rate']), 2).values[0]
    elif row ['Currency Code2'] == 'EUR':
        return round (-(row['Notional Amount2 CCY'] / exch_EUR['Exchange Rate']), 2).values[0]
    elif row ['Currency Code2'] == 'GBP':
        return row ['Notional Amount2 CCY'] * -1
    
# Apply the function ( future_pyble_gbp) to create a new column
df_derivative['FUTURE PAYABLE GBP'] = df_derivative.apply(future_pyble_gbp, axis=1)


# Column FUTURE RECEIVABLE GBP

def future_receivable_gbp(row):
    if row['Currency Code1'] == 'USD':
        return round (row['Notional Amount1 GBP'] / exch_usd['Exchange Rate'], 2).values[0]
    elif row ['Currency Code1'] == 'EUR':
        return round ((row['Notional Amount1 GBP'] / exch_EUR['Exchange Rate']), 2).values[0]
    elif row ['Currency Code1'] == 'GBP':
        return row ['Notional Amount1 GBP']
    
# Apply the function ( future_receivable_gbp) to create a new column
df_derivative['FUTURE RECEIVABLE GBP'] = df_derivative.apply(future_receivable_gbp, axis=1)

'''

# # Repay File Start

# Merge the DataFrames on column 'Currency Code1'
df_repay = df_repay.merge(df_exch, left_on='Currency', right_on='Currency Code', how='left')

# Drop the unwanted column 'Currency Code'
df_repay = df_repay.drop(['Currency Code'], axis =1 )

# # Add the dates to the repay file as separate columns
# Add new date columns to the data frame 
# These columns are used for future reference in the code

# Add a new column named Current Date
df_repay['Current Date'] = date_obj

# Add a new column named CASA_Notice_Acc_Date
df_repay['CASA_Notice_Acc_Date'] = casa_excl_notice_acc_dt

# Add the new column named Future Date
df_repay['Future Date'] = future_date

# # Convert the Dates in appropriate format so that the date columns can be worked

# Convert column Current Date to datetime
df_repay['Current Date'] = pd.to_datetime(df_repay['Current Date'], format='%d/%m/%Y')

# Convert column Future Date to datetime
df_repay['Future Date'] = pd.to_datetime(df_repay['Future Date'], format='%d/%m/%Y')

# Convert column CASA_Notice_Acc_Date to datetime
df_repay['CASA_Notice_Acc_Date'] = pd.to_datetime(df_repay['CASA_Notice_Acc_Date'], format='%d/%m/%Y')

# Convert column Interest Cash Flow Date to datetime
df_repay['Interest Cash Flow Date'] = pd.to_datetime(df_repay['Interest Cash Flow Date'], format='%d-%m-%Y')

# Convert column Repayment Date to datetime
df_repay['Repayment Date'] = pd.to_datetime(df_repay['Repayment Date'], format='%d-%m-%Y')

'''a = df_repay[df_repay['Interest Cash Flow Date'] <= df_repay['Current Date']]
a'''

'''b = df_repay[df_repay['Repayment Date'] <= df_repay['Current Date']]
b'''

# when the values in column 'Interest Cash Flow Date' is less than or equal to current date then replace such instances with CASA_Notice_Acc_Date 

mask_1 = (df_repay['Account type'] == 'TDA') & \
       (df_repay['Interest Cash Flow Date'] <= df_repay['Current Date']) & \
       (df_repay['Repayment Date'] <= df_repay['Current Date'])

# Update both 'Interest Cash Flow Date' and 'Repayment Date' columns simultaneously
#df_repay.loc[mask_1, ['Interest Cash Flow Date', 'Repayment Date']] = df_repay.loc[mask_1, 'CASA_Notice_Acc_Date']

# Update both 'Interest Cash Flow Date' and 'Repayment Date' columns simultaneously
df_repay.loc[mask_1, ['Interest Cash Flow Date', 'Repayment Date']] = \
    df_repay.loc[mask_1, 'CASA_Notice_Acc_Date'].dt.date

#df_repay['Repayment Date']

a = df_repay[df_repay['Repayment Date']<= df_repay['Current Date']]
#a

# # Change the Interest Cash Flow Date and Repayment Date back to the desired Format

# Convert column Interest Cash Flow Date to datetime
df_repay['Interest Cash Flow Date'] = pd.to_datetime(df_repay['Interest Cash Flow Date'])

# Convert column Repayment Date to datetime
df_repay['Repayment Date'] = pd.to_datetime(df_repay['Repayment Date'])

# Convert column Interest Cash Flow Date to datetime
df_repay['Interest Cash Flow Date'] = pd.to_datetime(df_repay['Interest Cash Flow Date'], format='%d-%m-%Y')

# Convert column Repayment Date to datetime
df_repay['Repayment Date'] = pd.to_datetime(df_repay['Repayment Date'])

#df_repay['Interest Cash Flow Date']
#df_repay['Repayment Date']

#df_repay.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/December Validation/December-23/repay_test2.csv', index = False)

# Column INTEREST AMOUNT GBP

df_repay['INTEREST AMOUNT GBP'] = df_repay['Interest Cash Flow Amount'] / df_repay['Exchange Rate']

# Round down the column upto 2 decimals
df_repay['INTEREST AMOUNT GBP'] = df_repay['INTEREST AMOUNT GBP'].round(2)
#df_repay['INTEREST AMOUNT GBP'].sum()

# Column PRINCIPLE AMOUNT GBP

df_repay['PRINCIPLE AMOUNT GBP'] = df_repay['Repayment Amount'] / df_repay['Exchange Rate']

# Round down the column upto 2 decimals
df_repay['PRINCIPLE AMOUNT GBP'] = df_repay['PRINCIPLE AMOUNT GBP'].round(2)
#df_repay['PRINCIPLE AMOUNT GBP'].sum().round(2)


'''
# Coulmn INTEREST AMOUNT GBP

def int_amt_gbp(row):
    if row['Currency'] == 'USD':
        return round (row['Interest Cash Flow Amount'] / exch_usd['Exchange Rate'], 2).values[0]
    elif row ['Currency'] == 'EUR':
        return round (row['Interest Cash Flow Amount'] / exch_EUR['Exchange Rate'], 2).values[0]
    elif row ['Currency'] == 'GBP':
        return row ['Interest Cash Flow Amount']
    
# Apply the function ( int_amt_gbp) to create a new column
df_repay['INTEREST AMOUNT GBP'] = df_repay.apply(int_amt_gbp, axis=1)



def principle_amt_gbp(row):
    if row['Currency'] == 'USD':
        return round (row['Repayment Amount'] / exch_usd['Exchange Rate'], 2).values[0]
    elif row ['Currency'] == 'EUR':
        return round (row['Repayment Amount'] / exch_EUR['Exchange Rate'], 2).values[0]
    elif row ['Currency'] == 'GBP':
        return row ['Repayment Amount']
    
# Apply the function ( principle_amt_gbp) to create a new column
df_repay['PRINCIPLE AMOUNT GBP'] = df_repay.apply(principle_amt_gbp, axis=1)

'''

# Converting the columns dtype to string 
df_repay['Customer Identifier'] = df_repay['Customer Identifier'].astype(str)

df_cust['Customer Identifier'] = df_cust['Customer Identifier'].astype(str)

# Create a dictionary mapping Customer Identifier to Retail Flag
customer_mapping = df_cust.set_index('Customer Identifier')['Retail Flag'].to_dict()

# Create a new column 'Retail Flag' in df_repay and fill it based on the mapping
df_repay['Retail Flag'] = df_repay['Customer Identifier'].map(customer_mapping).fillna(0)
#df_repay['Retail Flag'].value_counts()

# Create a dictionary mapping Customer Identifier to SME/Non SME
SME_mapping = df_cust.set_index('Customer Identifier')['SME Flag'].to_dict()

# Create a new column 'Retail Flag' in df_repay and fill it based on the mapping
df_repay['SME/NON SME'] = df_repay['Customer Identifier'].map(SME_mapping).fillna(0)

#Check the total occurances for all the uniques 
#df_repay['SME/NON SME'].value_counts()

# Create a dictionary mapping Customer Identifier to SME/Non SME
sector_code_mapping = df_cust.set_index('Customer Identifier')['Sector Code'].to_dict()

# Create a new column 'Retail Flag' in df_repay and fill it based on the mapping
df_repay['SECTOR CODE'] = df_repay['Customer Identifier'].map(sector_code_mapping).fillna(0).astype(int)

#Check the total occurances for all the uniques 
#df_repay['SECTOR CODE'].value_counts()

# Check the unique entries in column Account Type
#df_repay['Account type'].unique()

# Create a dictionary to map Account Type values to Type values
account_type_mapping = {
    'TDA': 'Deposits',
    'NSB': 'Deposits',
    'LAA': 'Loans',
    'SWAP': 'Treasury',
    'SECURITY': 'Investments',
    'COMMLOAN': 'Treasury'}

# Apply the mapping to create the 'Type' column in df_repay
df_repay['Type'] = df_repay['Account type'].map(account_type_mapping)
#df_repay['Type'].value_counts()

# Create a new column 'Inflow/Outflow' in df_repay
df_repay['Inflow/Outflow'] = df_repay.apply(lambda row: 'Inflow' if (row['INTEREST AMOUNT GBP'] + row['PRINCIPLE AMOUNT GBP'] > 0) and (row['Account type'] != 'NSB') else 'Outflow', axis=1)
#df_repay['Inflow/Outflow'].value_counts()

#df_repay['Deal Reference'].duplicated().sum()

# condition 1: (OR(D5="TDA",D5="NSB")
condition_1 = df_repay['Account type'].isin(['TDA', 'NSB'])
df_repay['condition_1'] = np.where(condition_1, 1, 0)

#df_repay['condition_1']

# Concatenat the values of the two columns into a new column

def combine_columns(row):
    return f"{row['Customer Identifier']}|{row['Deal Reference']}"

#Apply logic (combine_columns) to create the column LCR Sheet Lookup
df_repay['CUST_N_ACC_NUM'] = df_repay.apply(combine_columns, axis=1)

# Create a new data frame for a truncated version of FSCS data frame

# Select the 'CUSTOMER_ID', 'ACC_NUMBER'ananaconda columns from the DataFrame df_FSCSSUB 

df_trunc_FSCS = df_FSCSSUB[['CUSTOMER_ID', 'ACC_NUMBER','INDIVIDUAL + SME','CAT_1','LIAB_OUTF_CAT', 'NON-RETAIL DEPOSIT ROW250']]
#df_trunc_FSCS

# Concatenat the values of the two columns into a new column

def combine_columns(row):
    return f"{row['CUSTOMER_ID']}|{row['ACC_NUMBER']}"

#Apply logic (combine_columns) to create the column LCR Sheet Lookup
df_trunc_FSCS['CUST_N_ACC_NUM'] = df_trunc_FSCS.apply(combine_columns, axis=1)

# Merge the data frames based on common column "CUST_N_ACC_NUM" and "CUST_N_ACC_NUM"

df_repay = pd.merge(df_repay, df_trunc_FSCS, left_on='CUST_N_ACC_NUM', right_on='CUST_N_ACC_NUM', how='left')
#df_repay.head()

'''# Create a dictionary mapping Deal Reference to LIAB_OUTF_CAT
mapping_dict = df_FSCSSUB.set_index('ACC_NUMBER')['LIAB_OUTF_CAT'].to_dict()

# Use map to apply the mapping to df_repay
df_repay['col_frm_FSCS'] = df_repay['Deal Reference'].map(mapping_dict)
'''

# Coulumn Deposit outflow category
def deposit_outflow_cat(row):
    if row ['condition_1'] == 1:
        return row ['LIAB_OUTF_CAT']
    else:
        return row['Type']
    
    
# Apply the logic (deposit_outflow_cat) to create the 'Deposit outflow category' column in df_repay
df_repay['Deposit outflow category'] = df_repay.apply(deposit_outflow_cat, axis=1)

# Fill empty values in 'Deposit outflow category' with values from 'Type' column
df_repay['Deposit outflow category'].fillna(df_repay['Type'], inplace=True)
#df_repay['Deposit outflow category'].value_counts()

#merged_df.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/December Validation/December-23/merged_REpay_N_fscs.csv', index = False)
#merged_df.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/August ReDIS Files/03.10.2023/trial/merged_df3.csv', index = False)

'''# Create a mapping dictionary from 'ACC_NUMBER' to 'LIAB_OUTF_CAT'
mapping_dict = dict(zip(df_FSCSSUB['CUSTOMER_ID'], df_FSCSSUB['LIAB_OUTF_CAT']))

# Map values to df_repay based on 'Deal Reference'
df_repay['Type from FSCS'] = df_repay['Customer Identifier'].map(mapping_dict)


# Fill empty values with 0
#df_repay['NON-RETAIL DEPOSIT ROW250'].fillna(0, inplace=True)'''

'''# Create a dictionary to map 'Deal Reference' to 'LIAB_OUTF_CAT'
deal_ref_to_liab_outf_cat = df_FSCSSUB.set_index('ACC_NUMBER')['LIAB_OUTF_CAT'].to_dict()

# Update the 'Deposit outflow category' column in df_repay based on the mapping
df_repay['Type from FSCS'] = df_repay['Deal Reference'].map(deal_ref_to_liab_outf_cat).fillna('')'''

# Use the Deal Reference values to look up data in df_security
#df_repay.loc[investments_rows.index, 'HQLA'] = investments_rows['Deal Reference'].map(
#    df_security.set_index('Deal Reference')['HQ Indicator']).fillna(0)

# Column HQLA flag

def repay_hqla_flag(row):
    if row['Type'] == 'Investments':
        matching_security_rows = df_security[
            #(df_security['Deal Reference'] == row['Deal Reference']) &
            (df_security['Product Type'] == row['Deal Reference'])
        ]
        if not matching_security_rows.empty:
            return matching_security_rows['HQ Indicator'].iloc[0]
    return 0

# Apply the logic (repay_hqla_flag) to create the 'HQLA flag' column in df_repay
df_repay['HQLA flag'] = df_repay.apply(repay_hqla_flag, axis=1)
df_repay['HQLA flag'].fillna(0, inplace=True)
#df_repay['HQLA flag'].value_counts()

# Column LCR sheet Lookup

def combine_columns(row):
    return f"{row['Type']}|{row['Inflow/Outflow']}|{row['HQLA flag']}"

#Apply logic (combine_columns) to create the column LCR Sheet Lookup
df_repay['LCR sheet Lookup'] = df_repay.apply(combine_columns, axis=1)

'''a = df_repay[df_repay['LCR sheet Lookup'] == "Investments|Outflow|0"]
a'''
#df_repay['LCR sheet Lookup'].value_counts()
#df_repay.columns

# Reorder the columns in the data frame

# Drop the 'NON-RETAIL DEPOSIT ROW250' column
# The pop () method is used to remove the 'NON-RETAIL DEPOSIT ROW250' column from the DataFrame and save its values.
non_retail_row_250 = df_repay.pop('NON-RETAIL DEPOSIT ROW250')

# Insert the 'ORIG_BAL_BEFORE_INT' column after the 'ACC_CRNCY' column
df_repay.insert(df_repay.columns.get_loc('LCR sheet Lookup') + 1, 'NON-RETAIL DEPOSIT ROW250', non_retail_row_250)

# Fill empty values with 0
df_repay['NON-RETAIL DEPOSIT ROW250'].fillna(0, inplace=True)
#df_repay.head()
#df_repay['NON-RETAIL DEPOSIT ROW250'].value_counts()

#df_repay.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/December Validation/December-23/test_repay_28thFeb.csv', index = False)

'''# =IFERROR(VLOOKUP(H5,'Lombard FSCS'!$D:$AK,34,0),0)
# Create a mapping dictionary from 'ACC_NUMBER' to 'NON-RETAIL DEPOSIT ROW250'
mapping_dict = dict(zip(df_FSCSSUB['ACC_NUMBER'], df_FSCSSUB['NON-RETAIL DEPOSIT ROW250']))

# Map values to df_repay based on 'Deal Reference'
df_repay['NON-RETAIL DEPOSIT ROW250'] = df_repay['Deal Reference'].map(mapping_dict)

# Fill empty values with 0
df_repay['NON-RETAIL DEPOSIT ROW250'].fillna(0, inplace=True)'''

# Create a dictionary to map LCR sheet Lookup values to LCR sheet no values
LCR_shhet_mapping = {
    'Treasury|Inflow|0' : '74',
    'Treasury|Outflow|0' : '73',
    'Loans|Inflow|0' : '74',
    'Deposits|Inflow|0' : '74',
    'Deposits|Outflow|0' : '73',
    'Investments|Inflow|II A' : '72',
    'Investments|Inflow|0' : '74',
    'Investments|Inflow|II B' : '72',
    'Investments|Inflow|I' : '72'
    }

# Apply the mapping to create the 'Type' column in df_repay
df_repay['LCR sheet'] = df_repay['LCR sheet Lookup'].map(LCR_shhet_mapping)
#df_repay['LCR sheet'].value_counts()

# Addition of empty columns as disscussed
df_repay['Row for 72'] = ''
df_repay['Row for 73']= ''
df_repay['Row for 74']= ''
df_repay['LCR sheet Row']= ''
df_repay['Row for72']= ''
df_repay['Row for73']= ''
df_repay['Row for74']= ''
df_repay['LCR_Sheet_Row']= ''
df_repay['Blank Column']= ''

#df_repay.info()

# Condition to check if Sector Code is one of the specified values
condition = df_repay['SECTOR CODE'].isin([3350, 3370, 4539])

# Apply np.where to return 'YES' if condition is True, else 'NO'
df_repay['FINANCIAL/NON-FINANCIAL'] = np.where(condition, 'YES', 'NO')
#df_repay['FINANCIAL/NON-FINANCIAL'].value_counts()

# Now since we have the values from the FSCS file placed in the truncated FSCS version which is merged with the Repay file we can directly go ahead and perform the remaining vlookup operations between the FSCS and Repay file.

# Column INFLOW CATEGORY

# function for column Inflow Category.

def inflow_category (row):
    if row ['FINANCIAL/NON-FINANCIAL'] == "YES":
        return "FINANCIAL"
    elif row['SME/NON SME'] == "NON-SME":
        return "NON-RETAIL"
    elif row ['CAT_1'] == "NON Retail Deposit":
        return "NON-RETAIL"
    else:
        return "RETAIL"
    
#Apply logic (inflow_category) to create the column INFLOW CATEGORY
df_repay['INFLOW CATEGORY'] = df_repay.apply(inflow_category, axis=1)
#df_repay['INFLOW CATEGORY'].value_counts()
#df_repay.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/December Validation/December-23/test_repay_29thFeb.csv', index = False)

# Column Exempt Deposits

def exempt_deposits(row):
    if row ['INDIVIDUAL + SME'] == "NO":
        return "NON RETAIL Deposit"
    elif row['INDIVIDUAL + SME'] == "YES":
        return "RETAIL"
    #elif pd.isnull(row['INDIVIDUAL + SME']):
        #return "0"
    else:
        return "0"        
        
#Apply logic (exempt_deposits) to create the column Exempt Deposits
df_repay['Exempt Deposits'] = df_repay.apply(exempt_deposits, axis=1)
# If there's no match, fill the 'Exempt Deposits' column with 0
#df_repay['Exempt Deposits'].fillna(0, inplace=True)

df_repay['INDIVIDUAL + SME'].isnull().sum()

#=IFERROR(IF(VLOOKUP(B5,'Lombard FSCS'!$B$5:$AQ$113308,27,0)="NO","NON RETAIL Deposit","RETAIL"),"0")

#df_repay['INDIVIDUAL + SME'].value_counts()
#df_repay['Exempt Deposits'].value_counts()
#df_repay['Exempt Deposits'].isnull().sum()

# Create a function to apply the logic to get the column Investment Type from data frame FSCS

def get_investment_type(row):
    matching_security_rows = df_security[df_security['Product Type'] == row['Deal Reference']]
    if not matching_security_rows.empty:
        return matching_security_rows['Investment Type'].iloc[0]
    return None  # Return None if no match is found

# Apply the function to create the 'cvb' column in df_repay
df_repay['Investment Type from security'] = df_repay.apply(get_investment_type, axis=1)
#df_repay['Investment Type from security'].value_counts()

# Column Investment Type

def investement_type (row):
    if row['Type'] == "Investments":
        return row['Investment Type from security']
    else:
        return "0"
    
#Apply logic (investement_type) to create the column Investment Type
df_repay['Investment Type'] = df_repay.apply(investement_type, axis=1)
#df_repay['Investment Type'].value_counts()

# Column Original Maturity

def orig_maturity (row):
    if row['Account type'] == "NSB":
        return "35"
    else:
        return "0"

#Apply logic (orig_maturity) to create the column Original Maturity
df_repay['Original Maturity'] = df_repay.apply(orig_maturity, axis=1)

# Create new empty columns

df_repay['Original Maturity'] ='Column Not Used in Calculation'
df_repay['Residual Maturity'] = 'Column Not Used in Calculation'
df_repay['PROD ORIG MAT'] = 'Column Not Used in Calculation'
df_repay['PROD RESI MAT'] = 'Column Not Used in Calculation'
df_repay['C68 Eligibility'] = 'Column Not Used in Calculation'

# Reorder the columns in the data frame to reorder the following column

# Column Individual + SME

# The pop () method is used to remove the 'INDIVIDUAL + SME' column from the DataFrame and save its values.
individual_n_SME = df_repay.pop('INDIVIDUAL + SME')

# Insert the 'INDIVIDUAL + SME' column after the 'C68 Eligibility' column
df_repay.insert(df_repay.columns.get_loc('C68 Eligibility') + 1, 'INDIVIDUAL + SME', individual_n_SME)

# Column LCR Period INT

def lcR_period_int (row):
    if pd.isnull(row['Interest Cash Flow Date']):
        return ""
    elif date_obj < row['Interest Cash Flow Date'] <= row['Future Date']:
        return "YES"
    else:
        return "NO"
#Apply logic (lcR_period_int) to create the column LCR PERIOD INT
df_repay['LCR PERIOD INT'] = df_repay.apply(lcR_period_int, axis=1)

# Column LCR PERIOD PRINCIPAL

def lcR_period_principle (row):
    if pd.isnull(row['Repayment Date']):
        return ""
    elif date_obj < row['Repayment Date'] <= row['Future Date']:
        return "YES"
    else:
        return "NO"
#Apply logic (lcR_period_principle) to create the column LCR PERIOD PRINCIPAL
df_repay['LCR PERIOD PRINCIPAL'] = df_repay.apply(lcR_period_principle, axis=1)
#df_repay.info()

# Drop the unwanted columns from the Repay file

repay_col_to_drop = ['Current Date','Exchange Rate','condition_1', 'CUST_N_ACC_NUM','CUSTOMER_ID', 'ACC_NUMBER', 'CAT_1', 'LIAB_OUTF_CAT',
                     'Investment Type from security','CASA_Notice_Acc_Date','Future Date']
df_repay = df_repay.drop(columns=repay_col_to_drop)
#df_repay.head(5)

#df_repay.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/December Validation/December-23/test_repay_29thFeb1.csv', index = False)


# # SecurityPortfolio

# Create a copy of the Securityportfolio
df_securityportfolio_copy = df_securityportfolio.copy()

#df_securityportfolio_copy.columns

# The security Portfolio file comes in with an additional row  of data at the bottom which is not required for the calculation purpose.
# Thus this row needs to be searched based on the logic below and eliminated from the file/ data frame.

# Find rows containing the phrases and create a boolean mask
mask = df_securityportfolio_copy['Book Value'].str.contains('Total Face Value in USD|Total Book Value in USD', case=False, na=False)

# Use the boolean mask to filter and keep rows that do not contain the phrases
df_securityportfolio_copy = df_securityportfolio_copy[~mask]
#df_securityportfolio_copy

security_col_to_drop = ['Moodys Rating', 'Fitch Rating','SnP Rating', 'Average Purchase Price', 
                        'Market Price','MTM_USD','Coupon Rate','Appreciation Depreciation in USD', 
                        'percentage change in price','Modified Duration', 'MTM_MD_USD', 'PV01', 
                        'EOD Date', 'YTM']
df_securityportfolio_copy = df_securityportfolio_copy.drop(columns=security_col_to_drop)
#df_securityportfolio_copy.columns
#df_securityportfolio_copy.info()
#df_securityportfolio_copy.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/August ReDIS Files/28.09.2023/fx.csv', index = False)


# # Logic for Pivot Security

# Group the DataFrame by 'Product Type' and calculate the sum of 'Market Value CCY' and MTM GBP
Pivot_Security = df_security.groupby('Product Type').agg({'MTM GBP': 'sum', 'Market Value CCY': 'sum'}).reset_index()
#Pivot_Security

# Merge df_securityportfolio_copy with Pivot_Security on 'Deal Num' and 'Product Type'
df_securityportfolio_copy = df_securityportfolio_copy.merge(Pivot_Security, left_on='Deal Num', right_on='Product Type', how='left')
#df_securityportfolio_copy.columns

# Rename the columns as per the convience

# Rename the 'MTM GBP' column to 'Market Price GBP' and column  'Market Value CCY' to 'Market Price CCY'in df_securityportfolio_copy
df_securityportfolio_copy.rename(columns={'MTM GBP': 'Market Price GBP', 'Market Value CCY': 'Market Price CCY'}, inplace=True)
#df_securityportfolio_copy.columns

# Create a dictionary mapping 'Product Type' to 'HQ Indicator'
type_to_hq = df_security.set_index('Product Type')['HQ Indicator'].to_dict()

# Use the map function to create a new column 'HQLA Indicator' in df_securityportfolio_copy
df_securityportfolio_copy['HQLA Indicator'] = df_securityportfolio_copy['Deal Num'].map(type_to_hq)

# Fill empty cells with 0
df_securityportfolio_copy['HQLA Indicator'] = df_securityportfolio_copy['HQLA Indicator'].fillna('0')
#df_securityportfolio_copy['HQLA Indicator'].value_counts()
#df_securityportfolio_copy.info()

df_securityportfolio_copy['Book Value'] = df_securityportfolio_copy['Book Value'].astype(float) 

# Merge the DataFrames on 'Scrip Currency' and Currency Code'
df_securityportfolio_copy = df_securityportfolio_copy.merge(df_exch, left_on='Scrip Currency', right_on='Currency Code', how='left')

# Drop the unwanted column Currency Code
df_securityportfolio_copy = df_securityportfolio_copy.drop('Currency Code', axis =1)

# Column BV GBP
df_securityportfolio_copy['BV GBP'] = df_securityportfolio_copy['Book Value']/ df_securityportfolio_copy['Exchange Rate']

df_securityportfolio_copy['BV GBP'] = df_securityportfolio_copy['BV GBP'].round(2)
#df_securityportfolio_copy['BV GBP'].sum()

'''def bv_gbp(row):
    if row['Scrip Currency'] == 'GBP':
        return round( row['Book Value'], 2)
    elif row['Scrip Currency'] == 'USD':
        return round( row['Book Value'] / exch_usd['Exchange Rate'], 2).values[0]
    elif row['Scrip Currency'] == 'EUR':
        return round (row['Book Value'] / exch_EUR['Exchange Rate'], 2).values[0]
    
#Apply logic (bv_gbp) to create the column BV GBP
df_securityportfolio_copy['BV GBP'] = df_securityportfolio_copy.apply(bv_gbp, axis=1)'''

# Column NSFR VALUE GBP
def nsfr_val_gbp(row):
    if row ['Portfolio'] == "HTM":
        return row['BV GBP']
    else:
        return row['Market Price GBP']

#Apply logic (nsfr_val_gbp) to create the column NSFR VALUE GBP
df_securityportfolio_copy['NSFR VALUE GBP'] = df_securityportfolio_copy.apply(nsfr_val_gbp, axis=1)

df_securityportfolio_copy['NSFR VALUE GBP'].sum().round(2)

# Column NSFR VALUE CCY
def nsfr_val_ccy(row):
    if row ['Portfolio'] == "HTM":
        return row['Book Value']
    else:
        return row['Market Price CCY']

#Apply logic (nsfr_val_ccy) to create the column NSFR VALUE CCY
df_securityportfolio_copy['NSFR VALUE CCY'] = df_securityportfolio_copy.apply(nsfr_val_ccy, axis=1)

df_securityportfolio_copy['NSFR VALUE CCY'].sum().round(2)

df_securityportfolio_copy = df_securityportfolio_copy.drop(columns = ['Product Type','Exchange Rate'])
#df_securityportfolio_copy.info()

# # Limit File

#print('Limit File Info:', df_limit.info())
#print('Limit File columns:', df_limit.columns)

# Merge the DataFrames on 'Scrip Currency' and Currency Code' to get the Exchange Rates
df_limit = df_limit.merge(df_exch, left_on='Currency Code', right_on='Currency Code', how='left')

# Create a copy of customer data frame
df_cust_copy = df_cust.copy() 

# Check the columns in the copied data frame
df_cust_copy.columns

# Truncate df_cust_copy to have important columns
drop_columns = ['Branch id', 'Customer Name','Risk Code', 'Incorporation Code', 'Industry Code','Connected Code', 'Deposit Guarantee Scheme',
       'Development Bank Indicator', 'Retail Flag', 'Parent Code','Ultimate Parent Code', 'SME Flag',
       'Short Term Credit Rating Code - Moody',
       'Long Term Credit Rating Code - Moody',
       'Short Term Credit Rating Code - Fitch',
       'Long Term Credit Rating Code - Fitch',
       'Short Term Credit Rating Code - S and P',
       'Long Term Credit Rating Code - S and P',
       'Sovereign Risk Transfer Indicator', 'LECustGroup', 'LECustomerCode',
       'LEGroupCode', 'LEI CODE']

df_cust_copy = df_cust_copy.drop(columns=drop_columns)

# Convert the data dtype of the column Customer Identifier to string  
df_cust_copy['Customer Identifier'] = df_cust_copy['Customer Identifier'].astype(str)

# Convert the data type of the column Customer Identifier to string
df_limit['Customer Identifier'] = df_limit['Customer Identifier'].astype(str)

#merge the dataframe df_cust_copy and df_limit
# Merge the DataFrames on 'Customer Identifier' and Customer Identifier 
df_limit = df_limit.merge(df_cust_copy, left_on='Customer Identifier', right_on='Customer Identifier', how='left')

# Conert the data type of the column Sector Code to int
df_limit['Sector Code'] = df_limit['Sector Code'].astype(int)
#df_limit.isnull().sum()

# Add a new column named Future Date
df_limit['Future Date'] = future_date
#df_limit.info()

# Convert all the date related columns to the date and time format

# Column Start Date to the Date and time format
df_limit['Start Date'] = pd.to_datetime(df_limit['Start Date'], format='%d-%m-%Y')

# Column Maturity Date to the Date and time format
df_limit['Maturity Date'] = pd.to_datetime(df_limit['Maturity Date'], format='%d-%m-%Y')

# Column Future Date to the Date and time format
df_limit['Future Date'] = pd.to_datetime(df_limit['Future Date'], format='%d/%m/%Y')

# Column Drawn amount CCY
# Add the column Drawn amount CCY by performing the subtraction operation on 2 columns 
df_limit['Drawn amount CCY'] = df_limit['Facility Amount'] - df_limit['Undrawn Amount']
df_limit['Drawn amount CCY'] = df_limit['Drawn amount CCY'].round(2)

# Column Drawn amount GBP

df_limit['Drawn amount GBP'] = df_limit['Drawn amount CCY']/ df_limit['Exchange Rate']

# round down the column to 2 decimal points
df_limit['Drawn amount GBP'] = df_limit['Drawn amount GBP'].round(2)

# Column Undrawn GBP
df_limit['Undrawn GBP'] = df_limit['Undrawn Amount']/ df_limit['Exchange Rate']

# Round donw the column upto 2 decimnal points
df_limit['Undrawn GBP'] = df_limit['Undrawn GBP'].round(2)


# Add an empty column Scheme Code to the data frame
df_limit[' Scheme code '] = ''

# Column LCR Row

def lcr_row(row):
    if row['Reference Code'] in ["CEL05", "CEL06"]:
        return "C74070"
    else:
        return "0"

#Apply logic (lcr_row) to create the column LCR ROW
df_limit['LCR ROW'] = df_limit.apply(lcr_row, axis=1)
#df_limit['LCR ROW'].value_counts()

# Column LCR PERIOD

def lcr_period (row):
    if row['Maturity Date'] < row['Future Date']:
        return "YES"
    else:
        return "NO"
    
#Apply logic (lcr_period) to create the column LCR PERIOD
df_limit['LCR PERIOD'] = df_limit.apply(lcr_period, axis=1)
#df_limit['LCR PERIOD'].value_counts()

# Drop the 'Sector Code' column
# The pop () method is used to remove the 'Sector Code' column from the DataFrame and save its values.
secotr_code = df_limit.pop('Sector Code')

# Insert the 'Sector Code' column after the 'ACC_CRNCY' column
df_limit.insert(df_limit.columns.get_loc('LCR PERIOD') + 1, 'Sector Code', secotr_code)
#df_limit['Sector Code'].value_counts()

# Column FINANCIAL CUSTOMER

def cust_info(row):
    if row ['Sector Code'] in [3350, 3370, 4539]:
        return "FINANCIAL"
    else:
        return "NON-FINANCIAL"

#Apply logic (cust_info) to create the column FINANCIAL CUSTOMER
df_limit['FINANCIAL CUSTOMER'] = df_limit.apply(cust_info, axis=1)
#df_limit['FINANCIAL CUSTOMER'].value_counts()

drop_col = ['Exchange Rate', 'Future Date','Residency Code']
df_limit = df_limit.drop(columns=drop_col) 
#df_limit.info()


# # GL Bal File

# Merge the DataFrames on 'Currency Code'
df_glbal = df_glbal.merge(df_exch, left_on='Currency Code', right_on='Currency Code', how='left')

# Column AMT GBP
df_glbal['AMT GBP'] = df_glbal['Amount CCY'] / df_glbal['Exchange Rate']

# Round down the column upto 2 decimal poitns
df_glbal['AMT GBP'] = df_glbal['AMT GBP'].round(2)

def glbal_LCR_row (row): 
    if row ['GL NUMBER'] == "12030" and row['Acct NUMBER'] != "61198261203016":
        return "C74160"
    else:
        return "0"

#Apply logic (glbal_LCR_row) to create the column LCR ROW
df_glbal['LCR ROW'] = df_glbal.apply(glbal_LCR_row, axis=1)

df_glbal['GL NUMBER'] = df_glbal['GL NUMBER'].astype(str)

df_glbal['TEXT GL NUMBER'] = df_glbal['GL NUMBER'].str[0]
#df_glbal['LCR ROW'].value_counts()

#columns_to_drop = ['Exchange Rate']
df_glbal = df_glbal.drop('Exchange Rate', axis = 1)
#df_glbal['CIF ID'].isnull().value_counts()

#df_glbal.to_csv('filtered_glbal.csv')

#df_glbal.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/New folder/trial/July Liquidity File/Redis Files/trial.csv', index = False)

# Get user input for the parent folder
destination_folder = input("Enter the parent folder path: ")

# Define a dictionary to map DataFrame names to their corresponding variables
data_frames = {
    "filtered_GLBalanceSUB": df_glbal,
    "filtered_FSCSSUB": df_FSCSSUB,
    "filtered_SecuritySUB": df_security,
    "filtered_FXDealsSUB": df_fxdeals,
    "filtered_DerivativeSUB": df_derivative,
    "filtered_RepaySUB": df_repay,
    "filtered_SECURITYPORTFOLIOSUB": df_securityportfolio_copy,
    "filtered_CUSTSUB": df_cust,
    "filtered_LimitSUB": df_limit,
}

# Loop through the dictionary and save each DataFrame as .txt and .csv files
for file_name, data_frame in data_frames.items():
    # Save as .txt file
    txt_file_path = os.path.join(destination_folder, f"{file_name}.txt")
    data_frame.to_csv(txt_file_path, sep=',', index=False)
    
    # Save as .csv file
    txt_file_path = os.path.join(destination_folder, f"{file_name}.csv")
    data_frame.to_csv(txt_file_path, sep=',', index=False)
    
    ## Save as .xlsx file
    #txt_file_path = os.path.join(destination_folder, f"{file_name}.xlsx")
    #data_frame.to_excel(txt_file_path,index=False)

# Confirm the successful saving of the files
print("DataFrames saved in both .txt and .csv formats in the provided folder.")


#df_derivative.columns



# Convert 'Maturity Date' to datetime format for comparison and manipulation

df_security['Maturity Date'] = pd.to_datetime(df_security['Maturity Date'], dayfirst=True)
df_security['Start Date'] = pd.to_datetime(df_security['Start Date'], dayfirst=True)
df_security['Settlement Date'] = pd.to_datetime(df_security['Settlement Date'], dayfirst=True)
df_security['Next Reset Date'] = pd.to_datetime(df_security['Next Reset Date'], dayfirst=True)

df_fxdeals['Deal Date'] = pd.to_datetime(df_fxdeals['Deal Date'], dayfirst=True)
df_fxdeals['Maturity date'] = pd.to_datetime(df_fxdeals['Maturity date'], dayfirst=True)

df_derivative['Deal Date'] = pd.to_datetime(df_derivative['Deal Date'], dayfirst=True)
df_derivative['Maturity date'] = pd.to_datetime(df_derivative['Maturity date'], dayfirst=True)
df_derivative['Settlement Date'] = pd.to_datetime(df_derivative['Settlement Date'], dayfirst=True)
df_derivative['Settlement Date Rec'] = pd.to_datetime(df_derivative['Settlement Date Rec'], dayfirst=True)
df_derivative['Exch Principal Start'] = pd.to_datetime(df_derivative['Exch Principal Start'], dayfirst=True)
df_derivative['Exch Principal End'] = pd.to_datetime(df_derivative['Exch Principal End'], dayfirst=True)


#excel_file = "C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/Reports/Liquidity automation file_June Trial.xlsx"

#fscs_excel = pd.read_excel(excel_file, sheet_name= 'Lombard FSCS')


#fscs_excel.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/New folder/trial/July Liquidity File/Redis Files/excle_FSCSC.csv', index = False)

# import xlwings as xw


# Search for the LCR file in the parent folder
lcr_files = [file for file in all_files if file.startswith('Liquidity automation file_Python Values') and file.endswith('.xlsx')]

if len(lcr_files) == 1:
    # Assuming there is only one LCR file, open it directly
    lcr_file_path = os.path.join(parent_folder, lcr_files[0])
    # Open the Excel workbook
    wb = xw.Book(lcr_file_path)
    # Print the names of all worksheets
    print("Worksheets in the workbook:")
    for sheet in wb.sheets:
        print(sheet.name)
else:
    print("No or multiple LCR files found. Please ensure there is exactly one LCR file in the parent folder.")

#except Exception as ex:
 #   print(f"An error occurred: {ex}")


''' # Get the worksheet by name
ws_security = wb.sheets["Lombard Security"]
    
    # Print the worksheet name
    #print(f"Displaying worksheet: {ws_security.name}")
    
# Show the worksheet
ws_security.activate()'''

# Get user input for the path to the Excel file
#file_path = input("Enter the path to the Excel file: ")

# Open the Excel workbook
#wb = xw.Book(file_path)

# Print the names of all worksheets
#print("Worksheets in the workbook:")
#for sheet in wb.sheets:
#    print(sheet.name)

# Select the worksheet named 72

ws_72 = wb.sheets['72']

# Get user input for the date in dd/mm/yyyy format
#date_str = input("Enter the date in dd/mm/yyyy format: ")

# Validate the date format
try:
    # Convert the input string to a datetime object
    #date = datetime.strptime(date_str, '%d/%m/%Y')
    # Write the date to cell E5 of the worksheet
    #ws_72.range('E5').value = date
    ws_72.range('E5').value = date_obj
    print("Date successfully entered in cell E5.")
except ValueError:
    print("Invalid date format. Please enter the date in dd/mm/yyyy format.")

# Select the worksheet named "Lombard EXCH Current"
ws_exch = wb.sheets['Lombard EXCH Current']

# Delete all the contents of column C from the 2nd row
ws_exch.range('C2:C25').clear_contents()

# Copy data from DataFrame to worksheet starting from the 2nd row
exch_data = df_exch['Exchange Rate'].values.reshape(-1, 1)  # Reshape to a column vector
ws_exch.range('C2').value = exch_data


def clear_contents_from_2ndrow(sheet_name, start_row):
    # Select the worksheet by name
    ws = wb.sheets[sheet_name]
    
    # Get the last used row in the worksheet
    last_row = ws.cells(ws.api.Rows.Count, 1).end("up").row
    
    # Determine the range to clear
    range_to_clear = ws.range((start_row, 1), (last_row, ws.api.Columns.Count))
    
    # Clear the contents of the range
    range_to_clear.clear_contents()

# Open the Excel workbook
#file_path = input("Enter the path to the Excel file: ")
#wb = xw.Book(file_path)

# Call the function for each worksheet
worksheets_to_clear = ["Lombard Cust RPT", "Lombard Security", "Security Portfolio",
                       "Lombard FX", "Lombard Derivatives", "Lombard GL Bala"]  # Add more worksheet names as needed
start_row_to_clear = 2  # Specify the start row

for sheet_name in worksheets_to_clear:
    clear_contents_from_2ndrow(sheet_name, start_row_to_clear)


'''# Select the worksheet named "Lombard Security"
ws_security = wb.sheets['Lombard Security']

# Clear all contents in the worksheet
ws_security.clear()'''


def clear_contents_from_5throw(sheet_name, start_row):
    # Select the worksheet by name
    ws = wb.sheets[sheet_name]
    
    # Get the last used row in the worksheet
    last_row = ws.cells(ws.api.Rows.Count, 1).end("up").row
    
    # Determine the range to clear
    range_to_clear = ws.range((start_row, 1), (last_row, ws.api.Columns.Count))
    
    # Clear the contents of the range
    range_to_clear.clear_contents()

# Call the function for each worksheet
worksheets_to_clear = ["Lombard Repay", "Lombard FSCS"]  # Add more worksheet names as needed
start_row_to_clear = 5  # Specify the start row

for sheet_name in worksheets_to_clear:
    clear_contents_from_5throw(sheet_name, start_row_to_clear)

#df_security['Maturity Date']

def copy_dataframe_to_worksheet(df, ws, start_row):
    # Write the data (excluding headers) to the worksheet
    ws.range(f"A{start_row}").value = df.values

# Define a list of tuples containing DataFrame and corresponding worksheet names followed by the start row
dataframes_and_worksheets = [
    (df_cust, "Lombard Cust RPT", 2),
    (df_repay, "Lombard Repay", 5),
    (df_security, "Lombard Security", 2),
    (df_securityportfolio_copy, "Security Portfolio", 2),
    (df_FSCSSUB, "Lombard FSCS", 5),
    (df_fxdeals, "Lombard FX", 2),
    (df_derivative, "Lombard Derivatives", 2),
    (df_glbal, "Lombard GL Bala", 2)
]

# Iterate over each tuple and copy the corresponding DataFrame to the worksheet
for df, sheet_name, start_row in dataframes_and_worksheets:
    ws = wb.sheets[sheet_name]
    copy_dataframe_to_worksheet(df, ws, start_row)

df_security 

# Select the worksheet named "Lombard Repay"
ws_limit = wb.sheets["Lombard Limit"]

# Get the range starting from the 5th row
start_row = 3
start_column = 1  # Assuming you want to start from the first column
end_row = ws_limit.api.Rows.Count
end_column = 22
range_to_clear = ws_limit.range((start_row, start_column), (end_row, end_column))

# Clear the contents of the range
range_to_clear.clear_contents()

# Write the data (excluding headers) to the worksheet
ws_limit.range("A3").value = df_limit.values


#df_security['Maturity Date']

#xw.view(df_security)

# Save the workbook
wb.save()

# Close the workbook
wb.close()


# Record the end time
laptime = time.time()

# Calculate the elapsed time
LCR_completion_time = laptime - start_time

# Convert elapsed time to a human-readable format
hours, rem = divmod(LCR_completion_time, 3600)
minutes, seconds = divmod(rem, 60)

# Display the time taken
print("Time taken for LCR File Completion: {:0>2}:{:0>2}:{:05.2f}".format(int(hours), int(minutes), seconds))


# # Capital Dashboard Calculation Start #

# #### Read Dependency File


def read_dependency_coding_file(parent_folder, file_name='Capital_Dashboard_Dependency.xlsx'):
    """
    Read specific worksheets from the Capital_Dashboard_Dependency file in the specified parent folder.

    Args:
    parent_folder (str): Path of the parent folder.
    file_name (str): Name of the GL coding file. Default is 'Capital_Dashboard_Dependency.xlsx'.

    Returns:
    tuple: Tuple containing DataFrames for each worksheet (ratings_df).
    """
    
    file_path = os.path.join(parent_folder, file_name)

    # Check if the file exists before reading it
    if os.path.exists(file_path):
        # Use xlwings to load the Excel file
        app = xw.App(visible=False)  # Create an Excel app instance (hidden)
        wb = xw.Book(file_path)  # Open the Excel file

        # Initialize DataFrames for each worksheet
        df_rating = None
        #code_desc_df = None
        #acc_desc_df = None
        #schm_desc_df = None
        df_Inst_Y_N = None

        # Loop through each worksheet and load data into respective DataFrames
        for sheet in wb.sheets:
            if sheet.name == 'Ratings':
                df_rating = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
            elif sheet.name == 'BOE Sector Code':
                df_BOE_sector_code = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
            elif sheet.name == 'Risk Weight Logic':
                df_rwlogic = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
            elif sheet.name == 'Inst Y_N':
                df_Inst_Y_N = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
            #elif sheet.name == 'MPC Dates':
            #    mpc_df = sheet.used_range.options(pd.DataFrame, index=False, header=True).value

        wb.close()  # Close the workbook
        app.quit()  # Quit the Excel app
        
        return df_rating, df_BOE_sector_code,df_rwlogic, df_Inst_Y_N
    
    else:
        print(f"The file '{file_name}' does not exist in the specified folder.")
        return None

# CAll the Function:
df_rating, df_BOE_sector_code,df_rwlogic, df_Inst_Y_N = read_dependency_coding_file(parent_folder)

if df_rating is not None:
    print("Capital Dependency data loaded successfully.")
    # Proceed with using the DataFrames as needed
else:
    print("No Capital Dependency data loaded.")


#df_BOE_sector_code


#df_rating


# ### Limit File ###

#df_accbal.columns

#df_limit


# Merge based on the condition where 'ACCT NUMBER' matches 'Deal Reference'

merged_df = pd.merge(df_limit, df_accbal[['Scheme Code', 'Account Number']], 
                     left_on='Deal Reference', right_on='Account Number', how='left')

# Drop the 'Account Number' column as it is not needed
merged_df.drop(columns='Account Number', inplace=True)

# Now, merged_df contains the updated values where applicable

# Assign it back to df_limit if needed
df_limit = merged_df

# Define a custom function to apply the logic
def calculate_ccf(row):
    if row['Reference Code'] in ['CEL05', 'CEL06']:
        if row['Scheme Code'] in ['OD221', 'OD201']:
            return 0.2
        elif row['Scheme Code'] == 'OD222':
            return 0.5
        elif (row['Maturity Date'] - row['Start Date']).days <= 366:
            return 0.2
        else:
            return 0.5
    else:
        return None  # You can modify this to return a default value if needed

# Apply the custom function to create the 'CCF' column
df_limit['CCF'] = df_limit.apply(calculate_ccf, axis=1)


#df_limit.info()

#df_security.shape

#df_limit.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/Dailies_Python Values/trial/limit.csv', index = False)

# #### Security File

#df_security.columns

#df_rating.columns

def merge_security_with_ratings(df_security, df_rating):
    """
    Merge security data with ratings data based on different credit rating agencies.

    Args:
    df_security (DataFrame): DataFrame containing security data.
    df_rating (DataFrame): DataFrame containing credit ratings data.

    Returns:
    DataFrame: Merged DataFrame with updated columns for credit ratings.
    """
    # Merge 'Long Term Moody' with 'Moodys'
    df_security = pd.merge(df_security, df_rating[['Moodys', 'New #']], 
                           left_on='Long Term Moody', right_on='Moodys', how='left')
    
    # Rename columns and drop unnecessary column
    df_security.rename(columns={'New #': 'Moodys Number'}, inplace=True)
    df_security.drop(columns='Moodys', inplace=True)

    # Merge 'Long Term Fitch' with 'Fitch'
    df_security = pd.merge(df_security, df_rating[['Fitch', 'New #']], 
                           left_on='Long Term Fitch', right_on='Fitch', how='left')

    # Rename columns and drop unnecessary column
    df_security.rename(columns={'New #': 'Fitch Number'}, inplace=True)
    df_security.drop(columns='Fitch', inplace=True)

    # Merge 'Long Term S P' with 'S&P'
    df_security = pd.merge(df_security, df_rating[['S&P', 'New #']], 
                           left_on='Long Term S P', right_on='S&P', how='left')

    # Rename columns and drop unnecessary column
    df_security.rename(columns={'New #': 'S&P Number'}, inplace=True)
    df_security.drop(columns='S&P', inplace=True)

    return df_security

# Call the function to merge and update ratings information
df_security = merge_security_with_ratings(df_security, df_rating)

#df_security.info()

#df_cust.columns

# Merge based on the condition where 'ACCT NUMBER' matches 'Deal Reference'

merged_df = pd.merge(df_security, df_cust[['Customer Identifier', 'Customer Name']], 
                     left_on='Issuer Id', right_on='Customer Identifier', how='left')

# Drop the 'Account Number' column as it is not needed
merged_df.drop(columns='Customer Identifier', inplace=True)

# Now, merged_df contains the updated values where applicable

# Assign it back to df_security if needed
df_security = merged_df

#df_security.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/Dailies_Python Values/trial/security.csv', index = False)

def calculate_final_rating(df):
    """
    Calculate the 'Final Rating' based on non-empty cells in specified columns of the input DataFrame.

    Args:
    df (pandas.DataFrame): Input DataFrame containing columns 'Moodys Number', 'Fitch Number', 'S&P Number'.

    Returns:
    pandas.DataFrame: DataFrame with updated 'Final Rating' column.
    """
    # Count non-empty cells in specified columns
    non_empty_counts = df[['Moodys Number', 'Fitch Number', 'S&P Number']].count(axis=1)

    # Initialize 'Final Rating' column with zeros
    df['Final Rating'] = 0

    # Check conditions and update 'Final Rating' column accordingly
    for idx, count in enumerate(non_empty_counts):
        if count > 0:  # At least one non-empty cell
            if count == 2:
                df.loc[idx, 'Final Rating'] = np.nanmax(df.loc[idx, ['Moodys Number', 'Fitch Number', 'S&P Number']])
            else:
                non_empty_values = df.loc[idx, ['Moodys Number', 'Fitch Number', 'S&P Number']].dropna()
                df.loc[idx, 'Final Rating'] = np.round(np.nanmedian(non_empty_values), 0) if not non_empty_values.empty else 0
    
    return df

# Apply the function to calculate 'Final Rating'
df_security = calculate_final_rating(df_security)


'''# Count non-empty cells in specified columns
non_empty_counts = df_security[['Moodys Number', 'Fitch Number', 'S&P Number']].count(axis=1)

# Initialize 'Final Rating' column with zeros
df_security['Final Rating'] = 0

# Check conditions and update 'Final Rating' column accordingly
for idx, count in enumerate(non_empty_counts):
    if count > 0:  # At least one non-empty cell
        if count == 2:
            df_security.loc[idx, 'Final Rating'] = np.nanmax(df_security.loc[idx, ['Moodys Number', 'Fitch Number', 'S&P Number']])
        else:
            non_empty_values = df_security.loc[idx, ['Moodys Number', 'Fitch Number', 'S&P Number']].dropna()
            df_security.loc[idx, 'Final Rating'] = np.round(np.nanmedian(non_empty_values), 0) if not non_empty_values.empty else 0
'''



# Convert the column from float type to int64 type
df_rating['#'] = df_rating['#'].astype('int64')

#df_rating

# Merge data frame df_security and df_rating on column Final Rating and column #
merged_df = pd.merge(df_security,df_rating[['#','S&P','CQS']],
                      left_on='Final Rating', right_on='#', how='left')

# Drop the '#' column as it is not needed
merged_df.drop(columns='#', inplace=True)

# Rename columns and drop unnecessary column
merged_df.rename(columns={'S&P': 'Final Letter Rating'}, inplace=True)

df_security = merged_df

#df_security.columns


df_security['Maturity Date'] = pd.to_datetime(df_security['Maturity Date'], dayfirst=True)


#df_security['Maturity Date'] 

#date_obj


df_security['Residual Maturity Days'] = (df_security['Maturity Date'] - date_obj).dt.days


# Function to categorize the residual maturity
def categorize_maturity(days):
    if days < 90:  # Less than 3 months (90 days)
        return '<3M'
    elif days == 90:  # Exactly 3 months (90 days)
        return '3M'
    else:  # Greater than 3 months
        return '>3M'

# Apply the categorization function to create a new column
df_security['3M < Residual Maturity > 3M'] = df_security['Residual Maturity Days'].apply(categorize_maturity)

#df_BOE_sector_code

#df_cust.columns

#df_security.columns


# Merge data frame df_security and df_cust on columnIssuer ID and column Customer Identifier
merged_df = pd.merge(df_security,df_cust[['Customer Identifier','Sector Code']],
                      left_on='Issuer Id', right_on='Customer Identifier', how='left')

# Drop the '#' column as it is not needed
merged_df.drop(columns='Customer Identifier', inplace=True)

# Rename columns and drop unnecessary column
#merged_df.rename(columns={'S&P': 'Final Letter Rating'}, inplace=True)

df_security = merged_df


# convert the type of the column BOE Sector Code from float 64 to int64

#df_BOE_sector_code['BOE Sector Code'] = df_BOE_sector_code['BOE Sector Code'].astype('int64')


#df_security.info()


# Merge data frame df_security and df_BOE_sector_code on column Sector Code and column BOE Sector Code
merged_df = pd.merge(df_security,df_BOE_sector_code[['BOE Sector Code','Category']],
                      left_on='Sector Code', right_on='BOE Sector Code', how='left')

# Drop the '#' column as it is not needed
merged_df.drop(columns='BOE Sector Code', inplace=True)

# Rename columns and drop unnecessary column
#merged_df.rename(columns={'S&P': 'Final Letter Rating'}, inplace=True)

df_security = merged_df


#df_BOE_sector_code.columns

#df_cust.columns

# Merge data frame df_security and df_cust on columnIssuer ID and column Customer Identifier
merged_df = pd.merge(df_security,df_cust[['Customer Identifier','Incorporation Code']],
                      left_on='Issuer Id', right_on='Customer Identifier', how='left')

# Drop the '#' column as it is not needed
merged_df.drop(columns='Customer Identifier', inplace=True)

# Rename columns and drop unnecessary column
#merged_df.rename(columns={'S&P': 'Final Letter Rating'}, inplace=True)

df_security= merged_df

# truncate the Inst_y_n to select specific column 
trunc_inst_Y_N = df_Inst_Y_N[['CODE', 'Institution elig']]

# Merge The trunc_inst_Y_N and df_security 
df_security = df_security.merge(trunc_inst_Y_N,left_on= 'Incorporation Code',right_on='CODE')

# drop the unwantedcolumn
df_security.drop(columns=['CODE'], inplace= True)


'''# Function to categorize based on conditions
def categorize_corporate_institution(row):
    if row['Incorporation Code'] == 'AE':
        return 'Corporates'
    elif row['Incorporation Code'] == 'QA' and row['Category'] == 'FIN':
        return 'Corporates'
    elif row['Category'] in ['FIN', 'SOV']:
        return 'Institutions'
    elif row['Category'] == 'GOV':
        return 'Central governments/central banks'
    elif row['Category'] == 'MDB':
        return 'Multilateral developments banks'
    
    else:
        return None  # Or any default value

# Apply the categorization function to create a new column
df_security['Corporate / Institution'] = df_security.apply(categorize_corporate_institution, axis=1)'''


# Define a function to categorize based on conditions
'''
 # Old Logic which  works in version 2.2.5
def categorize_corporate_institution(row):
    if row['Incorporation Code'] == 'AE':
        return 'Corporates'
    elif row['Incorporation Code'] == 'QA' and row['Category'] == 'FIN':
        return 'Corporates'
    elif row['Category'] in ['FIN', 'SOV']:
        return 'Institutions'
    elif row['Category'] == 'GOV':
        return 'Central governments/central banks'
    elif row['Category'] == 'MDB':
        return 'Multilateral developments banks'
    else:
        return 'Corporates'  # Default value'''

def categorize_corporate_institution(row):
    if row['Category'] in ['NON FIN', 'Corp']:
        return 'Corporates'
    elif row['Category'] == 'GOV':
        return 'Central governments/central banks'
    elif row['Category'] == 'MDB':
        return 'Multilateral developments banks'
    elif row['Institution elig'] == 'Y':
        return 'Institution'
    else:
        return 'Corporates'  # Default value
# Apply the categorization function to create a new column
df_security['Corporate / Institution'] = df_security.apply(categorize_corporate_institution, axis=1)


# Create the 'Investment Combo' column by concatenating values from 'CQS', 'Corporate / Institution', and '3M < Residual Maturity > 3M'
df_security['Investment Combo'] = df_security['CQS'] + '-' + df_security['Corporate / Institution'] + '-' + df_security['3M < Residual Maturity > 3M']


#df_security.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/Dailies_Python Values/trial/security.csv', index = False)

#df_rwlogic


# Merge the DataFrames on 'Currency Code'

df_security = df_security.merge(df_rwlogic, left_on='Investment Combo', right_on='Combination', how='left')


# Remove the duplicate rows from the data frame

df_security = df_security.drop_duplicates(keep='first')





# ### Account Master

def find_matching_lst_file(parent_folder):
    """
    Find the file in the given parent folder whose name starts with "acctMast" and has the extension ".LST".

    Args:
    parent_folder (str): Path of the parent folder.

    Returns:
    str: File path of the matching file, or None if no matching file is found.
    """
    files_in_folder = os.listdir(parent_folder)
    matching_files = [file for file in files_in_folder if file.startswith("acctMast") and file.lower().endswith(".lst")]
    
    if matching_files:
        lst_file_path = os.path.join(parent_folder, matching_files[0])
        return lst_file_path
    else:
        return None


# Ask for user input for the parent folder path
#parent_folder = input("\nEnter the path of the parent folder: ")
matching_file_path = find_matching_lst_file(parent_folder)

if matching_file_path:
    # Determine the CSV file path for saving
    csv_file_path = os.path.splitext(matching_file_path)[0] + ".csv"

    # Convert and save the file
    with open(matching_file_path, 'r', encoding='utf-8', newline='') as lst_file, \
            open(csv_file_path, 'w', encoding='utf-8', newline='') as csv_file:
        reader = csv.reader(lst_file, delimiter='|')
        writer = csv.writer(csv_file)

        # Read and write the headers
        headers = next(reader)
        writer.writerow(headers)    

        # Read and write the data rows
        for row in reader:
            writer.writerow(row)

    print(f"\nFile '{matching_file_path}' converted to '{csv_file_path}'.")
else:
    print("\nNo matching file found in the parent folder.")
    
# Read the account master to a data frame 
df_acctmast  = pd.read_csv(csv_file_path, index_col=False)


# Account Master Filteration

def process_file(csv_file_path):
    """
    Process a CSV file by extracting date from the file name and adding 36 days to the 'END DATE' column
    when SCHM CODE is 'NSB15', 'NSB16', or 'NSB17'. Select desired columns and return the DataFrame.

    Args:
    csv_file_path (str): Path of the CSV file to process.

    Returns:
    pd.DataFrame: Processed DataFrame with selected columns.
    """
    # Load CSV into DataFrame
    df_acctmast = pd.read_csv(csv_file_path, index_col=False)

    # Extract date from file name and add 36 days
    csv_file_name = csv_file_path.split("/")[-1]  # Extract the file name from the file path
    date_regex = r"\d{4}-\d{2}-\d{2}"
    date_match = re.search(date_regex, csv_file_name)

    if date_match:
        # Get the date string from the file name
        date_str = date_match.group(0)

        # Convert the date string to a date object
        current_date = datetime.strptime(date_str, "%Y-%m-%d").date()

        # Calculate the date after 35 days
        new_date = current_date + timedelta(days=35)  # Includes the date of the file thus 36 days else 35 days

        # Format the new date as dd-mm-yyyy
        new_date_str = new_date.strftime("%d-%m-%Y")

        # Update the "END DATE" column when SCHM CODE is 'NSB15', 'NSB16', or 'NSB17'
        df_acctmast.loc[df_acctmast['SCHM CODE'].isin(['NSB15', 'NSB16', 'NSB17']), 'END DATE'] = new_date_str

    # Select the desired columns
    truncated_columns = df_acctmast[['SOL ID', 'ACCOUNT OWNERSHIP', 'GL SUB HEAD CODE',
                                     'ACCT RPT CODE', 'ACCT NUMBER', 'ACCT NAME',
                                     'CUSTOMER ID', 'ACCT CCY', 'SCHM CODE',
                                     'SCHM TYPE', 'EOD BALANCE', 'EOD BALANCE(BILR)',
                                     'CLR BAL AMT', 'UN CLR BAL AMT', 'FUTURE BAL AMT',
                                     'ACCT OPN DATE', 'CR INT RATE', 'START DATE',
                                     'END DATE', 'FREE CODE 6', 'FREE TEXT 1']]

    # Put the selected columns in a Pandas DataFrame
    truncated_columns = pd.DataFrame(truncated_columns)

    return truncated_columns

def filter_and_process(df):
    """
    Perform various filtering and processing operations on the DataFrame.

    Args:
    df (pd.DataFrame): Input DataFrame.

    Returns:
    pd.DataFrame: Processed DataFrame.
    """
    # Step 0: Change the data type of column ACCT NUMBER to string format to avoid overflow issues
    df['ACCT NUMBER'] = df['ACCT NUMBER'].astype(str)
    
    # Step 1: Filter rows where 'EOD BALANCE(BILR)' is not equal to zero
    Column_L_filteration = df[df['EOD BALANCE(BILR)'] != 0]

    # Step 2: Filter rows where 'GL SUB HEAD CODE' is not equal to the specified values
    exclude_values = [80040, 80060, 80080, 85040, 85060, 85080, 89000]
    Column_C_filteration = Column_L_filteration[~Column_L_filteration['GL SUB HEAD CODE'].isin(exclude_values)]

    # Step 3: Create a new column named 'A/L'
    Column_C_filteration['A/L'] = ''

    # Step 4: Reorder the dataframe by 'GL SUB HEAD CODE' in ascending order
    Column_C_filteration.sort_values(by='GL SUB HEAD CODE', inplace=True, ascending=True)

    # Step 5: Assign 'A' or 'L' based on the values in 'EOD BALANCE(BILR)'
    Column_C_filteration.loc[Column_C_filteration['EOD BALANCE(BILR)'] < 0, 'A/L'] = 'A'
    Column_C_filteration.loc[Column_C_filteration['EOD BALANCE(BILR)'] >= 0, 'A/L'] = 'L'

    # Step 6: Allocation based on the GL code 89020
    filtered_89020 = Column_C_filteration[Column_C_filteration['GL SUB HEAD CODE'] == 89020]
    balance_sum = filtered_89020['EOD BALANCE(BILR)'].sum()
    Column_C_filteration.loc[Column_C_filteration['GL SUB HEAD CODE'] == 89020, 'A/L'] = 'A' if balance_sum < 0 else 'L'

    # Step 7: Filter the dataframe based on specific conditions
    conditions = Column_C_filteration['GL SUB HEAD CODE'].isin([28020, 28030, 28050, 55020, 55030, 55050, 55070, 55100])
    negative_balance = Column_C_filteration['EOD BALANCE(BILR)'] < 0
    positive_balance = Column_C_filteration['EOD BALANCE(BILR)'] >= 0
    Column_C_filteration.loc[conditions & negative_balance, 'A/L'] = 'A'
    Column_C_filteration.loc[conditions & positive_balance, 'A/L'] = 'L'

    # Step 8: Filter the dataframe based on specific conditions for GL SUB HEAD CODE and ACCT NUMBER
    gl_sub_condition = Column_C_filteration['GL SUB HEAD CODE'] == 55070
    acct_number_condition = Column_C_filteration['ACCT NUMBER'].isin([
        '61198265507010', '61198405507010', '61198265507009', '61199785507007',
        '61198405507007', '61198265507004', '61198405507001', '61198265507001'
    ])
    Column_C_filteration.loc[gl_sub_condition & acct_number_condition, 'A/L'] = 'A'

    gl_sub_condition_new = Column_C_filteration['GL SUB HEAD CODE'] == 28050
    acct_number_condition_new = Column_C_filteration['ACCT NUMBER'].isin([
        '61198262805008', '61198402805008', '61198402805002', '61198262805002'
    ])
    Column_C_filteration.loc[gl_sub_condition_new & acct_number_condition_new, 'A/L'] = 'A'

    # Step 9: Filter the dataframe based on specific conditions for GL SUB HEAD CODE and ACCT NUMBER
    mask = (Column_C_filteration['GL SUB HEAD CODE'] == 55100) & (
            Column_C_filteration['ACCT NUMBER'].isin(['61198265510002', '61198405510002'])
    )
    filtered_data = Column_C_filteration[mask]

    net_balance = filtered_data['EOD BALANCE(BILR)'].sum()
    Column_C_filteration.loc[mask, 'A/L'] = 'A' if net_balance < 0 else 'L'

    # Step 10: Filter the dataframe based on SCHM CODE values
    Column_C_filteration.loc[Column_C_filteration['SCHM CODE'].isin(['INCOM', 'EXPEN', 'NSB15', 'NSB16', 'NSB17']), 'A/L'] = 'L'

    # Step 11: Remove data where SCHM CODE is 'CONAS' or 'CONLI'
    Column_C_filteration = Column_C_filteration[~Column_C_filteration['SCHM CODE'].isin(['CONAS', 'CONLI'])]

    # Step 12: Update 'A/L' based on GL codes
    gl_sub_head_codes_A = [10010, 12030, 12261, 14060, 16050, 16250, 21000, 21010, 21040,
                           21510, 21520, 26000, 26010, 27000, 27010, 27040, 27050, 27070,
                           27100, 27143, 28020, 28030, 38060, 38070, 54000, 54060, 58000,
                           52010, 52030, 52040, 52070, 53000]
    Column_C_filteration.loc[Column_C_filteration['GL SUB HEAD CODE'].isin(gl_sub_head_codes_A), 'A/L'] = 'A'

    gl_sub_head_codes_L = [30000, 31000, 34000, 36000, 36020, 36060, 36070, 38000, 38010,
                           38030, 38040, 42000, 42010, 42610, 46050, 52010, 52030, 52040,
                           52070, 53000, 54010, 54020, 54050, 55030, 55050, 59000, 60000,
                           62000, 62010, 62020, 62030, 62040, 62060, 62070, 62080, 62090,
                           62100, 62110, 63000, 63030, 64010, 64020, 64040, 64050, 64060,
                           64200, 64210, 64300, 65000, 70000, 70030, 70050, 70060, 74020,
                           76000, 76010, 76010, 76030, 76040, 76045, 76090, 76120, 76200,
                           78010, 78040, 78050]
    Column_C_filteration.loc[Column_C_filteration['GL SUB HEAD CODE'].isin(gl_sub_head_codes_L), 'A/L'] = 'L'

    # Step 13: Check for OD accounts and allocate 'A' or 'L'
    od_condition = Column_C_filteration['SCHM CODE'].str.startswith('OD')
    Column_C_filteration.loc[od_condition & (Column_C_filteration['EOD BALANCE(BILR)'] < 0), 'A/L'] = 'A'
    Column_C_filteration.loc[od_condition & (Column_C_filteration['EOD BALANCE(BILR)'] >= 0), 'A/L'] = 'L'

    # Convert date columns to datetime format
    Column_C_filteration['ACCT OPN DATE'] = pd.to_datetime(Column_C_filteration['ACCT OPN DATE'], dayfirst=True).dt.strftime('%d/%m/%Y')
    Column_C_filteration['START DATE'] = pd.to_datetime(Column_C_filteration['START DATE'], dayfirst=True).dt.strftime('%d/%m/%Y')
    Column_C_filteration['END DATE'] = pd.to_datetime(Column_C_filteration['END DATE'], dayfirst=True).dt.strftime('%d/%m/%Y')

    return Column_C_filteration

# Usage example
truncated_acctMast = process_file(csv_file_path)  # Assuming this function reads and preprocesses the CSV
truncated_acctMast = filter_and_process(truncated_acctMast)

'''
def filter_and_process(df):
    """
    Perform various filtering and processing operations on the DataFrame.

    Args:
    df (pd.DataFrame): Input DataFrame.

    Returns:
    pd.DataFrame: Processed DataFrame.
    """
    # Step 0: Change the data type of column ACCT NUMBER to intiger format
    df['ACCT NUMBER'] = df['ACCT NUMBER'].astype(int)
    
    # Step 1: Filter rows where 'EOD BALANCE(BILR)' is not equal to zero
    Column_L_filteration = df[df['EOD BALANCE(BILR)'] != 0]

    # Step 2: Filter rows where 'GL SUB HEAD CODE' is not equal to the specified values
    exclude_values = [80040, 80060, 80080, 85040, 85060, 85080, 89000]
    Column_C_filteration = Column_L_filteration[~Column_L_filteration['GL SUB HEAD CODE'].isin(exclude_values)]

    # Step 3: Create a new column named 'A/L'
    Column_C_filteration['A/L'] = ''

    # Step 4: Reorder the dataframe by 'GL SUB HEAD CODE' in ascending order
    Column_C_filteration.sort_values(by='GL SUB HEAD CODE', inplace=True, ascending=True)

    # Step 5: Assign 'A' or 'L' based on the values in 'EOD BALANCE(BILR)'
    Column_C_filteration.loc[Column_C_filteration['EOD BALANCE(BILR)'] < 0, 'A/L'] = 'A'
    Column_C_filteration.loc[Column_C_filteration['EOD BALANCE(BILR)'] >= 0, 'A/L'] = 'L'

    # Step 6: Allocation based on the GL code 89020
    filtered_89020 = Column_C_filteration[Column_C_filteration['GL SUB HEAD CODE'] == 89020]
    balance_sum = filtered_89020['EOD BALANCE(BILR)'].sum()
    Column_C_filteration.loc[Column_C_filteration['GL SUB HEAD CODE'] == 89020, 'A/L'] = 'A' if balance_sum < 0 else 'L'

    # Step 7: Filter the dataframe based on specific conditions
    conditions = Column_C_filteration['GL SUB HEAD CODE'].isin([28020, 28030, 28050, 55020, 55030, 55050, 55070, 55100])
    negative_balance = Column_C_filteration['EOD BALANCE(BILR)'] < 0
    positive_balance = Column_C_filteration['EOD BALANCE(BILR)'] >= 0
    Column_C_filteration.loc[conditions & negative_balance, 'A/L'] = 'A'
    Column_C_filteration.loc[conditions & positive_balance, 'A/L'] = 'L'

    # Step 8: Filter the dataframe based on specific conditions for GL SUB HEAD CODE and ACCT NUMBER
    gl_sub_condition = Column_C_filteration['GL SUB HEAD CODE'] == 55070
    acct_number_condition = Column_C_filteration['ACCT NUMBER'].isin([
        61198265507010, 61198405507010, 61198265507009, 61199785507007,
        61198405507007, 61198265507004, 61198405507001, 61198265507001
    ])
    Column_C_filteration.loc[gl_sub_condition & acct_number_condition, 'A/L'] = 'A'

    gl_sub_condition_new = Column_C_filteration['GL SUB HEAD CODE'] == 28050
    acct_number_condition_new = Column_C_filteration['ACCT NUMBER'].isin([
        61198262805008, 61198402805008, 61198402805002, 61198262805002
    ])
    Column_C_filteration.loc[gl_sub_condition_new & acct_number_condition_new, 'A/L'] = 'A'

    # Step 9: Filter the dataframe based on specific conditions for GL SUB HEAD CODE and ACCT NUMBER
    mask = (Column_C_filteration['GL SUB HEAD CODE'] == 55100) & (
            Column_C_filteration['ACCT NUMBER'].isin([61198265510002, 61198405510002])
    )
    filtered_data = Column_C_filteration[mask]

    net_balance = filtered_data['EOD BALANCE(BILR)'].sum()
    Column_C_filteration.loc[mask, 'A/L'] = 'A' if net_balance < 0 else 'L'

    # Step 10: Filter the dataframe based on SCHM CODE values
    Column_C_filteration.loc[Column_C_filteration['SCHM CODE'].isin(['INCOM', 'EXPEN', 'NSB15', 'NSB16', 'NSB17']), 'A/L'] = 'L'

    # Step 11: Remove data where SCHM CODE is 'CONAS' or 'CONLI'
    Column_C_filteration = Column_C_filteration[~Column_C_filteration['SCHM CODE'].isin(['CONAS', 'CONLI'])]

    # Step 12: Update 'A/L' based on GL codes
    gl_sub_head_codes_A = [10010, 12030, 12261, 14060, 16050, 16250, 21000, 21010, 21040,
                           21510, 21520, 26000, 26010, 27000, 27010, 27040, 27050, 27070,
                           27100, 27143, 28020, 28030, 38060, 38070, 54000, 54060, 58000, 89010,
                           52010, 52030, 52040, 52070, 53000]
    Column_C_filteration.loc[Column_C_filteration['GL SUB HEAD CODE'].isin(gl_sub_head_codes_A), 'A/L'] = 'A'

    gl_sub_head_codes_L = [30000, 31000, 34000, 36000, 36020, 36060, 36070, 38000, 38010,
                           38030, 38040, 42000, 42010, 42610, 46050, 52010, 52030, 52040,
                           52070, 53000, 54010, 54020, 54050, 55030, 55050, 59000, 60000,
                           62000, 62010, 62020, 62030, 62040, 62060, 62070, 62080, 62090,
                           62100, 62110, 63000, 63030, 64010, 64020, 64040, 64050, 64060,
                           64200, 64210, 64300, 65000, 70000, 70030, 70050, 70060, 74020,
                           76000, 76010, 76010, 76030, 76040, 76045, 76090, 76120, 76200,
                           78010, 78040, 78050]
    Column_C_filteration.loc[Column_C_filteration['GL SUB HEAD CODE'].isin(gl_sub_head_codes_L), 'A/L'] = 'L'

    # Step 13: Check for OD accounts and allocate 'A' or 'L'
    od_condition = Column_C_filteration['SCHM CODE'].str.startswith('OD')
    Column_C_filteration.loc[od_condition & (Column_C_filteration['EOD BALANCE(BILR)'] < 0), 'A/L'] = 'A'
    Column_C_filteration.loc[od_condition & (Column_C_filteration['EOD BALANCE(BILR)'] >= 0), 'A/L'] = 'L'

    # Convert date columns to datetime format
    Column_C_filteration['ACCT OPN DATE'] = pd.to_datetime(Column_C_filteration['ACCT OPN DATE'], dayfirst=True).dt.strftime('%d/%m/%Y')
    Column_C_filteration['START DATE'] = pd.to_datetime(Column_C_filteration['START DATE'], dayfirst=True).dt.strftime('%d/%m/%Y')
    Column_C_filteration['END DATE'] = pd.to_datetime(Column_C_filteration['END DATE'], dayfirst=True).dt.strftime('%d/%m/%Y')

    return Column_C_filteration

# Usage:

truncated_acctMast = process_file(csv_file_path)
truncated_acctMast = filter_and_process(truncated_acctMast)

'''


# Get user input for the parent folder
#destination_folder = input("Enter the parent folder path: ")

# Create a new copy for the Assets and Liabilities
Assets = truncated_acctMast[truncated_acctMast['A/L'] == 'A']
Liabilities = truncated_acctMast[truncated_acctMast['A/L'] == 'L']

#Assets.info()

# Convert the type of the column ACCT NUMBER from float 64 to int64

Assets['ACCT NUMBER'] = Assets['ACCT NUMBER'].astype('int64')

# Define the list of account numbers to be replaced
acct_nums_to_replace = [
    61198261406001, 61198401406001, 61198401406002, 61198261406002, 61198262805002,
    61198402805008, 61198402805002, 61198262805008, 61198405406001, 61198265406001,
    61198265507010, 61198405507010, 61198405507001, 61198265507001, 61198401605001, 
    61198401625001, 61198262805006, 61198261605001, 61199781605001
]

# Condition to check for the account numbers to replace
condition = Assets['ACCT NUMBER'].isin(acct_nums_to_replace)

# Update the 'BILR' column based on the condition
Assets.loc[condition, 'EOD BALANCE(BILR)'] = 0

# Define a dictionary to map DataFrame names to their corresponding variables
data_frames = {
    "Assets": Assets,
    "Liabilites": Liabilities,
    "Trunc_Account_Master": truncated_acctMast,
    
}

# Loop through the dictionary and save each DataFrame as .txt and .csv files
for file_name, data_frame in data_frames.items():
    # Save as .txt file
    txt_file_path = os.path.join(destination_folder, f"{file_name}.txt")
    data_frame.to_csv(txt_file_path, sep=',', index=False)
    
    # Save as .csv file
    txt_file_path = os.path.join(destination_folder, f"{file_name}.csv")
    data_frame.to_csv(txt_file_path, sep=',', index=False)
    
    ## Save as .xlsx file
    #txt_file_path = os.path.join(destination_folder, f"{file_name}.xlsx")
    #data_frame.to_excel(txt_file_path,index=False)

# Confirm the successful saving of the files
print("DataFrames saved in both .txt and .csv formats in the provided folder.")
#Assets.info()


# ### BTL LTV File 

def read_dependency_coding_file(parent_folder):
    """
    Read specific worksheets from the Excel file in the specified parent folder.

    Args:
    parent_folder (str): Path of the parent folder.

    Returns:
    tuple: Tuple containing DataFrames for each worksheet (ratings_df, df_BOE_sector_code, df_rwlogic).
    """
    # Find the file matching the specified pattern
    file_pattern = os.path.join(parent_folder, 'FRP Data*.xlsx')
    file_list = glob.glob(file_pattern)

    if len(file_list) == 0:
        print("No matching Excel file found.")
        return None

    # Assuming there's only one matching file, use the first one in the list
    file_path = file_list[0]

    # Use xlwings to load the Excel file
    app = xw.App(visible=False)  # Create an Excel app instance (hidden)
    wb = app.books.open(file_path)  # Open the Excel file

    # Initialize DataFrames for each worksheet
    df_frp_acc = None
    df_frp_mortgage = None
    

    # Loop through each worksheet and load data into respective DataFrames
    for sheet in wb.sheets:
        if sheet.name == 'FRP-Account':
            df_frp_acc = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        elif sheet.name == 'FRP-Mortgage':
            df_frp_mortgage = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        #elif sheet.name == 'Risk Weight Logic':
        #    df_rwlogic = sheet.used_range.options(pd.DataFrame, index=False, header=True).value

    wb.close()  # Close the workbook
    app.quit()  # Quit the Excel app

    return df_frp_acc, df_frp_mortgage

# Example usage:
#parent_folder = "path/to/parent/folder"

df_frp_acc, df_frp_mortgage = read_dependency_coding_file(parent_folder)
#dfs = read_dependency_coding_file(parent_folder)

if df_frp_acc is not None:
    print("Capital Dependency data loaded successfully.")
    # Proceed with using the DataFrames as needed
else:
    print("No Capital Dependency data loaded.")

#df_frp_acc
#df_frp_acc.info()

# Select the desired columns
trunc_frp_col = df_frp_acc[['Account Number', 'Account Short Name', 'Account Status Description', 'Total Balance', 
                            'Capital Balance','Months Since Active Arrears Start Date', 'Original Expiry Date','Loan Expiry Date', 'Redemption Date',
                            'Inception Market Valuation', 'Original Loan Amount', 'Latest Valuation Amount','Original LTV', 
                            'Current LTV', 'Customer ID', 'Customer Type']]

# Put the selected columns in a Pandas DataFrame
trunc_df_frp_acc = pd.DataFrame(trunc_frp_col)

#Remove all the negative and zero balances
trunc_df_frp_acc = trunc_df_frp_acc[trunc_df_frp_acc['Total Balance']>0]

# Loop through the dictionary and save each DataFrame as .txt and .csv files
for file_ext in ['.txt', '.csv']:
    # Construct the file path with the desired file name and extension
    file_path = os.path.join(destination_folder, f"BTLLTV{file_ext}")

    # Check the file extension and save accordingly
    if file_ext == '.txt':
        trunc_df_frp_acc.to_csv(file_path, sep=',', index=False)  # Save as .txt with comma-separated values
    elif file_ext == '.csv':
        trunc_df_frp_acc.to_csv(file_path, sep=',', index=False)  # Save as .csv

# Confirm the successful saving of the files
print("DataFrames saved in both .txt and .csv formats in the provided folder.")


def replace_and_sum_assets(Assets, trunc_df_frp_acc):
    """
    Sum the values in the 'Total Balance' column from 'trunc_df_frp_acc',
    fetch the row where 'ACCT NUMBER' is 61198262100001 from 'Assets',
    and replace the 'EOD BALANCE(BILR)' value in 'Assets' with the calculated sum.

    Parameters:
    Assets (pd.DataFrame): The 'Assets' DataFrame.
    trunc_df_frp_acc (pd.DataFrame): The 'trunc_df_frp_acc' DataFrame.

    Returns:
    pd.DataFrame: Updated 'Assets' DataFrame.
    """

    acct_number = 61198262100001

    # Ensure the specific 'ACCT NUMBER' exists in both DataFrames
    if acct_number not in Assets['ACCT NUMBER'].values:
        raise ValueError(f"ACCT NUMBER {acct_number} not found in 'Assets' DataFrame.")
    #if acct_number not in trunc_df_frp_acc['ACCT NUMBER'].values:
       # raise ValueError(f"ACCT NUMBER {acct_number} not found in 'trunc_df_frp_acc' DataFrame.")

    # Calculate the sum of 'Total Balance' from 'trunc_df_frp_acc'
    BTL_LTV_SUM = trunc_df_frp_acc['Total Balance'].sum()

    # Fetch the row/data when 'ACCT NUMBER' in 'Assets' is 61198262100001
    BTL_acccount = Assets[Assets['ACCT NUMBER'] == acct_number]

    # Calculate the final sum
    final_sum = BTL_acccount['EOD BALANCE(BILR)'].values[0] + BTL_LTV_SUM

    # Replace the value in 'EOD BALANCE(BILR)' for the row where 'ACCT NUMBER' is 61198262100001
    Assets.loc[Assets['ACCT NUMBER'] == acct_number, 'EOD BALANCE(BILR)'] = final_sum

    return Assets


# Perform the replacement and sum calculation
Assets = replace_and_sum_assets(Assets, trunc_df_frp_acc)

#trunc_df_frp_acc.info()
#Assets.columns


# ### MM Placements File (Loand Dep)


#df_repay.head(5)

#df_loandep.head(5)

# Filter the DataFrame based on the conditions
capital_df_loandep = df_loandep[~df_loandep['Product ID'].isin(['LAA', 'TDA'])]

#capital_df_loandep.info()


# Convert the Date columns to date time format
capital_df_loandep['Start Date'] = pd.to_datetime(capital_df_loandep['Start Date'], dayfirst=True)
capital_df_loandep['Maturity Date'] = pd.to_datetime(capital_df_loandep['Maturity Date'], dayfirst=True)
#capital_df_loandep.info()

# Merge the DataFrames on 'Currency Code'
capital_df_loandep = capital_df_loandep.merge(df_exch, left_on='Currency', right_on='Currency Code', how='left')

# Drop the column Currency code
columns_to_drop = ['Currency Code']
capital_df_loandep = capital_df_loandep.drop(columns=columns_to_drop)
#capital_df_loandep.info()

# Define the comparison function
def get_deal_status(start_date, date_obj):
    return 'STARTED' if start_date <= date_obj else 'NOT STARTED'
# Apply the function to create the 'Deal Status' column
capital_df_loandep['Deal Status'] = capital_df_loandep.apply(lambda row: get_deal_status(row['Start Date'], date_obj), axis=1)

#capital_df_loandep['Deal Status']
# Coulumnd Principle Amount (GBP)

def principal_amt_calc(row):
    if row['Currency'] != 'GBP':
        return round (row['Amount Currency'] / row['Exchange Rate'], 2)
    else:
        return round(row['Amount Currency'],2)
# Apply the  function  to create a new column
capital_df_loandep['Principal Amount (GBP)'] = capital_df_loandep.apply(principal_amt_calc, axis=1)

# Column Borrowing/Placement#

def borrow_place(row):
    if row['Amount Currency'] > 0:
        return "Placement"
    else:
        return "Borrowing"

# Apply the function to create a new column
capital_df_loandep['Borrowing/PLacement'] = capital_df_loandep.apply(borrow_place, axis=1)


capital_df_loandep['Residual Maturity Days'] = (capital_df_loandep['Maturity Date'] - date_obj).dt.days

# Ensure Customer Identifier and Counterparty Name are consistent for lookups
df_cust['Customer Identifier'] = df_cust['Customer Identifier'].astype(str)
capital_df_loandep['Cif id '] = capital_df_loandep['Cif id '].astype(str)

# Merge df_cust for Sector Code, Residency Code, and Incorporation Code
capital_df_loandep = capital_df_loandep.merge(
    df_cust[['Customer Identifier', 'Sector Code', 'Residency Code', 'Incorporation Code']],
    left_on='Cif id ',
    right_on='Customer Identifier',
    how='left'
)
# Apply conditional logic for Country Code
capital_df_loandep['Country Code'] = capital_df_loandep.apply(
    lambda row: row['Residency Code'] if row['Sector Code'] == 3500 else row['Incorporation Code'],
    axis=1
)

# Merge The trunc_inst_Y_N and df_security 
capital_df_loandep = capital_df_loandep.merge(trunc_inst_Y_N,left_on= 'Country Code', right_on='CODE')

# drop the unwantedcolumn
capital_df_loandep.drop(columns=['CODE','Residency Code', 'Incorporation Code'], inplace= True)


def categorize_rwmmp(row):
    if row['Institution elig'] == 'N':
        return 1
    elif row['Residual Maturity Days'] > 90:  # Greater than 90 days
        return 0.5
    elif row['Residual Maturity Days'] <= 90:  
        return 0.2

# Apply the categorization function to create a new column
capital_df_loandep['RW_MMP'] = capital_df_loandep.apply(categorize_rwmmp, axis=1)

#capital_df_loandep

def calculate_rw_exposure(row):
    return row['Principal Amount (GBP)'] * row['RW_MMP']

# Apply the function to create a new column 'RW_EXPOSURE'
capital_df_loandep['RW_EXPOSURE'] = capital_df_loandep.apply(calculate_rw_exposure, axis=1)

def calculate_rw_capital(row):
    return row['RW_EXPOSURE'] * 0.08

# Apply the function to create a new column 'RW_EXPOSURE'
capital_df_loandep['REQ_Capital'] = capital_df_loandep.apply(calculate_rw_capital, axis=1)

capital_df_loandep['ORG Maturity Days'] = (capital_df_loandep['Maturity Date'] - capital_df_loandep['Start Date']).dt.days
#capital_df_loandep.columns





# Reorder/ Restructre the data frame by keeping only required/essential columns in the dataframe

capital_df_loandep = capital_df_loandep[['Deal Reference', 'Cif id ','Currency', 'Amount Currency',
                                        'Start Date','Maturity Date','Deal Status', 'Principal Amount (GBP)',
                                         'Borrowing/PLacement', 'RW_MMP', 'RW_EXPOSURE', 'REQ_Capital',
                                         'ORG Maturity Days', 'Residual Maturity Days','Sector Code','Country Code','Institution elig']]

# Sorting the DataFrame in descending order based on 'Amount Currency'
capital_df_loandep = capital_df_loandep.sort_values(by='Amount Currency', ascending=False)

#capital_df_loandep.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/Dailies_Python Values/trial/calpitalLoandep1.csv', index = False)


# ### Pipeline Tab

def read_and_process_pipeline_file(parent_folder):
    """
    Read specific worksheets from the Excel file in the specified parent folder,
    filter the data, perform necessary calculations, and save the result as a CSV file.

    Args:
    parent_folder (str): Path of the parent folder.

    Returns:
    pd.DataFrame: Processed DataFrame containing pipeline data.
    """
    # Find the file matching the specified pattern
    file_pattern = os.path.join(parent_folder, 'Pipeline Summary Report*.xlsx')
    file_list = glob.glob(file_pattern)

    if len(file_list) == 0:
        print("No matching Excel file found.")
        return None

    # Assuming there's only one matching file, use the first one in the list
    file_path = file_list[0]

    # Use xlwings to load the Excel file
    app = xw.App(visible=False)  # Create an Excel app instance (hidden)
    wb = app.books.open(file_path)  # Open the Excel file

    # Initialize DataFrame for the worksheet
    df_pipeline = None

    # Loop through each worksheet and load data into respective DataFrame
    for sheet in wb.sheets:
        if sheet.name == 'Case Summary':
            df_pipeline = sheet.used_range.options(pd.DataFrame, index=False, header=True).value

    wb.close()  # Close the workbook
    app.quit()  # Quit the Excel app

    if df_pipeline is not None:
        print("Pipeline data loaded successfully.")

        # Filter the data frame to have only Legals Instructed and COT Received
        df_pipeline = df_pipeline[df_pipeline['Application Stage'].isin(['Legals Instructed', 'COT Received'])]

        # Perform the necessary calculations
        df_pipeline['CCF'] = 0.2
        df_pipeline['RW'] = 0.35
        df_pipeline['Exposure Post CCF'] = df_pipeline['Advance'] * df_pipeline['CCF']
        df_pipeline['RWA'] = df_pipeline['Exposure Post CCF'] * df_pipeline['RW']

        # Define the output CSV file path
        output_csv_path = os.path.join(parent_folder, 'BTL Pipeline.csv')

        # Save the DataFrame to CSV
        df_pipeline.to_csv(output_csv_path, index=False)

        print(f"Pipeline data saved successfully to {output_csv_path}.")
        return df_pipeline
    else:
        print("Error in loading Pipeline data.")
        return None

# Usage
#parent_folder = 'path_to_your_parent_folder'
df_pipeline = read_and_process_pipeline_file(parent_folder)

'''def read_dependency_coding_file(parent_folder):
    """
    Read specific worksheets from the Excel file in the specified parent folder.

    Args:
    parent_folder (str): Path of the parent folder.

    Returns:
    tuple: Tuple containing DataFrames for each worksheet (ratings_df, df_BOE_sector_code, df_rwlogic).
    """
    # Find the file matching the specified pattern
    file_pattern = os.path.join(parent_folder, 'Pipeline Summary Report*.xlsx')
    file_list = glob.glob(file_pattern)

    if len(file_list) == 0:
        print("No matching Excel file found.")
        return None

    # Assuming there's only one matching file, use the first one in the list
    file_path = file_list[0]

    # Use xlwings to load the Excel file
    app = xw.App(visible=False)  # Create an Excel app instance (hidden)
    wb = app.books.open(file_path)  # Open the Excel file

    # Initialize DataFrames for each worksheet
    df_pipeline = None
    #df_frp_mortgage = None
    

    # Loop through each worksheet and load data into respective DataFrames
    for sheet in wb.sheets:
        if sheet.name == 'Case Summary':
            df_pipeline = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        #elif sheet.name == 'FRP-Mortgage':
        #    df_frp_mortgage = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        #elif sheet.name == 'Risk Weight Logic':
        #    df_rwlogic = sheet.used_range.options(pd.DataFrame, index=False, header=True).value

    wb.close()  # Close the workbook
    app.quit()  # Quit the Excel app

    return df_pipeline

# Example usage:
#parent_folder = "path/to/parent/folder"

df_pipeline = read_dependency_coding_file(parent_folder)
#dfs = read_dependency_coding_file(parent_folder)

if df_pipeline is not None:
    print("Pipeline data loaded successfully.")
    # Proceed with using the DataFrames as needed
else:
    print("Error in loading Pipeline data.")

#df_pipeline


#df_pipeline.columns

# Filter the data frame to have only Legals Instruceted And COT received
df_pipeline = df_pipeline[df_pipeline['Application Stage'].isin(['Legals Instructed', 'COT Received'])]

df_pipeline['CCF'] = 0.2
df_pipeline['RW'] = 0.35
df_pipeline['Exposure Post CCF'] = df_pipeline['Advance']*0.2
df_pipeline['RWA'] = df_pipeline['Advance']*0.2*0.35

#df_pipeline.to_csv('C:/Users/sbiuser/Downloads/TRIAL & ERROR Files/Dailies_Python Values/trial/trial.csv', index = False)
'''

# ### Account Balance File

def update_accbal_values(df_accbal):
    """
    Update values in DataFrame df_accbal based on a specified condition.

    Args:
    df_accbal (pd.DataFrame): DataFrame containing account balance data.

    Returns:
    pd.DataFrame: Updated DataFrame with modified values based on the condition.
    """
    # Condition to check for Cif id 800001122
    condition = (
        (df_accbal['Cif id '] == 800001122) &
        (df_accbal['Account Number'] == 96306438) &
        ((df_accbal['Risk Weight'] != 100) |
        (df_accbal['Type of Asset'] != 'CORPORATE LOAN-NON CRE') |
        (df_accbal['Secured/unsecured'] != 'SECURED') |
        (df_accbal['Exposure Type'] != 'CORPORATE LOAN: NON PROPERTY'))
    )

    # Update the values based on the condition
    df_accbal.loc[condition, ['Risk Weight', 'Type of Asset', 'Secured/unsecured', 'Exposure Type']] = [100, 'CORPORATE LOAN-NON CRE', 'SECURED', 'CORPORATE LOAN: NON PROPERTY']

    return df_accbal

# Usage:
filtered_df_accbal = update_accbal_values(df_accbal)


capital_security = df_security[['Branch Id', 'Issuer Id', 'Product Type', 'MDB Flag', 'Deal Reference',
                                   'Investment Type', 'Trading Indicator', 'Encumbered Flag', 'Currency',
                                   'Nominal Amount CCY', 'Principal Amount CCY', 'Market Value CCY', 'MTM',
                                   'Coupon Rate', 'Accrued Interest CCY', 'Start Date', 'Maturity Date',
                                   'Settlement Date', 'Fxd Flt Indicator', 'Next Reset Date','Security ID',
                                   'LAB Indicator', 'HQ Indicator','Senior Tranche Indicator', 'Issued/purchased flag',
                                   'USGS Indicator','LTST Indicator', 'Own Funds Indicator', 'Own Funds Amount',
                                   'Rehypothetication Indicator', 'Watch List Indicator','Default Indicator',
                                   'LCR Government Protected Indicator','LCRGteeIntOrgID', 'LCRIntOrgID',
                                   'LCRRequirementsMet','LiquidAssetQualityType', 'LiquidFacilityType', 'Notional Amount',
                                   'OtherLCRLiabilities', 'OtherTransAssets', 'Short Term Moody',
                                   'Long Term Moody', 'Short Term Fitch', 'Long Term Fitch','Short Term S P',
                                   'Long Term S P', 'Purchase Cost', 'Purchase Price','Risk Weight', 'PD external',
                                   'LGD External', 'PD Internal','LGD INTERNAL', 'MTM GBP', 'Principle GBP','Book Value GBP',
                                   'Residual Maturity Days','Sector Code','Category','Moodys Number', 'Fitch Number',
                                   'S&P Number','Final Rating', 'Final Letter Rating', 'CQS','Incorporation Code',
                                   'Corporate / Institution', 'Investment Combo','RW', 'Customer Name']]

# Group the DataFrame by 'Product Type' and calculate the sum of 'Balance BILR'
#az = capital_security.groupby('Product Type')['Book Value GBP'].sum().reset_index()
#az



# ### Capital file Operation

# Find the file matching the specified pattern
file_pattern = os.path.join(parent_folder, 'New File_Capital RWA*.xlsx')
file_list = glob.glob(file_pattern)

if len(file_list) == 0:
    print("No matching Excel file found.")
    #return None

# Assuming there's only one matching file, use the first one in the list
file_path = file_list[0]

# Open the Excel workbook and make Excel visible
app = xw.App(visible=True)
capital_file_wb = app.books.open(file_path)

# Print the names of all worksheets
print("Worksheets in the workbook:")
for sheet in capital_file_wb.sheets:
    print(sheet.name)


# Update the Worksheet named Assets

def write_assets_to_worksheet(ws, df, start_row=5, start_column=1):
    """
    Write data from a DataFrame to a specified worksheet starting from a given cell.

    Args:
    ws (xlwings.Sheet): Worksheet to write data to.
    df (pd.DataFrame): DataFrame containing the data to be written.
    start_row (int): Starting row for writing data. Default is 5.
    start_column (int): Starting column for writing data. Default is 1.

    Returns:
    None
    """
    # Get the range starting from the specified row and column
    end_row = ws.cells.last_cell.row
    #end_row = start_row + len(df) - 1  # Number of rows based on DataFrame length
    end_column = start_column + df.shape[1] - 1  # Number of columns based on DataFrame shape
    range_to_clear = ws.range((start_row, start_column), (end_row, end_column))

    # Clear the contents of the range
    range_to_clear.clear_contents()

    # Write the data (excluding headers) to the worksheet
    ws.range((start_row, start_column)).value = df.values

# Example usage:
# Assuming ws_capital_asset and Assets are already defined
ws_capital_asset = capital_file_wb.sheets["Assets"]

# Define the columns to keep in the subset
columns_to_keep = ['SOL ID', 'ACCOUNT OWNERSHIP', 'GL SUB HEAD CODE', 'ACCT RPT CODE',
                   'ACCT NUMBER', 'ACCT NAME', 'CUSTOMER ID', 'ACCT CCY', 'SCHM CODE',
                   'SCHM TYPE', 'EOD BALANCE(BILR)']

# Extract only the desired columns from the Assets DataFrame
assets_subset = Assets[columns_to_keep]

# Use the defined function to write the subset to the worksheet
write_assets_to_worksheet(ws_capital_asset, assets_subset)

#Assets


# Update the Worksheet named Liabilities

def write_liabilites_to_worksheet(ws, start_row=2, start_column=1):
    """
    Clear the contents of a specified worksheet starting from a given cell.

    Args:
    ws (xlwings.Sheet): Worksheet to clear.
    start_row (int): Starting row for clearing data. Default is 6.
    start_column (int): Starting column for clearing data. Default is 1.

    Returns: the new values from data frame Liabilities
    None
    """
    # Get the last used row and column
    last_row = ws.cells.last_cell.row
    last_column = ws.cells.last_cell.column

    # Get the range starting from the specified row and column to the last row and column
    range_to_clear = ws.range((start_row, start_column), (last_row, last_column))

    # Clear the contents of the range
    range_to_clear.clear_contents()
    
    # Write the data (excluding headers) to the worksheet
    ws.range((start_row, start_column)).value = Liabilities.values

# Example usage:
# Assuming ws_MMP is the worksheet named "MM Placements"
ws_liabilities = capital_file_wb.sheets["Liabilities"]

# Clear the worksheet starting from the 6th row across all columns
write_liabilites_to_worksheet(ws_liabilities)


'''# Select the worksheet named "Lombard Repay"
ws_capital_asset = wb.sheets["Assets"]

# Get the range starting from the 5th row
start_row = 5
start_column = 1  # Assuming you want to start from the first column
end_row = ws_capital_asset.api.Rows.Count
end_column = 11
range_to_clear = ws_capital_asset.range((start_row, start_column), (end_row, end_column))

# Clear the contents of the range
range_to_clear.clear_contents()

# Extract only the desired columns from the Assets DataFrame
columns_to_keep = ['SOL ID', 'ACCOUNT OWNERSHIP', 'GL SUB HEAD CODE', 'ACCT RPT CODE',
                   'ACCT NUMBER', 'ACCT NAME', 'CUSTOMER ID', 'ACCT CCY', 'SCHM CODE',
                   'SCHM TYPE','EOD BALANCE(BILR)']
assets_subset = Assets[columns_to_keep]

# Write the data (excluding headers) to the worksheet
ws_capital_asset.range("A5").value = assets_subset.values'''

#df_limit.info()


# Update the Worksheet named LimitSub

def write_capital_limit_to_worksheet(ws, df, start_row=6, start_column=1, end_column=16):
    """
    Write data from a DataFrame to a specified worksheet starting from a given cell.

    Args:
    ws (xlwings.Sheet): Worksheet to write data to.
    df (pd.DataFrame): DataFrame containing the data to be written.
    start_row (int): Starting row for writing data. Default is 6.
    start_column (int): Starting column for writing data. Default is 1.

    Returns:
    None
    """
    
    # Get the range starting from the specified row and column
    end_row = start_row + len(df) - 1  # Number of rows based on DataFrame length
    #end_column = start_column + df.shape[1] - 1  # Number of columns based on DataFrame shape
    range_to_clear = ws.range((start_row, start_column), (end_row, end_column))

    # Clear the contents of the range
    range_to_clear.clear_contents()

    # Write the data (excluding headers) to the worksheet
    ws.range((start_row, start_column)).value = df.values

# Example usage:
# Assuming ws_capital_limit is the worksheet named "LimitSub"
ws_capital_limit = capital_file_wb.sheets["LimitSub"]

# Filter the DataFrame based on the conditions
capital_df_limit = df_limit[(df_limit['Reference Code'] == "CEL05") | (df_limit['Reference Code'] == "CEL06")]

# Define the columns to keep in the subset
columns_to_keep = ['Branch id', 'Customer Identifier', 'Deal Reference', 'Reference Code',
                   'Facility / Commitment Type', 'Parent Limit', 'Currency Code',
                   'Facility Amount', 'Undrawn Amount', 'Comitted / Uncomitted','Start Date',
                   'Maturity Date', 'Secured/Unsecured indicator','Pipeline Indicator','Scheme Code', 'CCF']

# Extract only the desired columns from the Assets DataFrame
limit_subset = capital_df_limit[columns_to_keep]

# Use the defined function to write the subset to the worksheet
write_capital_limit_to_worksheet(ws_capital_limit, limit_subset)

def write_capital_MMP_to_worksheet(ws, start_row=6, start_column=2):
    """
    Clear the contents of a specified worksheet starting from a given cell.

    Args:
    ws (xlwings.Sheet): Worksheet to clear.
    start_row (int): Starting row for clearing data. Default is 6.
    start_column (int): Starting column for clearing data. Default is 1.

    Returns: the new values from data frame capital_df_loandep
    None
    """
    # Get the last used row and column
    last_row = ws.cells.last_cell.row
    last_column = ws.cells.last_cell.column

    # Get the range starting from the specified row and column to the last row and column
    range_to_clear = ws.range((start_row, start_column), (last_row, last_column))

    # Clear the contents of the range
    range_to_clear.clear_contents()
    
    # Write the data (excluding headers) to the worksheet
    ws.range((start_row, start_column)).value = capital_df_loandep.values

# Example usage:
# Assuming ws_MMP is the worksheet named "MM Placements"
ws_MMP = capital_file_wb.sheets["MM Placements"]

# Clear the worksheet starting from the 6th row across all columns
write_capital_MMP_to_worksheet(ws_MMP)

#df_security

def clear_and_write_data_to_sheet(ws, df, start_row=5, start_column=1, end_column=73):
    """
    Clear the contents of a range in a specified worksheet and write DataFrame data to the worksheet.

    Args:
    ws (xlwings.Sheet): Worksheet to write data to.
    df (pd.DataFrame): DataFrame containing the data to be written.
    start_row (int): Starting row for writing data. Default is 5.
    start_column (int): Starting column for writing data. Default is 1.

    Returns:
    None
    """
    # Get the range starting from the specified row and column
    end_row = ws.cells.last_cell.row  
    #end_column = start_column + df.shape[1] - 1  # Number of columns based on DataFrame shape
    range_to_clear = ws.range((start_row, start_column), (end_row, end_column))

    # Clear the contents of the range
    range_to_clear.clear_contents()
    
    capital_df_security = df_security[['Branch Id', 'Issuer Id', 'Product Type', 'MDB Flag', 'Deal Reference',
                                   'Investment Type', 'Trading Indicator', 'Encumbered Flag', 'Currency',
                                   'Nominal Amount CCY', 'Principal Amount CCY', 'Market Value CCY', 'MTM',
                                   'Coupon Rate', 'Accrued Interest CCY', 'Start Date', 'Maturity Date',
                                   'Settlement Date', 'Fxd Flt Indicator', 'Next Reset Date','Security ID',
                                   'LAB Indicator', 'HQ Indicator','Senior Tranche Indicator', 'Issued/purchased flag',
                                   'USGS Indicator','LTST Indicator', 'Own Funds Indicator', 'Own Funds Amount',
                                   'Rehypothetication Indicator', 'Watch List Indicator','Default Indicator',
                                   'LCR Government Protected Indicator','LCRGteeIntOrgID', 'LCRIntOrgID',
                                   'LCRRequirementsMet','LiquidAssetQualityType', 'LiquidFacilityType', 'Notional Amount',
                                   'OtherLCRLiabilities', 'OtherTransAssets', 'Short Term Moody',
                                   'Long Term Moody', 'Short Term Fitch', 'Long Term Fitch','Short Term S P',
                                   'Long Term S P', 'Purchase Cost', 'Purchase Price','Risk Weight', 'PD external',
                                   'LGD External', 'PD Internal','LGD INTERNAL', 'MTM GBP', 'Principle GBP','Book Value GBP',
                                   'Residual Maturity Days','Sector Code','Category','Moodys Number', 'Fitch Number',
                                   'S&P Number','Final Rating', 'Final Letter Rating', 'CQS','Incorporation Code',
                                   'Corporate / Institution', 'Investment Combo','RW', 'Customer Name', 'Institution elig']]

    # Write the data (excluding headers) to the worksheet
    ws.range((start_row, start_column)).value = capital_df_security.values

def refresh_pivot_table(ws, pivot_table_name):
    """
    Refresh a pivot table in a specified worksheet.

    Args:
    ws (xlwings.Sheet): Worksheet containing the pivot table.
    pivot_table_name (str): Name of the pivot table to refresh.

    Returns:
    None
    """

    # Get the pivot table object
    pivot_table = ws.api.PivotTables(pivot_table_name)

    # Refresh the pivot table
    pivot_table.RefreshTable()

# Example usage:
# Assuming capital_file_wb is the Excel workbook and df_security contains the data
# parent_folder = "path/to/parent/folder"
# df_security = ...  # Define or read your DataFrame

# Open the Excel workbook
#capital_file_wb = xw.Book("path/to/your/workbook.xlsx")

# Select the worksheet named "Lombard Security"
ws_security = capital_file_wb.sheets["Lombard Security"]

# Clear existing data and write new data to the worksheet
clear_and_write_data_to_sheet(ws_security, df_security)

# Refresh the pivot table in the same worksheet
pivot_table_name = "PivotTable1"  # Adjust with the actual pivot table name
refresh_pivot_table(ws_security, pivot_table_name)


def write_custfile_to_worksheet(ws, start_row=2, start_column=1):
    """
    Clear the contents of a specified worksheet starting from a given cell.

    Args:
    ws (xlwings.Sheet): Worksheet to clear.
    start_row (int): Starting row for clearing data. Default is 2.
    start_column (int): Starting column for clearing data. Default is 1.

    Returns: the new values from data frame df_cust
    None
    """
    # Get the last used row and column
    last_row = ws.cells.last_cell.row
    last_column = ws.cells.last_cell.column

    # Get the range starting from the specified row and column to the last row and column
    range_to_clear = ws.range((start_row, start_column), (last_row, last_column))

    # Clear the contents of the range
    range_to_clear.clear_contents()
    
    # Write the data (excluding headers) to the worksheet
    ws.range((start_row, start_column)).value = df_cust.values

# Example usage:
# Assuming ws_cust is the worksheet named "Lombard Cust"
ws_cust = capital_file_wb.sheets["Lombard Cust"]

# Clear the worksheet starting from the 6th row across all columns
write_custfile_to_worksheet(ws_cust)


def write_BTLLTV_to_worksheet(ws, start_row=3, start_column=1, end_column=16 ):
    """
    Clear the contents of a specified worksheet starting from a given cell.

    Args:
    ws (xlwings.Sheet): Worksheet to clear.
    start_row (int): Starting row for clearing data. Default is 2.
    start_column (int): Starting column for clearing data. Default is 1.

    Returns: the new values from data frame trunc_df_frp_acc
    None
    """
    # Get the last used row and column
    last_row = ws.cells.last_cell.row
    #last_column = ws.cells.last_cell.column

    # Get the range starting from the specified row and column to the last row and column
    range_to_clear = ws.range((start_row, start_column), (last_row, end_column))

    # Clear the contents of the range
    range_to_clear.clear_contents()
    
    # Write the data (excluding headers) to the worksheet
    ws.range((start_row, start_column)).value = trunc_df_frp_acc.values

# Example usage:
# Assuming ws_MMP is the worksheet named "MM Placements"
ws_btl_ltv = capital_file_wb.sheets["BTL LTV"]

# Clear the worksheet starting from the 6th row across all columns
write_BTLLTV_to_worksheet(ws_btl_ltv)


def write_loand_dep_to_worksheet(ws, start_row=2, start_column=1):
    """
    Clear the contents of a specified worksheet starting from a given cell.

    Args:
    ws (xlwings.Sheet): Worksheet to clear.
    start_row (int): Starting row for clearing data. Default is 2.
    start_column (int): Starting column for clearing data. Default is 1.

    Returns: the new values from data frame df_loandep
    None
    """
    # Get the last used row and column
    last_row = ws.cells.last_cell.row
    last_column = ws.cells.last_cell.column

    # Get the range starting from the specified row and column to the last row and column
    range_to_clear = ws.range((start_row, start_column), (last_row, last_column))

    # Clear the contents of the range
    range_to_clear.clear_contents()
    
    # Write the data (excluding headers) to the worksheet
    ws.range((start_row, start_column)).value = df_loandep.values

# Example usage:
# Assuming ws_loan_dep is the worksheet named "Lombard Loan Dep Report"
ws_loan_dep = capital_file_wb.sheets["Lombard Loan Dep Report"]

# Clear the worksheet starting from the 6th row across all columns
write_loand_dep_to_worksheet(ws_loan_dep)



def write_accbal_to_worksheet(ws, start_row=2, start_column=1):
    """
    Clear the contents of a specified worksheet starting from a given cell.

    Args:
    ws (xlwings.Sheet): Worksheet to clear.
    start_row (int): Starting row for clearing data. Default is 2.
    start_column (int): Starting column for clearing data. Default is 1.

    Returns: the new values from data frame filtered_df_accbal
    None
    """
    # Get the last used row and column
    last_row = ws.cells.last_cell.row
    last_column = ws.cells.last_cell.column

    # Get the range starting from the specified row and column to the last row and column
    range_to_clear = ws.range((start_row, start_column), (last_row, last_column))

    # Clear the contents of the range
    range_to_clear.clear_contents()
    
    # Write the data (excluding headers) to the worksheet
    ws.range((start_row, start_column)).value = filtered_df_accbal.values

# Example usage:
# Assuming ws_accbal is the worksheet named "Lombard Account Balance Report_"
ws_accbal = capital_file_wb.sheets["Lombard Account Balance Report_"]

# Clear the worksheet starting from the 6th row across all columns
write_accbal_to_worksheet(ws_accbal)


'''# Select the worksheet named "Lombard Repay"
ws_exch = capital_file_wb.sheets["Date_ExchRates"]

# Validate the date format
try:
    # Convert the input string to a datetime object
    #date = datetime.strptime(date_str, '%d/%m/%Y')
    # Write the date to cell E5 of the worksheet
    #ws_72.range('E5').value = date
    ws_exch.range('A1').value = date_obj
    print("Date successfully entered in cell A1.")
except ValueError:
    print("Invalid date format. Please enter the date in dd/mm/yyyy format.")
    
    
    
After this delete the values from column C starting from row 6'''


def write_df_pipeline_to_worksheet(ws, start_row=4, start_column=1):
    """
    Clear the contents of a specified worksheet starting from a given cell.

    Args:
    ws (xlwings.Sheet): Worksheet to clear.
    start_row (int): Starting row for clearing data. Default is 6.
    start_column (int): Starting column for clearing data. Default is 1.

    Returns: the new values from data frame df_pipeline
    None
    """
    # Get the last used row and column
    last_row = ws.cells.last_cell.row
    last_column = ws.cells.last_cell.column

    # Get the range starting from the specified row and column to the last row and column
    range_to_clear = ws.range((start_row, start_column), (last_row, last_column))

    # Clear the contents of the range
    range_to_clear.clear_contents()
    
    # Write the data (excluding headers) to the worksheet
    ws.range((start_row, start_column)).value = df_pipeline.values

# Example usage:
# Assuming ws_MMP is the worksheet named "BTL Pipeline"
ws_btl_pipeline = capital_file_wb.sheets["BTL Pipeline"]

# Clear the worksheet starting from the 6th row across all columns
write_df_pipeline_to_worksheet(ws_btl_pipeline)


def write_exch_to_worksheet(wb, sheet_name, date_obj, df_exch):
    """
    Enter a date in cell A1 of the specified worksheet and clear values in column C starting from row 6.

    Args:
    wb (xlwings.Book): Excel workbook object.
    sheet_name (str): Name of the worksheet where operations will be performed.
    date_obj (datetime.datetime): Date object to be entered in cell A1.

    Returns:
    None
    """
    # Select the worksheet
    ws_exch = wb.sheets[sheet_name]

    try:
        # Write the date to cell A1 of the worksheet
        ws_exch.range('A1').value = date_obj
        print("Date successfully entered in cell A1.")

        # Get the range starting from row 6 in column C
        start_row = 6
        start_column = 3  # Column C
        end_row = start_row + len(df_exch) - 1  # Calculate the end row based on DataFrame length
        range_to_clear = ws_exch.range((start_row, start_column), (end_row, start_column))


        # Clear the contents of the range
        range_to_clear.clear_contents()
        
        # Paste values from the 'Exchange Rate' column to column C starting from row 6
        for i, rate in enumerate(df_exch['Exchange Rate'], start=start_row):
            ws_exch.range((i, start_column)).value = rate
    except Exception as e:
        print(f"Error: {e}")

# Usage:
        
# Call the function
write_exch_to_worksheet(capital_file_wb, "Date_ExchRates", date_obj, df_exch)

# Calculate all formulas in the workbook
#app.calculate()

# Save the workbook
capital_file_wb.save()

# Close the workbook
capital_file_wb.close()
#app.quit()

# Record the end time
enditime = time.time()

# Calculate the elapsed time
elapsed_time = enditime - start_time

# Convert elapsed time to a human-readable format
hours, rem = divmod(elapsed_time, 3600)
minutes, seconds = divmod(rem, 60)

# Display the time taken
print("Total Time taken: {:0>2}:{:0>2}:{:05.2f}".format(int(hours), int(minutes), seconds))

# Convert elapsed time to a human-readable format
lcr_hours, lcr_rem = divmod(LCR_completion_time, 3600)
lcr_minutes, lcr_seconds = divmod(rem, 60)

# Display the time taken
print("Time taken for LCR File: {:0>2}:{:0>2}:{:05.2f}".format(int(lcr_hours), int(lcr_minutes), lcr_seconds))
