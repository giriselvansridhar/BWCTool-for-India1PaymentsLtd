from collections import defaultdict
import json
import os
import pickle
import pandas as pd
import shutil

def get_valid_prefixes(file_path):#7
    """Read the list of valid prefixes from a specified file."""
    with open(file_path, 'r') as file:#8 Open file
        prefixes = file.read().split(',')#9 Remove commas
    return [prefix.strip() for prefix in prefixes] #Returns the list

def is_valid_file(file, prefixes):
    """Check if the file name starts with any of the valid prefixes."""
    valid = any(file.startswith(prefix) for prefix in prefixes)
    if not valid:
        print(f"File {file} does not start with a valid prefix. Skipping...")
    return valid

def move_to_error_directory(file_path, error_directory):
    """Move a problematic file to the error directory."""
    try:
        shutil.move(file_path, os.path.join(error_directory, os.path.basename(file_path)))
        print(f"File {file_path} moved to error directory: {error_directory}")
    except Exception as e:
        print(f"Error moving file {file_path} to error directory: {error_directory}. Error: {e}")

def read_excel_file(file_path):
    """Read an Excel file into a DataFrame."""
    try:
        if file_path.endswith('.xlsx'):
            return pd.read_excel(file_path,header=1)        
        elif file_path.endswith('.xls'):
            return pd.read_excel(file_path,header=1)
        elif file_path.endswith('.xlsb'):
            return pd.read_excel(file_path, engine='pyxlsb', header=5)
        else:
            raise ValueError(f"Unsupported file type for file: {file_path}")
    except Exception as e:
        raise ValueError(f"Error reading Excel file {file_path}: {e}")
def read_excel_file_1(file_path,header1):
    """Read an Excel file into a DataFrame."""
    try:
        if file_path.endswith('.xlsx'):
            return pd.read_excel(file_path,header=1)        
        elif file_path.endswith('.xls'):
            return pd.read_excel(file_path,header=1)
        elif file_path.endswith('.xlsb'):
            return pd.read_excel(file_path, engine='pyxlsb')
        else:
            raise ValueError(f"Unsupported file type for file: {file_path}")
    except Exception as e:
        raise ValueError(f"Error reading Excel file {file_path}: {e}")



def process_files(directory, error_directory, prefixes):
    """Process all files in the directory, reading valid Excel files into DataFrames."""
    try:
        files = os.listdir(directory)
        dataframes = {}

        for file in files:
            file_path = os.path.join(directory, file)
            try:
                if is_valid_file(file, prefixes):
                    df = read_excel_file(file_path)
                    dataframes[file] = df
                else:
                    move_to_error_directory(file_path, error_directory)
            except Exception as e:
                print(f"Error processing file {file}: {e}")
                move_to_error_directory(file_path, error_directory)

        return dataframes
    
    except Exception as e:
        print(f"Error processing files in directory {directory}: {e}")


def GetTotalCreditAmount(fileDetails, fileName, prop_reader):
    if fileDetails[prop_reader.getDate()].dtype != '<M8[ns]':  
        fileDetails[prop_reader.getDate()] = pd.to_datetime('1899-12-30') + pd.to_timedelta(fileDetails[prop_reader.getDate()], unit='D')
    
    if prop_reader.getCrDrSeparator():
        fileDetails = fileDetails.sort_values(by=prop_reader.getDate())
        inputRemarks = fileDetails[prop_reader.getRemarks()].unique()
        inputDates = fileDetails[prop_reader.getDate()].unique()

        results = []
        for date in inputDates:
            filtered_by_date = fileDetails[fileDetails[prop_reader.getDate()] == date]
            for remark in inputRemarks:
                filtered_by_remark = filtered_by_date[filtered_by_date[prop_reader.getRemarks()] == remark]
                # print("filtered_by_remark",filtered_by_remark.columns )
                col_seperator = prop_reader.getCrDrSeparatorColName()
                # print("col_seperator",col_seperator )
                filtered_by_DrSeperator = filtered_by_remark[filtered_by_remark[col_seperator]== "CR" ]                
                if not filtered_by_DrSeperator.empty:
                    # print("filtered_by_DrSeperator", filtered_by_DrSeperator)
                    debitsum = filtered_by_DrSeperator[prop_reader.getDebitAmount()].sum()
                    results.append([date, remark, debitsum])

                    
                    
    
    else:
        fileDetails = fileDetails.sort_values(by=prop_reader.getDate())
        inputRemarks = fileDetails[prop_reader.getRemarks()].unique()
        inputDates = fileDetails[prop_reader.getDate()].unique()
        results = []

        for date in inputDates:
            filtered_by_date = fileDetails[fileDetails[prop_reader.getDate()] == date]
            for remark in inputRemarks:
                creditsum = filtered_by_date.loc[
                    filtered_by_date[prop_reader.getRemarks()] == remark, 
                    prop_reader.getDebitAmount()
                ].astype(float).sum()
                results.append([date, remark, creditsum])
    
    return results


def GetTotalDebitAmount(fileDetails, fileName, prop_reader):
    if fileDetails[prop_reader.getDate()].dtype != '<M8[ns]':  
        fileDetails[prop_reader.getDate()] = pd.to_datetime('1899-12-30') + pd.to_timedelta(fileDetails[prop_reader.getDate()], unit='D')
    
    if prop_reader.getCrDrSeparator():
        fileDetails = fileDetails.sort_values(by=prop_reader.getDate())
        inputRemarks = fileDetails[prop_reader.getRemarks()].unique()
        inputDates = fileDetails[prop_reader.getDate()].unique()

        results = []
        for date in inputDates:
            filtered_by_date = fileDetails[fileDetails[prop_reader.getDate()] == date]
            for remark in inputRemarks:
                filtered_by_remark = filtered_by_date[filtered_by_date[prop_reader.getRemarks()] == remark]
                # print("filtered_by_remark",filtered_by_remark.columns )
                col_seperator = prop_reader.getCrDrSeparatorColName()
                # print("col_seperator",col_seperator )
                filtered_by_DrSeperator = filtered_by_remark[filtered_by_remark[col_seperator]== "DR" ]                
                if not filtered_by_DrSeperator.empty:
                    # print("filtered_by_DrSeperator", filtered_by_DrSeperator)
                    debitsum = filtered_by_DrSeperator[prop_reader.getDebitAmount()].sum()
                    
                    results.append([date, remark, debitsum])
            
    
    else:
        fileDetails = fileDetails.sort_values(by=prop_reader.getDate())
        inputRemarks = fileDetails[prop_reader.getRemarks()].unique()
        inputDates = fileDetails[prop_reader.getDate()].unique()
        results = []

        for date in inputDates:
            filtered_by_date = fileDetails[fileDetails[prop_reader.getDate()] == date]
            for remark in inputRemarks:
                debitsum = filtered_by_date.loc[
                    filtered_by_date[prop_reader.getRemarks()] == remark, 
                    prop_reader.getDebitAmount()
                ].astype(float).sum()
                results.append([date, remark, debitsum])
    
    return results

        
def RemarkTheHeadingCreditList(CreditsSumsLists,RemarksHeadingMapping,fileName):
    filename=fileName.split('.', 1)[0]
    #print(filename)
    filtered_df = RemarksHeadingMapping[(RemarksHeadingMapping['BankName'] == filename) & (RemarksHeadingMapping['Transaction Type'] == 'CR')]
    
    #print(filtered_df)


    # Create a mapping dictionary from dataframe Remarks to Heading
    remarks_to_heading = dict(zip(filtered_df['Remarks'], filtered_df['Heading']))
    print(remarks_to_heading)



# Create new nested list with both conditions
    new_nested_list = []
    for item in CreditsSumsLists:
        timestamp, remark, value = item
        if value != 0 and remark in remarks_to_heading:
        

            heading = remarks_to_heading[remark]
            new_nested_list.append([timestamp, remark, value, heading,fileName])


    
    # Dictionary to store aggregated amounts
    aggregated_transactions = defaultdict(lambda: defaultdict(float))

# Aggregate amounts based on date and description
    for transaction in new_nested_list:
        date, description, amount = transaction[0], transaction[3], transaction[2]
        aggregated_transactions[date][description] += amount

# Convert aggregated transactions back to a list of lists if needed
    aggregated_list = [[date, description, amount] for date, descriptions in aggregated_transactions.items() for description, amount in descriptions.items()]
    return aggregated_list         



    

def Code_Inwards_Generator(Inwardsumwithheader,fileName):
    new_data = []
    current_date = None
    current_dict = {}
    filename=fileName.split('.', 1)[0]

    for entry in Inwardsumwithheader:
        date, transaction_type, amount = entry
        if date != current_date:
            if current_dict:  # if the current_dict is not empty, add it to new_data
                new_data.append(current_dict)
            current_date = date
            current_dict = {('Details', 'Bank'): filename, ('Details', 'Date'): date}
        
        if ('Inwards', transaction_type) in current_dict:
            current_dict[('Inwards', transaction_type)] += amount
        else:
            current_dict[('Inwards', transaction_type)] = amount

    if current_dict:  # add the last collected dictionary
        new_data.append(current_dict)
    

    return new_data

    



def Code_Outwards_Generator(Outwardsumwithheader,fileName):
     new_data = []
     current_date = None
     current_dict = {}
     filename=fileName.split('.', 1)[0]
     for entry in Outwardsumwithheader:
        date, transaction_type, amount = entry
        if date != current_date:
            if current_dict:  # if the current_dict is not empty, add it to new_data
                new_data.append(current_dict)
            current_date = date
            current_dict = {('Details', 'Bank'): filename, ('Details', 'Date'): date}
        
        if ('Inwards', transaction_type) in current_dict:
            current_dict[('Outwards', transaction_type)] += amount
        else:
            current_dict[('Outwards', transaction_type)] = amount

     if current_dict:  # add the last collected dictionary
        new_data.append(current_dict)
     return  new_data

   



    
    






def RemarkTheHeadingDebitList(DebitsSumLists,RemarksHeadingMapping,fileName):
  
    filename=fileName.split('.', 1)[0]
    
    
    
    filtered_df = RemarksHeadingMapping[(RemarksHeadingMapping['BankName'] == filename) & (RemarksHeadingMapping['Transaction Type'] == 'DR')]
    remarks_to_heading = dict(zip(filtered_df['Remarks'], filtered_df['Heading']))


# Create new nested list with both conditions
    new_nested_list = []
    for item in DebitsSumLists:
        timestamp, remark, value = item
        if value != 0 and remark in remarks_to_heading:
            heading = remarks_to_heading[remark]
            new_nested_list.append([timestamp, remark, value, heading,fileName])


    # Dictionary to store aggregated amounts
    aggregated_transactions = defaultdict(lambda: defaultdict(float))

# Aggregate amounts based on date and description
    for transaction in new_nested_list:
        date, description, amount = transaction[0], transaction[3], transaction[2]
        aggregated_transactions[date][description] += amount

# Convert aggregated transactions back to a list of lists if needed
    aggregated_list = [[date, description, amount] for date, descriptions in aggregated_transactions.items() for description, amount in descriptions.items()]

    return aggregated_list       
    
    
    
    
    


                



        







def read_excel_file_remarks_header_map(file_path):
    try:
        if file_path.endswith('.xlsx'):
            return pd.read_excel(file_path)        
        elif file_path.endswith('.xls'):
            return pd.read_excel(file_path)
        elif file_path.endswith('.xlsb'):
            return pd.read_excel(file_path)
        else:
            raise ValueError(f"Unsupported file type for file: {file_path}")
    except Exception as e:
        raise ValueError(f"Error reading Excel file {file_path}: {e}")


# Function to save data to pickle file
def dump_to_pickle(data_list, file_name):
    try:
        # Ensure directory exists
        os.makedirs(os.path.dirname(file_name), exist_ok=True)
        
        # Check if file exists to append data or create new
        if os.path.exists(file_name):
            with open(file_name, 'rb') as f:
                existing_data = pickle.load(f)
                data_list.extend(existing_data)
        
        # Dump data to pickle file
        with open(file_name, 'wb') as f:
            pickle.dump(data_list, f)
        # print(f"Data saved to {file_name}")
    except Exception as e:
        print(f"Error saving to {file_name}: {e}")

# Function to load data from pickle file
def load_and_delete_pickle(file_name):
    try:
        with open(file_name, 'rb') as f:
            data = pickle.load(f)
        
        # Delete the file after loading
        os.remove(file_name)
        print(f"File '{file_name}' deleted successfully.")
        
        return data
    except FileNotFoundError:
        print(f"File '{file_name}' not found.")
        return []
    except Exception as e:
        print(f"Error loading from {file_name}: {e}")
        return []
    

def merge_dicts(list_of_dicts):
    merged_dict = defaultdict(lambda: defaultdict(float))
    for d in list_of_dicts:
        key = (d[('Details', 'Bank')], d[('Details', 'Date')])
        for k, v in d.items():
            if k not in [('Details', 'Bank'), ('Details', 'Date')]:
                merged_dict[key][k] += v
            else:
                merged_dict[key][k] = v
    return [dict(v) for v in merged_dict.values()]    
