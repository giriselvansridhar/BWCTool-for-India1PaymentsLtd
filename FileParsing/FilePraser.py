
import json
import os
from Models.JsonPropertyReader import JsonPropertyReader

def fetchpropertydetails(input_filename, input_filepath, validation_filename, validation_filepath):
    # Check if the validation JSON file exists
    validation_file_path = os.path.join(validation_filepath, validation_filename)
    if not os.path.exists(validation_file_path):
        print(f"Error: JSON file '{validation_file_path}' does not exist.")
        return None
    
    try:
        # Load JSON data from validation file
        with open(validation_file_path, 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
            
            # Search for the filename in the JSON data
            for bank in data.get("banks", []):
                if input_filename.startswith(bank["FileName"]):
                    # Format the property details into a string
                    property_details = (
                        f"HeadingRowNo: {bank['HeadingRowNo']}\n"
                        f"Date: {bank['Date']}\n"
                        f"DebitAmount: {bank['DebitAmount']}\n"
                        f"CreditAmount: {bank['CreditAmount']}\n"
                        f"Remarks: {bank['Remarks']}\n"
                        f"CrDrSeparator: {bank['CrDrSeparator']}\n"
                        f"CrDrSeparatorColName: {bank['CrDrSeparatorColName']}"
                    )

                    # Create an instance of JsonPropertyReader
                    prop_reader = JsonPropertyReader()
                    prop_reader.setFileName(bank.get("FileName", bank['FileName']))
                    prop_reader.setHeadingRowNo(bank.get("HeadingRowNo", bank['HeadingRowNo']))
                    prop_reader.setDate(bank.get("Date", bank['Date']))
                    prop_reader.setDebitAmount(bank.get("DebitAmount", bank['DebitAmount']))
                    prop_reader.setCreditAmount(bank.get("CreditAmount", bank['CreditAmount']))
                    prop_reader.setRemarks(bank.get("Remarks", bank['Remarks']))
                    prop_reader.setCrDrSeparator(bank.get("CrDrSeparator", bank['CrDrSeparator']))
                    prop_reader.setCrDrSeparatorColName(bank.get("CrDrSeparatorColName", bank['CrDrSeparatorColName']))
                    return prop_reader  
            
            
            # If filename is not found in JSON data
            print(f"Error: Filename '{input_filename}' not found in JSON '{validation_filename}'.")
            return None
    
    except Exception as e:
        print(f"Error reading JSON file '{validation_file_path}': {str(e)}")
        return None

