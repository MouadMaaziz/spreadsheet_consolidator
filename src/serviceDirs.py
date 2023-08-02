import pandas as pd
import os
import glob


def extract_year_service(spreadsheet):
    """
    This function takes as input the name of a spreadsheet named: <service>_<localityType>_<year>
    and stores those in a variable.
    """
    service,_,year = spreadsheet.split("_")[-3:]
    return service, year[:4]


def serviceDirs(pwd, results_dir):
    """ 
    This function is built to:
    1. merge douars with sous_douars into one list per year and per service,
    2. move spreadsheets into different directories accordingly to the type of service 
    """
    data = {}
    # Iterating over the spreadsheets in the directory
    for file_path in glob.glob(f"{pwd}/*.xls*"):
        spreadsheet_name = os.path.basename(file_path)
        service, year = extract_year_service(spreadsheet_name)
        
        # read the file and only take the needed fields.
        with open(file_path, 'rb') as file:
            df = pd.read_excel(file, usecols=['Name','Total_Length'], dtype={'Name': str, 'Total_Length': float})
        
        # creating a directory for the service if it does not exist.
        if not os.path.exists(results_dir):
            os.mkdir(results_dir)

        if not os.path.exists(os.path.join(results_dir, f'{service}')):
            os.mkdir(os.path.join(results_dir, f'{service}'))

        # storing the dataframe into a dictionary
        if (service,year) in data: # If this evaluate to true, then the opposite locality type is there.
            data[(service,year)] = pd.concat([data[(service,year)],df], ignore_index=True)
        else:
            data[(service,year)] = df

    # cleaning the data and saving the final spreadsheet.
    for key in data.keys():
        df = data[key]
        service, year = key
        df[['Code_localite', f'Code_etab_{service}_{year}']] = df['Name'].str.split(' - ', n=1, expand=True)
        df[f'{service}_{year}'] = df['Total_Length']
        def type_localite(x):  return "Douar" if len(x)==12 else "Sous-Douar"
        df['type'] = df['Code_localite'].apply(type_localite)
        df.drop(columns='Name', inplace=True)
        df = df[['type', 'Code_localite', f'Code_etab_{service}_{year}', f'{service}_{year}']]
        df.to_excel(os.path.join(results_dir,service, f"{service}_{year}.xlsx"))
        print(f"passed {service}_{year}")
