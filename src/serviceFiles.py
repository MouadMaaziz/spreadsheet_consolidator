import os
import pandas as pd



def serviceFiles(results_dir):
    """ Create a spreadsheet for each service and save it in the result directory"""
    for service in os.listdir(results_dir):
        # No point of using glob here since all the files within each folder are .xls.
        if os.path.isdir(os.path.join(results_dir, service)):
            print(f"executing {service}...")
            files = os.listdir(os.path.join(results_dir, service)) 
            dfs = [pd.read_excel(os.path.join(results_dir, service, f), dtype=object) for f in files]
            results=pd.DataFrame(dfs[0])
            for df in (dfs[1:]):
                results = pd.merge(results, df.iloc[:,2:], on='Code_localite', how='left')
                results.drop_duplicates(inplace=True)
            results.to_excel(os.path.join(results_dir, f"{service}.xlsx"), index=False)
