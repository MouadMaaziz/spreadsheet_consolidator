import os
import pandas as pd
import glob



def consolidate(results_dir):
    files = [x for x in glob.glob(f"{results_dir}/*.xls*")]
    dfs = [pd.read_excel(os.path.join(results_dir, f), dtype=object) for f in files]
    results = pd.DataFrame(dfs[0])
    for df in (dfs[1:]):
        results = pd.merge(results, df.iloc[:,2:], on='Code_localite', how='left')
        results.drop_duplicates(inplace=True)  
    results.to_excel(os.path.join(results_dir, f"services_16_20.xlsx"), index=False)
    