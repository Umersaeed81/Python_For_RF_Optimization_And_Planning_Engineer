# Import Libraries
import os
import zipfile
import numpy as np
import pandas as pd
from glob import glob
from datetime import datetime, timedelta
import warnings
warnings.simplefilter("ignore")

# Set Path for GSM Cell DA KPIs
path_cell_hourly = 'D:/Advance_Data_Sets/KPIs_Analysis/Hourly_KPIs_Cell_Level/LTE'
os.chdir(path_cell_hourly)

# List all files in the path
file_list = glob('*.zip')

# Calculate the date for three days ago
required_date = datetime.now() - timedelta(days=7)
required_date_str = required_date.strftime("%d%m%Y")

# Filter files that match the required date or the two days before
filtered_file_list = [file for file in file_list if any((required_date + timedelta(days=i)).\
                                     strftime("%d%m%Y") in file for i in range(1))]

if filtered_file_list:
    print("Filtered File List:", filtered_file_list)
else:
    print(f"No file found with the required dates in the last three days.")

# Define a function to read csv files from zip archives
def read_csv_from_zip(file, **kwargs):
    with zipfile.ZipFile(file) as z:
        # Get list of csv files in the zip archive
        csv_files = [name for name in z.namelist() if name.endswith('.csv')]
        # Concatenate all csv files into a single dataframe
        dfs = [pd.read_csv(z.open(name), **kwargs) for name in csv_files]
        return pd.concat(dfs, ignore_index=True)

df0 = pd.concat(
    read_csv_from_zip(file, skiprows=range(6), skipfooter=1, engine='python',
                      parse_dates=["Date","Time"], na_values=['NIL','/0'])\
                      for file in filtered_file_list)

def export_df_to_parquet_by_date(df, output_folder, compression='brotli'):
    import os

    # 🧼 Step 1: Optimize memory usage before export
    # 🔁 Convert repetitive strings to category
    for col in df.select_dtypes(include='object').columns:
        if df[col].nunique() / len(df) < 0.5:  # Heuristic: If >50% values repeat
            df[col] = df[col].astype('category')

    # 🔽 Downcast float and integer columns to reduce size
    df[df.select_dtypes(include='float').columns] = df.select_dtypes(include='float').apply(
        pd.to_numeric, downcast='float'
    )
    df[df.select_dtypes(include='int').columns] = df.select_dtypes(include='int').apply(
        pd.to_numeric, downcast='integer'
    )

    # 📁 Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # 🔢 Extract numeric prefixes from existing files for safe numbering
    files = os.listdir(output_folder)
    numeric_prefixes = [
        int(f.split('_')[0]) for f in files if f.split('_')[0].isdigit()
    ]
    new_max_num = max(numeric_prefixes) + 1 if numeric_prefixes else 0

    # 🗓️ Export one Parquet file per date with optimized DataFrame
    for unique_date in df['Date'].unique():
        unique_date_str = unique_date.strftime('%d%m%Y')  # Format date as DDMMYYYY
        output_file = os.path.join(
            output_folder,
            f'{new_max_num:02d}_LTE_KPIs_For_Reporting_LTE_DA{unique_date_str}.parquet'
        )

        # 🔍 Filter rows for the current date
        df_filtered = df[df['Date'] == unique_date]

        # 💾 Save as compressed Parquet
        df_filtered.to_parquet(output_file, index=False, compression=compression)

        new_max_num += 1  # Increment file counter

export_df_to_parquet_by_date(
    df0,                  # 🧹 Final cleaned and labeled DataFrame
    'E:/test_path',               # 📁 Output directory to save Parquet files (partitioned by date)
    compression='brotli'
)
