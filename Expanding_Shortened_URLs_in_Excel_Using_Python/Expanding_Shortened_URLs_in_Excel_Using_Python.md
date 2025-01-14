#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

--------------------------------------------------------------------------------

# Expanding Shortened URLs in Excel Using Python: A Step-by-Step Guide

Shortened URLs are commonly used to simplify long links, making them easier to share. However, in scenarios like analytics or auditing, it’s often necessary to expand these URLs to their original form. In this article, we outline a Python-based solution to automate this process, providing a seamless way to expand shortened URLs stored in an Excel file and save the results into a new file.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Expanding_Shortened_URLs_in_Excel_Using_Python/00_Shortened_URLs.png?raw=true)

## Key Features of the Solution

1. **Automated Expansion**: Uses Python to expand shortened URLs via HTTP requests.

2. **Incremental Writing**: Writes results to the output file row by row, ensuring efficient memory usage.

3. **Progress Tracking**: Includes a progress bar for monitoring the operation.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Expanding_Shortened_URLs_in_Excel_Using_Python/01_Shortened_URLs.png?raw=true)

# Pseudocode

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Expanding_Shortened_URLs_in_Excel_Using_Python/02_Shortened_URLs.png?raw=true)

### 1. Import Required Libraries

- Import `os` for file and directory management.
- Import `requests` for handling HTTP requests.
- Import `pandas` for reading and manipulating Excel data.
- Import `tqdm` for adding progress bars to loops.

### 2. Set the Target Directory

- Define the path where the input Excel file is located.
- Change the working directory to this path.

### 3. Define a Function to Expand Shortened URLs

- Accept a shortened URL as input.
- Make a `HEAD` request to the URL, allowing redirects.
- If successful, return the expanded URL.
- If an error occurs, return the error message.

### 4. Load the Excel File

- Read the Excel file (`input_file.xlsx`) into a DataFrame.

### 5. Check for the Required Column

- If the column Shortened_URL is not present in the DataFrame:
    - Print a message indicating the column is missing.
- Otherwise:
    - Proceed to process the data.

### 6. Initialize a Progress Bar

- Use `tqdm` to track the progress of URL expansion.

### 7. Open an Excel Writer

- Create an Excel writer object to save the results incrementally.
- Write an empty DataFrame with column headers (`Shortened_URL` and `Expanded_URL`) to initialize the output file.

### 8. Process Each Row of the DataFrame

- Iterate through each row of the DataFrame using `tqdm` for progress tracking.
    - Extract the shortened URL from the `Shortened_URL` column.
    - Expand the URL using the `expand_url` function.
    - Store the expanded URL in a new column, `Expanded_URL`.

### 9. Append Each Processed Row to the Output File

- Convert the current row to a DataFrame.
- Append the row to the output Excel file without headers.

### 10. Complete the Process

- Close the Excel writer.
- Print a message confirming that the expanded URLs have been saved.

## Import Required Libraries

Utilize essential Python libraries for handling URLs, regular expressions, data manipulation, and progress visualization


```python
import os  # Library for interacting with the operating system, useful for file and directory management.
import requests  # Library to make HTTP requests, useful for working with APIs and fetching web content.
import pandas as pd  # Library for data analysis and manipulation, especially for tabular data.
from tqdm import tqdm  # Library to display progress bars for loops and long-running processes.
```

### `os`:

- Provides tools for interacting with the operating system.
- Commonly used for file and directory management, such as creating, deleting, and navigating directories.
- Key functions:
    - `os.path.join()`: Combines directory paths in a platform-independent way.
    - `os.listdir()`: Lists all files and folders in a directory.
    - `os.makedirs()`: Creates directories recursively.

### `requests`:

- Facilitates HTTP requests to interact with APIs or fetch web content.
- Supports methods like `GET`, `POST`, and `HEAD`.
- Key attributes:
    - `url`: The final resolved URL after redirects.
    - `status_code`: HTTP status code of the response.
    - `headers`: HTTP headers from the response.

### `pandas`:

- Used for structured data manipulation and analysis.
- Key object: `DataFrame`, which represents tabular data.
- Features include data cleaning, aggregation, and export to file formats like Excel and CSV.

### `tqdm`:

- Adds progress bars to loops and other iterative processes for better visualization of progress.
- `tqdm(iterable)`: Wraps an iterable like a list or range with a progress bar.
- Customizable with options such as `desc` for descriptions and `unit` for progress units.

## Setting the Working Directory


```python
# Define the path to the target directory
path = 'D:/Advance_Data_Sets/google_map'
# Change the current working directory to the specified path
os.chdir(path)
```

## Function to Expand Shortened URLs


```python
# Function to expand shortened URLs
def expand_url(short_url):
    try:
        response = requests.head(short_url, allow_redirects=True)
        return response.url
    except requests.RequestException as e:
        return f"Error: {e}"
```

## Load Data from Excel File

- Read the Excel file containing URLs into a pandas `DataFrame`.


```python
# Read the Excel file into a DataFrame
df = pd.read_excel("input_file.xlsx")
```

## Incrementally Expanding Shortened URLs and Saving to Excel

This code checks if a specific column named `Shortened_URL` exists in an Excel file. If the column is present, it processes each shortened URL in the column, expands it to its full URL, and saves the expanded URLs alongside the original ones into a new Excel file called `expanded_urls.xlsx`.

The process is done row by row, and progress is tracked using a progress bar for better visibility. The data is saved incrementally to the Excel file to ensure all processed rows are written without losing data in case of interruptions.


```python
# Check if the required column exists in the DataFrame
if 'Shortened_URL' not in df.columns:
    print("Column 'Shortened_URL' not found in the Excel file.")  # Notify user if column is missing
else:
    # Initialize tqdm for progress tracking with a description
    tqdm.pandas(desc="Expanding URLs")
    
    # Open an Excel writer object to save the expanded data progressively
    with pd.ExcelWriter("expanded_urls.xlsx", engine='openpyxl', mode='w') as writer:
        # Write an empty DataFrame with headers to initialize the output file
        df[['Shortened_URL']].assign(Expanded_URL="").head(0).to_excel(writer, index=False)

        # Iterate over rows of the DataFrame with progress tracking
        for idx, row in tqdm(df.iterrows(), total=len(df), desc="Processing URLs"):
            short_url = row['Shortened_URL']  # Extract the shortened URL from the row
            expanded_url = expand_url(short_url)  # Expand the shortened URL using the custom function
            df.loc[idx, 'Expanded_URL'] = expanded_url  # Save the expanded URL back to the DataFrame
            
            # Append the current row with expanded URL to the Excel file without headers
            pd.DataFrame([df.loc[idx]]).to_excel(writer, index=False, header=False, startrow=idx + 1)

    print(f"Expanded URLs saved to expanded_urls.xlsx.")  # Confirm completion
```

    Processing URLs: 100%|███████████████████████████████████████████████████████████████| 134/134 [03:55<00:00,  1.76s/it]

    Expanded URLs saved to expanded_urls.xlsx.
    

    
    

## Conclusion

This solution simplifies the process of expanding shortened URLs, making it particularly useful for data analysis and reporting tasks. Check out the GitHub repository for the full code, and feel free to adapt it to your specific needs!
