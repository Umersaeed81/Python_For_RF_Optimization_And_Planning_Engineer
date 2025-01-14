#  [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)
Sr. RF Planning & Optimization Engineer<br>
BSc Telecommunications Engineering, School of Engineering<br>
MS Data Science, School of Business and Economics<br>
**University of Management & Technology**<br>
**Mobile:**     +923018412180<br>
**Email:**  umersaeed81@hotmail.com<br>
**Address:** Dream Gardens,Defence Road, Lahore<br>

--------------------------------------------------------------------------------
# Extracting Coordinates from Google Maps URLs and Exporting to Excel

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Extracting_Coordinates_from_Google_Maps_URLs/00_Extracting_Coordinates.png?raw=true)

## Introduction

Analyzing geographical data often involves extracting latitude and longitude coordinates from URLs. This article explains a Python-based solution to efficiently extract these coordinates from Google Maps URLs and export the results to an Excel file. The process leverages the requests, re, pandas, and tqdm libraries for URL handling, pattern matching, data management, and progress visualization, respectively.

This solution is particularly useful for tasks involving large datasets of Google Maps URLs, ensuring accuracy and automating repetitive tasks.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Extracting_Coordinates_from_Google_Maps_URLs/01_Extracting_Coordinates.png?raw=true)

## Objectives

1. Extract latitude and longitude coordinates from various Google Maps URL formats:
    - URLs containing @lat,lng.
    - URLs containing q=lat,lng.
    - Shortened URLs (e.g., goo.gl).
2. Process large datasets of URLs stored in an Excel file.
3. Incrementally save processed data to avoid loss in case of interruptions.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Extracting_Coordinates_from_Google_Maps_URLs/02_Extracting_Coordinates.png?raw=true)

## Flow Chart

The process workflow for extracting coordinates from Google Maps URLs is outlined below. The flow begins with loading the URLs from an Excel file, followed by resolving shortened URLs, matching patterns for latitude and longitude, and updating the data in a structured format. Each step is designed to ensure accurate data extraction and error handling. The final output is saved to an Excel file, making it easy to integrate the results into subsequent analyses.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Extracting_Coordinates_from_Google_Maps_URLs/03_Extracting_Coordinates.png?raw=true)

## Import Required Libraries

Utilize essential Python libraries for handling URLs, regular expressions, data manipulation, and progress visualization


```python
import os  # Library for interacting with the operating system, useful for file and directory management.
import re  # Regular expressions library for pattern matching and text processing.
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

### `re`(Regular Expressions):

- Provides tools for pattern matching and text processing.
- Useful for extracting patterns like latitude and longitude from strings.
- Key functions:
    - `search()`: Finds a match anywhere in the string.
    - `match()`: Checks for a match at the beginning of the string.
    - `group()`: Retrieves specific parts of the matched pattern.

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

## Define the Function to Extract Coordinates

Create a function that extracts latitude and longitude from the resolved URLs.

- **Input**: A Google Maps URL.
- **Output**: A tuple containing latitude and longitude, or `(None, None)` if no coordinates are found.
- **Steps**:
    1. Resolve shortened URLs using `requests.head()`.
    2. Search for coordinates in `@lat,lng` format using `re.search()`.
    3. If no match is found, search for coordinates in `q=lat,lng` format.
    4. Return extracted coordinates or `None` if no match exists.


```python
def extract_coordinates(google_maps_url):
    """
    Extract latitude and longitude from a Google Maps URL.
    Supports '@lat,lng', 'q=lat,lng', and short URLs like goo.gl.
    """
    try:
        # Resolve the short URL to the full URL
        response = requests.head(google_maps_url, allow_redirects=True, timeout=10)
        resolved_url = response.url  # Extracts the final resolved URL
        
        # Check for '@lat,lng' format in the URL
        match_at = re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", resolved_url)
        if match_at:
            return match_at.group(1), match_at.group(2)  # Return latitude and longitude
        
        # Check for 'q=lat,lng' format in the URL
        match_q = re.search(r"q=(-?\d+\.\d+),(-?\d+\.\d+)", resolved_url)
        if match_q:
            return match_q.group(1), match_q.group(2)  # Return latitude and longitude
        
        # Return None if no coordinates are found
        return None, None
    except Exception as e:
        print(f"Error processing URL {google_maps_url}: {e}")
        return None, None
```

### 1. `requests.head()`:

- Sends an HTTP HEAD request to fetch response headers.
- Attributes:
    - `allow_redirects`: Follows redirects if `True`.
    - `timeout`: Sets a timeout for the request.
    - `response.url`: Extracts the resolved URL after following redirects.

### 2. `re.search():`

- Searches for a pattern in a string.
- If a match is found, returns a match object.
- Match object methods:
    - `group(1)`: Retrieves the first captured group in the pattern.

### 3. `Error Handling:`

- Catches exceptions with try-except.
- Provides meaningful error messages and prevents crashes.

## Setting the Working Directory

This block of code sets the target directory and changes the working directory to the specified path.


```python
# Define the path to the target directory
path = 'D:/Advance_Data_Sets/google_map'
# Change the current working directory to the specified path
os.chdir(path)
```

## Load Data from Excel File

- Read the Excel file containing Google Maps URLs into a pandas `DataFrame`.
- Validate that the file includes a column named `URL`. Raise an error if the column is missing.


```python
# Read the Excel file containing URLs
df = pd.read_excel('input_file.xlsx')  # Load data into a DataFrame

# Ensure the URLs are in a specific column, e.g., 'URL'
if 'URL' not in df.columns:
    raise ValueError("The input Excel file must have a column named 'URL'.")
```

### 1. `pd.read_excel():`

- Reads an Excel file and returns a `DataFrame`.
- Arguments:
    - `filename`: Path to the Excel file.
    - Returns: `DataFrame` representing the data

### 2. `df.columns:`

- Lists column names of the DataFrame.
- Used to check for the presence of specific columns.

### 3. `raise ValueError:`

- Raises an error with a custom message if a condition is not met.

## Prepare Output File

Define the name of the output Excel file where processed data will be saved. This ensures data can be incrementally written to avoid loss in case of interruptions.


```python
output_file = 'output_coordinates.xlsx'  # Define the output file name
```

Sets the name of the file where processed data with coordinates will be saved.

## Process URLs and Extract Coordinates

- Iterate through each URL in the URL column.
- For each URL:
    1. Call the extraction function to retrieve latitude and longitude.
    2. Update the `DataFrame` with the extracted values.
    3. Save the updated `DataFrame` to the output file incrementally.
- Add a progress bar using `tqdm` to provide real-time feedback during the process.


```python
# Process each URL to extract coordinates with a progress bar
for index, url in tqdm(enumerate(df['URL']), desc="Processing URLs", unit="URL"):
    lat, lng = extract_coordinates(url)  # Extract coordinates
    df.at[index, 'Latitude'] = lat  # Add latitude to DataFrame
    df.at[index, 'Longitude'] = lng  # Add longitude to DataFrame
    
    # Incremental save to ensure data is not lost
    df.to_excel(output_file, index=False)
```

    Processing URLs: 2738URL [1:40:35,  2.20s/URL]
    

### 1. `enumerate():`

- Adds an index to the iteration.
- Useful for updating rows in the DataFrame.

### 2. `df.at:`

Updates a specific cell in the DataFrame using index and column name.

### 3. `tqdm():`

- Wraps an iterable to display progress during execution.
- Key arguments:
    - `desc`: Description of the task.
    - `unit`: Unit of progress (e.g., "URL").

### 4. `df.to_excel():`

- Writes the `DataFrame` to an Excel file.
- Arguments:
    - `index`: Whether to include row indices in the file.

## Final Save

After processing all URLs, ensure the complete `DataFrame` is saved to the output file.


```python
# Final export to ensure all data is saved
df.to_excel(output_file, index=False)
```

Ensures the DataFrame is fully saved after processing all URLs.

# Advantages of the Approach

- **Automation:** Eliminates manual effort by automating the extraction and saving processes.
- **Error Handling:** Handles exceptions gracefully, ensuring interruptions do not halt the process entirely.
- **Incremental Saves:** Reduces the risk of data loss by saving progress after processing each URL.
- **Scalability:** Handles large datasets efficiently, thanks to progress tracking with tqdm.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Extracting_Coordinates_from_Google_Maps_URLs/04_Extracting_Coordinates.png?raw=true)

# Conclusion

This Python-based solution simplifies the task of extracting coordinates from Google Maps URLs and exporting the results to Excel. By leveraging well-known libraries, it ensures accuracy, efficiency, and user-friendly feedback during execution.

# Future Enhancements

- Support for additional URL formats or geolocation data sources.
- Integration with mapping libraries like `folium` for visualizing coordinates on a map.
- Batch processing of multiple Excel files.

![](https://github.com/Umersaeed81/Python_For_RF_Optimization_And_Planning_Engineer/blob/main/Extracting_Coordinates_from_Google_Maps_URLs/05_Extracting_Coordinates.png?raw=true)


```python

```
