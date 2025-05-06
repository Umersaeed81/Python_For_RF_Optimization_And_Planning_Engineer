```python
import pandas as pd
from datetime import timedelta
```


```python
# Input DataFrame
df = pd.DataFrame({
    'Location': ['A', 'B','C','D','E','F','G','H'],
    'Date': pd.to_datetime(['2025-04-21', '2025-04-21','2025-04-23','2025-04-24','2025-04-29','2025-04-30','2025-05-01','2025-05-02'])
})
```


```python
def generate_labeled_dates(row):
    location = row['Location']
    current_date = row['Date']
    today = pd.Timestamp.today().normalize()

    # Step 1: Generate post dates (up to 7 days before today)
    post_dates = []
    check_date = today - timedelta(days=1)
    while check_date > current_date and len(post_dates) < 7:
        post_dates.append(check_date)
        check_date -= timedelta(days=1)
    post_dates.reverse()  # Ascending order

    # Step 2: Get the weekdays from post_dates
    post_weekdays = [d.weekday() for d in post_dates]  # 0=Mon, ..., 6=Sun

    # Step 3: For each post weekday, find last occurrence before current_date
    pre_dates = []
    for wd in post_weekdays:
        days_back = (current_date.weekday() - wd + 7) % 7 or 7
        pre_dates.append(current_date - timedelta(days=days_back))

    # Step 4: Combine
    dates = pre_dates + [current_date] + post_dates
    labels = ['pre'] * len(pre_dates) + ['current'] + ['post'] * len(post_dates)
    locations = [location] * len(dates)

    return pd.DataFrame({'Location': locations, 'Date': dates, 'Label': labels})

# Final output
final_df = pd.concat([generate_labeled_dates(row) for _, row in df.iterrows()], ignore_index=True)
```


```python
# 🗓️ Extract the day name (e.g., Monday) from the 'Date' column and 🏷️ store it in a new 'Day' column
final_df['Day'] = final_df['Date'].dt.day_name()  
```


```python
import os
working_directory = 'D:/Advance_Data_Sets/Asim'  # 📂 Define the target working directory
os.chdir(working_directory)  # 🔄 Change the current working directory to the specified path
```


```python
final_df.to_csv('gg.csv',index=False)
```


```python

```
