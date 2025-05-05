```python
import pandas as pd
from datetime import timedelta
```


```python
# Input DataFrame
df = pd.DataFrame({
    'Location': ['A', 'B','C','D','E','F','G','H'],
    'Date': pd.to_datetime(['2025-04-21', '2025-04-21','2025-04-23','2025-04-24','2025-04-29','2025-04-30','2025-05-1','2025-05-2'])
})
```


```python
def generate_labeled_dates(row):
    location = row['Location']
    current_date = row['Date']
    today =  pd.Timestamp.today().normalize()  # Automatically gets today's date

    # Calculate available post dates
    post_dates = []
    check_date = today - timedelta(days=1)
    while check_date > current_date and len(post_dates) < 7:
        post_dates.append(check_date)
        check_date -= timedelta(days=1)
    post_dates.reverse()  # So they are in ascending order

    # Now match pre dates count with post
    pre_dates = [current_date - timedelta(days=i) for i in range(len(post_dates), 0, -1)]

    # Combine
    dates = pre_dates + [current_date] + post_dates
    labels = ['pre'] * len(pre_dates) + ['current'] + ['post'] * len(post_dates)
    locations = [location] * len(dates)

    return pd.DataFrame({'Location': locations, 'Date': dates, 'Label': labels})
```


```python
# Apply to all rows
final_df = pd.concat([generate_labeled_dates(row) for _, row in df.iterrows()], ignore_index=True)

```


```python
# Output
final_df
```


```python
final_df.to_csv('dd.csv',index=False)
```


```python

```
