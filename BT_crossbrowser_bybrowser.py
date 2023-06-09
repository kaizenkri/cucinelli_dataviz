import os
import pandas as pd
import numpy as np

# Define the list of browsers
browsers = ['chrome', 'safari', 'other']

# The path where your files are located
base_path = '/Users/cristianmarchisio/Python/cucinelli_dataviz/'

# Pages of interest
pages_of_interest = ['product-list', 'homepage', 'Discovery', 'product-detail', 'Inspiration']

# Metrics of interest
metrics_of_interest = ['Page Views', 'Onload (s)', 'First Byte (s)', 'First Contentful Paint (s)',
                       'Largest Contentful Paint (s)', 'First Input Delay Duration (s)', 'Cumulative Layout Shift']

# Initialize a list to store the data
data = []

for browser in browsers:
    old_file_path = os.path.join(base_path, f'BT_crossbrowser_{browser}_old.xlsx')
    new_file_path = os.path.join(base_path, f'BT_crossbrowser_{browser}_new.xlsx')

    df_old = pd.read_excel(old_file_path)
    df_new = pd.read_excel(new_file_path)

    df_old = df_old[df_old['Page Name'].isin(pages_of_interest)]
    df_new = df_new[df_new['Page Name'].isin(pages_of_interest)]

    for page in pages_of_interest:
        # Subset the data for the current page
        df_old_page = df_old[df_old['Page Name'] == page]
        df_new_page = df_new[df_new['Page Name'] == page]

        for metric in metrics_of_interest:
            # Calculate the percentage change for the current metric
            pct_change = ((df_new_page[metric] - df_old_page[metric]) / df_old_page[metric]) * 100

            # Handle the cases when the old metric is 0 to avoid division by zero error
            pct_change = pct_change.replace([np.inf, -np.inf], np.nan)

            # Combine the new value and the percentage change
            cell_value = f'{df_new_page[metric].values[0]} ({pct_change.values[0]:.2f}%)' if not np.isnan(pct_change.values[0]) else df_new_page[metric].values[0]

            # Add the percentage change to the data list
            data.append([metric, page, browser, cell_value])

# Convert the data list to a DataFrame
df = pd.DataFrame(data, columns=['Metric', 'Page', 'Browser', 'Pct Change'])

# Reorder the browsers and then set the index
df['Browser'] = pd.Categorical(df['Browser'], ['chrome', 'safari', 'other'])
df['Page'] = pd.Categorical(df['Page'], ['homepage', 'product-list', 'product-detail', 'Discovery', 'Inspiration'])
df.set_index(['Metric', 'Page', 'Browser'], inplace=True)

# Unstack the DataFrame to get the desired format and sort the pages as requested
df = df.unstack(level=[1,2]).sort_index(axis=1, level=[1,2], sort_remaining=True)

# Create the output folder if it doesn't exist
os.makedirs(os.path.join(base_path, 'output'), exist_ok=True)

# Save the final DataFrame to an Excel file
df.to_excel(os.path.join(base_path, 'output', 'BT_crossbrowser_comparison.xlsx'))

print("The output has been saved in the 'output' folder.")
