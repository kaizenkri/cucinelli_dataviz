import pandas as pd
import numpy as np
import csv

# Function to convert to integer
def convert_to_int(x):
    try:
        return int(str(x).replace('.', ''))  # remove period before converting to int
    except Exception as e:
        print(f"Failed to convert value: {x}, type: {type(x)}, error: {str(e)}")
        return np.nan

def clean_and_convert(x):
    try:
        return int(pd.to_numeric(x, errors='coerce').dropna())  # convert to numeric and drop NaN values
    except Exception as e:
        print(f"Failed to convert value: {x}, type: {type(x)}, error: {str(e)}")
        return np.nan

def process_excel(file_path, label_col, sessioni_col, sheet_name):
    # Read Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Convert 'Sessioni' column to integer
    df[sessioni_col] = df[sessioni_col].apply(convert_to_int)

    # Converting other columns to numeric types
    numeric_cols = [col for col in df.columns if col != sessioni_col]
    df[numeric_cols] = df[numeric_cols].applymap(clean_and_convert)

    return df

# Load and process Excel files for Sources dimension
df_new_sources = process_excel("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_sources_new.xlsx", 'Sorgente/Mezzo', 'Sessioni', 'Set di dati1')
df_old_sources = process_excel("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_sources_old.xlsx", 'Sorgente/Mezzo', 'Sessioni', 'Set di dati1')

# Load and process Excel files for Page Categories dimension
df_new_pagecat = process_excel("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_pagecat_new.xlsx", 'pageCategory', 'Sessioni', 'Set di dati1')
df_old_pagecat = process_excel("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_pagecat_old.xlsx", 'pageCategory', 'Sessioni', 'Set di dati1')

# ... Rest of your code ...


# Calculating differences for Sources dimension
df_diff_sources = pd.DataFrame()
df_diff_sources[df_new_sources.columns[0]] = df_new_sources[df_new_sources.columns[0]]

# Calculate percentage differences for Sessioni
df_diff_sources['Sessioni'] = df_new_sources['Sessioni'].astype(str) + ' (' + ((df_new_sources['Sessioni'] - df_old_sources['Sessioni']) / df_old_sources['Sessioni'] * 100).round(2).astype(str) + '%)'

for col in df_new_sources.columns[2:]:
    df_diff_sources[col] = df_new_sources[col].astype(str) + '% (' + (df_new_sources[col] - df_old_sources[col]).round(2).astype(str) + '%)'

# Transposing the DataFrame for Sources dimension
df_diff_T_sources = df_diff_sources.T
df_diff_T_sources.columns = df_diff_T_sources.iloc[0]
df_diff_T_sources = df_diff_T_sources[1:]

# Sort the DataFrame columns by Sessioni row in descending order for Sources dimension
df_diff_T_sources = df_diff_T_sources[df_diff_T_sources.loc['Sessioni'].str.extract(r"(\d+)", expand=False).astype(int).sort_values(ascending=False).index]

# Remove the index name for Sources dimension
df_diff_T_sources.index.name = None

# Displaying the DataFrame of differences for Sources dimension
print("=== Differences for Sources ===")
pd.set_option('display.max_rows', None)  # To display all rows
print(df_diff_T_sources)

# Saving the DataFrame to Excel for Sources dimension
df_diff_T_sources.to_excel('Cucinelli_GA_sources_delta.xlsx')


# Calculating differences for Page Categories dimension
df_diff_pagecat = pd.DataFrame()
df_diff_pagecat[df_new_pagecat.columns[0]] = df_new_pagecat[df_new_pagecat.columns[0]]

# Calculate percentage differences for Sessioni
df_diff_pagecat['Sessioni'] = df_new_pagecat['Sessioni'].astype(str) + ' (' + ((df_new_pagecat['Sessioni'] - df_old_pagecat['Sessioni']) / df_old_pagecat['Sessioni'] * 100).round(2).astype(str) + '%)'

for col in df_new_pagecat.columns[2:]:
    df_diff_pagecat[col] = df_new_pagecat[col].astype(str) + '% (' + (df_new_pagecat[col] - df_old_pagecat[col]).round(2).astype(str) + '%)'

# Transposing the DataFrame for Page Categories dimension
df_diff_T_pagecat = df_diff_pagecat.T
df_diff_T_pagecat.columns = df_diff_T_pagecat.iloc[0]
df_diff_T_pagecat = df_diff_T_pagecat[1:]

# Sort the DataFrame columns by Sessioni row in descending order for Page Categories dimension
df_diff_T_pagecat = df_diff_T_pagecat[df_diff_T_pagecat.loc['Sessioni'].str.extract(r"(\d+)", expand=False).astype(int).sort_values(ascending=False).index]

# Remove the index name for Page Categories dimension
df_diff_T_pagecat.index.name = None

# Displaying the DataFrame of differences for Page Categories dimension
print("=== Differences for Page Categories ===")
pd.set_option('display.max_rows', None)  # To display all rows
print(df_diff_T_pagecat)

# Saving the DataFrame to Excel for Page Categories dimension
df_diff_T_pagecat.to_excel('Cucinelli_GA_pagecat_delta.xlsx')