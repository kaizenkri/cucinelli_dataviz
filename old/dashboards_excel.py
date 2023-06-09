import pandas as pd
import numpy as np

# Function to convert to integer
def convert_to_int(x):
    try:
        return int(str(x).replace('.', ''))  # remove period before converting to int
    except Exception as e:
        print(f"Failed to convert value: {x}, type: {type(x)}, error: {str(e)}")
        return np.nan

def process_excel(file_path, label_col, sessioni_col, sheet_name='Set di dati1'):
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype={sessioni_col: str})

    # Removing unwanted rows
    df = df.iloc[:-1]

    # Convert 'Sessioni' column to integer
    df[sessioni_col] = df[sessioni_col].apply(convert_to_int)

    # Converting other columns to numeric types
    numeric_cols = [col for col in df.columns if col not in [label_col, sessioni_col]]
    df[numeric_cols] = df[numeric_cols].applymap(lambda x: float(x.replace(',', '.').replace('%', '')) if isinstance(x, str) else x)

    return df

# Load and process Excel files
df_new = process_excel("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_pagecat_new.xlsx", 'pageCategory', 'Sessioni')
df_old = process_excel("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_pagecat_old.xlsx", 'pageCategory', 'Sessioni')

# Calculating differences
df_diff = pd.DataFrame()
df_diff[df_new.columns[0]] = df_new[df_new.columns[0]]

# Calculate percentage differences for Sessioni
df_diff['Sessioni'] = df_new['Sessioni'].astype(str) + ' (' + ((df_new['Sessioni'] - df_old['Sessioni']) / df_old['Sessioni'] * 100).round(2).astype(str) + '%)'

# ...
for col in df_new.columns[2:]:
    df_diff[col] = (df_new[col]*100).astype(str) + '% (' + ((df_new[col] - df_old[col])*100).round(2).astype(str) + '%)'

# ...


# Transposing the DataFrame
df_diff_T = df_diff.T
df_diff_T.columns = df_diff_T.iloc[0]
df_diff_T = df_diff_T[1:]

# Sort the DataFrame columns by Sessioni row in descending order
df_diff_T = df_diff_T[df_diff_T.loc['Sessioni'].str.extract(r"(\d+)", expand=False).astype(int).sort_values(ascending=False).index]

# Remove the index name
df_diff_T.index.name = None

# Displaying the DataFrame of differences
pd.set_option('display.max_rows', None)  # To display all rows
print(df_diff_T)

# Saving the DataFrame to Excel
df_diff_T.to_excel('Cucinelli_GA_pagecat_delta.xlsx')
