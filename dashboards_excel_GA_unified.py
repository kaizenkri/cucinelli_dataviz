import pandas as pd
import numpy as np
import os

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

def get_output_path(input_file_path, output_file_name):
    directory = os.path.dirname(input_file_path)
    output_dir = os.path.join(directory, 'output')
    os.makedirs(output_dir, exist_ok=True)  # create the output directory if it doesn't exist
    output_path = os.path.join(output_dir, output_file_name)
    return output_path

# Load and process Excel files for Sources dimension
df_new_sources = process_excel("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_sources_new.xlsx", 'Sorgente/Mezzo', 'Sessioni')
df_old_sources = process_excel("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_sources_old.xlsx", 'Sorgente/Mezzo', 'Sessioni')

df_diff_sources = pd.DataFrame()
df_diff_sources[df_new_sources.columns[0]] = df_new_sources[df_new_sources.columns[0]]

df_diff_sources['Sessioni'] = df_new_sources['Sessioni'].astype(str) + ' (' + ((df_new_sources['Sessioni'] - df_old_sources['Sessioni']) / df_old_sources['Sessioni'] * 100).round(2).astype(str) + '%)'

for col in df_new_sources.columns[2:]:
    df_diff_sources[col] = (df_new_sources[col]*100).astype(str) + '% (' + ((df_new_sources[col] - df_old_sources[col])*100).round(2).astype(str) + '%)'

# Transpose the DataFrame for Sources dimension
df_diff_T_sources = df_diff_sources.transpose()

# Save the transposed DataFrame for Sources dimension
df_diff_T_sources.to_excel(get_output_path("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_sources_new.xlsx", 'Cucinelli_GA_sources_delta.xlsx'))

# Load and process Excel files for Page Categories dimension
df_new_pagecat = process_excel("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_pagecat_new.xlsx", 'pageCategory', 'Sessioni')
df_old_pagecat = process_excel("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_pagecat_old.xlsx", 'pageCategory', 'Sessioni')

df_diff_pagecat = pd.DataFrame()
df_diff_pagecat[df_new_pagecat.columns[0]] = df_new_pagecat[df_new_pagecat.columns[0]]

df_diff_pagecat['Sessioni'] = df_new_pagecat['Sessioni'].astype(str) + ' (' + ((df_new_pagecat['Sessioni'] - df_old_pagecat['Sessioni']) / df_old_pagecat['Sessioni'] * 100).round(2).astype(str) + '%)'

for col in df_new_pagecat.columns[2:]:
    df_diff_pagecat[col] = (df_new_pagecat[col]*100).astype(str) + '% (' + ((df_new_pagecat[col] - df_old_pagecat[col])*100).round(2).astype(str) + '%)'

# Transpose the DataFrame for Page Categories dimension
df_diff_T_pagecat = df_diff_pagecat.transpose()
# Save the transposed DataFrame for Page Categories dimension
df_diff_T_pagecat.to_excel(get_output_path("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_pagecat_new.xlsx", 'Cucinelli_GA_pagecat_delta.xlsx'))

print("Differences for Sources saved to: ", get_output_path("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_sources_new.xlsx", 'Cucinelli_GA_sources_delta.xlsx'))
print("Differences for Page Categories saved to: ", get_output_path("/Users/cristianmarchisio/Python/cucinelli_dataviz/GA_pagecat_new.xlsx", 'Cucinelli_GA_pagecat_delta.xlsx'))

