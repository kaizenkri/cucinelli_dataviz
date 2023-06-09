import os
import pandas as pd

def process_files(device_type):
    # Percorsi dei file di input
    file_path = f'/Users/cristianmarchisio/Python/cucinelli_dataviz/BT_crossbrowser_chrome_old.xlsx'
    new_file_path = f'/Users/cristianmarchisio/Python/cucinelli_dataviz/BT_crossbrowser_chrome_new.xlsx'

    # Leggi i due file Excel
    df_old = pd.read_excel(file_path)
    df_new = pd.read_excel(new_file_path)

    # Seleziona le pagine di interesse
    pages_of_interest = ['product-list', 'homepage', 'Discovery', 'product-detail', 'Inspiration']

    # Filtra i DataFrame in base alle pagine di interesse
    df_old = df_old[df_old['Page Name'].isin(pages_of_interest)]
    df_new = df_new[df_new['Page Name'].isin(pages_of_interest)]

    # Seleziona le colonne di interesse
    columns_of_interest = ['Page Views', 'Onload (s)', 'First Byte (s)', 'First Contentful Paint (s)',
                           'Largest Contentful Paint (s)', 'First Input Delay Duration (s)', 'Cumulative Layout Shift']

    # Crea un dataframe per il confronto
    comparison_df = pd.DataFrame()

    # Aggiungi la colonna 'Page Name' al dataframe di confronto
    comparison_df['Page Name'] = df_new['Page Name']

    # Aggiungi i dati presi da entrambi i file per le colonne di interesse
    for column in columns_of_interest:
        new_column_name = column
        comparison_df[new_column_name] = df_new[column]
        comparison_df[column] = df_new[column].astype(str) + " (" + (((df_new[column] - df_old[column]) / df_old[column]) * 100).round(2).astype(str) + "%)"

    # Trasponi il DataFrame
    comparison_df = comparison_df.T

    # Percorso della cartella di output
    output_folder = os.path.join(os.path.dirname(file_path), 'output')

    # Crea la cartella di output se non esiste già
    os.makedirs(output_folder, exist_ok=True)

    # Salva il DataFrame trasposto come file Excel nella cartella di output
    output_file = os.path.join(output_folder, f'BT_crossbrowser_{device_type}.xlsx')
    comparison_df.to_excel(output_file)

    print(f"L'output è stato salvato in: {output_file}")

# Processa i file 'mobile'
process_files('mobile')

# Processa i file 'desktop'
process_files('desktop')