import streamlit as st
import pandas as pd
from xlsx2csv import Xlsx2csv
from io import StringIO, BytesIO
import os

def convert_xlsx_to_csv(file):
    try:
        output = StringIO()
        Xlsx2csv(file, outputencoding="utf-8").convert(output)
        output.seek(0)
        df = pd.read_csv(output)
        return df
    except Exception as e:
        st.error(f"Si è verificato un errore durante la conversione: {str(e)}")
        return None

def process_csv(data, discount_percentage):
    new_data = []
    current_model = None
    current_sizes = []
    current_price = None
    current_confirmed = []
    current_model_name = None
    current_color_description = None
    current_upc = []
    current_product_type = None

    for index, row in data.iterrows():
        if 'Modello/Colore:' in row.values:
            if current_model is not None:
                for size, confirmed, upc in zip(current_sizes, current_confirmed, current_upc):
                    new_data.append([current_model, size, current_price, confirmed, current_model_name, current_color_description, upc, discount_percentage, current_product_type])
            current_model = row[row.values.tolist().index('Modello/Colore:') + 1]
            current_price = row[row.values.tolist().index('Prezzo all\'ingrosso') + 1]
            current_sizes = []
            current_confirmed = []
            current_upc = []
        elif 'Nome del modello:' in row.values:
            current_model_name = row[row.values.tolist().index('Nome del modello:') + 1]
        elif 'Descrizione colore:' in row.values:
            current_color_description = row[row.values.tolist().index('Descrizione colore:') + 1]
        elif 'Tipo di prodotto:' in row.values:
            current_product_type = row[row.values.tolist().index('Tipo di prodotto:') + 1]
        elif pd.notna(row[0]) and row[0] not in ['Misura', 'Totale qtà:', '']:
            current_sizes.append(str(row[0]))
            current_confirmed.append(str(row[5]))
            current_upc.append(str(row[1]))

    if current_model is not None:
        for size, confirmed, upc in zip(current_sizes, current_confirmed, current_upc):
            new_data.append([current_model, size, current_price, confirmed, current_model_name, current_color_description, upc, discount_percentage, current_product_type])

    filtered_data_final = [entry for entry in new_data if entry[1] not in ['Riga articolo:', 'Nome del modello:', 'Descrizione colore:', 'Tipo di prodotto:', '']]

    final_df_filtered_complete = pd.DataFrame(filtered_data_final, columns=['Modello/Colore', 'Misura', 'Prezzo all\'ingrosso', 'Confermati', 'Nome del modello', 'Descrizione colore', 'Codice a Barre (UPC)', 'Percentuale sconto', 'Tipo di prodotto'])

    final_df_filtered_complete['Codice'] = final_df_filtered_complete['Modello/Colore'].apply(lambda x: x.split('-')[0])
    final_df_filtered_complete['Colore'] = final_df_filtered_complete['Modello/Colore'].apply(lambda x: x.split('-')[1])

    final_df_filtered_complete['Prezzo all\'ingrosso'] = final_df_filtered_complete['Prezzo all\'ingrosso'].str.replace('€', '').str.replace(',', '.').astype(float)
    final_df_filtered_complete['Confermati'] = pd.to_numeric(final_df_filtered_complete['Confermati'], errors='coerce').fillna(0).astype(int)

    final_df_filtered_complete['Prezzo finale'] = final_df_filtered_complete.apply(
        lambda row: row['Prezzo all\'ingrosso'] * (1 - float(row['Percentuale sconto']) / 100), axis=1
    )

    final_df_filtered_complete['Prezzo totale'] = final_df_filtered_complete.apply(
        lambda row: row['Prezzo finale'] * row['Confermati'], axis=1
    )

    final_df_filtered_complete = final_df_filtered_complete[final_df_filtered_complete['Confermati'] != 0]
    final_df_filtered_complete = final_df_filtered_complete[~final_df_filtered_complete['Misura'].str.contains('Prezzi:|Tutti i prezzi al netto di I.V.A. e spese di spedizione e altre tasse che si', na=False)]

    final_df_filtered_complete.reset_index(drop=True, inplace=True)
    final_df_filtered_complete = final_df_filtered_complete.fillna('')

    final_df_filtered_complete = final_df_filtered_complete[['Modello/Colore', 'Descrizione colore', 'Codice', 'Nome del modello', 'Tipo di prodotto', 'Colore', 'Misura', 'Codice a Barre (UPC)', 'Confermati', 'Prezzo all\'ingrosso', 'Percentuale sconto', 'Prezzo finale', 'Prezzo totale']]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df_filtered_complete.to_excel(writer, index=False)

    return output.getvalue()

# Interfaccia Streamlit
st.title("Convertitore da XLSX a CSV e Processor con Sconto")

# Caricamento del file XLSX
uploaded_file = st.file_uploader("Carica un file XLSX", type="xlsx")

if uploaded_file is not None:
    # Estrai il nome del file senza estensione
    original_filename = os.path.splitext(uploaded_file.name)[0]

    # Converti il file XLSX in CSV
    df = convert_xlsx_to_csv(uploaded_file)

    if df is not None:
        st.write("Anteprima del CSV convertito:")
        st.write(df)

        # Input per la percentuale di sconto
        discount_percentage = st.number_input("Inserisci la percentuale di sconto", min_value=0.0, max_value=100.0, step=0.1)

        if st.button("Elabora e Scarica"):
            # Processa il CSV e calcola il risultato
            processed_file = process_csv(df, discount_percentage)

            # Nome del file processato
            processed_filename = f"{original_filename}_processed.xlsx"

            # Permetti il download del file Excel elaborato
            st.download_button(
                label="Scarica il file elaborato",
                data=processed_file,
                file_name=processed_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
