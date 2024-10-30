import streamlit as st
import pandas as pd
from xlsx2csv import Xlsx2csv
from io import StringIO, BytesIO
import re
import os

# Funzione per estrarre ID_ORDINE dal nome del file
def extract_order_id(filename):
    match = re.search(r"_(\d+)_", filename)
    if match:
        return match.group(1)
    return ""

# Funzione per convertire XLSX in CSV
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

# Funzione per processare il CSV e applicare il calcolo dello sconto
def process_csv(data, discount_percentage, order_id, view_option):
    new_data = []
    current_model = None
    current_sizes = []
    current_price = None
    current_confirmed = []
    current_shipped = []  # Variabile per i valori "Spediti"
    current_model_name = None
    current_color_description = None
    current_upc = []
    current_product_type = None

    for index, row in data.iterrows():
        if 'Modello/Colore:' in row.values:
            if current_model is not None:
                for size, confirmed, shipped, upc in zip(current_sizes, current_confirmed, current_shipped, current_upc):
                    new_data.append([current_model, size, current_price, confirmed, shipped, current_model_name, current_color_description, upc, discount_percentage, current_product_type, order_id])
            current_model = row[row.values.tolist().index('Modello/Colore:') + 1]
            current_price = row[row.values.tolist().index('Prezzo all\'ingrosso') + 1]
            current_sizes = []
            current_confirmed = []
            current_shipped = []  # Inizializza lista per spediti
            current_upc = []
        elif 'Nome del modello:' in row.values:
            current_model_name = row[row.values.tolist().index('Nome del modello:') + 1]
        elif 'Descrizione colore:' in row.values:
            current_color_description = row[row.values.tolist().index('Descrizione colore:') + 1]
        elif 'Tipo di prodotto:' in row.values:
            current_product_type = row[row.values.tolist().index('Tipo di prodotto:') + 1]
        elif pd.notna(row[0]) and row[0] not in ['Misura', 'Totale qtà:', '']:
            current_sizes.append(str(row[0]))
            current_confirmed.append(str(row[5]))  # Assicurati che l'indice sia corretto per "Confermati"
            current_shipped.append(str(row[8]))  # Usa l'indice corretto per "Spediti"
            current_upc.append(str(row[1]))

    if current_model is not None:
        for size, confirmed, shipped, upc in zip(current_sizes, current_confirmed, current_shipped, current_upc):
            new_data.append([current_model, size, current_price, confirmed, shipped, current_model_name, current_color_description, upc, discount_percentage, current_product_type, order_id])

    # Filtra i dati finali
    filtered_data_final = [entry for entry in new_data if entry[1] not in ['Riga articolo:', 'Nome del modello:', 'Descrizione colore:', 'Tipo di prodotto:', '']]

    # Crea il DataFrame finale con tutte le colonne richieste
    final_df_filtered_complete = pd.DataFrame(
        filtered_data_final,
        columns=['Modello/Colore', 'Misura', 'Prezzo all\'ingrosso', 'Confermati', 'Spediti', 'Nome del modello', 'Descrizione colore', 'Codice a Barre (UPC)', 'Percentuale sconto', 'Tipo di prodotto', 'ID_ORDINE']
    )

    # Suddivisione del codice modello e colore
    final_df_filtered_complete['Codice'] = final_df_filtered_complete['Modello/Colore'].apply(lambda x: x.split('-')[0])
    final_df_filtered_complete['Colore'] = final_df_filtered_complete['Modello/Colore'].apply(lambda x: x.split('-')[1])

    # Conversioni e calcoli
    final_df_filtered_complete['Prezzo all\'ingrosso'] = final_df_filtered_complete['Prezzo all\'ingrosso'].str.replace('€', '').str.replace(',', '.').astype(float)
    final_df_filtered_complete['Confermati'] = pd.to_numeric(final_df_filtered_complete['Confermati'], errors='coerce').fillna(0).astype(int)
    final_df_filtered_complete['Spediti'] = pd.to_numeric(final_df_filtered_complete['Spediti'], errors='coerce').fillna(0).astype(int)

    # Calcola il prezzo finale
    final_df_filtered_complete['Prezzo finale'] = final_df_filtered_complete.apply(
        lambda row: row['Prezzo all\'ingrosso'] * (1 - float(row['Percentuale sconto']) / 100), axis=1
    )

    # Calcola il totale confermati e spediti
    final_df_filtered_complete['TOT CONFERMATI'] = final_df_filtered_complete.apply(
        lambda row: row['Prezzo finale'] * row['Confermati'], axis=1
    )
    final_df_filtered_complete['TOT SPEDITI'] = final_df_filtered_complete.apply(
        lambda row: row['Prezzo finale'] * row['Spediti'], axis=1
    )

    # Seleziona le colonne in base all'opzione scelta
    if view_option == "CONFERMATI":
        final_df_filtered_complete = final_df_filtered_complete[['Modello/Colore', 'Descrizione colore', 'Codice', 'Nome del modello', 'Tipo di prodotto', 'Colore', 'Misura', 'Codice a Barre (UPC)', 'ID_ORDINE', 'Confermati', 'Prezzo all\'ingrosso', 'Percentuale sconto', 'Prezzo finale', 'TOT CONFERMATI']]
    else:
        final_df_filtered_complete = final_df_filtered_complete[['Modello/Colore', 'Descrizione colore', 'Codice', 'Nome del modello', 'Tipo di prodotto', 'Colore', 'Misura', 'Codice a Barre (UPC)', 'ID_ORDINE', 'Spediti', 'Prezzo all\'ingrosso', 'Percentuale sconto', 'Prezzo finale', 'TOT SPEDITI']]

    # Esportazione del DataFrame in Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df_filtered_complete.to_excel(writer, index=False)

    return output.getvalue(), final_df_filtered_complete

# Interfaccia Streamlit
st.title("Nike order details")

# Caricamento del file XLSX
uploaded_file = st.file_uploader("Carica un file XLSX", type="xlsx")

if uploaded_file is not None:
    # Estrai il nome del file senza estensione e prova a ottenere l'ID_ORDINE
    original_filename = os.path.splitext(uploaded_file.name)[0]
    extracted_order_id = extract_order_id(original_filename)

    # Campo per l'ID ordine, precompilato con l'ID estratto se disponibile
    order_id = st.text_input("ID_ORDINE", value=extracted_order_id)

    # Opzione di visualizzazione per "CONFERMATI" o "SPEDITI"
    view_option = st.radio("Seleziona l'opzione di visualizzazione:", ("CONFERMATI", "SPEDITI"))

    # Converti il file XLSX in CSV
    df = convert_xlsx_to_csv(uploaded_file)

    if df is not None:
        # Input per la percentuale di sconto
        discount_percentage = st.number_input("Inserisci la percentuale di sconto sul prezzo whl", min_value=0.0, max_value=100.0, step=0.1)

        if st.button("Elabora"):
            # Processa il CSV e calcola il risultato
            processed_file, final_df = process_csv(df, discount_percentage, order_id, view_option)

            # Mostra l'anteprima del file elaborato
            st.write("Anteprima del file elaborato:")
            st.write(final_df)

            # Nome del file processato
            processed_filename = f"{original_filename}_processed.xlsx"

            # Permetti il download del file Excel elaborato
            st.download_button(
                label="Scarica il file elaborato",
                data=processed_file,
                file_name=processed_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
