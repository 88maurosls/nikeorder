import streamlit as st
import pandas as pd
from xlsx2csv import Xlsx2csv
from io import StringIO

def convert_xlsx_to_csv(file):
    try:
        # Converti il file XLSX in CSV usando xlsx2csv
        output = StringIO()
        Xlsx2csv(file, outputencoding="utf-8").convert(output)
        output.seek(0)
        df = pd.read_csv(output)
        return df
    except Exception as e:
        st.error(f"Si è verificato un errore: {str(e)}")
        return None

# Titolo dell'app
st.title("Convertitore da XLSX a CSV")

# Caricamento del file XLSX
uploaded_file = st.file_uploader("Carica un file XLSX", type="xlsx")

if uploaded_file is not None:
    # Converti il file XLSX in CSV
    df = convert_xlsx_to_csv(uploaded_file)

    if df is not None:
        st.write("Anteprima del CSV convertito:")
        st.write(df)

        # Scarica il file CSV
        csv = df.to_csv(index=False)
        st.download_button(label="Scarica CSV", data=csv, file_name="converted_file.csv", mime="text/csv")

