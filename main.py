import streamlit as st
import pandas as pd
from thefuzz import process
import os

def load_value_list():
    try:
        return pd.read_excel('value_list.xlsx')
    except FileNotFoundError:
        st.error("Le fichier value_list.xlsx n'a pas été trouvé dans le répertoire courant.")
        return None

def find_closest_match(val, choices):
    if pd.isna(val):
        return None
    match, score = process.extractOne(str(val), choices)
    return match

def process_file(uploaded_file, value_list_df):
    if uploaded_file is None:
        return None
    
    try:
        df = pd.read_excel(uploaded_file)
        
        if "VAL_NAT_CRP *" not in df.columns:
            st.error("La colonne 'VAL_NAT_CRP *' n'existe pas dans le fichier.")
            return None
            
        choices = value_list_df["Valeur CRP"].astype(str).tolist()
        df["Matched_Value"] = df["VAL_NAT_CRP *"].apply(
            lambda x: find_closest_match(x, choices)
        )
        
        return df
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier: {str(e)}")
        return None

def main():
    st.title("Matching Fuzzy des Valeurs CRP")
    
    value_list_df = load_value_list()
    if value_list_df is None:
        return
        
    uploaded_file = st.file_uploader(
        "Choisissez un fichier Excel (.xlsx)", 
        type="xlsx"
    )
    
    if uploaded_file:
        result_df = process_file(uploaded_file, value_list_df)
        
        if result_df is not None:
            st.write("Aperçu des résultats:")
            st.dataframe(result_df)
            
            # Option de téléchargement
            csv = result_df.to_csv(index=False)
            st.download_button(
                "Télécharger les résultats (CSV)",
                csv,
                "resultats_matching.csv",
                "text/csv",
                key='download-csv'
            )

if __name__ == "__main__":
    main()