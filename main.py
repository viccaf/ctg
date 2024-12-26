import streamlit as st
import pandas as pd
from thefuzz import process
import os
import numpy as np

def load_value_list():
    try:
        return pd.read_excel('value_list.xlsx')
    except FileNotFoundError:
        st.error("Le fichier value_list.xlsx n'a pas été trouvé dans le répertoire courant.")
        return None

def find_closest_match(val, choices):
    if pd.isna(val):
        return None, 0
    match, score = process.extractOne(str(val), choices)
    return (match, score) if score >= 90 else ("REFUSÉ", score)

def process_file(uploaded_file, value_list_df):
    if uploaded_file is None:
        return None
    
    try:
        df = pd.read_excel(uploaded_file)
        
        if "VAL_NAT_CRP *" not in df.columns:
            st.error("La colonne 'VAL_NAT_CRP *' n'existe pas dans le fichier.")
            return None
            
        choices = value_list_df["Valeur CRP"].astype(str).tolist()
        
        # Créer les colonnes pour les matches et scores
        matches_and_scores = df["VAL_NAT_CRP *"].apply(
            lambda x: find_closest_match(x, choices)
        )
        
        df["Matched_Value"] = matches_and_scores.apply(lambda x: x[0])
        df["Match_Score"] = matches_and_scores.apply(lambda x: x[1])
        
        return df
    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier: {str(e)}")
        return None

def display_statistics(df):
    total = len(df)
    refused = len(df[df["Matched_Value"] == "REFUSÉ"])
    accepted = total - refused
    corrected = len(df[
        (df["Matched_Value"] != "REFUSÉ") & 
        (df["VAL_NAT_CRP *"] != df["Matched_Value"])
    ])

    st.write("### Statistiques")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Mots acceptés", f"{accepted} ({(accepted/total*100):.1f}%)")
    with col2:
        st.metric("Mots corrigés", f"{corrected} ({(corrected/total*100):.1f}%)")
    with col3:
        st.metric("Mots refusés", f"{refused} ({(refused/total*100):.1f}%)")

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
            display_statistics(result_df)
            
            st.write("### Aperçu des résultats")
            # Formater le DataFrame pour l'affichage
            def color_refused(val):
                return 'background-color: red' if val == "REFUSÉ" else ''
            
            styled_df = result_df.style.apply(
                lambda x: ['background-color: red' if v == "REFUSÉ" else '' 
                          for v in x], 
                subset=['Matched_Value']
            )
            
            st.dataframe(styled_df)
            
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