import streamlit as st
import pandas as pd
import os
import time
from openpyxl.utils import get_column_letter

# Configuration
FICHIER_EXCEL = "data_bases.xlsx"

COLONNES = ["Client", "Statut Client", "Incident Déclaré", "Agence", 
            "Intervenant", "Date Declaration Incident", "Date Debut Intervention", "Heure Debut Intervention", 
            "Date Fin Intervention", "Heure Fin Intervention", "Etat Incident", "Durée en Jours", "Durée de Remise en Etat",
            "Commentaire"]

# Fonction pour charger ou initialiser le fichier Excel
def charger_donnees():
    if os.path.exists(FICHIER_EXCEL):
        # Lecture du fichier Excel existant
        df = pd.read_excel(FICHIER_EXCEL, sheet_name='Source')
        # Vérification des colonnes
        for col in COLONNES:
            if col not in df.columns:
                df[col] = ""
        df = df[COLONNES]
    else:
        # Création d'un DataFrame vide avec les colonnes définies
        df = pd.DataFrame(columns=COLONNES)
        # Sauvegarde initiale dans le fichier Excel
        sauvegarder_donnees(df)
    return df

# Fonction pour sauvegarder les données
def sauvegarder_donnees(df):
    # On utilise ExcelWriter pour garder le contrôle sur la mise en page
    with pd.ExcelWriter(FICHIER_EXCEL, engine='openpyxl') as writer:
        # Sauvegarde classique des données
        df.to_excel(writer, index=False, sheet_name='Source')
        
        # Récupération de la feuille Excel générée
        worksheet = writer.sheets['Source']
        
        # Boucle pour ajuster la largeur de chaque colonne
        for i, col in enumerate(df.columns):
            # Sécurité : Si le DataFrame est vide, on évite l'erreur NaN
            taille_donnees = df[col].astype(str).map(len).max() if not df.empty else 0
            taille_en_tete = len(str(col))
            taille_max = max(taille_donnees, taille_en_tete)
            
            # Convertit l'index de la boucle (0, 1, 2) en lettre Excel (A, B, C)
            lettre_col = get_column_letter(i + 1)
            
            # Applique la largeur (+3 pour que le texte ne soit pas trop collé aux bords)
            worksheet.column_dimensions[lettre_col].width = taille_max + 3

# ------ Interface avec streamlit ------

# Style de la page
st.set_page_config(page_title="Formulaire de Saisie Dynamique", layout="wide")

st.title("📝 Formulaire de Saisie")
st.write("Les informations saisies ici seront automatiquement enregistrées dans Excel.")

# Chargement de données
df_actuel = charger_donnees()

# CREATION DU FORMULAIRE DE SAISIE
with st.form("formulaire_saisie", clear_on_submit=True):
    
    st.subheader("Nouvel Enregistrement")
    st.markdown(
        """
        <hr style="
            height: 4px; 
            border: none; 
            background: linear-gradient(to right, red, yellow); 
            border-radius: 5px;
        "/>
        """, 
        unsafe_allow_html=True
    )
    
    # Champs du formulaire
    col1, col2 = st.columns(2)
    default_client = ["BNI", "BACI", "SGCI", "AFG BANK", "CORIS"]
    
    with col1:
        client_selection = st.selectbox(
            "Client *",
            options=default_client + ["Autre (Nouveau client)..."], 
            index=None,
            placeholder="Sélectionnez un client"
        )
        
        # Si l'utilisateur choisit "Autre", on affiche un champ texte pour l'écrire
        if client_selection == "Autre (Nouveau client)...":
            client = st.text_input("Veuillez saisir le nom du nouveau client *")
        else:
            client = client_selection

    with col2:
        statut_client = st.selectbox("Statut Client *", ["Garantie", "Contrat", "Hors Contrat"])
    
    col3, col4 = st.columns(2)
    with col3:
        incident = st.selectbox("Incident Déclaré *", ["MAINTENANCE CURATIVE", "MAINTENANCE PREVENTIVE", "INSTALLATION", "AUTRE"])
        if incident == "AUTRE":
            incident_declare = st.text_input("Veuillez préciser l'incident déclaré")
        else:
            incident_declare = incident

    with col4:
        agence = st.text_input("Agence *")
    
    st.divider()
    intervenant = st.text_input("Intervenant *")

    col5, col6, col7 = st.columns(3)
    with col5:
        date_declaration_incident = st.date_input("Date de Déclaration de l'Incident")
    with col6:
        date_debut_intervention = st.date_input("Date de Début de l'Intervention")
    with col7:
        date_fin_intervention = st.date_input("Date de Fin de l'Intervention")

    col8, col9 = st.columns(2)
    with col8:
        heure_debut_intervention = st.time_input("Heure de Début de l'Intervention")
    with col9:
        heure_fin_intervention = st.time_input("Heure de Fin de l'Intervention")

    st.divider()

    etat_incident = st.selectbox("État de l'Incident *", ["En cours", "Resolu", "Non Résolu"])
    duree_en_jours = st.number_input("Durée en Jours", min_value=0)
    duree_de_remise_en_etat = st.number_input("Durée de Remise en État (minutes)", min_value=0)
    commentaire = st.text_area("Commentaire ou observation")
    
    # Bouton de validation du formulaire
    soumis = st.form_submit_button("Enregistrer les données", type="primary", help="Cliquez ici pour enregistrer les données dans le fichier Excel.")

    # Traitement
    if soumis:

        if not client or not statut_client or not incident_declare or not agence or not intervenant:
            st.error("⚠️ Veuillez remplir tous les champs obligatoires.")
        else:
            # Conversion des dates et heures en texte
            date_declaration_incident_str = date_declaration_incident.strftime("%d-%m-%Y")
            date_debut_intervention_str = date_debut_intervention.strftime("%d-%m-%Y")
            date_fin_intervention_str = date_fin_intervention.strftime("%d-%m-%Y")

            heure_debut_intervention_str = heure_debut_intervention.strftime("%H:%M")
            heure_fin_intervention_str = heure_fin_intervention.strftime("%H:%M")

            # Création d'une nouvelle ligne
            nouvelle_saisie = pd.DataFrame({
                "Client": [client],
                "Statut Client": [statut_client],
                "Incident Déclaré": [incident_declare],
                "Agence": [agence],
                "Intervenant": [intervenant],
                "Date de Déclaration de l'Incident": [date_declaration_incident_str],
                "Date de Début de l'Intervention": [date_debut_intervention_str],
                "Heure de Début de l'Intervention": [heure_debut_intervention_str],
                "Date de Fin de l'Intervention": [date_fin_intervention_str],
                "Heure de Fin de l'Intervention": [heure_fin_intervention_str],
                "État de l'Incident": [etat_incident],
                "Durée en Jours": [duree_en_jours],
                "Durée de Remise en État (minutes)": [duree_de_remise_en_etat],
                "Commentaire": [commentaire]
            })
            
            # Concaténation et sauvegarde
            df_mis_a_jour = pd.concat([df_actuel, nouvelle_saisie], ignore_index=True)
            df_mis_a_jour = df_mis_a_jour[COLONNES] # ordre sur les colonnes
            
            sauvegarder_donnees(df_mis_a_jour)
            
            st.success("✅ Données enregistrées avec succès !")
            time.sleep(2)  # Petite pause de 2 secondes
            st.rerun()     # Actualise l'interface

# 3. Éléments HORS du formulaire (Téléchargement et Affichage)
st.divider()

# Bouton de téléchargement correctement placé en dehors du st.form
if os.path.exists(FICHIER_EXCEL):
    with open(FICHIER_EXCEL, "rb") as fichier:
        st.download_button(
            label="📥 Télécharger la base de données (Excel)",
            data=fichier,
            file_name="Base_Interventions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with st.expander("Voir les données enregistrées (Excel)"):
    st.subheader("📊 Base de données Excel actuelle")
    st.dataframe(
        df_actuel,
        use_container_width=True,
        hide_index=True
    )