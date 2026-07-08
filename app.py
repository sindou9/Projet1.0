import streamlit as st
import pandas as pd
import os, time
import openpyxl
from openpyxl.utils import get_column_letter

# Configuration
FICHIER_EXCEL = "base_de_donnees.xlsx"
COLONNES = ["Client", "Statut Client", "Incident Declaré", "Agence", 
            "Intervenant", "Date Declaration Incident", "Date Debut Intervention", "Heure Debut Intervention", 
            "Date Fin Intervention", "Heure Fin Intervention", "Etat Incident", "Durée en Jours", "Durée De Remise en Etat",
            "Commentaire"]

# Fonction pour charger ou initialiser le fichier Excel
def charger_donnees():
    if os.path.exists(FICHIER_EXCEL):
        # Si le fichier existe, on le lit
        return pd.read_excel(FICHIER_EXCEL)
    else:
        # Sinon, on crée
        return pd.DataFrame(columns=COLONNES)

# Fonction pour sauvegarder les données
def sauvegarder_donnees(df):
    # On utilise ExcelWriter pour garder le contrôle sur la mise en page
    with pd.ExcelWriter(FICHIER_EXCEL, engine='openpyxl') as writer:
        # 1. Sauvegarde classique des données
        df.to_excel(writer, index=False, sheet_name='Source')
        
        # 2. Récupération de la feuille Excel générée
        worksheet = writer.sheets['Source']
        
        # 3. Boucle pour ajuster la largeur de chaque colonne
        for i, col in enumerate(df.columns):
            # Calcule la longueur max entre le nom de la colonne et son contenu
            taille_donnees = df[col].astype(str).map(len).max()
            taille_en_tete = len(str(col))
            taille_max = max(taille_donnees, taille_en_tete)
            
            # Convertit l'index de la boucle (0, 1, 2) en lettre Excel (A, B, C)
            lettre_col = get_column_letter(i + 1)
            
            # Applique la largeur (+2 pour que le texte ne soit pas trop collé aux bords)
            worksheet.column_dimensions[lettre_col].width = taille_max + 2


def exporter_excel():
    # Vérifie si le fichier existe
    if os.path.exists(FICHIER_EXCEL):
        # Crée un lien de téléchargement pour le fichier Excel
        with open(FICHIER_EXCEL, "rb") as f:
            st.download_button(
                label="📥 Télécharger la base de données Excel",
                data=f,
                file_name=FICHIER_EXCEL,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Le fichier Excel n'existe pas encore. Veuillez d'abord saisir des données.")

# ------ Interface avec streamlit

# Style de la page
st.set_page_config(page_title="Formulaire de Saisie Dynamique", layout="wide")
st.markdown("""
    <style>
    .stApp {
        background-color: auto;
    }
    </style>
    """, unsafe_allow_html=True)



st.title("📝 Formulaire de Saisie")
st.write("Les informations saisies ici seront automatiquement enregistrées dans Excel.")

# Chargement de l'état actuel des données
df_actuel = charger_donnees()

# 4. Création du formulaire
with st.form("formulaire_saisie", clear_on_submit=True):
    
    st.subheader("Nouvel Enregistrement")
    st.markdown(
    """
    <hr style="
        height: 4px; 
        border: auto; 
        background: linear-gradient(to right, red, yellow); 
        border-radius: 5px;
    "/>
    """, 
    unsafe_allow_html=True
    )
    
    # Champs du formulaire
    col1, col2 = st.columns(2)
    with col1:
        client = st.selectbox("Client *", ["BNI", "BACI", "SGCI", "AFG BANK", "CORIS"])
    with col2:
        statut_client = st.selectbox("Statut Client", ["Garantie", "Contrat", "Hors Contrat"])
    

    col1, col2 = st.columns(2)
    with col1:
        incident_declare = st.selectbox("Incident Déclaré", ["DEPANNAGE", "ENTRETIEN"])
    with col2:
        agence = st.text_input("Agence")
    
    st.divider()
    intervenant = st.text_input("Intervenant")

    col1, col2, col3 = st.columns(3)
    with col1:
        date_declaration_incident = st.date_input("Date Déclaration Incident")
    with col2:
        date_debut_intervention = st.date_input("Date Début Intervention")
    with col3:
        date_fin_intervention = st.date_input("Date Fin Intervention")

    col1, col2 = st.columns(2)
    with col1:
        heure_debut_intervention = st.time_input("Heure Début Intervention")
    with col2:
        heure_fin_intervention = st.time_input("Heure Fin Intervention")

    st.divider()

    etat_incident = st.selectbox("État Incident", ["En cours", "Resolu", "Non Résolu"])
    duree_en_jours = st.number_input("Durée en Jours", min_value=0)
    duree_de_remise_en_etat = st.number_input("Durée de Remise en État", min_value=0)
    commentaire = st.text_area("Commentaire ou observation")
    
    # Bouton de validation
    
    soumis = st.form_submit_button("Enregistrer les données", type="primary", help="Cliquez ici pour enregistrer les données dans le fichier Excel.")

    # 5. Logique de traitement après soumission
    if soumis:
        if client == "":
            st.error("Veuillez remplir le champ obligatoire (Client).")
        else:
            # Création d'une nouvelle ligne avec les données saisies
            date_declaration_incident = date_declaration_incident.strftime("%Y-%m-%d")
            date_debut_intervention = date_debut_intervention.strftime("%Y-%m-%d")
            date_fin_intervention = date_fin_intervention.strftime("%Y-%m-%d")
            heure_debut_intervention = heure_debut_intervention.strftime("%H:%M:%S")
            heure_fin_intervention = heure_fin_intervention.strftime("%H:%M:%S")

            nouvelle_saisie = pd.DataFrame({
                "Client": [client],
                "Statut Client": [statut_client],
                "Incident Declaré": [incident_declare],
                "Agence": [agence],
                "Intervenant": [intervenant],
                "Date Declaration Incident": [date_declaration_incident],
                "Date Debut Intervention": [date_debut_intervention],
                "Heure Debut Intervention": [heure_debut_intervention],
                "Date Fin Intervention": [date_fin_intervention],
                "Heure Fin Intervention": [heure_fin_intervention],
                "Etat Incident": [etat_incident],
                "Durée en Jours": [duree_en_jours],
                "Durée De Remise en Etat": [duree_de_remise_en_etat],
                "Commentaire": [commentaire]
            })
            
            # Concaténation avec les données existantes
            df_mis_a_jour = pd.concat([df_actuel, nouvelle_saisie], ignore_index=True)
            df_mis_a_jour = df_mis_a_jour[COLONNES]
            # Sauvegarde physique dans le fichier Excel
            sauvegarder_donnees(df_mis_a_jour)
            
            st.success("Données enregistrées avec succès !")
            # Mise à jour de l'affichage en direct
            df_actuel = df_mis_a_jour

            time.sleep(3)  # Petite pause 
            st.rerun()

# Afficher la base de données en bas de page pour vérifier
st.divider()
with st.expander("Voir les données enregistrées (Excel)"):
    st.subheader("📊 Base de données Excel actuelle")
    st.dataframe(df_actuel,
                 use_container_width=True,
                 hide_index=True)