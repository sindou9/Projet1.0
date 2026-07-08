import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection

# 1. Configuration de l'application
st.set_page_config(page_title="Formulaire de Saisie Dynamique", layout="wide")

# Lien de votre Google Sheets (Lien transmis dans votre message)
FICHIER_EXCEL = "https://docs.google.com/spreadsheets/d/1Go4fyX6NFsnOjsiZzz9lF_wvHSVJ7IqrxMk2RVbwJek/edit?usp=sharing

COLONNES = [
    "Client", "Statut Client", "Incident Declaré", "Agence", 
    "Intervenant", "Date Declaration Incident", "Date Debut Intervention", "Heure Debut Intervention", 
    "Date Fin Intervention", "Heure Fin Intervention", "Etat Incident", "Durée en Jours", "Durée De Remise en Etat",
    "Commentaire"
]

# 2. Fonction pour charger les données depuis Google Sheets
def charger_donnees():
    try:
        # Initialisation de la connexion gsheets
        conn = st.connection("gsheets", type=GSheetsConnection)
        # Lecture globale de la feuille (ttl=0d force l'actualisation à chaque rafraîchissement)
        data = conn.read(spreadsheet=FICHIER_EXCEL, ttl="0d")
        
        # Si la feuille est totalement vide, on retourne un tableau structuré vide
        if data.empty:
            return pd.DataFrame(columns=COLONNES)
            
        # Filtrer pour s'assurer d'avoir exactement les bonnes colonnes
        return data[COLONNES]
    except Exception as e:
        # Si échec, on crée un tableau vide avec la bonne structure pour ne pas bloquer l'interface
        return pd.DataFrame(columns=COLONNES)

# ------ Interface Utilisateur Streamlit ------

st.title("📝 Formulaire de Saisie")
st.write("Les informations saisies ici seront automatiquement enregistrées en ligne sur Google Sheets.")

# Chargement immédiat des données du cloud
df_actuel = charger_donnees()

# Formulaire de saisie principal
with st.form("formulaire_saisie", clear_on_submit=True):
    
    st.subheader("Nouvel Enregistrement")
    st.markdown(
        '<hr style="height: 4px; background: linear-gradient(to right, red, yellow); border-radius: 5px;"/>', 
        unsafe_allow_html=True
    )
    
    # Ligne 1 : Client & Statut
    col1, col2 = st.columns(2)
    with col1:
        client = st.selectbox("Client *", ["BNI", "BACI", "SGCI", "AFG BANK", "CORIS"])
    with col2:
        statut_client = st.selectbox("Statut Client", ["Garantie", "Contrat", "Hors Contrat"])
    
    # Ligne 2 : Type incident & Agence
    col1, col2 = st.columns(2)
    with col1:
        incident_declare = st.selectbox("Incident Déclaré", ["DEPANNAGE", "ENTRETIEN"])
    with col2:
        agence = st.text_input("Agence")
    
    st.divider()
    intervenant = st.text_input("Intervenant")

    # Ligne 3 : Les dates (Sélection visuelle)
    col1, col2, col3 = st.columns(3)
    with col1:
        date_declaration_incident = st.date_input("Date Déclaration Incident")
    with col2:
        date_debut_intervention = st.date_input("Date Début Intervention")
    with col3:
        date_fin_intervention = st.date_input("Date Fin Intervention")

    # Ligne 4 : Les heures
    col1, col2 = st.columns(2)
    with col1:
        heure_debut_intervention = st.time_input("Heure Début Intervention")
    with col2:
        heure_fin_intervention = st.time_input("Heure Fin Intervention")

    st.divider()

    # Ligne 5 : Métriques et commentaires
    col1, col2, col3 = st.columns(3)
    with col1:
        etat_incident = st.selectbox("État Incident", ["En cours", "Resolu", "Non Résolu"])
    with col2:
        duree_en_jours = st.number_input("Durée en Jours", min_value=0, step=1)
    with col3:
        duree_de_remise_en_etat = st.number_input("Durée de Remise en État", min_value=0, step=1)
        
    commentaire = st.text_area("Commentaire ou observation")
    
    # Bouton d'envoi
    soumis = st.form_submit_button(
        "Enregistrer les données", 
        type="primary", 
        help="Cliquez ici pour enregistrer directement dans le cloud Google Sheets."
    )

    # Logique après clic sur le bouton
    if soumis:
        if not client:
            st.error("Veuillez renseigner le champ obligatoire (Client).")
        else:
            # 1. Formatage propre des objets date/heure en chaînes textuelles standard
            str_date_dec = date_declaration_incident.strftime("%Y-%m-%d")
            str_date_deb = date_debut_intervention.strftime("%Y-%m-%d")
            str_date_fin = date_fin_intervention.strftime("%Y-%m-%d")
            str_heure_deb = heure_debut_intervention.strftime("%H:%M:%S")
            str_heure_fin = heure_fin_intervention.strftime("%H:%M:%S")

            # 2. Modélisation de la nouvelle ligne entrante
            nouvelle_saisie = pd.DataFrame([{
                "Client": client,
                "Statut Client": statut_client,
                "Incident Declaré": incident_declare,
                "Agence": agence,
                "Intervenant": intervenant,
                "Date Declaration Incident": str_date_dec,
                "Date Debut Intervention": str_date_deb,
                "Heure Debut Intervention": str_heure_deb,
                "Date Fin Intervention": str_date_fin,
                "Heure Fin Intervention": str_heure_fin,
                "Etat Incident": etat_incident,
                "Durée en Jours": int(duree_en_jours),
                "Durée De Remise en Etat": int(duree_de_remise_en_etat),
                "Commentaire": commentaire
            }])
            
            # 3. Fusion avec l'historique récupéré
            df_mis_a_jour = pd.concat([df_actuel, nouvelle_saisie], ignore_index=True)
            df_mis_a_jour = df_mis_a_jour[COLONNES]
            
            # 4. Envoi de la mise à jour globale vers Google Sheets
            try:
                conn = st.connection("gsheets", type=GSheetsConnection)
                conn.update(spreadsheet=FICHIER_EXCEL, data=df_mis_a_jour)
                
                st.success("Données enregistrées avec succès sur le Cloud Google Sheets !")
                st.rerun() # Recharge l'application pour afficher la ligne ajoutée instantanément
            except Exception as error_cloud:
                st.error("Impossible d'écrire sur Google Sheets.")
                st.info("Avez-vous bien partagé votre Google Sheets en mode 'Tous les utilisateurs disposant du lien peuvent modifier' comme Éditeur ?")

# Affichage synchrone de la base pour vérification visuelle
st.divider()
with st.expander("Voir les données enregistrées en direct (Google Sheets)"):
    st.subheader("📊 Base de données Cloud actuelle")
    st.dataframe(df_actuel, use_container_width=True, hide_index=True)
