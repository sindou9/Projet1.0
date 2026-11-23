# Suivi des interventions (Streamlit)

Petite application Streamlit pour saisir et suivre des interventions techniques
(maintenance curative ou préventive, installation, etc.). Chaque saisie est
enregistrée dans un fichier Excel.

## Fonctionnalités

- Formulaire de saisie : client, statut (garantie, contrat, hors contrat), type d'incident, agence, intervenant, dates et heures, état, durée, commentaire.
- Enregistrement automatique dans `base_de_donnees.xlsx` (feuille `Source`).
- Téléchargement de la base au format Excel et affichage des données saisies.

## Lancer l'application

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Contenu du dépôt

| Fichier | Rôle |
|---|---|
| `app.py` | Application principale (stockage Excel local) |
| `app_gsheets.py` | Variante qui stocke les données dans Google Sheets (`st-gsheets-connection`, nécessite des secrets Streamlit) |
| `base_de_donnees.xlsx` | Base Excel utilisée par `app.py` |
| `requirements.txt` | Dépendances Python |

## Données

Le fichier `base_de_donnees.xlsx` est versionné dans ce dépôt public. N'y mets
que des données fictives ou anonymisées.
