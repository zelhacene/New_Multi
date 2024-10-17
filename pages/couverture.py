import pandas as pd
import streamlit as st
import calendar
from io import BytesIO
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

# Charger le DataFrame depuis le fichier Excel
df_cn = pd.read_excel(r"H:\Drive partagés\ACHAT 2024\Streamlit_app\Formule.xlsx", sheet_name="Consommation")

# Filtrer les lignes avec des descriptions spécifiques
df_cn = df_cn[~df_cn['Description'].isin([
    "Vitesse (Boites\\Heures)",
    "Nombre de Boites produit par jours",
    "Nombre de Boites par carton",
    "Nombre de Carton par palette",
    "Nombre d intercalaires par palette",
    "Nombre de palettes produits par jour",
    "Contenance (Litre)",
    "Litre de sirop fini produit en 24"
])]

# Liste des lignes disponibles
lignes_disponibles = ['Ligne 1/ 200 ml', 'Ligne 2/ 125 ml', 'Ligne 3/ 125 ml', 'Ligne 4/ 125 ml']

# Interface utilisateur avec Streamlit
st.title("Calcul couverture des Matières Premières")
selected_lignes = st.multiselect("Sélectionnez les lignes en marche", lignes_disponibles)

# Filtrer le DataFrame en fonction des lignes sélectionnées
filtered_df = df_cn[['Description'] + selected_lignes].copy()

# Créer une nouvelle colonne pour calculer le total pour les lignes sélectionnées
filtered_df['Consommation'] = filtered_df[selected_lignes].sum(axis=1)
filtered_df['Consommation Mensuelle'] = filtered_df['Consommation'] * 30

# Arrondir la colonne 'Consommation' à deux décimales
filtered_df['Consommation'] = filtered_df['Consommation'].round(2)
filtered_df['Consommation Mensuelle'] = filtered_df['Consommation Mensuelle'].round(2)

aujourd_hui = datetime.now()
Mois = aujourd_hui.month
Year = aujourd_hui.year
mois_en_lettres = calendar.month_name[Mois]

# Construisez le chemin du fichier Excel
chemin_fichier_excel = fr"H:\Drive partagés\ACHAT 2024\SITUATION STOCK\Suivi situation stock {mois_en_lettres} {Year}.xlsx"

# Charger le DataFrame depuis le fichier Excel
df_st = pd.read_excel(chemin_fichier_excel, sheet_name="situation")
df_st = df_st[['Description', 'St_Final']].dropna(subset=['Description'])
df_st = df_st[df_st['Description'] != "TOTAL DU MOIS"]

# Merge les deux tableaux
df_final = pd.merge(filtered_df, df_st, on='Description', how='left')

# Ajouter une nouvelle colonne 'Couverture/Jours' à df_final
df_final['Couverture/Jours'] = df_final['St_Final'] / df_final['Consommation']
df_final.fillna(0, inplace=True)

# Appliquer le format 123 333.00
df_final['St_Final'] = df_final['St_Final'].astype(int)
df_final['Consommation'] = df_final['Consommation'].astype(int)
df_final['Couverture/Jours'] = df_final['Couverture/Jours'].round(2)

# Afficher le tableau avec les colonnes "Description" et le total pour les lignes sélectionnées
st.table(df_final[['Description', 'St_Final', 'Consommation', 'Consommation Mensuelle', 'Couverture/Jours']])

# Fonction pour imprimer le tableau en format PDF
def print_to_pdf(data):
    st.markdown("**Impression en PDF en cours...**")

    # Convertir le DataFrame en une liste de listes
    table_data = [list(data.columns)] + data.values.tolist()

    # Créer un document PDF
    pdf_filename = "exported_table.pdf"
    pdf_buffer = BytesIO()
    pdf = SimpleDocTemplate(pdf_buffer, pagesize=letter)

    # Créer la table et définir le style
    table = Table(table_data)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # En-tête de colonne en gris
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Texte de l'en-tête de colonne en blanc
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Alignement au centre
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Police en gras pour l'en-tête de colonne
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Ajouter un espace en bas de l'en-tête de colonne
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # Couleur de fond des lignes impaires
    ])
    table.setStyle(style)

    # Ajouter la table au document PDF
    pdf.build([table])

    # Récupérer les données du PDF
    pdf_data = pdf_buffer.getvalue()

    # Télécharger le fichier PDF généré
    st.download_button(
        label="Télécharger PDF",
        data=pdf_data,
        key='pdf_export',
        file_name=pdf_filename,
        mime='application/pdf'
    )

# Bouton pour imprimer le tableau en format PDF
if st.button("Imprimer en PDF"):
    print_to_pdf(df_final)

# Afficher le graphique à barres
st.bar_chart(df_final.set_index('Description')['Couverture/Jours'])
