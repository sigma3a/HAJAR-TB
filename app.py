import pandas as pd
import streamlit as st
import plotly.express as px




# Charger le fichier Excel
@st.cache_data
def load_data(sheet_name):
    df = pd.read_excel("Suivi-M3 PIAFE LSH.xlsx", sheet_name=sheet_name, header=[0, 1])
    cl = pd.read_excel("Suivi-M3 PIAFE LSH.xlsx", sheet_name='Check-List', header=[0, 1])
    return df, cl

# Interface Streamlit
st.title('Document Checker for CIN')

# Ajouter un selectbox pour choisir la feuille
sheet = st.selectbox("Choisissez une feuille", ["Laayoune", "Essmara"])

# Charger les données en fonction de la feuille sélectionnée
df, cl = load_data(sheet)

# Fonction pour compter les lignes qui ont commencé le processus de création
def BeganCreation(df):
    count = 0
    for i in range(len(df)):
        na_flags = df.loc[i].isna()
        # Itérer à travers la série
        for value in na_flags:
            if not value:  # S'il n'y a pas de valeurs NaN (c'est-à-dire que value est False)
                count += 1
                break
    return count

# Fonction pour compter les lignes qui ont terminé le processus de création
def CompletedCreation(df):
    count = 0
    for i in range(len(df)):
        if not df.loc[i].isna().any():
            count += 1
    return count

# Calculer les comptes pour chaque catégorie
Count_RCPP_started = BeganCreation(df[[('RC-PP', i) for i in range(1, 9)]])
Count_RCPP_completed = CompletedCreation(df[[('RC-PP', i) for i in range(1, 9)]])
Count_AE_started = BeganCreation(df[[('AE', i) for i in range(1, 8)]])
Count_AE_completed = CompletedCreation(df[[('AE', i) for i in range(1, 8)]])
Count_SARL_started = BeganCreation(df[[('SARL', i) for i in range(1, 13)]])
Count_SARL_completed = CompletedCreation(df[[('SARL', i) for i in range(1, 13)]])
Count_COOP_started = BeganCreation(df[[('COOP', i) for i in range(1, 11)]])
Count_COOP_completed = CompletedCreation(df[[('COOP', i) for i in range(1, 11)]])

TotalCompleted = Count_RCPP_completed + Count_AE_completed + Count_SARL_completed + Count_COOP_completed
TotalStarted = Count_RCPP_started + Count_AE_started + Count_SARL_started + Count_COOP_started

# Afficher les comptes avec des sous-titres et des séparateurs
st.header("Statistiques de Création")
st.subheader("RC-PP")
st.write(f'Débuté : {Count_RCPP_started}')
st.write(f'Terminé : {Count_RCPP_completed}')
st.subheader("AE")
st.write(f'Débuté : {Count_AE_started}')
st.write(f'Terminé : {Count_AE_completed}')
st.subheader("SARL")
st.write(f'Débuté : {Count_SARL_started}')
st.write(f'Terminé : {Count_SARL_completed}')
st.subheader("COOP")
st.write(f'Débuté : {Count_COOP_started}')
st.write(f'Terminé : {Count_COOP_completed}')

st.subheader("Total")
st.write(f'Total débuté : {TotalStarted}') 
st.write(f'Total terminé : {TotalCompleted}') 

# Créer un DataFrame pour le graphique
data = {
    'Catégorie': ['RC-PP', 'AE', 'SARL', 'COOP', 'Total'],
    'Débuté': [Count_RCPP_started, Count_AE_started, Count_SARL_started, Count_COOP_started, TotalStarted],
    'Terminé': [Count_RCPP_completed, Count_AE_completed, Count_SARL_completed, Count_COOP_completed, TotalCompleted]
}

plot_df = pd.DataFrame(data)

# Tracer les données
fig = px.bar(plot_df, x='Catégorie', y=['Débuté', 'Terminé'], barmode='group', title='Créations Débutées vs Terminées')

st.plotly_chart(fig)

# Fonction pour trouver la première colonne non-NaN
def first_non_na_column(row, columns):
    non_na_columns = row[columns].dropna().index.tolist()
    return non_na_columns[0] if non_na_columns else None

# Fonction pour déterminer les documents restants
def DocRemaining(df, cl):
    rows_with_na_columns = {}
    for index, row in df.iterrows():
        # Obtenir les colonnes contenant des valeurs NaN
        na_columns = row[row.isna()].index.tolist()
        if not na_columns:
            return "Tous les documents sont fournis", []
        # Stocker le résultat dans le dictionnaire
        rows_with_na_columns[index] = na_columns
        remaining_docs = [cl[col].iloc[0] for col in rows_with_na_columns[index]]
        return 'Les documents restants sont:', remaining_docs
    return "Tous les documents sont fournis", []

# Fonction pour déterminer où chercher les documents restants en fonction du statut
def WhereToLook(df, cl, statut):
    if statut == 'COOP':
        lookHere = df[[('COOP', i) for i in range(1, 11)]]
    elif statut == 'RC-PP':
        lookHere = df[[('RC-PP', i) for i in range(1, 9)]]
    elif statut == 'SARL':
        lookHere = df[[('SARL', i) for i in range(1, 13)]]
    elif statut == 'Auto-entrepreneur':
        lookHere = df[[('AE', i) for i in range(1, 8)]]
    else:
        return "Tous les documents sont nécessaires", []

    return DocRemaining(lookHere, cl)

# Entrée du CIN
cin = st.text_input("Donner le CIN")

# Bouton pour vérifier les documents
if st.button('Vérifier les Documents'):
    if cin:
        ddf = df[df[('Information', 'CIN')] == cin]
        # Afficher les informations du propriétaire du CIN
        if not ddf.empty:
            st.header('Information du propriétaire du CIN:')
            owner_info = ddf.iloc[0][['Information']].dropna()
            for col in owner_info.index:
                st.write(f'{col[1]}: {owner_info[col]}')
        # Définir les colonnes à vérifier pour chaque catégorie
        st.write("\n\n")
        st.header('Statut des Documents:')
        PP = [('RC-PP', i) for i in range(1, 9)]
        AE = [('AE', i) for i in range(1, 8)]
        SARL = [('SARL', i) for i in range(1, 13)]
        COOP = [('COOP', i) for i in range(1, 11)]

        if not ddf.empty:
            statut = first_non_na_column(ddf.iloc[0], PP + AE + SARL + COOP)
            if statut:
                statut = statut[0]
                message, docs = WhereToLook(ddf, cl, statut)
                st.write(message)
                for doc in docs:
                    st.write(doc)
            else:
                st.write("Tous les documents sont nécessaires")
        else:
            st.write("CIN not found in the dataset.")
    else:
        st.write("Please enter a CIN.")
