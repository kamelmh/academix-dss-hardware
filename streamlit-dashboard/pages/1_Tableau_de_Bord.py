import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import random

st.set_page_config(page_title="Tableau de Bord", page_icon="📊", layout="wide")

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1e40af;
        text-align: center;
        margin-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_sample_data():
    products = pd.DataFrame({
        'ID': range(1, 41),
        'Nom': [
            'Vis Inox 4x40', 'Boulon Galva 6x80', 'Rondelle M8', 'Ecrou M6',
            'Clou 80mm', 'Clou 100mm', 'Fer à Beton 8mm', 'Fer à Beton 12mm',
            'Tôle Acier 2mm', 'Profil U 40x20', 'Profil T 30x30', 'Tube Rond 40mm',
            'Câble Électrique 2.5mm', 'Câble Électrique 4mm', 'Gaine ICTA 20mm', 'Gaine ICTA 32mm',
            'Switch 2P', 'Disjoncteur 16A', 'Disjoncteur 25A', 'Prise 16A',
            'Joints EPDM 10mm', 'Joints Cuivre 15mm', 'Ruban Teflon', 'Pâte Étanchéité',
            'Marteau 500g', 'Clé Anglaise 12"', 'Tournevis Ph 2', 'Pince Multiprise',
            'Perceuse Bosch', 'Meuleuse 125mm', 'Niveau 60cm', 'Mètre Ruban 5m',
            'Peinture Blanche 20L', 'Peinture Grise 20L', 'Solvant 5L', 'Pinceau 50mm',
            'Scie Circulaire', 'Scie à Main', 'Foret HSS 10mm', 'Foret HSS 13mm'
        ],
        'Catégorie': [
            'Visserie', 'Visserie', 'Visserie', 'Visserie',
            'Clouterie', 'Clouterie', 'Fer', 'Fer',
            'Tôlerie', 'Profilés', 'Profilés', 'Profilés',
            'Électricité', 'Électricité', 'Électricité', 'Électricité',
            'Électricité', 'Électricité', 'Électricité', 'Électricité',
            'Plomberie', 'Plomberie', 'Plomberie', 'Plomberie',
            'Outillage', 'Outillage', 'Outillage', 'Outillage',
            'Outillage', 'Outillage', 'Outillage', 'Outillage',
            'Peinture', 'Peinture', 'Peinture', 'Peinture',
            'Outillage', 'Outillage', 'Outillage', 'Outillage'
        ],
        'Prix_Unitaire': [
            15, 25, 5, 8, 12, 15, 180, 280,
            1200, 350, 300, 450, 85, 120, 25, 45,
            180, 350, 420, 95, 35, 45, 25, 65,
            850, 650, 180, 320, 15000, 8500, 2500, 450,
            4500, 4500, 2800, 120, 12000, 3500, 85, 95
        ],
        'Stock_Actuel': [
            15000, 8500, 25000, 18000, 12000, 9500, 450, 280,
            85, 120, 95, 75, 350, 220, 500, 280,
            150, 85, 65, 200, 450, 320, 250, 180,
            45, 35, 60, 40, 12, 18, 25, 50,
            25, 30, 45, 80, 8, 15, 200, 180
        ],
        'Seuil_Alerte': [
            5000, 3000, 8000, 6000, 4000, 3000, 150, 100,
            30, 40, 35, 25, 100, 75, 150, 90,
            50, 30, 25, 70, 150, 100, 80, 60,
            15, 12, 20, 15, 5, 6, 10, 20,
            10, 12, 15, 30, 3, 5, 70, 60
        ],
        'Fournisseur': [
            'Serral', 'Serral', 'Serral', 'Serral',
            'Brikol', 'Brikol', 'Acier Plus', 'Acier Plus',
            'Acier Plus', 'Profil Algeria', 'Profil Algeria', 'Profil Algeria',
            'CâbleTech', 'CâbleTech', 'CâbleTech', 'CâbleTech',
            'ÉlecPro', 'ÉlecPro', 'ÉlecPro', 'ÉlecPro',
            'PlombTech', 'PlombTech', 'PlombTech', 'PlombTech',
            'OutillagePro', 'OutillagePro', 'OutillagePro', 'OutillagePro',
            'Bosch', 'Bosch', 'Stanley', 'Stanley',
            'PeintureDZ', 'PeintureDZ', 'PeintureDZ', 'PeintureDZ',
            'Makita', 'Stanley', 'Bosch', 'Bosch'
        ]
    })

    products['Valeur_Stock'] = products['Prix_Unitaire'] * products['Stock_Actuel']
    products['CMUP'] = products['Prix_Unitaire'] * 0.92
    products['Rotation'] = [random.randint(2, 15) for _ in range(len(products))]

    total_value = products['Valeur_Stock'].sum()
    products = products.sort_values('Valeur_Stock', ascending=False)
    products['Cumul_Pct'] = products['Valeur_Stock'].cumsum() / total_value * 100
    products['ABC'] = products['Cumul_Pct'].apply(
        lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C')
    )

    products['Statut'] = products.apply(
        lambda x: 'Critique' if x['Stock_Actuel'] <= x['Seuil_Alerte'] * 0.5 else (
            'Alerte' if x['Stock_Actuel'] <= x['Seuil_Alerte'] else 'Normal'
        ), axis=1
    )

    transactions = pd.DataFrame({
        'Date': [(datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(20)],
        'Produit': [products['Nom'].iloc[i % len(products)] for i in range(20)],
        'Type': [random.choice(['Entrée', 'Sortie']) for _ in range(20)],
        'Quantité': [random.randint(10, 500) for _ in range(20)],
        'Montant': [random.randint(5000, 500000) for _ in range(20)]
    })

    return products, transactions

products, transactions = load_sample_data()

st.markdown('<h1 class="main-header">📊 Academix DSS — Tableau de Bord</h1>', unsafe_allow_html=True)
st.caption("Démo interactive — Gestion de stock pour quincailleries en Algérie")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric(label="📦 Articles", value=f"{len(products)}", delta="+2 cette semaine")
with col2:
    st.metric(label="💰 Valeur Stock", value=f"{products['Valeur_Stock'].sum():,.0f} DZD", delta="+5.2%")
with col3:
    st.metric(label="⚠️ Alertes", value=f"{len(products[products['Statut'] != 'Normal'])}", delta="-3")
with col4:
    st.metric(label="🔄 Rotation Moy.", value=f"{products['Rotation'].mean():.1f}", delta="+0.8")

st.divider()

col1, col2 = st.columns(2)
with col1:
    st.subheader("📈 Valeur Stock par Catégorie")
    cat_data = products.groupby('Catégorie')['Valeur_Stock'].sum().reset_index()
    fig = px.pie(cat_data, values='Valeur_Stock', names='Catégorie',
                 color_discrete_sequence=px.colors.qualitative.Set3)
    fig.update_traces(textposition='inside', textinfo='percent+label')
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("🎯 Classification ABC")
    abc_data = products.groupby('ABC').agg({'ID': 'count', 'Valeur_Stock': 'sum'}).reset_index()
    abc_data.columns = ['Classe', 'Nb_Articles', 'Valeur']
    fig = go.Figure()
    fig.add_trace(go.Bar(x=abc_data['Classe'], y=abc_data['Nb_Articles'], name='Nb Articles', marker_color='#3b82f6'))
    fig.add_trace(go.Bar(x=abc_data['Classe'], y=abc_data['Valeur'] / 1000000, name='Valeur (M DZD)', marker_color='#10b981', yaxis='y2'))
    fig.update_layout(yaxis=dict(title='Nombre d\'articles'), yaxis2=dict(title='Valeur (M DZD)', overlaying='y', side='right'), barmode='group')
    st.plotly_chart(fig, use_container_width=True)

col1, col2 = st.columns(2)
with col1:
    st.subheader("📊 Statut du Stock")
    status_counts = products['Statut'].value_counts()
    fig = go.Figure(data=[go.Pie(labels=status_counts.index, values=status_counts.values,
                                  marker_colors=['#10b981', '#f59e0b', '#ef4444'], textinfo='label+value')])
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("🔄 Top 10 — Rotation")
    top_rotation = products.nlargest(10, 'Rotation')
    fig = px.bar(top_rotation, x='Nom', y='Rotation', color='Catégorie',
                 color_discrete_sequence=px.colors.qualitative.Set2)
    fig.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)

st.divider()
st.subheader("⚠️ Alertes Stock")
alerts = products[products['Statut'] != 'Normal'].sort_values('Stock_Actuel')
if len(alerts) > 0:
    for _, row in alerts.iterrows():
        icon = "🔴" if row['Statut'] == 'Critique' else "🟡"
        st.write(f"{icon} **{row['Nom']}** — Stock: {row['Stock_Actuel']} | Seuil: {row['Seuil_Alerte']} | Fournisseur: {row['Fournisseur']}")
else:
    st.success("✅ Tous les stocks sont normaux")

st.divider()
st.subheader("📋 Transactions Récentes")
st.dataframe(transactions, use_container_width=True)

st.divider()
st.markdown("### 💡 Vous aimez ce que vous voyez ?")
st.markdown("Testez Academix DSS avec vos propres données de stock.")
st.page_link("pages/2_Contact.py", label="📞 Demander une démo gratuite", icon="→")

st.divider()
st.caption("© 2026 Academix DSS — Système de Gestion de Stock pour Quincailleries en Algérie")
