import streamlit as st

st.set_page_config(
    page_title="Academix DSS",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

st.markdown('<h1 class="main-header">📊 Academix DSS</h1>', unsafe_allow_html=True)
st.markdown("### Système de Gestion de Stock pour Quincailleries en Algérie")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📊 Tableau de Bord")
    st.markdown("Visualisez vos stocks, alertes et performances en temps réel.")
    st.page_link("pages/1_Tableau_de_Bord.py", label="→ Ouvrir le tableau de bord")

with col2:
    st.markdown("### 📞 Demander une Démo")
    st.markdown("Testez Academix DSS avec vos propres données de stock.")
    st.page_link("pages/2_Contact.py", label="→ Demander une démo gratuite")

st.divider()

st.subheader("🎯 Fonctionnalités")
features = [
    ("📦", "Gestion de Stock", "Suivi en temps réel"),
    ("💰", "Calcul CMUP", "Automatique"),
    ("📊", "Classification ABC", "Intelligente"),
    ("⚠️", "Alertes Stock", "Préventives"),
    ("🧾", "Facturation", "Rapide"),
    ("📈", "Tableaux de bord", "Interactifs"),
]
cols = st.columns(3)
for i, (icon, title, desc) in enumerate(features):
    with cols[i % 3]:
        st.markdown(f"**{icon} {title}** — {desc}")

st.divider()

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("📦 Articles", "40", "+2")
with col2:
    st.metric("💰 Valeur Stock", "53.9M DZD", "+5.2%")
with col3:
    st.metric("⚠️ Alertes", "0", "-3")

st.divider()

st.markdown("""
### 💳 Tarification
| Offre | Prix | Articles |
|-------|------|----------|
| 🥉 Amateur | **15 000 DZD** | Jusqu'à 100 |
| 🥈 Professionnel | **25 000 DZD** | Jusqu'à 500 |
| 🥇 Entreprise | **45 000 DZD** | Illimité |
""")

st.info("🎁 **Offre de lancement** — Première démo gratuite + 1 mois d'essai")

st.divider()
st.caption("📧 kamelmahi71@gmail.com | 📞 +213 676 77 38 92 | © 2026 Academix DSS")
