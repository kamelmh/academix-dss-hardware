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
    /* Mobile responsive */
    @media (max-width: 768px) {
        .main-header { font-size: 1.8rem; }
        .stMetric { padding: 0.5rem; }
        .stColumns { flex-wrap: wrap; }
    }
    /* Feature cards */
    .feature-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        margin: 0.5rem 0;
        transition: transform 0.2s;
    }
    .feature-card:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    /* CTA button */
    .stButton > button {
        background: #1e40af;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        width: 100%;
    }
    .stButton > button:hover { background: #1e3a5f; }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-header">📊 Academix DSS</h1>', unsafe_allow_html=True)
st.markdown("### Système de Gestion de Stock pour Quincailleries en Algérie")

st.divider()

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("### 🎯 Démo Interactive")
    st.markdown("Explorez les fonctionnalités pas à pas.")
    st.page_link("pages/3_Demo_Interactif.py", label="→ Lancer la démo interactive")

with col2:
    st.markdown("### 📊 Tableau de Bord")
    st.markdown("Visualisez vos stocks, alertes et performances en temps réel.")
    st.page_link("pages/1_Tableau_de_Bord.py", label="→ Ouvrir le tableau de bord")

with col3:
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
