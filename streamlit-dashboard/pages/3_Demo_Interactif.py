import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import time

st.set_page_config(page_title="Démo Interactive", page_icon="🎯", layout="wide")

st.markdown("""
<style>
    .step-box {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1rem;
    }
    .step-number {
        font-size: 2rem;
        font-weight: bold;
        color: #fbbf24;
    }
    .benefit-box {
        background: #f0fdf4;
        border-left: 4px solid #10b981;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
    }
    .problem-box {
        background: #fef2f2;
        border-left: 4px solid #ef4444;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #1e40af;'>🎯 Démo Interactive — Academix DSS</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #6b7280;'>Découvrez comment Academix DSS transforme la gestion de votre quincaillerie</p>", unsafe_allow_html=True)

st.divider()

step = st.radio(
    "Navigation",
    ["🏠 Accueil", "📊 Problème", "✨ Solution", "📈 Résultats", "📞 Contact"],
    horizontal=True,
    label_visibility="collapsed"
)

if step == "🏠 Accueil":
    st.markdown("## Bienvenue dans Academix DSS")
    st.markdown("""
    **Le premier système de gestion de stock** spécialement conçu pour les quincailleries en Algérie.

    ✅ Alimenté par des **vrais données** de quincaillerie
    ✅ Interface **simple** — pas besoin d'informaticien
    ✅ **Sauvegarde automatique** sur Excel
    """)

    col1, col2 = st.columns(2)
    with col1:
        st.info("💡 **Temps de formation** : Moins de 1 heure")
    with col2:
        st.info("🔒 **Sécurité** : Mot de passe protégé")

    st.markdown("---")
    st.markdown("### 🚀 Commencez l'exploration")
    st.markdown("Cliquez sur **📊 Problème** pour voir les défis que Academix résout.")

elif step == "📊 Problème":
    st.markdown("## Les problèmes des quincailleries")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="problem-box">
            <h4>❌ Gestion manuelle</h4>
            <p>Cahiers, calculatrices, perte de temps</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="problem-box">
            <h4>❌ Ruptures de stock</h4>
            <p>Client demande un produit → pas en stock → client va chez le concurrent</p>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="problem-box">
            <h4>❌ Surstockage</h4>
            <p>Argent immobilisé dans du stock qui ne se vend pas</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="problem-box">
            <h4>❌ Pas de historique</h4>
            <p>Impossible de savoir qui a acheté quoi, quand</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 💭 Combien vous coûte ces problèmes ?")

    nb_products = st.slider("Nombre d'articles dans votre magasin", 50, 500, 200)
    avg_loss = st.slider("Perte moyenne par rupture de stock (DZD)", 1000, 50000, 10000)
    ruptures_mois = st.slider("Nombre de ruptures de stock par mois", 1, 20, 5)

    perte_annuelle = ruptures_mois * 12 * avg_loss
    st.error(f"**Perte annuelle estimée : {perte_annuelle:,.0f} DZD** — soit {perte_annuelle/12:,.0f} DZD/mois")

    st.markdown("---")
    st.markdown("### 👉 Cliquez sur **✨ Solution** pour voir comment Academix résout ces problèmes")

elif step == "✨ Solution":
    st.markdown("## La solution Academix DSS")

    feature = st.selectbox("Choisissez une fonctionnalité", [
        "📦 Gestion des articles",
        "📊 Tableau de bord",
        "⚠️ Alertes automatiques",
        "📈 Classification ABC",
        "🔄 Mouvements de stock",
        "👥 Gestion fournisseurs"
    ])

    if feature == "📦 Gestion des articles":
        st.markdown("### 📦 Gestion des articles")
        st.markdown("""
        - Ajoutez, modifiez, supprimez des articles en quelques clics
        - Catégories : Visserie, Électricité, Plomberie, Outillage...
        - Prix unitaire, seuil d'alerte, fournisseur
        """)

        st.code("""
        Exemple d'article :
        Nom: Vis Inox 4x40
        Catégorie: Visserie
        Prix: 15 DZD
        Stock: 15,000 unités
        Seuil: 5,000 unités
        Fournisseur: Serral
        """)

    elif feature == "📊 Tableau de Bord":
        st.markdown("### 📊 Tableau de bord en temps réel")

        @st.cache_data
        def show_sample_kpis():
            return {
                "articles": 40,
                "valeur_stock": 53_896_681,
                "alertes": 8,
                "rotation_moy": 8.5
            }
        kpis = show_sample_kpis()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("📦 Articles", kpis["articles"])
        c2.metric("💰 Valeur Stock", f"{kpis['valeur_stock']:,.0f} DZD")
        c3.metric("⚠️ Alertes", kpis["alertes"])
        c4.metric("🔄 Rotation", kpis["rotation_moy"])

        st.success("✅ **Tout est visible en un coup d'œil** — plus besoin de chercher dans des cahiers")

    elif feature == "⚠️ Alertes automatiques":
        st.markdown("### ⚠️ Alertes automatiques")
        st.markdown("""
        Le système vous avertit quand :
        - 🔴 **Stock critique** — En dessous de 50% du seuil
        - 🟡 **Stock bas** — En dessous du seuil d'alerte
        - 🟢 **Stock normal** — Tout va bien
        """)

        st.warning("⚠️ **Exemple** : Vis Inox 4x40 — Stock: 4,500 / Seuil: 5,000 — **Commandez maintenant !**")

    elif feature == "📈 Classification ABC":
        st.markdown("### 📈 Classification ABC")
        st.markdown("""
        - **A** (rouge) — 20% des articles = 80% de la valeur → surveillance maximale
        - **B** (jaune) — 30% des articles = 15% de la valeur → surveillance moyenne
        - **C** (vert) — 50% des articles = 5% de la valeur → surveillance minimale
        """)

        abc_data = pd.DataFrame({
            'Classe': ['A', 'B', 'C'],
            'Articles': [8, 12, 20],
            'Valeur': [43_117_345, 8_084_503, 2_694_833]
        })
        fig = px.bar(abc_data, x='Classe', y='Valeur', color='Classe',
                     color_discrete_map={'A': '#ef4444', 'B': '#f59e0b', 'C': '#10b981'},
                     title="Valeur stock par classe ABC")
        st.plotly_chart(fig, use_container_width=True)

    elif feature == "🔄 Mouvements de stock":
        st.markdown("### 🔄 Mouvements de stock")
        st.markdown("""
        Chaque entrée et sortie est enregistrée avec :
        - 📅 Date
        - 📦 Produit
        - ➡️ Type (Entrée/Sortie)
        - 🔢 Quantité
        - 💰 Montant
        - 👤 Opérateur
        """)

        st.dataframe(pd.DataFrame({
            'Date': ['2026-08-01', '2026-08-01', '2026-07-31', '2026-07-30'],
            'Produit': ['Vis Inox 4x40', 'Boulon Galva 6x80', 'Clou 80mm', 'Disjoncteur 16A'],
            'Type': ['Entrée', 'Sortie', 'Entrée', 'Sortie'],
            'Quantité': [5000, 2000, 3000, 50],
            'Montant (DZD)': [75_000, 50_000, 36_000, 17_500]
        }), use_container_width=True)

    elif feature == "👥 Gestion fournisseurs":
        st.markdown("### 👥 Gestion fournisseurs")
        st.markdown("""
        - 📇 Répertoire complet de vos fournisseurs
        - 📞 Contacts, adresses, délais de livraison
        - ⭐ Notation et historique
        - 🔗 Liés aux articles
        """)

        st.dataframe(pd.DataFrame({
            'Fournisseur': ['Serral', 'Brikol', 'Acier Plus', 'CâbleTech', 'ÉlecPro'],
            'Spécialité': ['Visserie', 'Clouterie', 'Fer', 'Électricité', 'Électricité'],
            'Délai (jours)': [3, 5, 7, 2, 4],
            'Note': ['⭐⭐⭐⭐⭐', '⭐⭐⭐⭐', '⭐⭐⭐⭐', '⭐⭐⭐⭐⭐', '⭐⭐⭐⭐']
        }), use_container_width=True)

    st.markdown("---")
    st.markdown("### 👉 Cliquez sur **📈 Résultats** pour voir les bénéfices")

elif step == "📈 Résultats":
    st.markdown("## Les résultats concrets")

    st.markdown("""
    <div class="benefit-box">
        <h4>⏱️ Gain de temps</h4>
        <p>Remplace 2 heures/jour de calcul manuel par 15 minutes de vérification</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="benefit-box">
        <h4>💰 Réduction des pertes</h4>
        <p>Évitez les ruptures de stock — ne perdez plus de clients</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="benefit-box">
        <h4>📊 Meilleures décisions</h4>
        <p>Savoir exactement quoi commander, quand, et en quelle quantité</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="benefit-box">
        <h4>🔒 Sécurité</h4>
        <p>Données protégées par mot de passe, sauvegarde automatique</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📊 Simulation d'économies")

    ca mensuel = st.slider("Chiffre d'affaires mensuel (DZD)", 500_000, 10_000_000, 2_000_000)
    reduction_pertes = 30
    reduction_temps = 75

    economie_pertes = ca mensuel * 0.05 * reduction_pertes / 100
    economie_temps = 60 * 25 * 30 / 1000  # 60h/mois × 2500 DZD/h × 30% gain
    economie_totale = economie_pertes + economie_temps

    c1, c2, c3 = st.columns(3)
    c1.metric("💰 Économie stock", f"{economie_pertes:,.0f} DZD/mois")
    c2.metric("⏱️ Économie temps", f"{economie_temps:,.0f} DZD/mois")
    c3.metric("🎯 Total", f"{economie_totale:,.0f} DZD/mois", delta="Économie mensuelle")

    st.success(f"**Économie annuelle estimée : {economie_totale * 12:,.0f} DZD**")

    st.markdown("---")
    st.markdown("### 👉 Cliquez sur **📞 Contact** pour demander une démo gratuite")

elif step == "📞 Contact":
    st.markdown("## 📞 Contactez-nous")

    st.markdown("""
    **Prêt à essayer Academix DSS ?**

    📧 **Email** : kamelmahi71@gmail.com
    📱 **WhatsApp** : +213 676 77 38 92
    💻 **GitHub** : [kamelmh/academix-dss-hardware](https://github.com/kamelmh/academix-dss-hardware)
    """)

    st.markdown("---")
    st.markdown("### 📝 Envoyez-nous un message")

    with st.form("contact_form"):
        name = st.text_input("Votre nom")
        email = st.text_input("Votre email")
        phone = st.text_input("Votre téléphone (optionnel)")
        message = st.text_area("Votre message")
        submitted = st.form_submit_button("📤 Envoyer")

        if submitted:
            if name and email and message:
                st.success("✅ Message envoyé ! Nous vous répondrons dans les 24 heures.")
            else:
                st.error("❌ Veuillez remplir au moins votre nom, email et message.")

    st.markdown("---")
    st.markdown("### 🎁 Offre spéciale")
    st.info("🌟 **Démo gratuite** — Testez Academix DSS avec vos propres données pendant 7 jours, sans engagement.")

st.divider()
st.caption("© 2026 Academix DSS — Système de Gestion de Stock pour Quincailleries en Algérie")
