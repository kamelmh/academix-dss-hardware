import streamlit as st
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(page_title="Demander une Démo", page_icon="📞", layout="wide")

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

st.markdown('<h1 class="main-header">📞 Demander une Démo</h1>', unsafe_allow_html=True)
st.caption("Remplissez le formulaire et recevez un accès démo gratuit")

col_form, col_info = st.columns([3, 2])

with col_form:
    with st.form("contact_form"):
        st.subheader("📋 Vos Informations")

        name = st.text_input("Nom complet *", placeholder="Ahmed Benali")
        phone = st.text_input("Téléphone *", placeholder="0555 12 34 56")
        email = st.text_input("Email", placeholder="ahmed@quincaillerie.dz")
        store_name = st.text_input("Nom de la quincaillerie *", placeholder="Quincaillerie Benali")
        store_size = st.selectbox("Taille de la quincaillerie", [
            "Petite (< 100 articles)",
            "Moyenne (100-500 articles)",
            "Grande (> 500 articles)"
        ])
        source = st.selectbox("Comment nous avez-vous trouvés ?", [
            "Facebook", "Bouche à oreille", "Google", "Autre"
        ])
        message = st.text_area("Message (optionnel)", placeholder="Décrivez vos besoins...")

        submitted = st.form_submit_button("🚀 Envoyer la Demande", use_container_width=True)

        if submitted:
            if not name or not phone or not store_name:
                st.error("Veuillez remplir les champs obligatoires (Nom, Téléphone, Quincaillerie)")
            else:
                try:
                    smtp_user = "kamelmahi71@gmail.com"
                    smtp_pass = st.secrets.get("gmail_app_password", "")

                    if not smtp_pass:
                        st.warning("📧 Mode démo — configurez `gmail_app_password` dans Streamlit Secrets pour l'envoi automatique.")
                        st.success(f"""
                        ✅ **Demande enregistrée !**

                        | Champ | Valeur |
                        |-------|--------|
                        | **Nom** | {name} |
                        | **Téléphone** | {phone} |
                        | **Email** | {email} |
                        | **Quincaillerie** | {store_name} |
                        | **Taille** | {store_size} |
                        | **Source** | {source} |
                        | **Message** | {message} |

                        📞 Nous vous contacterons dans les **24h** via WhatsApp au **{phone}**
                        """)
                    else:
                        msg = MIMEMultipart('alternative')
                        msg['Subject'] = f"🆕 Demande de démo — {store_name}"
                        msg['From'] = smtp_user
                        msg['To'] = smtp_user

                        html = f"""
                        <h2>🆕 Nouvelle demande de démo Academix DSS</h2>
                        <table style="border-collapse:collapse;width:100%">
                            <tr><td style="padding:8px;border:1px solid #ddd"><b>Nom</b></td><td style="padding:8px;border:1px solid #ddd">{name}</td></tr>
                            <tr><td style="padding:8px;border:1px solid #ddd"><b>Téléphone</b></td><td style="padding:8px;border:1px solid #ddd">{phone}</td></tr>
                            <tr><td style="padding:8px;border:1px solid #ddd"><b>Email</b></td><td style="padding:8px;border:1px solid #ddd">{email}</td></tr>
                            <tr><td style="padding:8px;border:1px solid #ddd"><b>Quincaillerie</b></td><td style="padding:8px;border:1px solid #ddd">{store_name}</td></tr>
                            <tr><td style="padding:8px;border:1px solid #ddd"><b>Taille</b></td><td style="padding:8px;border:1px solid #ddd">{store_size}</td></tr>
                            <tr><td style="padding:8px;border:1px solid #ddd"><b>Source</b></td><td style="padding:8px;border:1px solid #ddd">{source}</td></tr>
                            <tr><td style="padding:8px;border:1px solid #ddd"><b>Message</b></td><td style="padding:8px;border:1px solid #ddd">{message}</td></tr>
                        </table>
                        <p>📞 <b>Action:</b> Contactez {name} via WhatsApp au {phone}</p>
                        """
                        msg.attach(MIMEText(html, 'html'))

                        context = ssl.create_default_context()
                        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
                            server.login(smtp_user, smtp_pass)
                            server.sendmail(smtp_user, smtp_user, msg.as_string())

                        st.success(f"""
                        ✅ **Demande envoyée avec succès !**

                        Nous vous contacterons dans les **24h** via WhatsApp au **{phone}**.
                        """)
                        st.balloons()

                except Exception as e:
                    st.error(f"Erreur: {e}")
                    st.info("Contactez-nous directement: kamelmahi71@gmail.com | WhatsApp: +213 676 77 38 92")

with col_info:
    st.subheader("🎯 Pourquoi Academix DSS ?")

    features = [
        ("📦", "Gestion de Stock", "Suivi en temps réel de tous vos articles"),
        ("💰", "Calcul CMUP", "Coût moyen unitaire pondéré automatique"),
        ("📊", "Classification ABC", "Identifiez vos articles les plus importants"),
        ("⚠️", "Alertes Stock", "Soyez averti avant les ruptures"),
        ("🧾", "Facturation", "Générez vos factures en un clic"),
        ("📈", "Tableaux de bord", "Vue d'ensemble de votre activité"),
    ]

    for icon, title, desc in features:
        st.markdown(f"**{icon} {title}** — {desc}")

    st.divider()

    st.subheader("💳 Tarification")
    st.markdown("""
    | Offre | Prix | Articles |
    |-------|------|----------|
    | 🥉 Amateur | **15 000 DZD** | Jusqu'à 100 |
    | 🥈 Professionnel | **25 000 DZD** | Jusqu'à 500 |
    | 🥇 Entreprise | **45 000 DZD** | Illimité |
    """)

    st.info("🎁 **Offre de lancement** — Première démo gratuite, 1 mois d'essai")

    st.divider()

    st.subheader("📱 Contact direct")
    st.markdown("""
    - 📧 **Email:** kamelmahi71@gmail.com
    - 📞 **WhatsApp:** [+213 676 77 38 92](https://wa.me/213676773892)
    """)

    st.divider()

    st.subheader("📊 Voir la démo")
    st.page_link("pages/1_Tableau_de_Bord.py", label="📊 Tableau de bord interactif")

st.divider()
st.caption("© 2026 Academix DSS — Système de Gestion de Stock pour Quincailleries en Algérie")
