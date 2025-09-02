# 📂 scripts/dashboard.py

import os
import glob
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from io import BytesIO
import streamlit as st
from fpdf import FPDF
import zipfile
import tempfile
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import schedule
import time
import threading
from dotenv import load_dotenv

# --------------------
# 🔐 Login functionaliteit
# --------------------
def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        st.title("🔐 MailMind Login")

        username = st.text_input("Gebruikersnaam")
        password = st.text_input("Wachtwoord", type="password")

        if st.button("Inloggen"):
            valid_username = st.secrets["auth"]["username"]
            valid_password = st.secrets["auth"]["password"]
            if username == valid_username and password == valid_password:
                st.session_state.logged_in = True
                st.success("✅ Ingelogd")
                st.rerun()
            else:
                st.error("❌ Ongeldige inloggegevens")

        st.stop()

check_login()

# --------------------
# ⚙️ Basisinstellingen
# --------------------
BASE_DIR = os.path.dirname(os.path.dirname(__file__))
LOGS_DIR = os.path.join(BASE_DIR, "logs")

# 🚀 Dashboard gebruikt eigen .env
load_dotenv(os.path.join(BASE_DIR, "Streamlit-dashboard", ".env.dashboard"))

st.set_page_config(page_title="📬 MailMind Dashboard", layout="wide", page_icon="📬")

# 🌟 Branding + auto-refresh
col_logo, col_title = st.columns([1, 4])
with col_logo:
    st.image("assets/mailmind_logo.png", width=80)
with col_title:
    st.markdown(
        """
        # 📬 MailMind Dashboard  
        Welkom bij je dagelijkse e-mailoverzicht.  
        Volg prestaties van AI, analyseer fallbacks en exporteer rapportages.
        """
    )

st_autorefresh = st.sidebar.checkbox("🔄 Auto-refresh elke minuut")
if st_autorefresh:
    st.experimental_set_query_params(refresh=datetime.now().strftime("%H:%M:%S"))
    st.experimental_rerun()

# --------------------
# 📑 Tabs
# --------------------
tab_stats, tab_logs, tab_graphs, tab_trends, tab_export, tab_config = st.tabs(
    ["📈 Statistieken", "📄 Logs", "📊 Grafieken", "📆 Trends", "⬇️ Export", "⚙️ Config"]
)

# --------------------
# 📂 Data ophalen
# --------------------
def get_log_files(mode, ref_date, all_logs_toggle):
    if all_logs_toggle:
        return glob.glob(os.path.join(LOGS_DIR, "mail_log_*.xlsx"))
    if mode == "Dag":
        return [os.path.join(LOGS_DIR, f"mail_log_{ref_date.strftime('%Y-%m-%d')}.xlsx")]
    elif mode == "Week":
        start = ref_date - pd.to_timedelta(ref_date.weekday(), unit="D")
        return [
            os.path.join(LOGS_DIR, f"mail_log_{(start + pd.to_timedelta(i, unit='D')).strftime('%Y-%m-%d')}.xlsx")
            for i in range(7)
        ]
    elif mode == "Maand":
        month_str = ref_date.strftime("%Y-%m")
        return glob.glob(os.path.join(LOGS_DIR, f"mail_log_{month_str}-*.xlsx"))

st.sidebar.markdown("### 📂 Filters")
all_logs_toggle = st.sidebar.checkbox("📊 Toon totaaloverzicht van alle logs")

if st.sidebar.button("❌ Reset filters"):
    st.session_state.clear()
    st.rerun()

col_date, col_mode = st.columns([2, 1])
with col_mode:
    period_mode = st.radio("📅 Weergave", ["Dag", "Week", "Maand"], horizontal=True)
with col_date:
    selected_date = st.date_input("Datum", value=datetime.today())

log_files = get_log_files(period_mode, selected_date, all_logs_toggle)

df_list = []
for path in log_files:
    if os.path.exists(path):
        try:
            df_temp = pd.read_excel(path)
            df_temp["Bestand"] = os.path.basename(path)
            df_list.append(df_temp)
        except:
            continue

df = pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()

# --------------------
# 📈 Statistieken tab
# --------------------
with tab_stats:
    if not df.empty:
        df["Categorie"] = df["Categorie"].fillna("Onbekend")
        df["Beantwoord"] = df["Beantwoord"].fillna("Nee")
        if "Reden" not in df.columns:
            df["Reden"] = None
        df["Tijdstip"] = pd.to_datetime(df["Tijdstip"], errors="coerce")
        df["Uur"] = df["Tijdstip"].dt.hour

        total_mails = len(df)
        unique_senders = df["Afzender"].nunique()
        answered_count = df[df["Beantwoord"] == "Ja"].shape[0]
        answered_pct = (answered_count / total_mails * 100) if total_mails > 0 else 0
        complaints_count = df[df["Categorie"] == "Klacht"].shape[0]
        fallback_count = df["Reden"].notna().sum()
        fallback_pct = (fallback_count / total_mails * 100) if total_mails > 0 else 0
        ai_count = answered_count

        col1, col2, col3 = st.columns(3)
        col4, col5, col6 = st.columns(3)
        col1.metric("📬 Totaal e-mails", total_mails)
        col2.metric("👤 Unieke afzenders", unique_senders)
        col3.metric("✅ Beantwoord door AI", f"{ai_count} ({answered_pct:.0f}%)")
        col4.metric("📤 Fallbacks", f"{fallback_count} ({fallback_pct:.0f}%)")
        col5.metric("🚨 Klachten", complaints_count)
        col6.metric("🤖 AI vs Fallback", f"{ai_count}/{fallback_count}")

        if complaints_count > total_mails * 0.1:
            st.error("⚠️ Hoog percentage klachten!")
        if fallback_pct > 50:
            st.warning("📤 Meer dan 50% mails ging via fallback.")

    else:
        st.warning("❌ Geen gegevens beschikbaar")

# --------------------
# 📄 Logs tab
# --------------------
with tab_logs:
    if not df.empty:
        selected_cats = st.multiselect("Categorieën", sorted(df["Categorie"].unique()))
        selected_senders = st.multiselect("Afzenders", sorted(df["Afzender"].unique()))
        selected_reasons = st.multiselect("Fallback-redenen", sorted(df["Reden"].dropna().unique()))

        filtered_df = df.copy()
        if selected_cats:
            filtered_df = filtered_df[filtered_df["Categorie"].isin(selected_cats)]
        if selected_senders:
            filtered_df = filtered_df[filtered_df["Afzender"].isin(selected_senders)]
        if selected_reasons:
            filtered_df = filtered_df[filtered_df["Reden"].isin(selected_reasons)]

        st.dataframe(filtered_df)

        if not filtered_df["Reden"].isna().all():
            st.markdown("### 📊 Fallbacks per reden")
            st.bar_chart(filtered_df["Reden"].value_counts())
    else:
        st.info("Geen logs gevonden voor deze periode.")

# --------------------
# 📊 Grafieken tab
# --------------------
with tab_graphs:
    if not df.empty:
        st.markdown("### 📊 E-mails per categorie")
        st.bar_chart(df["Categorie"].value_counts())

        st.markdown("### 🥧 Verdeling per categorie")
        st.pyplot(df["Categorie"].value_counts().plot(kind="pie", autopct="%1.1f%%").get_figure())

        st.markdown("### 🤖 AI vs Fallback")
        counts = pd.Series({"AI": ai_count, "Fallback": fallback_count})
        st.bar_chart(counts)

        st.markdown("### ✅ Beantwoord-status")
        st.bar_chart(df["Beantwoord"].value_counts())

        st.markdown("### 🕒 E-mails per uur")
        st.bar_chart(df["Uur"].value_counts().sort_index())
    else:
        st.info("Geen grafieken beschikbaar.")

# --------------------
# 📆 Trends tab
# --------------------
with tab_trends:
    if not df.empty:
        df["Datum"] = df["Tijdstip"].dt.date
        st.markdown("### 📈 Dagelijks aantal e-mails")
        st.line_chart(df.groupby("Datum").size())

        complaints_daily = df[df["Categorie"] == "Klacht"].groupby("Datum").size()
        if not complaints_daily.empty:
            st.markdown("### 🚨 Klachten per dag")
            st.line_chart(complaints_daily)
    else:
        st.info("Geen trendgegevens beschikbaar.")

# --------------------
# ⬇️ Export tab
# --------------------
with tab_export:
    if not df.empty:
        excel_buffer = BytesIO()
        df.to_excel(excel_buffer, index=False)
        st.download_button("⬇️ Download Excel", excel_buffer.getvalue(), "emails.xlsx")

        if st.button("⬇️ Genereer PDF-rapport"):
            pdf = FPDF()
            pdf.add_page()
            logo_path = os.path.join(BASE_DIR, "Streamlit-dashboard", "assets", "mailmind_logo.png")
            if os.path.exists(logo_path):
                pdf.image(logo_path, x=10, y=8, w=25)
            pdf.set_font("Arial", "B", 16)
            pdf.cell(200, 10, "MailMind Rapport", ln=True, align="C")

            pdf.set_font("Arial", "", 12)
            pdf.cell(200, 10, f"Datum: {datetime.now().strftime('%Y-%m-%d')}", ln=True)
            pdf.ln(10)
            pdf.cell(200, 10, f"Totaal e-mails: {total_mails}", ln=True)
            pdf.cell(200, 10, f"AI beantwoord: {ai_count} ({answered_pct:.0f}%)", ln=True)
            pdf.cell(200, 10, f"Fallbacks: {fallback_count} ({fallback_pct:.0f}%)", ln=True)
            pdf.cell(200, 10, f"Klachten: {complaints_count}", ln=True)

            pdf_buffer = BytesIO()
            pdf.output(pdf_buffer)
            st.download_button("⬇️ Download PDF", pdf_buffer.getvalue(), "rapport.pdf")

        if st.button("⬇️ Download grafieken (PNG)"):
            tmpdir = tempfile.mkdtemp()
            figs = {
                "categorie.png": df["Categorie"].value_counts().plot(kind="bar").get_figure(),
                "ai_vs_fallback.png": pd.Series({"AI": ai_count, "Fallback": fallback_count}).plot(kind="bar").get_figure(),
            }
            zip_path = os.path.join(tmpdir, "grafieken.zip")
            with zipfile.ZipFile(zip_path, "w") as zf:
                for name, fig in figs.items():
                    buf = BytesIO()
                    fig.savefig(buf, format="png")
                    zf.writestr(name, buf.getvalue())
            with open(zip_path, "rb") as f:
                st.download_button("⬇️ Download ZIP", f.read(), "grafieken.zip")
    else:
        st.info("Geen data om te exporteren.")

# --------------------
# 📧 Automatische mailfunctie
# --------------------
def send_report_via_email():
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, "MailMind Rapport", ln=True, align="C")

    pdf.set_font("Arial", "", 12)
    pdf.cell(200, 10, f"Datum: {datetime.now().strftime('%Y-%m-%d')}", ln=True)
    pdf.ln(10)
    pdf.cell(200, 10, f"Totaal e-mails: {total_mails}", ln=True)
    pdf.cell(200, 10, f"AI beantwoord: {ai_count} ({answered_pct:.0f}%)", ln=True)
    pdf.cell(200, 10, f"Fallbacks: {fallback_count} ({fallback_pct:.0f}%)", ln=True)
    pdf.cell(200, 10, f"Klachten: {complaints_count}", ln=True)

    pdf_buffer = BytesIO()
    pdf.output(pdf_buffer)

    msg = MIMEMultipart()
    msg["From"] = os.getenv("SMTP_USER")
    msg["To"] = os.getenv("REPORT_EMAIL")
    msg["Subject"] = f"MailMind Rapport - {datetime.now().strftime('%Y-%m-%d')}"

    body = MIMEText("Beste manager,\n\nIn de bijlage vindt u het dagelijkse MailMind rapport.\n\nGroeten,\nMailMind", "plain")
    msg.attach(body)

    attachment = MIMEApplication(pdf_buffer.getvalue(), _subtype="pdf")
    attachment.add_header("Content-Disposition", "attachment", filename="rapport.pdf")
    msg.attach(attachment)

    try:
        with smtplib.SMTP(os.getenv("SMTP_SERVER"), int(os.getenv("SMTP_PORT"))) as server:
            server.starttls()
            server.login(os.getenv("SMTP_USER"), os.getenv("SMTP_PASS"))
            server.sendmail(msg["From"], msg["To"], msg.as_string())
        print("✅ Rapport gemaild naar manager")
    except Exception as e:
        print("❌ Fout bij mailen rapport:", e)

schedule.every().day.at("08:00").do(send_report_via_email)

def run_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(60)

threading.Thread(target=run_scheduler, daemon=True).start()

# --------------------
# ⚙️ Config tab
# --------------------
with tab_config:
    st.subheader("⚙️ Configuratie")
    st.markdown("Hier kun je later instellingen beheren (bijvoorbeeld logo uploaden, kleuren aanpassen, alerts instellen).")
