
import os
import glob
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from io import BytesIO
import streamlit as st

# 📁 Pad naar logs
BASE_DIR = os.path.dirname(os.path.dirname(__file__))
LOGS_DIR = os.path.join(BASE_DIR, "logs")

st.set_page_config(page_title="📬 MailMind Dashboard", layout="wide")

# 🌟 Header
st.markdown("""# 📬 MailMind Dashboard  
Welkom bij je dagelijkse e-mailoverzicht. Bekijk trends, statistieken en download eenvoudig rapportages.""")

if st.button("🔁 Herlaad dashboard"):
    st.rerun()

# 📅 Selectie
st.markdown("---")
st.markdown("### 📆 Selecteer een datum of periode")
col_date, col_mode = st.columns([2, 1])
with col_mode:
    period_mode = st.radio("📅 Weergave", ["Dag", "Week", "Maand"], horizontal=True)
with col_date:
    selected_date = st.date_input("Datum", value=datetime.today())

# 📂 Sidebar-opties
st.sidebar.markdown("---")
all_logs_toggle = st.sidebar.checkbox("📊 Toon totaaloverzicht van alle logs")

if st.sidebar.button("❌ Reset filters"):
    st.session_state["selected_cats"] = []
    st.session_state["selected_senders"] = []
    st.rerun()

# 🔍 Bestanden ophalen
def get_log_files(mode, ref_date):
    if all_logs_toggle:
        return glob.glob(os.path.join(LOGS_DIR, "mail_log_*.xlsx"))
    if mode == "Dag":
        return [os.path.join(LOGS_DIR, f"mail_log_{ref_date.strftime('%Y-%m-%d')}.xlsx")]
    elif mode == "Week":
        start = ref_date - pd.to_timedelta(ref_date.weekday(), unit="D")
        return [os.path.join(LOGS_DIR, f"mail_log_{(start + pd.to_timedelta(i, unit='D')).strftime('%Y-%m-%d')}.xlsx") for i in range(7)]
    elif mode == "Maand":
        month_str = ref_date.strftime('%Y-%m')
        return glob.glob(os.path.join(LOGS_DIR, f"mail_log_{month_str}-*.xlsx"))

log_files = get_log_files(period_mode, selected_date)

# 📥 Data inladen
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

if not df.empty:
    df["Categorie"] = df["Categorie"].fillna("Onbekend")
    df["Beantwoord"] = df["Beantwoord"].fillna("Nee")
    df["Tijdstip"] = pd.to_datetime(df["Tijdstip"], errors="coerce")
    df["Uur"] = df["Tijdstip"].dt.hour
    if "Antwoord" in df.columns:
        df.loc[df["Antwoord"].notna() & (df["Antwoord"].astype(str).str.strip() != ""), "Beantwoord"] = "Ja"

    with st.sidebar:
        st.header("📂 Filters")
        selected_cats = st.multiselect("Categorieën", sorted(df["Categorie"].unique()),
                                       default=st.session_state.get("selected_cats", []),
                                       key="selected_cats", placeholder="Kies een optie")
        selected_senders = st.multiselect("Afzenders", sorted(df["Afzender"].unique()),
                                          default=st.session_state.get("selected_senders", []),
                                          key="selected_senders", placeholder="Kies een optie")

    filtered_df = df.copy()
    if selected_cats:
        filtered_df = filtered_df[filtered_df["Categorie"].isin(selected_cats)]
    if selected_senders:
        filtered_df = filtered_df[filtered_df["Afzender"].isin(selected_senders)]

    # 📈 Statistieken
    st.markdown("---")
    st.markdown("## 📈 Statistieken")
    total_mails = len(filtered_df)
    unique_senders = filtered_df["Afzender"].nunique()
    answered_count = filtered_df[filtered_df["Beantwoord"] == "Ja"].shape[0]
    answered_pct = (answered_count / total_mails * 100) if total_mails > 0 else 0
    complaints_count = filtered_df[filtered_df["Categorie"] == "Klacht"].shape[0]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📬 Totaal e-mails", total_mails)
    col2.metric("👤 Unieke afzenders", unique_senders)
    col3.metric("✅ Beantwoord", f"{answered_count} ({answered_pct:.0f}%)")
    col4.metric("🚨 Klachten", complaints_count)

    # ➕ Gemiddelden
    if not all_logs_toggle:
        all_files = glob.glob(os.path.join(LOGS_DIR, "mail_log_*.xlsx"))
        all_data = []
        for f in all_files:
            try:
                d = pd.read_excel(f)
                d["Beantwoord"] = d["Beantwoord"].fillna("Nee")
                d["Categorie"] = d["Categorie"].fillna("Onbekend")
                all_data.append(d)
            except:
                continue
        if all_data:
            df_all = pd.concat(all_data, ignore_index=True)
            total_all = len(df_all)
            avg_answered = df_all[df_all["Beantwoord"] == "Ja"].shape[0] / total_all * 100
            avg_complaints = df_all[df_all["Categorie"] == "Klacht"].shape[0] / total_all * 100
            st.markdown(f"<div style='margin-top:-10px; font-size:0.9em; color:gray;'>Gemiddeld (alle logs): {avg_answered:.0f}% beantwoord • {avg_complaints:.0f}% klachten</div>", unsafe_allow_html=True)

    # 📄 Tabel en download
    st.markdown("---")
    st.markdown("## 📄 Geselecteerde log")

    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        filtered_df.to_excel(writer, index=False, sheet_name="Log")
        workbook = writer.book
        worksheet = writer.sheets["Log"]
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        for col_num, value in enumerate(filtered_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            max_len = max(filtered_df[value].astype(str).map(len).max(), len(value))
            worksheet.set_column(col_num, col_num, max_len + 2)
        green_fill = workbook.add_format({'bg_color': '#C6EFCE'})
        red_fill = workbook.add_format({'bg_color': '#FFC7CE'})
        antwoord_idx = filtered_df.columns.get_loc("Beantwoord")
        for row_num, val in enumerate(filtered_df["Beantwoord"], start=1):
            kleur = green_fill if val == "Ja" else red_fill
            worksheet.write(row_num, antwoord_idx, val, kleur)

    def highlight_row(row):
        if row["Categorie"] == "Klacht":
            return ["background-color: #fff3cd"] * len(row)
        elif row["Beantwoord"] == "Nee":
            return ["background-color: #f8d7da"] * len(row)
        return [""] * len(row)

    st.dataframe(filtered_df.style.apply(highlight_row, axis=1))
    st.download_button("⬇️ Download Excel-bestand", data=excel_buffer.getvalue(),
                       file_name="filtered_emails_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # 📊 Grafieken
    def render_and_download(fig, title, filename):
        st.pyplot(fig)
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight")
        st.download_button(f"⬇️ Download PNG – {title}", data=buf.getvalue(), file_name=filename, mime="image/png")

    st.markdown("---")
    st.markdown("## 📊 Verdeling per categorie")
    cat_counts = filtered_df["Categorie"].value_counts()
    fig1, ax1 = plt.subplots(figsize=(10, 6))
    cat_counts.plot(kind="bar", ax=ax1, color="#4e79a7")
    ax1.set_title("E-mails per categorie")
    render_and_download(fig1, "Categorieën", "categorie_balk.png")

    st.markdown("## 🥧 Categorieverhouding")
    fig2, ax2 = plt.subplots()
    cat_counts.plot(kind="pie", autopct="%1.1f%%", startangle=90, ax=ax2)
    ax2.set_ylabel("")
    ax2.set_title("Verhouding per categorie")
    render_and_download(fig2, "Taartdiagram", "categorie_taart.png")

    st.markdown("## ✅ Beantwoord-status")
    answered_counts = filtered_df["Beantwoord"].value_counts()
    fig3, ax3 = plt.subplots()
    answered_counts.plot(kind="bar", ax=ax3, color=["#e15759", "#59a14f"])
    ax3.set_title("Beantwoord-status")
    render_and_download(fig3, "Beantwoord", "beantwoord_status.png")

    st.markdown("## 🕒 Tijdlijn: E-mails per uur")
    if not filtered_df["Uur"].isna().all():
        hourly = filtered_df["Uur"].value_counts().sort_index()
        fig4, ax4 = plt.subplots()
        hourly.plot(kind="bar", ax=ax4, color="#f28e2b")
        ax4.set_title("E-mails per uur")
        render_and_download(fig4, "Tijdlijn", "tijdlijn_per_uur.png")

else:
    st.warning("❌ Geen gegevens beschikbaar voor de geselecteerde periode.")