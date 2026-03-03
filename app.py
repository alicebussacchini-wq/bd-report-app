import streamlit as st
import anthropic
import PyPDF2
import io
import os
import requests
import json
import base64
import time
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

st.set_page_config(page_title="Taxi Report", page_icon="📊", layout="wide")

ARCHIVIO_PATH = r"C:\Users\1103540\bd-report-app\archivio"
os.makedirs(ARCHIVIO_PATH, exist_ok=True)

st.markdown("""
<style>
    .stApp { background-color: #1a1a1a; color: #f0f0f0; }
    .hl-header { display: flex; align-items: center; justify-content: space-between; padding: 20px 0; border-bottom: 3px solid #c8e04a; margin-bottom: 30px; }
    .hl-logo { height: 70px; }
    .hl-title { font-size: 22px; color: #c8e04a; font-weight: 600; text-align: right; }
    .hl-subtitle { font-size: 13px; color: #999; text-align: right; }
    .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin: 16px 0; }
    .kpi-card { background: #2a2a2a; border: 1px solid #c8e04a; border-radius: 10px; padding: 16px 12px; text-align: center; }
    .kpi-label { font-size: 11px; color: #999; text-transform: uppercase; letter-spacing: 1px; }
    .kpi-value { font-size: 20px; font-weight: 700; color: #c8e04a; margin-top: 6px; }
    .section-box { background: #242424; border-left: 4px solid #c8e04a; border-radius: 8px; padding: 20px; margin: 12px 0; }
    .section-title { font-size: 16px; font-weight: 700; color: #c8e04a; margin-bottom: 12px; }
    .section-text { font-size: 14px; color: #ddd; line-height: 1.7; }
    .ma-item { background: #2a2a2a; border-radius: 8px; padding: 14px; margin: 8px 0; border: 1px solid #444; }
    .ma-anno { background: #c8e04a; color: #1a1a1a; font-weight: 700; padding: 2px 10px; border-radius: 20px; font-size: 12px; display: inline-block; }
    .ma-tipo { color: #999; font-size: 12px; margin: 6px 0; text-transform: uppercase; }
    .ma-desc { color: #ddd; font-size: 14px; }
    .stButton > button { background: #c8e04a; color: #1a1a1a; font-weight: 700; border: none; padding: 12px 32px; border-radius: 8px; font-size: 16px; width: 100%; cursor: pointer; }
    .stButton > button:hover { background: #b5cc3a; }
    h1, h2, h3 { color: #f0f0f0; }
    .stExpander { background: #242424; border: 1px solid #444; border-radius: 8px; }
    .archivio-card { background: #242424; border: 1px solid #444; border-radius: 10px; padding: 18px; margin: 10px 0; }
    .archivio-nome { font-size: 18px; font-weight: 700; color: #c8e04a; }
    .archivio-data { font-size: 12px; color: #999; margin: 4px 0 12px 0; }
    .archivio-kpi { display: flex; gap: 12px; flex-wrap: wrap; }
    .archivio-kpi-item { background: #1a1a1a; border-radius: 6px; padding: 6px 12px; font-size: 12px; }
    .archivio-kpi-label { color: #999; }
    .archivio-kpi-val { color: #c8e04a; font-weight: 700; margin-left: 6px; }
</style>
""", unsafe_allow_html=True)

def get_logo_base64():
    # Funziona sia in locale che su Streamlit Cloud
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

logo_b64 = get_logo_base64()
logo_html = f'<img src="data:image/jpeg;base64,{logo_b64}" class="hl-logo">' if logo_b64 else "<span style='color:#c8e04a;font-size:24px;font-weight:700;'>Hogan Lovells</span>"
data_oggi = datetime.now().strftime("%d %B %Y")

st.markdown(f"""
<div class="hl-header">
    {logo_html}
    <div>
        <div class="hl-title">Taxi Report Generator</div>
        <div class="hl-subtitle">Generato il {data_oggi}</div>
    </div>
</div>
""", unsafe_allow_html=True)

def estrai_testo_url(url):
    try:
        return requests.get(url, timeout=10).text[:5000]
    except:
        return "Impossibile recuperare il contenuto dell'URL."

def get_sheet():
    credenziali = {
        "type": st.secrets["gcp_service_account"]["type"],
        "project_id": st.secrets["gcp_service_account"]["project_id"],
        "private_key_id": st.secrets["gcp_service_account"]["private_key_id"],
        "private_key": st.secrets["gcp_service_account"]["private_key"],
        "client_email": st.secrets["gcp_service_account"]["client_email"],
        "client_id": st.secrets["gcp_service_account"]["client_id"],
        "auth_uri": st.secrets["gcp_service_account"]["auth_uri"],
        "token_uri": st.secrets["gcp_service_account"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["gcp_service_account"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["gcp_service_account"]["client_x509_cert_url"]
    }
    scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(credenziali, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open("Taxi Report Archivio").sheet1
    return sheet

def salva_report(nome_azienda, report_json, documenti_files):
    try:
        sheet = get_sheet()
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
        sheet.append_row([
            timestamp,
            nome_azienda,
            json.dumps(report_json, ensure_ascii=False)
        ])
    except Exception as e:
        st.warning(f"Errore salvataggio archivio: {e}")

def carica_archivio():
    try:
        sheet = get_sheet()
        righe = sheet.get_all_values()
        reports = []
        for riga in reversed(righe):
            if len(riga) >= 3 and riga[2]:
                try:
                    report = json.loads(riga[2])
                    reports.append({
                        "data": riga[0],
                        "nome": riga[1],
                        "report": report,
                        "riga": righe.index(riga) + 1
                    })
                except:
                    pass
        return reports
    except Exception as e:
        st.warning(f"Errore caricamento archivio: {e}")
        return []

if "pagina" not in st.session_state:
    st.session_state["pagina"] = "genera"

col_nav1, col_nav2 = st.columns(2)
with col_nav1:
    if st.button("➕ Genera nuovo report"):
        st.session_state["pagina"] = "genera"
with col_nav2:
    if st.button("🗂️ Archivio report"):
        st.session_state["pagina"] = "archivio"

st.markdown("---")

if st.session_state["pagina"] == "genera":

    lingua = st.radio("🌐 Lingua del report / Report language", ["Italiano", "English"], horizontal=True)
    st.markdown("### Carica i documenti")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.caption("📄 Bilancio consolidato (PDF)")
        bilancio = st.file_uploader("Bilancio PDF", type=["pdf"], key="bilancio", label_visibility="collapsed")

    with col2:
        st.caption("📈 Export Mergermarket")
        mergermarket = st.file_uploader("Mergermarket", type=["pdf", "csv"], key="merger", label_visibility="collapsed")

    with col3:
        st.caption("🌐 URL Sito / Press Release")
        url_azienda = st.text_input("URL", placeholder="https://...", label_visibility="collapsed")

    with col4:
        st.caption("🏛️ Visura Camerale")
        visura = st.file_uploader("Visura PDF", type=["pdf"], key="visura", label_visibility="collapsed")

    testi_documenti = {}
    documenti_binari = {}

    if bilancio:
        contenuto = bilancio.read()
        pdf_b64 = base64.b64encode(contenuto).decode()
        client_vision = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
        st.info("Analisi bilancio in corso...")
        time.sleep(30)
        try:
            risposta = client_vision.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=4000,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "document",
                            "source": {
                                "type": "base64",
                                "media_type": "application/pdf",
                                "data": pdf_b64
                            }
                        },
                        {
                            "type": "text",
                            "text": "Estrai tutti i dati finanziari da questo bilancio: ricavi totali, EBITDA, utile netto, totale attivo, patrimonio netto, debiti finanziari. Mantieni i valori numerici esatti con le unità di misura."
                        }
                    ]
                }]
            )
            testo = risposta.content[0].text
        except Exception as e:
            testo = ""
            st.warning(f"Errore analisi bilancio: {e}")
        testi_documenti["Bilancio Consolidato"] = testo
        documenti_binari[bilancio.name] = contenuto
        st.success("✅ Bilancio analizzato")

    if mergermarket:
        contenuto = mergermarket.read()
        if mergermarket.name.endswith(".pdf"):
            reader = PyPDF2.PdfReader(io.BytesIO(contenuto))
            testo = ""
            for i, p in enumerate(reader.pages):
                testo += f"\n--- PAGINA {i+1} ---\n{p.extract_text() or ''}"
            testi_documenti["Mergermarket"] = testo
        else:
            testi_documenti["Mergermarket"] = contenuto.decode("utf-8")
        documenti_binari[mergermarket.name] = contenuto
        st.success("✅ Mergermarket caricato")

    if url_azienda:
        testi_documenti["Sito Aziendale"] = estrai_testo_url(url_azienda)
        st.success("✅ URL acquisito")

    if visura:
        contenuto = visura.read()
        pdf_b64_visura = base64.b64encode(contenuto).decode()
        client_vision2 = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
        try:
            risposta_visura = client_vision2.messages.create(
                model="claude-opus-4-5",
                max_tokens=4000,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "document",
                            "source": {
                                "type": "base64",
                                "media_type": "application/pdf",
                                "data": pdf_b64_visura
                            }
                        },
                        {
                            "type": "text",
                            "text": "Estrai tutte le informazioni rilevanti da questa visura camerale: ragione sociale, sede, codice fiscale, soci, amministratori, capitale sociale, oggetto sociale."
                        }
                    ]
                }]
            )
            testo_visura = risposta_visura.content[0].text
        except Exception as e:
            reader = PyPDF2.PdfReader(io.BytesIO(contenuto))
            testo_visura = ""
            for i, p in enumerate(reader.pages):
                testo_visura += f"\n--- PAGINA {i+1} ---\n{p.extract_text() or ''}"
        testi_documenti["Visura Camerale"] = testo_visura
        documenti_binari[visura.name] = contenuto
        st.success("✅ Visura caricata")

    st.markdown("---")
    nome_azienda = st.text_input("Nome dell'azienda", placeholder="es. Eco Eridania S.p.A.")

    if st.button("🚀 Genera Report", disabled=len(testi_documenti) == 0):
        if not nome_azienda:
            st.warning("Inserisci il nome dell'azienda.")
        else:
            with st.spinner("Claude sta analizzando i documenti..."):
                testo_completo = ""
                for fonte, testo in testi_documenti.items():
                    testo_completo += f"\n\n--- FONTE: {fonte} ---\n{testo[:10000]}"
                 
                lingua_prompt = "in inglese" if lingua == "English" else "in italiano"
                prompt = f"""Sei un analista M&A e finance di uno studio legale internazionale.
Analizza i seguenti documenti relativi all'azienda {nome_azienda} e produci un report strutturato in JSON, con tutti i testi {lingua_prompt}.

DOCUMENTI:
{testo_completo}

Rispondi SOLO con un oggetto JSON valido, senza backtick, senza testo aggiuntivo:

{{
  "nome_azienda": "",
  "overview": "",
  "core_business": "",
  "mercati": "",
  "dati_finanziari": {{
    "ricavi": "",
    "ebitda": "",
    "utile_netto": "",
    "totale_attivo": "",
    "patrimonio_netto": "",
    "anno_riferimento": ""
  }},
  "struttura_debito": {{
    "indebitamento_totale": "",
    "debito_bancario": "",
    "obbligazioni": "",
    "debito_netto": "",
    "leva_finanziaria": "",
    "scadenze_principali": "",
    "note": ""
  }},
  "ownership": {{
    "azionista_principale": "",
    "quota_principale": "",
    "altri_azionisti": "",
    "struttura_controllo": "",
    "note": ""
  }},
  "operazioni_ma": [
    {{"anno": "", "tipo": "", "descrizione": ""}}
  ],
  "note_aggiuntive": ""
}}

Se un dato non è disponibile scrivi N/D. Non inventare dati."""

                client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
                messaggio = client.messages.create(
                    model="claude-opus-4-5",
                    max_tokens=4000,
                    messages=[{"role": "user", "content": prompt}]
                )
                risposta = messaggio.content[0].text

                try:
                    risposta_pulita = risposta.strip()
                    if risposta_pulita.startswith("```"):
                        risposta_pulita = risposta_pulita.split("```")[1]
                        if risposta_pulita.startswith("json"):
                            risposta_pulita = risposta_pulita[4:]
                    risposta_pulita = risposta_pulita.strip()
                    report = json.loads(risposta_pulita)
                    salva_report(nome_azienda, report, documenti_binari)
                    st.session_state["report"] = report
                    st.success("✅ Report generato e salvato!")
                    st.rerun()
                except Exception as e:
                    st.error("Errore nel parsing. Risposta grezza:")
                    st.text(risposta)

    if "report" in st.session_state:
        report = st.session_state["report"]
        nome = report.get("nome_azienda", "")
        fin = report.get("dati_finanziari", {})

        st.markdown(f"## 📋 {nome}")
        st.markdown(f"""
        <div class="kpi-grid">
            <div class="kpi-card"><div class="kpi-label">Ricavi</div><div class="kpi-value">{fin.get('ricavi','N/D')}</div></div>
            <div class="kpi-card"><div class="kpi-label">EBITDA</div><div class="kpi-value">{fin.get('ebitda','N/D')}</div></div>
            <div class="kpi-card"><div class="kpi-label">Utile Netto</div><div class="kpi-value">{fin.get('utile_netto','N/D')}</div></div>
            <div class="kpi-card"><div class="kpi-label">Totale Attivo</div><div class="kpi-value">{fin.get('totale_attivo','N/D')}</div></div>
            <div class="kpi-card"><div class="kpi-label">Patrimonio Netto</div><div class="kpi-value">{fin.get('patrimonio_netto','N/D')}</div></div>
            <div class="kpi-card"><div class="kpi-label">Anno Rif.</div><div class="kpi-value">{fin.get('anno_riferimento','N/D')}</div></div>
        </div>
        """, unsafe_allow_html=True)

        with st.expander("🏢 Overview", expanded=True):
            st.markdown(f'<div class="section-box"><div class="section-title">Overview</div><div class="section-text">{report.get("overview","N/D")}</div></div>', unsafe_allow_html=True)

        with st.expander("⚙️ Core Business & Mercati"):
            st.markdown(f'<div class="section-box"><div class="section-title">Core Business</div><div class="section-text">{report.get("core_business","N/D")}</div></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="section-box"><div class="section-title">Mercati</div><div class="section-text">{report.get("mercati","N/D")}</div></div>', unsafe_allow_html=True)

        with st.expander("💰 Struttura del Debito"):
            debito = report.get("struttura_debito", {})
            if isinstance(debito, dict):
                st.markdown(f"""
                <div class="kpi-grid">
                    <div class="kpi-card"><div class="kpi-label">Indebitamento Totale</div><div class="kpi-value">{debito.get('indebitamento_totale','N/D')}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Debito Bancario</div><div class="kpi-value">{debito.get('debito_bancario','N/D')}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Obbligazioni</div><div class="kpi-value">{debito.get('obbligazioni','N/D')}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Debito Netto</div><div class="kpi-value">{debito.get('debito_netto','N/D')}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Leva Finanziaria</div><div class="kpi-value">{debito.get('leva_finanziaria','N/D')}</div></div>
                </div>
                """, unsafe_allow_html=True)
                if debito.get('scadenze_principali', 'N/D') != 'N/D':
                    st.markdown(f'<div class="section-box"><div class="section-title">Scadenze Principali</div><div class="section-text">{debito.get("scadenze_principali","N/D")}</div></div>', unsafe_allow_html=True)
                if debito.get('note', 'N/D') != 'N/D':
                    st.markdown(f'<div class="section-box"><div class="section-title">Note</div><div class="section-text">{debito.get("note","N/D")}</div></div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="section-box"><div class="section-text">{debito}</div></div>', unsafe_allow_html=True)

        with st.expander("👥 Struttura Ownership"):
            ownership = report.get("ownership", {})
            if isinstance(ownership, dict):
                st.markdown(f"""
                <div class="kpi-grid">
                    <div class="kpi-card"><div class="kpi-label">Azionista Principale</div><div class="kpi-value">{ownership.get('azionista_principale','N/D')}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Quota</div><div class="kpi-value">{ownership.get('quota_principale','N/D')}</div></div>
                </div>
                """, unsafe_allow_html=True)
                if ownership.get('altri_azionisti', 'N/D') != 'N/D':
                    st.markdown(f'<div class="section-box"><div class="section-title">Altri Azionisti</div><div class="section-text">{ownership.get("altri_azionisti","N/D")}</div></div>', unsafe_allow_html=True)
                if ownership.get('struttura_controllo', 'N/D') != 'N/D':
                    st.markdown(f'<div class="section-box"><div class="section-title">Struttura di Controllo</div><div class="section-text">{ownership.get("struttura_controllo","N/D")}</div></div>', unsafe_allow_html=True)
                if ownership.get('note', 'N/D') != 'N/D':
                    st.markdown(f'<div class="section-box"><div class="section-title">Note</div><div class="section-text">{ownership.get("note","N/D")}</div></div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="section-box"><div class="section-text">{ownership}</div></div>', unsafe_allow_html=True)

        with st.expander("🔀 Operazioni M&A"):
            operazioni = report.get("operazioni_ma", [])
            if operazioni and operazioni[0].get("descrizione", "N/D") != "N/D":
                for op in operazioni:
                    st.markdown(f"""
                    <div class="ma-item">
                        <span class="ma-anno">{op.get('anno','')}</span>
                        <div class="ma-tipo">{op.get('tipo','')}</div>
                        <div class="ma-desc">{op.get('descrizione','')}</div>
                    </div>""", unsafe_allow_html=True)
            else:
                st.write("Nessuna operazione rilevata.")

        with st.expander("📝 Note Aggiuntive"):
            st.markdown(f'<div class="section-box"><div class="section-text">{report.get("note_aggiuntive","N/D")}</div></div>', unsafe_allow_html=True)

elif st.session_state["pagina"] == "archivio":

    st.markdown("### 🗂️ Archivio Report")
    reports = carica_archivio()

    if not reports:
        st.info("Nessun report salvato. Genera il tuo primo report!")
    else:
        for item in reports:
            r = item["report"]
            fin = r.get("dati_finanziari", {})

            st.markdown(f"""
            <div class="archivio-card">
                <div class="archivio-nome">{r.get('nome_azienda','')}</div>
                <div class="archivio-data">📅 {item['data']}</div>
                <div class="archivio-kpi">
                    <div class="archivio-kpi-item"><span class="archivio-kpi-label">Ricavi</span><span class="archivio-kpi-val">{fin.get('ricavi','N/D')}</span></div>
                    <div class="archivio-kpi-item"><span class="archivio-kpi-label">EBITDA</span><span class="archivio-kpi-val">{fin.get('ebitda','N/D')}</span></div>
                    <div class="archivio-kpi-item"><span class="archivio-kpi-label">Utile Netto</span><span class="archivio-kpi-val">{fin.get('utile_netto','N/D')}</span></div>
                    <div class="archivio-kpi-item"><span class="archivio-kpi-label">Anno</span><span class="archivio-kpi-val">{fin.get('anno_riferimento','N/D')}</span></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            col_a, col_b, col_c = st.columns([1, 1, 4])
            with col_a:
                if st.button("📖 Apri", key=f"apri_{item['riga']}"):
                    st.session_state["report"] = r
                    st.session_state["pagina"] = "genera"
                    st.rerun()
            with col_b:
                if st.button("🗑️ Elimina", key=f"elimina_{item['riga']}"):
                    try:
                        sheet = get_sheet()
                        sheet.delete_rows(item["riga"])
                        st.success("Report eliminato.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Errore: {e}")