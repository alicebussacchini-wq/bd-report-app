import streamlit as st
import anthropic
import pdfplumber
import PyPDF2
import io
import os
import requests
import json
import base64
import time
import gspread
from pypdf import PdfReader, PdfWriter
from google.oauth2.service_account import Credentials
from datetime import datetime

# ── Configurazione pagina ─────────────────────────────────────────────────────

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
    .pdf-badge { border-radius: 8px; padding: 10px 14px; margin: 8px 0; font-size: 13px; }
    .pdf-badge-ok { background: #1a2e1a; border: 1px solid #4caf50; color: #4caf50; }
    .pdf-badge-warn { background: #2e2a1a; border: 1px solid #ff9800; color: #ff9800; }
    .pdf-badge-scan { background: #1a1a2e; border: 1px solid #2196f3; color: #90caf9; }
</style>
""", unsafe_allow_html=True)

# ── Logo e header ─────────────────────────────────────────────────────────────

def get_logo_base64():
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

# ── Costanti keyword per ricerca sezioni finanziarie ─────────────────────────

FINANCIAL_KEYWORDS = [
    ("stato patrimoniale",          10),
    ("conto economico",             10),
    ("rendiconto finanziario",      10),
    ("prospetto delle variazioni",   8),
    ("nota integrativa",             5),
    ("posizione finanziaria netta",  9),
    ("indebitamento finanziario",    9),
    ("pfn",                          6),
    ("ricavi delle vendite",         6),
    ("ricavi",                       4),
    ("ebitda",                       8),
    ("ebit",                         7),
    ("utile netto",                  6),
    ("perdita netto",                6),
    ("risultato netto",              6),
    ("risultato operativo",          6),
    ("ammortamenti",                 5),
    ("totale attivo",                6),
    ("totale passivo",               6),
    ("patrimonio netto",             7),
    ("capitale sociale",             4),
    ("debiti finanziari",            7),
    ("disponibilità liquide",        5),
    ("(migliaia di euro)",           8),
    ("(milioni di euro)",            8),
    ("in migliaia",                  7),
    ("in milioni",                   7),
    ("esercizio chiuso",             6),
    ("31 dicembre",                  5),
    ("31/12/",                       5),
]

SCORE_THRESHOLD = 8
CONTEXT_PAGES = 1

# ── Funzioni PDF intelligente ─────────────────────────────────────────────────

def _score_page(text: str) -> int:
    text_lower = text.lower()
    score = 0
    for keyword, weight in FINANCIAL_KEYWORDS:
        if keyword in text_lower:
            score += weight
    return score


def analizza_pdf(pdf_bytes: bytes) -> dict:
    """
    Analizza un PDF e trova le pagine con dati finanziari.
    Restituisce info sull'analisi e la modalità di estrazione.
    """
    result = {
        "mode": "text",
        "total_pages": 0,
        "selected_pages": [],
        "page_scores": [],
        "avg_chars": 0.0,
        "sections_found": [],
    }

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            total = len(pdf.pages)
            result["total_pages"] = total

            texts, scores, char_counts = [], [], []
            for page in pdf.pages:
                t = page.extract_text() or ""
                texts.append(t)
                scores.append(_score_page(t))
                char_counts.append(len(t))

            result["page_scores"] = scores
            avg_chars = sum(char_counts) / total if total > 0 else 0
            result["avg_chars"] = avg_chars

            # PDF scansionato: meno di 100 caratteri per pagina in media
            if avg_chars < 100:
                result["mode"] = "scanned"
                result["selected_pages"] = list(range(total))
                return result

            # Selezione pagine rilevanti + contesto
            relevant = set()
            for i, score in enumerate(scores):
                if score >= SCORE_THRESHOLD:
                    for j in range(
                        max(0, i - CONTEXT_PAGES),
                        min(total, i + CONTEXT_PAGES + 1)
                    ):
                        relevant.add(j)

            # Fallback: nessuna sezione trovata → prendiamo la metà centrale
            if not relevant:
                start = max(0, total // 4)
                end = min(total, 3 * total // 4)
                relevant = set(range(start, end))
                result["mode"] = "fallback"

            result["selected_pages"] = sorted(relevant)

            # Sezioni identificate
            all_text = " ".join(texts).lower()
            sections = []
            for kw, label in [
                ("stato patrimoniale", "Stato Patrimoniale"),
                ("conto economico", "Conto Economico"),
                ("rendiconto finanziario", "Rendiconto Finanziario"),
                ("posizione finanziaria netta", "PFN"),
                ("nota integrativa", "Nota Integrativa"),
            ]:
                if kw in all_text:
                    sections.append(label)
            result["sections_found"] = sections

    except Exception as e:
        # Se pdfplumber fallisce, trattalo come scansionato
        result["mode"] = "scanned"
        result["error"] = str(e)
        try:
            reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
            result["total_pages"] = len(reader.pages)
            result["selected_pages"] = list(range(len(reader.pages)))
        except:
            pass

    return result


def costruisci_pdf_chirurgico(pdf_bytes: bytes, page_indices: list) -> bytes:
    """Costruisce un PDF con solo le pagine selezionate (0-indexed)."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for i in page_indices:
        if 0 <= i < len(reader.pages):
            writer.add_page(reader.pages[i])
    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()


def prepara_bilancio(pdf_bytes: bytes) -> tuple:
    """
    Entry point principale per preparare un bilancio per Claude.
    Restituisce (pdf_da_inviare_bytes, info_dict).
    """
    info = analizza_pdf(pdf_bytes)

    if info["mode"] == "scanned":
        return pdf_bytes, info

    surgical = costruisci_pdf_chirurgico(pdf_bytes, info["selected_pages"])
    info["pagine_inviate"] = len(info["selected_pages"])
    info["riduzione_pct"] = round(
        (1 - len(surgical) / len(pdf_bytes)) * 100, 1
    ) if len(pdf_bytes) > 0 else 0

    return surgical, info

# ── Funzioni Google Sheets ────────────────────────────────────────────────────

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

# ── Navigazione ───────────────────────────────────────────────────────────────

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

# ══════════════════════════════════════════════════════════════════════════════
# PAGINA: GENERA REPORT
# ══════════════════════════════════════════════════════════════════════════════

if st.session_state["pagina"] == "genera":

    lingua = st.radio("🌐 Lingua del report", ["Italiano", "English"], horizontal=True)

    # ── SEZIONE BILANCI (multi-upload) ────────────────────────────────────────

    st.markdown("### 📂 Carica i bilanci")
    st.caption(
        "Puoi caricare fino a 5 bilanci PDF completi contemporaneamente. "
        "Il sistema trova automaticamente le pagine finanziarie rilevanti — "
        "non serve estrarre le pagine manualmente."
    )

    bilanci_files = st.file_uploader(
        "Seleziona uno o più bilanci PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key="multi_bilanci",
        label_visibility="collapsed",
    )

    # Form metadati + analisi per ogni bilancio caricato
    bilanci_pronti = []   # lista di dict pronti per Claude

    if bilanci_files:
        if len(bilanci_files) > 5:
            st.error("Massimo 5 bilanci alla volta.")
            bilanci_files = bilanci_files[:5]

        st.markdown("#### Conferma i dettagli")

        for idx, f in enumerate(bilanci_files):
            with st.container():
                st.markdown(f"**📄 {f.name}**")
                col_m1, col_m2 = st.columns([3, 1])
                with col_m1:
                    rs = st.text_input(
                        "Ragione sociale",
                        key=f"rs_{idx}",
                        placeholder="es. Zambon S.p.A.",
                    )
                with col_m2:
                    anno = st.selectbox(
                        "Anno bilancio",
                        options=[str(y) for y in range(2025, 2017, -1)],
                        key=f"anno_{idx}",
                    )

                # Analisi pagine (cached in session_state per evitare ri-analisi)
                cache_key = f"pdf_analysis_{f.name}_{f.size}"
                if cache_key not in st.session_state:
                    with st.spinner(f"Analisi pagine di {f.name}..."):
                        raw_bytes = f.read()
                        surgical_bytes, info = prepara_bilancio(raw_bytes)
                        st.session_state[cache_key] = {
                            "surgical": surgical_bytes,
                            "info": info,
                        }
                else:
                    surgical_bytes = st.session_state[cache_key]["surgical"]
                    info = st.session_state[cache_key]["info"]

                # Badge risultato analisi
                mode = info.get("mode", "text")
                total = info.get("total_pages", 0)
                selected = len(info.get("selected_pages", []))
                sections = ", ".join(info.get("sections_found", [])) or "—"
                riduzione = info.get("riduzione_pct", 0)

                if mode == "scanned":
                    st.markdown(
                        f'<div class="pdf-badge pdf-badge-scan">'
                        f'📷 PDF scansionato — {total} pagine totali inviate a Claude direttamente.'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                elif mode == "fallback":
                    st.markdown(
                        f'<div class="pdf-badge pdf-badge-warn">'
                        f'⚠️ Layout atipico — invio {selected} pagine centrali su {total} totali.'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        f'<div class="pdf-badge pdf-badge-ok">'
                        f'✅ {selected} pagine rilevanti su {total} totali '
                        f'({riduzione}% riduzione) — Sezioni: {sections}'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

                # Pannello debug opzionale
                with st.expander("🔍 Dettaglio analisi pagine", expanded=False):
                    scores = info.get("page_scores", [])
                    selected_set = set(info.get("selected_pages", []))
                    if scores:
                        st.caption(f"Soglia punteggio: {SCORE_THRESHOLD} | "
                                   f"Caratteri medi per pagina: {info.get('avg_chars', 0):.0f}")
                        cols_dbg = st.columns(10)
                        for i, score in enumerate(scores[:50]):
                            col = cols_dbg[i % 10]
                            color = "#c8e04a" if i in selected_set else "#555"
                            col.markdown(
                                f"<div style='text-align:center;font-size:10px;"
                                f"color:{color};border:1px solid {color};"
                                f"border-radius:4px;padding:3px;margin:2px;'>"
                                f"<b>{i+1}</b><br>{score}</div>",
                                unsafe_allow_html=True,
                            )
                        if len(scores) > 50:
                            st.caption(f"... e altre {len(scores)-50} pagine.")

                if rs.strip():
                    bilanci_pronti.append({
                        "ragione_sociale": rs.strip(),
                        "anno": anno,
                        "pdf_b64": base64.b64encode(surgical_bytes).decode(),
                        "filename": f.name,
                        "info": info,
                    })
                else:
                    st.warning("⬆️ Inserisci la ragione sociale per includere questo bilancio.")

                st.markdown("---")

    # ── ALTRI DOCUMENTI (invariati) ───────────────────────────────────────────

    st.markdown("### 📎 Altri documenti (opzionale)")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.caption("📈 Export Mergermarket")
        mergermarket = st.file_uploader(
            "Mergermarket", type=["pdf", "csv"],
            key="merger", label_visibility="collapsed"
        )

    with col2:
        st.caption("🌐 URL Sito / Press Release")
        url_azienda = st.text_input(
            "URL", placeholder="https://...", label_visibility="collapsed"
        )

    with col3:
        st.caption("🏛️ Visura Camerale")
        visura = st.file_uploader(
            "Visura PDF", type=["pdf"],
            key="visura", label_visibility="collapsed"
        )

    # Raccolta testi documenti supplementari
    testi_supplementari = {}
    documenti_binari = {}

    if mergermarket:
        contenuto = mergermarket.read()
        if mergermarket.name.endswith(".pdf"):
            reader = PyPDF2.PdfReader(io.BytesIO(contenuto))
            testo = "".join(p.extract_text() or "" for p in reader.pages)
        else:
            testo = contenuto.decode("utf-8")
        testi_supplementari["Mergermarket"] = testo
        documenti_binari[mergermarket.name] = contenuto
        st.success("✅ Mergermarket caricato")

    if url_azienda:
        testi_supplementari["Sito Aziendale"] = estrai_testo_url(url_azienda)
        st.success("✅ URL acquisito")

    if visura:
        contenuto = visura.read()
        try:
            client_temp = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
            resp_visura = client_temp.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=2000,
                messages=[{"role": "user", "content": [
                    {"type": "document", "source": {
                        "type": "base64", "media_type": "application/pdf",
                        "data": base64.b64encode(contenuto).decode()
                    }},
                    {"type": "text", "text": "Estrai tutte le informazioni rilevanti: "
                     "ragione sociale, sede, codice fiscale, soci, amministratori, "
                     "capitale sociale, oggetto sociale."}
                ]}]
            )
            testo_visura = resp_visura.content[0].text
        except:
            reader = PyPDF2.PdfReader(io.BytesIO(contenuto))
            testo_visura = "".join(p.extract_text() or "" for p in reader.pages)
        testi_supplementari["Visura Camerale"] = testo_visura
        documenti_binari[visura.name] = contenuto
        st.success("✅ Visura caricata")

    # ── PULSANTE GENERA ───────────────────────────────────────────────────────

    st.markdown("---")

    # Abilita il bottone solo se c'è almeno un bilancio pronto
    # oppure almeno un documento supplementare
    ha_contenuto = len(bilanci_pronti) > 0 or len(testi_supplementari) > 0

    if not ha_contenuto:
        st.info("Carica almeno un bilancio o documento per procedere.")

    if st.button("🚀 Genera Report", disabled=not ha_contenuto):

        lingua_prompt = "in inglese" if lingua == "English" else "in italiano"
        client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

        # Reset lista report per questa nuova generazione
        st.session_state["reports_generati"] = []

        # ── Caso A: uno o più bilanci caricati → report per ciascuno ─────────
        if bilanci_pronti:
            totale = len(bilanci_pronti)
            progress = st.progress(0, text="Preparazione...")

            for idx, item in enumerate(bilanci_pronti):
                nome = item["ragione_sociale"]
                anno = item["anno"]
                info = item["info"]
                mode = info.get("mode", "text")
                pagine_inviate = len(info.get("selected_pages", []))
                totale_pagine = info.get("total_pages", 0)

                progress.progress(
                    idx / totale,
                    text=f"Analisi {idx+1}/{totale}: {nome} ({anno})..."
                )

                # Testo supplementare per questo report
                testo_suppl = ""
                for fonte, testo in testi_supplementari.items():
                    testo_suppl += f"\n\n--- {fonte} ---\n{testo[:5000]}"

                nota_estrazione = (
                    f"Il bilancio è scansionato — leggi i dati dalle immagini."
                    if mode == "scanned" else
                    f"Sono state selezionate {pagine_inviate} pagine rilevanti "
                    f"su {totale_pagine} totali del bilancio completo."
                )

                prompt_testo = f"""Sei un analista M&A e finance di uno studio legale internazionale.
Analizza il bilancio allegato di **{nome}** (esercizio {anno}) e, se presenti, i documenti supplementari.
{nota_estrazione}
Produci un report strutturato in JSON, con tutti i testi {lingua_prompt}.
{testo_suppl}

Rispondi SOLO con un oggetto JSON valido, senza backtick, senza testo aggiuntivo:

{{
  "nome_azienda": "{nome}",
  "overview": "",
  "core_business": "",
  "mercati": "",
  "dati_finanziari": {{
    "ricavi": "",
    "ebitda": "",
    "utile_netto": "",
    "totale_attivo": "",
    "patrimonio_netto": "",
    "anno_riferimento": "{anno}"
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

                for tentativo in range(3):
                    try:
                        messaggio = client.messages.create(
                            model="claude-haiku-4-5-20251001",
                            max_tokens=4000,
                            messages=[{"role": "user", "content": [
                                {
                                    "type": "document",
                                    "source": {
                                        "type": "base64",
                                        "media_type": "application/pdf",
                                        "data": item["pdf_b64"],
                                    }
                                },
                                {"type": "text", "text": prompt_testo}
                            ]}]
                        )
                        risposta = messaggio.content[0].text.strip()
                        if risposta.startswith("```"):
                            risposta = risposta.split("```")[1]
                            if risposta.startswith("json"):
                                risposta = risposta[4:]
                        report = json.loads(risposta.strip())
                        salva_report(nome, report, documenti_binari)

                        # Salva in session_state (lista per multi-report)
                        if "reports_generati" not in st.session_state:
                            st.session_state["reports_generati"] = []
                        st.session_state["reports_generati"].append({
                            "nome": nome,
                            "anno": anno,
                            "report": report,
                        })
                        # Compatibilità con il codice esistente
                        st.session_state["report"] = report
                        break

                    except Exception as e:
                        if tentativo < 2:
                            st.warning(f"{nome}: tentativo {tentativo+1} fallito, riprovo...")
                            time.sleep(30)
                        else:
                            st.error(f"❌ Errore per {nome}: {e}")

            progress.progress(1.0, text="Completato!")
            st.success(f"✅ {len(bilanci_pronti)} report generati e salvati!")

        # ── Caso B: solo documenti supplementari (comportamento originale) ────
        else:
            nome_azienda = st.text_input(
                "Nome dell'azienda", placeholder="es. Eco Eridania S.p.A.",
                key="nome_solo_suppl"
            )
            if not nome_azienda:
                st.warning("Inserisci il nome dell'azienda.")
            else:
                with st.spinner("Claude sta analizzando i documenti..."):
                    testo_completo = ""
                    for fonte, testo in testi_supplementari.items():
                        testo_completo += f"\n\n--- FONTE: {fonte} ---\n{testo[:10000]}"

                    lingua_prompt2 = "in inglese" if lingua == "English" else "in italiano"
                    prompt2 = f"""Sei un analista M&A e finance di uno studio legale internazionale.
Analizza i seguenti documenti relativi all'azienda {nome_azienda} e produci un report in JSON {lingua_prompt2}.

{testo_completo}

Rispondi SOLO con un oggetto JSON valido:

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

                    try:
                        messaggio = client.messages.create(
                            model="claude-haiku-4-5-20251001",
                            max_tokens=4000,
                            messages=[{"role": "user", "content": prompt2}]
                        )
                        risposta = messaggio.content[0].text.strip()
                        if risposta.startswith("```"):
                            risposta = risposta.split("```")[1]
                            if risposta.startswith("json"):
                                risposta = risposta[4:]
                        report = json.loads(risposta.strip())
                        salva_report(nome_azienda, report, documenti_binari)
                        st.session_state["reports_generati"] = [{
                            "nome": nome_azienda,
                            "anno": "",
                            "report": report,
                        }]
                        st.success("✅ Report generato e salvato!")
                    except Exception as e:
                        st.error(f"Errore: {type(e).__name__}: {str(e)}")

    # ── VISUALIZZAZIONE REPORT ────────────────────────────────────────────────

    # Mostra tutti i report generati in questa sessione (multi-bilancio)
    if "reports_generati" in st.session_state and st.session_state["reports_generati"]:
        for item_r in st.session_state["reports_generati"]:
            report = item_r["report"]
            nome_r = item_r["nome"]
            anno_r = item_r["anno"]
            fin = report.get("dati_finanziari", {})

            st.markdown(f"## 📋 {nome_r} — {anno_r}")
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

            st.markdown("---")

    # Compatibilità con singolo report (sessione precedente o caso B)
    elif "report" in st.session_state and "reports_generati" not in st.session_state:
        report = st.session_state["report"]
        fin = report.get("dati_finanziari", {})

        st.markdown(f"## 📋 {report.get('nome_azienda', '')}")
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

# ══════════════════════════════════════════════════════════════════════════════
# PAGINA: ARCHIVIO
# ══════════════════════════════════════════════════════════════════════════════

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
                    st.session_state.pop("reports_generati", None)
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
