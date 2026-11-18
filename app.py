import streamlit as st
import anthropic
import PyPDF2
import io
import os
import requests
import json
import base64
import hashlib
import html
import time
import gspread
from pypdf import PdfReader, PdfWriter
from google.oauth2.service_account import Credentials
from datetime import datetime as _datetime
from zoneinfo import ZoneInfo
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

import importlib
import bilancio_engine as be

# Dopo un aggiornamento del codice Streamlit riesegue app.py ma può tenere in
# memoria la versione precedente di bilancio_engine: se il file è cambiato
# rispetto a quello caricato, lo si ricarica.
try:
    _mtime_engine = os.path.getmtime(be.__file__)
    if getattr(be, "_mtime_caricato", None) != _mtime_engine:
        be = importlib.reload(be)
        be._mtime_caricato = _mtime_engine
except Exception:
    pass


class datetime(_datetime):
    """Come datetime, ma now() restituisce l'ora italiana: il server di
    Streamlit Cloud usa l'ora UTC e i report risultavano indietro di 1-2 ore."""
    @classmethod
    def now(cls, tz=None):
        try:
            return _datetime.now(tz or ZoneInfo("Europe/Rome"))
        except Exception:
            return _datetime.now(tz)

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

# ── Normalizzazione e validazione KPI ────────────────────────────────────────

# Fattori di conversione verso EURO INTERI
UNIT_FACTORS = {
    "euro": 1,
    "eur": 1,
    "€": 1,
    "migliaia di euro": 1_000,
    "migliaia": 1_000,
    "k€": 1_000,
    "milioni di euro": 1_000_000,
    "milioni": 1_000_000,
    "mln": 1_000_000,
    "mln€": 1_000_000,
    "miliardi di euro": 1_000_000_000,
    "miliardi": 1_000_000_000,
    "thousand": 1_000,
    "/000": 1_000,
    "000": 1_000,
    "million": 1_000_000,
    "billion": 1_000_000_000,
}

KPI_NUMERICI = [
    "ricavi", "ebitda", "ebit", "utile_netto", "ammortamenti",
    "totale_attivo", "patrimonio_netto", "cassa",
    "indebitamento_totale", "debito_bancario", "obbligazioni", "debito_netto",
]

def normalizza_kpi(report: dict) -> dict:
    """
    Normalizza tutti i KPI numerici a euro interi.
    Aggiunge campo '_normalizzato': True e '_unita_originale' al report.
    """
    fin = report.get("dati_finanziari", {})
    debito = report.get("struttura_debito", {})

    unita_raw = str(fin.get("unita", "migliaia di euro")).lower().strip()
    fattore = UNIT_FACTORS.get(unita_raw, None)

    # Prova match parziale se non trova esatto. Le chiavi più lunghe hanno la
    # precedenza: "in migliaia di euro" contiene anche "euro", ma vale 1.000.
    if fattore is None:
        for chiave in sorted(UNIT_FACTORS, key=len, reverse=True):
            if chiave in unita_raw:
                fattore = UNIT_FACTORS[chiave]
                break
    if fattore is None:
        fattore = 1_000  # default migliaia se non riconosciuto

    def converti(v):
        if v is None or v == "N/D" or v == "":
            return None
        try:
            return float(v) * fattore
        except (TypeError, ValueError):
            pass
        # Prova a estrarre numero e unità dalla stringa (es. "€ 1.2 miliardi")
        import re
        s = str(v).lower().strip()
        # Fattore dalla stringa stessa (per report vecchi senza campo "unita")
        str_factor = 1
        if "miliard" in s:
            str_factor = 1_000_000_000
        elif "milion" in s or " mln" in s or "mln " in s:
            str_factor = 1_000_000
        elif "migliaia" in s or " k" in s or "k€" in s:
            str_factor = 1_000
        # Se abbiamo già un fattore dalla stringa, non moltiplicare anche per fattore unità
        effective_factor = str_factor if str_factor > 1 else fattore
        nums = re.findall(r"-?[\d][\d.,]*", s)
        if nums:
            num_str = nums[0].replace(".", "").replace(",", ".")
            try:
                return float(num_str) * effective_factor
            except ValueError:
                pass
        return None

    # Normalizza dati finanziari
    fin_norm = {k: v for k, v in fin.items()}
    for campo in ["ricavi", "ebitda", "ebit", "utile_netto", "ammortamenti",
                  "totale_attivo", "patrimonio_netto", "cassa"]:
        fin_norm[campo] = converti(fin.get(campo))
    fin_norm["unita"] = "euro"
    fin_norm["_unita_originale"] = unita_raw
    fin_norm["_fattore"] = fattore

    # Normalizza struttura debito
    debito_norm = {k: v for k, v in debito.items()}
    for campo in ["indebitamento_totale", "debito_bancario", "obbligazioni", "debito_netto"]:
        debito_norm[campo] = converti(debito.get(campo))

    report["dati_finanziari"] = fin_norm
    report["struttura_debito"] = debito_norm
    report["_normalizzato"] = True
    return report


def valida_kpi(report: dict) -> list:
    """
    Esegue check di validazione sui KPI normalizzati.
    Restituisce una lista di warning (dict con 'livello', 'campo', 'messaggio').
    Livelli: 'error' (identità contabile violata), 'warning' (anomalia).
    """
    warnings = []
    fin = report.get("dati_finanziari", {})
    debito = report.get("struttura_debito", {})

    def get(d, k):
        v = d.get(k)
        return float(v) if v is not None else None

    ricavi      = get(fin, "ricavi")
    ebitda      = get(fin, "ebitda")
    ebit        = get(fin, "ebit")
    amm         = get(fin, "ammortamenti")
    utile       = get(fin, "utile_netto")
    attivo      = get(fin, "totale_attivo")
    pn          = get(fin, "patrimonio_netto")
    cassa       = get(fin, "cassa")
    deb_netto   = get(debito, "debito_netto")
    deb_tot     = get(debito, "indebitamento_totale")

    # ── Check identità contabili ─────────────────────────────────────────────

    # EBITDA = EBIT + Ammortamenti (tolleranza 5%)
    if (ebitda is not None and ebit is not None and amm is not None
            and report.get("ebitda_metodo") != "calcolato"):
        atteso = ebit + amm
        if atteso != 0 and abs(ebitda - atteso) / abs(atteso) > 0.05:
            warnings.append({
                "livello": "error",
                "campo": "EBITDA",
                "messaggio": f"EBITDA ({_fmt(ebitda)}) ≠ EBIT ({_fmt(ebit)}) + Amm. ({_fmt(amm)}) = {_fmt(atteso)}"
            })

    # Patrimonio Netto deve essere < Totale Attivo
    if pn is not None and attivo is not None and attivo > 0:
        if pn > attivo:
            warnings.append({
                "livello": "error",
                "campo": "Patrimonio Netto",
                "messaggio": f"Patrimonio netto ({_fmt(pn)}) > Totale attivo ({_fmt(attivo)}) — impossibile"
            })

    # ── Check anomalie di valore ─────────────────────────────────────────────

    # Margine EBITDA fuori range ragionevole (-50% / +80%)
    if ebitda is not None and ricavi is not None and ricavi > 0:
        margine = ebitda / ricavi
        if margine < -0.5 or margine > 0.8:
            warnings.append({
                "livello": "warning",
                "campo": "Margine EBITDA",
                "messaggio": f"Margine EBITDA {margine*100:.1f}% fuori range normale (-50%/+80%) — verificare unità"
            })

    # Cassa > Totale Attivo (impossibile)
    if cassa is not None and attivo is not None and attivo > 0:
        if cassa > attivo:
            warnings.append({
                "livello": "error",
                "campo": "Cassa",
                "messaggio": f"Cassa ({_fmt(cassa)}) > Totale attivo ({_fmt(attivo)}) — errore di estrazione"
            })

    # Ricavi negativi (quasi sempre errore)
    if ricavi is not None and ricavi < 0:
        warnings.append({
            "livello": "warning",
            "campo": "Ricavi",
            "messaggio": f"Ricavi negativi ({_fmt(ricavi)}) — verificare il dato estratto"
        })

    # Debito netto molto alto vs patrimonio netto (leva > 10x)
    if deb_netto is not None and pn is not None and pn > 0:
        leva = deb_netto / pn
        if leva > 10:
            warnings.append({
                "livello": "warning",
                "campo": "Leva finanziaria",
                "messaggio": f"Leva {leva:.1f}x molto elevata — possibile errore di estrazione"
            })

    # ── Check struttura debito (categorizzazione corretta) ──────────────────

    deb_banc = get(debito, "debito_bancario")
    obblig   = get(debito, "obbligazioni")

    # Indebitamento totale deve essere >= debito bancario + obbligazioni
    # (perché la voce D del passivo INCLUDE D.1, D.2 e D.4)
    if deb_tot is not None and deb_banc is not None:
        somma_finanziaria = deb_banc + (obblig or 0)
        # Tolleranza 1% per arrotondamenti
        if somma_finanziaria > 0 and deb_tot < somma_finanziaria * 0.99:
            warnings.append({
                "livello": "error",
                "campo": "Indebitamento Totale",
                "messaggio": (
                    f"Indebitamento totale ({_fmt(deb_tot)}) < debito bancario ({_fmt(deb_banc)}) "
                    f"+ obbligazioni ({_fmt(obblig or 0)}) = {_fmt(somma_finanziaria)}. "
                    f"Probabilmente è stato letto solo il debito finanziario invece del Totale Debiti "
                    f"(voce D del passivo)."
                )
            })

    # Indebitamento totale sospettosamente uguale al debito bancario (lettura sbagliata)
    if deb_tot is not None and deb_banc is not None and deb_banc > 0:
        if abs(deb_tot - deb_banc) / deb_banc < 0.01:
            warnings.append({
                "livello": "warning",
                "campo": "Indebitamento Totale",
                "messaggio": (
                    f"Indebitamento totale ({_fmt(deb_tot)}) coincide con debito bancario. "
                    f"Verifica: il primo dovrebbe essere il Totale Debiti del passivo (voce D), "
                    f"che include anche fornitori, tributari, previdenziali, ecc."
                )
            })

    # Indebitamento totale > Totale Attivo (impossibile)
    if deb_tot is not None and attivo is not None and attivo > 0:
        if deb_tot > attivo:
            warnings.append({
                "livello": "error",
                "campo": "Indebitamento Totale",
                "messaggio": f"Indebitamento totale ({_fmt(deb_tot)}) > Totale attivo ({_fmt(attivo)}) — impossibile"
            })

    # Patrimonio netto + Indebitamento totale dovrebbero essere ≈ Totale Attivo
    # (PN + Debiti + Fondi + TFR + Ratei = Totale Passivo = Totale Attivo)
    # Se PN + Debiti > Totale Attivo, c'è un errore
    if (pn is not None and deb_tot is not None and attivo is not None
            and attivo > 0):
        somma = pn + deb_tot
        if somma > attivo * 1.05:  # tolleranza 5%
            warnings.append({
                "livello": "warning",
                "campo": "Quadratura Stato Patrimoniale",
                "messaggio": (
                    f"Patrimonio Netto ({_fmt(pn)}) + Indebitamento ({_fmt(deb_tot)}) = "
                    f"{_fmt(somma)} > Totale Attivo ({_fmt(attivo)}). "
                    f"Verifica le voci del passivo."
                )
            })

    return warnings


def confronta_anni(report_nuovo: dict, report_vecchio: dict) -> list:
    """
    Confronta due report dello stesso soggetto (anni diversi).
    Restituisce warning per variazioni YoY anomale (>200%).
    """
    warnings = []
    fin_n = report_nuovo.get("dati_finanziari", {})
    fin_v = report_vecchio.get("dati_finanziari", {})

    for campo in ["ricavi", "ebitda", "utile_netto", "totale_attivo", "patrimonio_netto"]:
        v_nuovo = fin_n.get(campo)
        v_vecchio = fin_v.get(campo)
        if v_nuovo is None or v_vecchio is None:
            continue
        try:
            v_nuovo, v_vecchio = float(v_nuovo), float(v_vecchio)
        except (TypeError, ValueError):
            continue
        if v_vecchio == 0:
            continue
        variazione = (v_nuovo - v_vecchio) / abs(v_vecchio)
        if abs(variazione) > 2.0:  # >200% YoY
            warnings.append({
                "livello": "warning",
                "campo": campo.replace("_", " ").title(),
                "messaggio": f"Variazione YoY {variazione*100:+.0f}% — possibile errore di estrazione o cambio unità"
            })
    return warnings


# Formato dei numeri: italiano di default; inglese per i report in inglese.
_FORMATO = {"inglese": False}


def imposta_formato(report: dict):
    """Da chiamare prima di mostrare o esportare un report."""
    _FORMATO["inglese"] = isinstance(report, dict) and report.get("_lingua") == "en"


def _dec(valore: float, decimali: int) -> str:
    testo = f"{valore:,.{decimali}f}"
    if _FORMATO["inglese"]:
        return testo
    return testo.replace(",", "§").replace(".", ",").replace("§", ".")


def _fmt(v):
    """Formatta un numero in euro in modo leggibile (€ 15,51 mld; € 679,0 mln)."""
    if v is None:
        return "N/D"
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    inglese = _FORMATO["inglese"]
    if abs(v) >= 1_000_000_000:
        return f"€ {_dec(v/1_000_000_000, 2)}{'B' if inglese else ' mld'}"
    if abs(v) >= 1_000_000:
        return f"€ {_dec(v/1_000_000, 1)}{'M' if inglese else ' mln'}"
    if abs(v) >= 1_000:
        return f"€ {_dec(v/1_000, 0)}{'K' if inglese else ' mila'}"
    return f"€ {_dec(v, 0)}"


def fmt_pct(valore: float, decimali: int = 1, segno: bool = False) -> str:
    testo = _dec(valore, decimali)
    if segno and valore > 0:
        testo = "+" + testo
    return testo + "%"


def confronto_esercizi(report: dict) -> list:
    """
    Righe per il confronto con l'esercizio precedente:
    (etichetta, valore corrente, valore precedente, variazione), già formattate.
    Lista vuota se l'esercizio precedente non è disponibile.
    """
    prec = report.get("dati_anno_precedente") or {}
    if not report.get("_normalizzato") or not prec.get("_normalizzato"):
        return []
    fin = report.get("dati_finanziari") or {}
    deb = report.get("struttura_debito") if isinstance(report.get("struttura_debito"), dict) else {}

    def num(v):
        try:
            return float(v) if v is not None else None
        except (TypeError, ValueError):
            return None

    def variazione(c, p):
        # con valori negativi (perdite, cassa netta) la percentuale è fuorviante
        if c is None or p is None or p <= 0 or c < 0:
            return "—"
        return fmt_pct((c - p) / abs(p) * 100, 1, segno=True)

    righe = []
    def riga(etichetta, c, p):
        righe.append((etichetta, _fmt(c) if c is not None else "N/D",
                      _fmt(p) if p is not None else "N/D", variazione(c, p)))

    for etichetta, sezione, campo in [
        ("Ricavi", fin, "ricavi"), ("EBITDA", fin, "ebitda"), (None, None, "margine"),
        ("EBIT", fin, "ebit"), ("Utile Netto", fin, "utile_netto"),
        ("Totale Attivo", fin, "totale_attivo"), ("Patrimonio Netto", fin, "patrimonio_netto"),
        ("Cassa", fin, "cassa"), ("Indebitamento Totale", deb, "indebitamento_totale"),
        ("Debito Bancario", deb, "debito_bancario"), ("Obbligazioni", deb, "obbligazioni"),
        ("Debito Netto", deb, "debito_netto"), (None, None, "leva"),
    ]:
        if campo == "margine":
            def m(e, r):
                return e / r * 100 if e is not None and r else None
            mc, mp = m(num(fin.get("ebitda")), num(fin.get("ricavi"))), m(num(prec.get("ebitda")), num(prec.get("ricavi")))
            delta = "—"
            if mc is not None and mp is not None:
                delta = ("+" if mc - mp > 0 else "") + _dec(mc - mp, 1) + " pp"
            righe.append(("Margine EBITDA", fmt_pct(mc) if mc is not None else "N/D",
                          fmt_pct(mp) if mp is not None else "N/D", delta))
        elif campo == "leva":
            def l(d, e):
                return d / e if d is not None and e and e > 0 and d > 0 else None
            lc = num(deb.get("leva_finanziaria"))
            lp = l(num(prec.get("debito_netto")), num(prec.get("ebitda")))
            righe.append(("Leva (PFN/EBITDA)", _dec(lc, 2) + "x" if lc is not None else "N/D",
                          _dec(lp, 2) + "x" if lp is not None else "n.s.", "—"))
        else:
            riga(etichetta, num(sezione.get(campo)), num(prec.get(campo)))
    return righe


def anno_precedente(report: dict, anno_corrente: str = "") -> str:
    prec = report.get("dati_anno_precedente") or {}
    if str(prec.get("anno") or "").strip():
        return str(prec["anno"]).strip()
    try:
        return str(int(str(anno_corrente)[:4]) - 1)
    except (TypeError, ValueError):
        return "Esercizio precedente"


def voci_note(testo) -> list:
    """Le note sul debito sono voci separate da " | "."""
    if testo in (None, "", "N/D"):
        return []
    return [v.strip() for v in str(testo).split(" | ") if v.strip()]


def fmt_kpi(report: dict, campo: str, sezione: str = "dati_finanziari") -> str:
    """
    Formatta un KPI per la visualizzazione.
    Se il report è normalizzato usa _fmt, altrimenti usa il valore grezzo.
    """
    sezione_dict = report.get(sezione, {})
    v = sezione_dict.get(campo)
    if report.get("_normalizzato") and v is not None:
        return _fmt(v)
    if v is None:
        return "N/D"
    return str(v)


def render_pannello_validazione(warnings: list):
    """Renderizza il pannello di qualità estrazione con semaforo."""
    if not warnings:
        st.markdown(
            '<div style="background:#1a2e1a;border:1px solid #4caf50;border-radius:8px;'
            'padding:10px 14px;margin:8px 0;color:#4caf50;font-size:13px;">'
            '✅ <b>Validazione OK</b> — tutti i check superati</div>',
            unsafe_allow_html=True
        )
        return

    errori = [w for w in warnings if w["livello"] == "error"]
    avvisi = [w for w in warnings if w["livello"] == "warning"]

    if errori:
        colore, icona, label = "#c62828", "🔴", "Errori rilevati"
        bg = "#2e1a1a"
    else:
        colore, icona, label = "#e65100", "🟡", "Anomalie da verificare"
        bg = "#2e2a1a"

    items_html = "".join(
        f'<div style="margin:4px 0;font-size:12px;">'
        f'{"🔴" if w["livello"]=="error" else "🟡"} '
        f'<b>{w["campo"]}</b>: {w["messaggio"]}</div>'
        for w in warnings
    )
    st.markdown(
        f'<div style="background:{bg};border:1px solid {colore};border-radius:8px;'
        f'padding:10px 14px;margin:8px 0;">'
        f'<div style="color:{colore};font-weight:700;margin-bottom:6px;">{icona} {label}</div>'
        f'{items_html}</div>',
        unsafe_allow_html=True
    )


# ── Supporto analisi bilanci (la logica è in bilancio_engine.py) ──────────────

def _cache_sessione(prefisso: str, contenuto: bytes, funzione):
    """Esegue `funzione` una sola volta per contenuto e ne conserva il risultato.
    Senza questa cache ogni clic o tasto premuto rifaceva l'operazione
    (comprese le chiamate a pagamento a Claude per la visura)."""
    chiave = f"{prefisso}_{hashlib.md5(contenuto).hexdigest()}"
    if chiave not in st.session_state:
        st.session_state[chiave] = funzione()
    return st.session_state[chiave]


def _intervalli(pagine: list) -> str:
    """[1,2,3,7,9,10] -> "1-3, 7, 9-10" """
    if not pagine:
        return "—"
    pagine = sorted(set(pagine))
    gruppi, inizio, prec = [], pagine[0], pagine[0]
    for p in pagine[1:]:
        if p == prec + 1:
            prec = p
            continue
        gruppi.append((inizio, prec))
        inizio = prec = p
    gruppi.append((inizio, prec))
    return ", ".join(str(a) if a == b else f"{a}-{b}" for a, b in gruppi)


def fmt_leva(debito: dict) -> str:
    v = debito.get("leva_finanziaria") if isinstance(debito, dict) else None
    if v in (None, "", "N/D"):
        return "N/D"
    try:
        return _dec(float(v), 2) + "x"
    except (TypeError, ValueError):
        return str(v)


def render_fonti_verifica(report: dict):
    """Pannello con la pagina di origine di ogni numero e l'esito della seconda lettura."""
    estr = report.get("_estrazione") or {}
    fonti = report.get("fonti") or {}
    registro = report.get("_verifica") or []
    note_calc = report.get("_note_calcolo") or []
    if not (estr or fonti or registro):
        return

    corretti = [r for r in registro if r.get("esito") == "corretto"]
    dubbi = [r for r in registro if r.get("esito") in ("non_verificabile", "da_verificare", "non_eseguita")]
    testi = [r for r in registro if r.get("esito") == "testo_da_verificare"]
    testi_corretti = [r for r in registro if r.get("esito") == "testo_corretto"]
    confermati = [r for r in registro if r.get("esito") == "ok"]

    if registro:
        if not (corretti or dubbi or testi):
            colore, bg, titolo = "#4caf50", "#1a2e1a", (
                f"✅ <b>Seconda lettura</b> — {len(confermati)} valori riletti e confermati"
                + (f", {len(testi_corretti)} frasi corrette nei testi" if testi_corretti else ""))
        else:
            colore, bg = "#e65100", "#2e2a1a"
            parti = [f"{len(confermati)} confermati"]
            if corretti:
                parti.append(f"{len(corretti)} corretti")
            if testi_corretti:
                parti.append(f"{len(testi_corretti)} frasi corrette")
            if dubbi:
                parti.append(f"{len(dubbi)} non verificabili")
            if testi:
                parti.append(f"{len(testi)} affermazioni da controllare")
            titolo = "🟡 <b>Seconda lettura</b> — " + ", ".join(parti)
        st.markdown(
            f'<div style="background:{bg};border:1px solid {colore};border-radius:8px;'
            f'padding:10px 14px;margin:8px 0;color:{colore};font-size:13px;">{titolo}</div>',
            unsafe_allow_html=True
        )

    with st.expander("🔎 Fonti dei dati e controllo", expanded=bool(corretti or dubbi or testi)):
        if estr:
            st.caption(
                f"Perimetro: {report.get('perimetro') or 'N/D'} · "
                f"Tipo bilancio: {report.get('tipo_bilancio') or 'N/D'} · "
                f"Pagine analizzate: {_intervalli(estr.get('pagine_inviate', []))} "
                f"su {estr.get('pagine_totali', '?')} ({estr.get('criterio', '')}) · "
                f"Modello: {estr.get('modello', 'N/D')}"
            )

        fin = report.get("dati_finanziari") or {}
        deb = report.get("struttura_debito") if isinstance(report.get("struttura_debito"), dict) else {}
        righe = []
        for campo in be.KPI_TUTTI:
            valore = (fin if campo in be.KPI_FINANZIARI else deb).get(campo)
            fonte = fonti.get(campo) or {}
            if valore is None and not fonte.get("voce"):
                continue
            righe.append({
                "Dato": be.ETICHETTE_KPI[campo],
                "Valore": _fmt(valore) if report.get("_normalizzato") else str(valore),
                "Pagina": str(fonte.get("pagina") or "—"),
                "Voce letta nel bilancio": fonte.get("voce") or "—",
                "Valore letto / calcolo": fonte.get("calcolo") or fonte.get("letto") or "—",
            })
        if righe:
            st.dataframe(righe, hide_index=True, use_container_width=True)

        def _lista(titolo, colore, voci):
            if not voci:
                return
            corpo = "".join(f'<div style="margin:4px 0;font-size:12px;">{v}</div>' for v in voci)
            st.markdown(
                f'<div style="border-left:3px solid {colore};padding:6px 12px;margin:8px 0;">'
                f'<div style="color:{colore};font-weight:700;font-size:13px;">{titolo}</div>{corpo}</div>',
                unsafe_allow_html=True
            )

        esc = html.escape
        _lista("Valori corretti dalla seconda lettura", "#ff9800", [
            f"<b>{esc(be.ETICHETTE_KPI.get(r['campo'], r['campo']))}</b>: "
            f"{esc(str(r.get('prima')))} → {esc(str(r.get('dopo')))} "
            f"(unità del bilancio) — {esc(str(r.get('motivo', '')))}"
            for r in corretti
        ])
        _lista("Valori che la seconda lettura non ha potuto verificare", "#ff9800", [
            f"<b>{esc(be.ETICHETTE_KPI.get(r['campo'], r['campo']))}</b>: {esc(str(r.get('motivo', '')))}"
            for r in dubbi
        ])
        _lista("Frasi corrette dalla seconda lettura", "#90caf9", [
            f"<b>{esc(str(r.get('campo', '')))}</b>: «{esc(str(r.get('frase', '')))}» → "
            f"«{esc(str(r.get('correzione', '')) or '(eliminata)')}» — {esc(str(r.get('motivo', '')))}"
            for r in testi_corretti
        ])
        _lista("Affermazioni senza riscontro nei documenti — da controllare", "#ff9800", [
            f"<b>{esc(str(r.get('campo', '')))}</b>: «{esc(str(r.get('frase', '')))}» — {esc(str(r.get('motivo', '')))}"
            for r in testi
        ])
        _lista("Ricalcoli del programma", "#90caf9", [esc(str(n)) for n in note_calc])

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
                    # Normalizza sempre i KPI al caricamento,
                    # così i report vecchi (pre-M2) vengono convertiti on-the-fly
                    if not report.get("_normalizzato"):
                        report = normalizza_kpi(report)
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


# ── Esportazione Excel ────────────────────────────────────────────────────────

# Colori HL
HL_VERDE    = "C8E04A"
HL_NERO     = "1A1A1A"
HL_GRIGIO_S = "2A2A2A"
HL_GRIGIO_C = "444444"
HL_BIANCO   = "F0F0F0"

def _cell_style(ws, row, col, value, bold=False, bg=None, fg="F0F0F0",
                align="left", border=False, num_format=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name="Calibri", size=10, bold=bold,
                     color=fg if fg else "000000")
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center",
                               wrap_text=True)
    if border:
        side = Side(style="thin", color=HL_GRIGIO_C)
        cell.border = Border(left=side, right=side, top=side, bottom=side)
    if num_format:
        cell.number_format = num_format
    return cell


def genera_excel(reports: list) -> bytes:
    """
    Genera un file Excel con un foglio per ogni report
    più un foglio KPI Summary comparativo se ci sono più report.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # rimuovi foglio vuoto default

    # ── Foglio KPI Summary (sempre presente) ─────────────────────────────────
    ws_sum = wb.create_sheet("KPI Summary")
    ws_sum.sheet_view.showGridLines = False
    ws_sum.tab_color = HL_VERDE

    # Header riga 1
    ws_sum.row_dimensions[1].height = 30
    titolo_cell = ws_sum.cell(row=1, column=1, value="TAXI REPORT — KPI Summary")
    titolo_cell.font = Font(name="Calibri", size=14, bold=True, color=HL_VERDE)
    titolo_cell.fill = PatternFill("solid", fgColor=HL_NERO)
    ws_sum.merge_cells(start_row=1, start_column=1,
                       end_row=1, end_column=2 + len(reports))

    # Sottotitolo
    ws_sum.row_dimensions[2].height = 16
    sub = ws_sum.cell(row=2, column=1,
                      value=f"Generato il {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    sub.font = Font(name="Calibri", size=9, color="999999")
    sub.fill = PatternFill("solid", fgColor=HL_NERO)
    ws_sum.merge_cells(start_row=2, start_column=1,
                       end_row=2, end_column=2 + len(reports))

    # Intestazioni colonne: KPI | Azienda1 (anno) | Azienda2 (anno) ...
    ws_sum.row_dimensions[4].height = 28
    _cell_style(ws_sum, 4, 1, "KPI", bold=True, bg=HL_NERO, fg=HL_VERDE,
                align="left", border=True)
    _cell_style(ws_sum, 4, 2, "Unità", bold=True, bg=HL_NERO, fg=HL_VERDE,
                align="center", border=True)
    for j, r in enumerate(reports):
        label = f"{r['nome']} ({r['anno']})"
        _cell_style(ws_sum, 4, 3 + j, label, bold=True, bg=HL_NERO,
                    fg=HL_VERDE, align="center", border=True)

    # Righe KPI
    kpi_rows = [
        ("Ricavi",           "dati_finanziari", "ricavi",            "€ #,##0"),
        ("EBITDA",           "dati_finanziari", "ebitda",            "€ #,##0"),
        ("Margine EBITDA",   None,              "_margine_ebitda",   "0.0%"),
        ("EBIT",             "dati_finanziari", "ebit",              "€ #,##0"),
        ("Utile Netto",      "dati_finanziari", "utile_netto",       "€ #,##0"),
        ("Ammortamenti",     "dati_finanziari", "ammortamenti",      "€ #,##0"),
        ("Totale Attivo",    "dati_finanziari", "totale_attivo",     "€ #,##0"),
        ("Patrimonio Netto", "dati_finanziari", "patrimonio_netto",  "€ #,##0"),
        ("Cassa",            "dati_finanziari", "cassa",             "€ #,##0"),
        ("Debito Netto",     "struttura_debito","debito_netto",      "€ #,##0"),
        ("Indebitamento Tot","struttura_debito","indebitamento_totale","€ #,##0"),
        ("Leva (x)",         "struttura_debito","leva_finanziaria",  "0.0"),
    ]

    for i, (label, sezione, campo, fmt) in enumerate(kpi_rows):
        row_idx = 5 + i
        ws_sum.row_dimensions[row_idx].height = 20
        bg = HL_GRIGIO_S if i % 2 == 0 else HL_NERO
        _cell_style(ws_sum, row_idx, 1, label, bg=bg, fg=HL_BIANCO,
                    border=True)
        _cell_style(ws_sum, row_idx, 2, "EUR" if "€" in fmt else "x" if "0.0" == fmt else "%",
                    bg=bg, fg="999999", align="center", border=True)

        for j, r in enumerate(reports):
            report = r["report"]
            val = None
            if campo == "_margine_ebitda":
                ric = report.get("dati_finanziari", {}).get("ricavi")
                ebt = report.get("dati_finanziari", {}).get("ebitda")
                if ric and ebt:
                    try:
                        val = float(ebt) / float(ric)
                    except:
                        pass
            elif sezione:
                raw = report.get(sezione, {}).get(campo)
                if raw is not None:
                    try:
                        val = float(raw)
                    except:
                        pass

            cell = _cell_style(ws_sum, row_idx, 3 + j, val, bg=bg,
                               fg=HL_BIANCO, align="right", border=True,
                               num_format=fmt)

    # Warning validazione se presenti
    warn_row = 5 + len(kpi_rows) + 2
    has_warnings = any(r["report"].get("_validation_warnings") for r in reports)
    if has_warnings:
        w_cell = ws_sum.cell(row=warn_row, column=1,
                             value="⚠️ Note validazione")
        w_cell.font = Font(name="Calibri", size=10, bold=True, color="E65100")
        w_cell.fill = PatternFill("solid", fgColor="2E2A1A")
        ws_sum.merge_cells(start_row=warn_row, start_column=1,
                           end_row=warn_row, end_column=2 + len(reports))
        for r in reports:
            for w in r["report"].get("_validation_warnings", []):
                warn_row += 1
                msg = f"{r['nome']} ({r['anno']}) — {w['campo']}: {w['messaggio']}"
                wc = ws_sum.cell(row=warn_row, column=1, value=msg)
                wc.font = Font(name="Calibri", size=9,
                               color="C62828" if w["livello"] == "error" else "E65100")
                wc.fill = PatternFill("solid", fgColor=HL_NERO)
                ws_sum.merge_cells(start_row=warn_row, start_column=1,
                                   end_row=warn_row, end_column=2 + len(reports))

    # Larghezze colonne
    ws_sum.column_dimensions["A"].width = 22
    ws_sum.column_dimensions["B"].width = 8
    for j in range(len(reports)):
        ws_sum.column_dimensions[get_column_letter(3 + j)].width = 22

    # ── Un foglio per ogni report (dati completi) ─────────────────────────────
    for r in reports:
        nome_sheet = f"{r['nome'][:20]} {r['anno']}"
        ws = wb.create_sheet(nome_sheet)
        ws.sheet_view.showGridLines = False

        report = r["report"]
        fin    = report.get("dati_finanziari", {})
        debito = report.get("struttura_debito", {})

        # Intestazione
        ws.row_dimensions[1].height = 30
        h = ws.cell(row=1, column=1,
                    value=f"{r['nome']} — Bilancio {r['anno']}")
        h.font = Font(name="Calibri", size=13, bold=True, color=HL_VERDE)
        h.fill = PatternFill("solid", fgColor=HL_NERO)
        ws.merge_cells("A1:C1")

        # Sezione Dati Finanziari
        _cell_style(ws, 3, 1, "DATI FINANZIARI", bold=True, bg=HL_GRIGIO_S,
                    fg=HL_VERDE, border=True)
        _cell_style(ws, 3, 2, "Valore (EUR)", bold=True, bg=HL_GRIGIO_S,
                    fg=HL_VERDE, align="right", border=True)
        prec_xl = report.get("dati_anno_precedente") or {}
        ha_prec = bool(prec_xl.get("_normalizzato"))
        titolo_prec = f"{anno_precedente(report, r['anno'])} (EUR)" if ha_prec else "Note"
        chiavi_xl = {"Ricavi": "ricavi", "EBITDA": "ebitda", "EBIT": "ebit", "Utile Netto": "utile_netto",
                     "Ammortamenti": "ammortamenti", "Totale Attivo": "totale_attivo",
                     "Patrimonio Netto": "patrimonio_netto", "Cassa": "cassa",
                     "Indebitamento Totale": "indebitamento_totale", "Debito Bancario": "debito_bancario",
                     "Obbligazioni": "obbligazioni", "Debito Netto": "debito_netto"}
        _cell_style(ws, 3, 3, titolo_prec, bold=True, bg=HL_GRIGIO_S,
                    fg=HL_VERDE, align="right" if ha_prec else "left", border=True)

        fin_campi = [
            ("Ricavi",            fin.get("ricavi")),
            ("EBITDA",            fin.get("ebitda")),
            ("EBIT",              fin.get("ebit")),
            ("Utile Netto",       fin.get("utile_netto")),
            ("Ammortamenti",      fin.get("ammortamenti")),
            ("Totale Attivo",     fin.get("totale_attivo")),
            ("Patrimonio Netto",  fin.get("patrimonio_netto")),
            ("Cassa",             fin.get("cassa")),
        ]
        for i, (label, val) in enumerate(fin_campi):
            row = 4 + i
            bg = HL_GRIGIO_S if i % 2 == 0 else HL_NERO
            ws.row_dimensions[row].height = 18
            _cell_style(ws, row, 1, label, bg=bg, fg=HL_BIANCO, border=True)
            num = None
            if val is not None:
                try:
                    num = float(val)
                except:
                    pass
            _cell_style(ws, row, 2, num, bg=bg, fg=HL_BIANCO,
                        align="right", border=True, num_format="€ #,##0")
            if ha_prec:
                _cell_style(ws, row, 3, prec_xl.get(chiavi_xl.get(label)), bg=bg, fg=HL_BIANCO,
                            align="right", border=True, num_format="€ #,##0")
            else:
                _cell_style(ws, row, 3, "", bg=bg, fg="999999", border=True)

        # Sezione Struttura Debito
        deb_start = 4 + len(fin_campi) + 2
        _cell_style(ws, deb_start, 1, "STRUTTURA DEBITO", bold=True,
                    bg=HL_GRIGIO_S, fg=HL_VERDE, border=True)
        _cell_style(ws, deb_start, 2, "Valore (EUR)", bold=True,
                    bg=HL_GRIGIO_S, fg=HL_VERDE, align="right", border=True)
        _cell_style(ws, deb_start, 3, titolo_prec, bold=True, bg=HL_GRIGIO_S,
                    fg=HL_VERDE, align="right" if ha_prec else "left", border=True)

        deb_campi = [
            ("Indebitamento Totale", debito.get("indebitamento_totale")),
            ("Debito Bancario",      debito.get("debito_bancario")),
            ("Obbligazioni",         debito.get("obbligazioni")),
            ("Debito Netto",         debito.get("debito_netto")),
            ("Leva Finanziaria",     debito.get("leva_finanziaria")),
        ]
        for i, (label, val) in enumerate(deb_campi):
            row = deb_start + 1 + i
            bg = HL_GRIGIO_S if i % 2 == 0 else HL_NERO
            ws.row_dimensions[row].height = 18
            _cell_style(ws, row, 1, label, bg=bg, fg=HL_BIANCO, border=True)
            num = None
            if val is not None:
                try:
                    num = float(val)
                except:
                    pass
            fmt = "0.0" if label == "Leva Finanziaria" else "€ #,##0"
            _cell_style(ws, row, 2, num, bg=bg, fg=HL_BIANCO,
                        align="right", border=True, num_format=fmt)
            if ha_prec and label in chiavi_xl:
                _cell_style(ws, row, 3, prec_xl.get(chiavi_xl[label]), bg=bg, fg=HL_BIANCO,
                            align="right", border=True, num_format="€ #,##0")
            else:
                _cell_style(ws, row, 3, "", bg=bg, fg="999999", border=True)

        # Sezione testi
        testi_start = deb_start + len(deb_campi) + 2
        for label, campo in [
            ("OVERVIEW", "overview"),
            ("CORE BUSINESS", "core_business"),
            ("MERCATI", "mercati"),
            ("NOTE AGGIUNTIVE", "note_aggiuntive"),
        ]:
            val = report.get(campo, "N/D") or "N/D"
            _cell_style(ws, testi_start, 1, label, bold=True,
                        bg=HL_GRIGIO_S, fg=HL_VERDE, border=True)
            ws.merge_cells(start_row=testi_start, start_column=1,
                           end_row=testi_start, end_column=3)
            testi_start += 1
            tc = ws.cell(row=testi_start, column=1, value=val)
            tc.font = Font(name="Calibri", size=9, color=HL_BIANCO)
            tc.fill = PatternFill("solid", fgColor=HL_NERO)
            tc.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[testi_start].height = 60
            ws.merge_cells(start_row=testi_start, start_column=1,
                           end_row=testi_start, end_column=3)
            testi_start += 2

        # Larghezze
        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 35

    # ── Foglio Fonti: pagina e voce di origine di ogni numero ────────────────
    if any(r["report"].get("fonti") for r in reports):
        ws_f = wb.create_sheet("Fonti")
        ws_f.sheet_view.showGridLines = False
        intestazioni = ["Azienda", "Anno", "Dato", "Valore (EUR)", "Pagina PDF",
                        "Voce letta nel bilancio", "Valore letto / calcolo", "Controllo"]
        for c, titolo in enumerate(intestazioni, start=1):
            _cell_style(ws_f, 1, c, titolo, bold=True, bg=HL_GRIGIO_S, fg=HL_VERDE, border=True)
        etichette_esito = {
            "ok": "Confermato dalla seconda lettura",
            "corretto": "Corretto dalla seconda lettura",
            "non_verificabile": "Non verificabile",
            "da_verificare": "Da verificare",
        }
        riga = 2
        for r in reports:
            report = r["report"]
            fonti = report.get("fonti") or {}
            fin = report.get("dati_finanziari") or {}
            debito = report.get("struttura_debito") if isinstance(report.get("struttura_debito"), dict) else {}
            esiti = {e.get("campo"): e.get("esito") for e in report.get("_verifica") or []}
            for campo in be.KPI_TUTTI:
                valore = (fin if campo in be.KPI_FINANZIARI else debito).get(campo)
                fonte = fonti.get(campo) or {}
                if valore is None and not fonte.get("voce"):
                    continue
                bg = HL_NERO if riga % 2 == 0 else HL_GRIGIO_S
                try:
                    valore_num = float(valore) if valore is not None else None
                except (TypeError, ValueError):
                    valore_num = None
                celle = [
                    (r["nome"], None, "left"), (r["anno"], None, "left"),
                    (be.ETICHETTE_KPI[campo], None, "left"),
                    (valore_num, "€ #,##0", "right"),
                    (fonte.get("pagina"), None, "right"),
                    (fonte.get("voce") or "", None, "left"),
                    (fonte.get("calcolo") or fonte.get("letto") or "", None, "left"),
                    (etichette_esito.get(esiti.get(campo), ""), None, "left"),
                ]
                for c, (val, fmt, align) in enumerate(celle, start=1):
                    _cell_style(ws_f, riga, c, val, bg=bg, fg=HL_BIANCO, align=align,
                                border=True, num_format=fmt)
                riga += 1
        for lettera, larghezza in zip("ABCDEFGH", [26, 8, 22, 18, 11, 48, 40, 30]):
            ws_f.column_dimensions[lettera].width = larghezza

    # Output bytes
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Esportazione Word ─────────────────────────────────────────────────────────

def genera_word(reports: list) -> bytes:
    """
    Genera un documento Word con un report per azienda.
    Stile HL: header con logo/titolo, tabelle KPI, sezioni testo.
    """
    doc = Document()

    # Imposta margini
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # Stili base
    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Calibri"
    style_normal.font.size = Pt(10)

    def add_heading(text, level=1, color_hex="C8E04A"):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.name = "Calibri"
        run.font.bold = True
        run.font.size = Pt(14 if level == 1 else 11)
        run.font.color.rgb = RGBColor.from_string(color_hex)
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after  = Pt(4)
        return p

    def _scrivi_cella(cella, testo, grassetto=False, dimensione=10, grigio=False):
        """Scrive il testo nella cella; una lista diventa un elenco di voci."""
        voci = testo if isinstance(testo, list) else [testo]
        for n, voce in enumerate(voci):
            paragrafo = cella.paragraphs[0] if n == 0 else cella.add_paragraph()
            run = paragrafo.add_run(("• " if len(voci) > 1 else "") + str(voce))
            run.font.size = Pt(dimensione)
            run.font.bold = grassetto
            if grigio:
                run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
            paragrafo.paragraph_format.space_after = Pt(2 if len(voci) > 1 else 0)

    def add_kpi_table(doc, kpi_list):
        """Tabella a due colonne (dato | valore). Un valore lista diventa un elenco puntato."""
        table = doc.add_table(rows=0, cols=2)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = False
        for label, value in kpi_list:
            row = table.add_row()
            _scrivi_cella(row.cells[0], label, grassetto=True, dimensione=9, grigio=True)
            if isinstance(value, list):
                _scrivi_cella(row.cells[1], value, dimensione=9)
            else:
                lungo = len(str(value)) > 60
                _scrivi_cella(row.cells[1], str(value) if value is not None else "N/D",
                              grassetto=not lungo, dimensione=9 if lungo else 10)
            row.cells[0].width, row.cells[1].width = Cm(5), Cm(11)
        return table

    def add_confronto_table(doc, righe, anno_corrente, anno_prec):
        """Tabella a quattro colonne: dato | esercizio | esercizio precedente | variazione."""
        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = False
        larghezze = [Cm(5), Cm(4), Cm(4), Cm(3)]
        for c, titolo in enumerate(["", str(anno_corrente), str(anno_prec), "Variazione"]):
            _scrivi_cella(table.rows[0].cells[c], titolo, grassetto=True, dimensione=9, grigio=True)
            table.rows[0].cells[c].width = larghezze[c]
        for etichetta, corrente, precedente, variazione in righe:
            celle = table.add_row().cells
            _scrivi_cella(celle[0], etichetta, grassetto=True, dimensione=9, grigio=True)
            _scrivi_cella(celle[1], corrente, grassetto=True, dimensione=10)
            _scrivi_cella(celle[2], precedente, dimensione=10)
            _scrivi_cella(celle[3], variazione, dimensione=9)
            for c in range(4):
                celle[c].width = larghezze[c]
        return table

    for idx, r in enumerate(reports):
        report = r["report"]
        imposta_formato(report)
        fin    = report.get("dati_finanziari", {})
        debito = report.get("struttura_debito", {})
        righe_cfr = confronto_esercizi(report)
        indice_debito = next((i for i, x in enumerate(righe_cfr) if x[0] == "Indebitamento Totale"), len(righe_cfr))

        # ── Copertina sezione ────────────────────────────────────────────────
        if idx > 0:
            doc.add_page_break()

        # Titolo azienda
        p_title = doc.add_paragraph()
        run = p_title.add_run(r["nome"].upper())
        run.font.name   = "Calibri"
        run.font.bold   = True
        run.font.size   = Pt(20)
        run.font.color.rgb = RGBColor.from_string("C8E04A")
        p_title.paragraph_format.space_after = Pt(2)

        p_sub = doc.add_paragraph()
        run2 = p_sub.add_run(f"Bilancio {r['anno']}  |  Generato il {datetime.now().strftime('%d/%m/%Y')}")
        run2.font.name  = "Calibri"
        run2.font.size  = Pt(9)
        run2.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        p_sub.paragraph_format.space_after = Pt(16)

        # Linea separatore
        p_hr = doc.add_paragraph("─" * 80)
        p_hr.runs[0].font.color.rgb = RGBColor.from_string("C8E04A")
        p_hr.runs[0].font.size = Pt(7)
        p_hr.paragraph_format.space_after = Pt(12)

        # ── Warning validazione ───────────────────────────────────────────────
        warnings = report.get("_validation_warnings", [])
        if warnings:
            add_heading("⚠️ Note di validazione", level=2, color_hex="E65100")
            for w in warnings:
                p = doc.add_paragraph(style="List Bullet")
                run = p.add_run(f"{w['campo']}: {w['messaggio']}")
                run.font.size = Pt(9)
                run.font.color.rgb = (RGBColor(0xC6, 0x28, 0x28)
                                      if w["livello"] == "error"
                                      else RGBColor(0xE6, 0x51, 0x00))

        # ── Dati finanziari ───────────────────────────────────────────────────
        add_heading("Dati Finanziari", level=2)
        kpi_fin = [
            ("Ricavi",            _fmt(fin.get("ricavi"))),
            ("EBITDA",            _fmt(fin.get("ebitda"))),
            ("Margine EBITDA",    fmt_pct(float(fin['ebitda'])/float(fin['ricavi'])*100)
                                  if fin.get("ebitda") and fin.get("ricavi") else "N/D"),
            ("EBIT",              _fmt(fin.get("ebit"))),
            ("Utile Netto",       _fmt(fin.get("utile_netto"))),
            ("Totale Attivo",     _fmt(fin.get("totale_attivo"))),
            ("Patrimonio Netto",  _fmt(fin.get("patrimonio_netto"))),
            ("Cassa",             _fmt(fin.get("cassa"))),
        ]
        if righe_cfr:
            add_confronto_table(doc, righe_cfr[:indice_debito], r["anno"], anno_precedente(report, r["anno"]))
        else:
            add_kpi_table(doc, kpi_fin)

        # ── Struttura debito ──────────────────────────────────────────────────
        add_heading("Struttura del Debito", level=2)
        kpi_deb = [
            ("Indebitamento Totale", _fmt(debito.get("indebitamento_totale"))),
            ("Debito Bancario",      _fmt(debito.get("debito_bancario"))),
            ("Obbligazioni",         _fmt(debito.get("obbligazioni"))),
            ("Debito Netto",         _fmt(debito.get("debito_netto"))),
            ("Leva Finanziaria (PFN/EBITDA)", fmt_leva(debito)),
            ("Scadenze",             debito.get("scadenze_principali", "N/D") or "N/D"),
        ]
        testi_debito = [("Scadenze", debito.get("scadenze_principali", "N/D") or "N/D")] \
            if isinstance(debito, dict) else []
        if isinstance(debito, dict) and voci_note(debito.get("note")):
            testi_debito.append(("Note", voci_note(debito.get("note"))))
        if righe_cfr:
            add_confronto_table(doc, righe_cfr[indice_debito:], r["anno"], anno_precedente(report, r["anno"]))
            doc.add_paragraph().paragraph_format.space_after = Pt(0)
            add_kpi_table(doc, testi_debito)
        else:
            add_kpi_table(doc, kpi_deb[:-1] + testi_debito)

        # ── Sezioni testuali ──────────────────────────────────────────────────
        for label, campo in [
            ("Overview",       "overview"),
            ("Core Business",  "core_business"),
            ("Mercati",        "mercati"),
            ("Ownership",      None),
            ("Operazioni M&A", None),
            ("Note Aggiuntive","note_aggiuntive"),
        ]:
            if campo:
                testo = report.get(campo, "N/D") or "N/D"
            elif label == "Ownership":
                own = report.get("ownership", {})
                if isinstance(own, dict):
                    vuoto = (None, "", "N/D")
                    parti = []
                    if own.get("azionista_principale") not in vuoto:
                        quota = own.get("quota_principale")
                        parti.append(
                            f"Azionista principale: {own['azionista_principale']}"
                            + (f" ({quota})" if quota not in vuoto else "") + "."
                        )
                    if own.get("altri_azionisti") not in vuoto:
                        parti.append(f"Altri azionisti: {own['altri_azionisti']}")
                    for chiave in ("struttura_controllo", "note"):
                        if own.get(chiave) not in vuoto:
                            parti.append(str(own[chiave]))
                    testo = " ".join(parti) or "N/D"
                else:
                    testo = str(own)
            else:  # Operazioni M&A: un paragrafo per operazione
                ops = [o for o in report.get("operazioni_ma", [])
                       if isinstance(o, dict) and o.get("descrizione", "N/D") not in ("", "N/D")]
                if ops:
                    add_heading(label, level=2)
                    for o in ops:
                        riga = " — ".join(
                            str(x) for x in (o.get("anno"), o.get("tipo"), o.get("descrizione")) if x
                        )
                        if o.get("fonte"):
                            riga += f" (fonte: {o['fonte']})"
                        p = doc.add_paragraph(riga)
                        p.runs[0].font.size = Pt(10)
                        p.paragraph_format.space_after = Pt(4)
                        for passaggio in o.get("passaggi") or []:
                            pb = doc.add_paragraph(style="List Bullet")
                            pb.add_run(str(passaggio)).font.size = Pt(9)
                            pb.paragraph_format.space_after = Pt(1)
                continue

            if testo and testo != "N/D":
                add_heading(label, level=2)
                p = doc.add_paragraph(testo)
                p.runs[0].font.size = Pt(10)
                p.paragraph_format.space_after = Pt(6)

        # ── Fonti dei dati ────────────────────────────────────────────────────
        fonti = report.get("fonti") or {}
        if fonti:
            esiti = {e.get("campo"): e.get("esito") for e in report.get("_verifica") or []}
            etichette_esito = {"ok": "confermato", "corretto": "corretto in verifica",
                               "non_verificabile": "non verificabile", "da_verificare": "da verificare"}
            righe_fonti = []
            for campo in be.KPI_TUTTI:
                fonte = fonti.get(campo) or {}
                if not (fonte.get("pagina") or fonte.get("voce") or fonte.get("calcolo")):
                    continue
                righe_fonti.append((
                    be.ETICHETTE_KPI[campo],
                    str(fonte.get("pagina") or "—"),
                    str(fonte.get("voce") or "—"),
                    str(fonte.get("calcolo") or fonte.get("letto") or "—"),
                    etichette_esito.get(esiti.get(campo), ""),
                ))
            if righe_fonti:
                add_heading("Fonti dei dati", level=2)
                tabella = doc.add_table(rows=1, cols=5)
                tabella.style = "Table Grid"
                tabella.alignment = WD_TABLE_ALIGNMENT.LEFT
                for c, titolo in enumerate(["Dato", "Pagina PDF", "Voce letta nel bilancio",
                                            "Valore letto / calcolo", "Controllo"]):
                    cella = tabella.rows[0].cells[c]
                    cella.text = titolo
                    cella.paragraphs[0].runs[0].font.bold = True
                    cella.paragraphs[0].runs[0].font.size = Pt(8)
                    cella.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
                for valori in righe_fonti:
                    celle = tabella.add_row().cells
                    for c, valore in enumerate(valori):
                        celle[c].text = valore
                        if celle[c].paragraphs[0].runs:
                            celle[c].paragraphs[0].runs[0].font.size = Pt(8)
                estr = report.get("_estrazione") or {}
                if estr:
                    p = doc.add_paragraph(
                        f"Pagine del PDF analizzate: {_intervalli(estr.get('pagine_inviate', []))} "
                        f"su {estr.get('pagine_totali', '?')}. I numeri di pagina sono quelli del file PDF, "
                        f"non quelli stampati sul documento."
                    )
                    p.runs[0].font.size = Pt(8)
                    p.runs[0].font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    # ── Footer nota metodologica ──────────────────────────────────────────────
    doc.add_page_break()
    add_heading("Note Metodologiche", level=1, color_hex="999999")
    p = doc.add_paragraph(
        "I dati finanziari sono stati estratti automaticamente dal bilancio tramite "
        "analisi AI (Claude). I KPI sono stati normalizzati in euro. "
        "Si raccomanda di verificare i valori rispetto al documento originale "
        "prima di qualsiasi utilizzo professionale. "
        f"Report generato il {datetime.now().strftime('%d/%m/%Y alle %H:%M')}."
    )
    p.runs[0].font.size = Pt(9)
    p.runs[0].font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# BENCHMARK DI SETTORE (statico)
# ══════════════════════════════════════════════════════════════════════════════

BENCHMARK_SETTORE = {
    "Farmaceutico / Healthcare": {
        "margine_ebitda": (0.18, 0.28),
        "margine_utile":  (0.08, 0.15),
        "leva":           (1.0,  3.0),
    },
    "Manifatturiero": {
        "margine_ebitda": (0.08, 0.16),
        "margine_utile":  (0.03, 0.08),
        "leva":           (1.5,  4.0),
    },
    "Retail / Grande Distribuzione": {
        "margine_ebitda": (0.04, 0.10),
        "margine_utile":  (0.01, 0.04),
        "leva":           (2.0,  5.0),
    },
    "Technology / Software": {
        "margine_ebitda": (0.15, 0.35),
        "margine_utile":  (0.08, 0.20),
        "leva":           (0.0,  2.0),
    },
    "Energia / Utilities": {
        "margine_ebitda": (0.20, 0.40),
        "margine_utile":  (0.05, 0.12),
        "leva":           (3.0,  6.0),
    },
    "Costruzioni / Real Estate": {
        "margine_ebitda": (0.08, 0.18),
        "margine_utile":  (0.03, 0.08),
        "leva":           (2.0,  6.0),
    },
    "Alimentare / Bevande": {
        "margine_ebitda": (0.10, 0.20),
        "margine_utile":  (0.04, 0.10),
        "leva":           (1.0,  3.5),
    },
}

def _benchmark_cell(valore, bm_range):
    """Restituisce (testo, colore) per una cella benchmark."""
    if valore is None or bm_range is None:
        return "—", "#666"
    lo, hi = bm_range
    if lo <= valore <= hi:
        return "✅ In range", "#4caf50"
    elif valore < lo:
        diff = (lo - valore) / lo * 100
        return f"🔴 -{diff:.0f}% vs mediana", "#ef5350"
    else:
        diff = (valore - hi) / hi * 100
        return f"🟡 +{diff:.0f}% vs mediana", "#ffa726"

# ── Navigazione ───────────────────────────────────────────────────────────────

if "pagina" not in st.session_state:
    st.session_state["pagina"] = "genera"

col_nav1, col_nav2, col_nav3, col_nav4 = st.columns(4)
with col_nav1:
    if st.button("➕ Genera nuovo report", use_container_width=True):
        st.session_state["pagina"] = "genera"
with col_nav2:
    if st.button("🗂️ Archivio report", use_container_width=True):
        st.session_state["pagina"] = "archivio"
with col_nav3:
    if st.button("📊 Confronta aziende", use_container_width=True):
        st.session_state["pagina"] = "archivio"  # apre archivio sul tab confronta
with col_nav4:
    if st.button("📄 Carica report esistente", use_container_width=True):
        st.session_state["pagina"] = "carica_esistente"

st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# PAGINA: GENERA REPORT
# ══════════════════════════════════════════════════════════════════════════════

if st.session_state["pagina"] == "genera":

    col_l1, col_l2 = st.columns(2)
    with col_l1:
        lingua = st.radio("🌐 Lingua del report", ["Italiano", "English"], horizontal=True)
    with col_l2:
        dettaglio_scelto = st.radio(
            "📏 Livello di dettaglio", list(be.DETTAGLI.keys()),
            format_func=lambda k: be.DETTAGLI[k], horizontal=True, key="opt_dettaglio",
            help="Sintetico: report di circa due pagine con l'essenziale. "
                 "Completo: tutte le informazioni trovate nei documenti.",
        )

    # ── SEZIONE BILANCI (multi-upload) ────────────────────────────────────────

    st.markdown("### 📂 Carica i bilanci")
    st.caption(
        "Puoi caricare fino a 5 bilanci PDF completi contemporaneamente "
        f"(massimo {be.DIMENSIONE_MAX_MB} MB ciascuno). "
        "Il sistema trova automaticamente le pagine rilevanti — "
        "non serve estrarre le pagine manualmente."
    )

    bilanci_files = st.file_uploader(
        "Seleziona uno o più bilanci PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key="multi_bilanci",
        label_visibility="collapsed",
    )

    with st.expander("⚙️ Opzioni di analisi", expanded=False):
        col_o1, col_o2, col_o3 = st.columns(3)
        with col_o1:
            perimetro_scelto = st.selectbox(
                "Dati da usare",
                options=list(be.PERIMETRI.keys()),
                format_func=lambda k: be.PERIMETRI[k],
                key="opt_perimetro",
                help="Se il PDF contiene sia il consolidato sia il bilancio d'esercizio, "
                     "indica quale usare. Tutti i numeri del report useranno lo stesso perimetro.",
            )
        with col_o2:
            pagine_max = st.selectbox(
                "Pagine massime per bilancio",
                options=[None, 25, 40, 60, 100],
                format_func=lambda v: "Automatico (60, oppure 100 oltre le 250 pagine)" if v is None else str(v),
                index=0,
                key="opt_pagine_max_v2",
                help="Più pagine = meno dati mancanti, ma analisi più lenta e costosa. "
                     "Se l'account API non regge la richiesta, il sistema riduce le pagine da solo.",
            )
        with col_o3:
            verifica_attiva = st.checkbox(
                "Seconda lettura di controllo",
                value=True,
                key="opt_verifica",
                help="Dopo l'estrazione, Claude rilegge il bilancio e ricontrolla ogni numero "
                     "(colonna dell'anno, perimetro, unità, voce). Consigliata.",
            )

    # Form metadati + analisi per ogni bilancio caricato
    bilanci_pronti = []   # lista di dict pronti per Claude
    chiavi_attive = set()
    anni_disponibili = [str(y) for y in range(datetime.now().year, 2017, -1)]

    if bilanci_files:
        if len(bilanci_files) > 5:
            st.error("Massimo 5 bilanci alla volta.")
            bilanci_files = bilanci_files[:5]

        st.markdown("#### Conferma i dettagli")

        for idx, f in enumerate(bilanci_files):
            with st.container():
                st.markdown(f"**📄 {f.name}**")
                col_m1, col_m2, col_m3 = st.columns([3, 1, 2])
                with col_m1:
                    rs = st.text_input(
                        "Ragione sociale",
                        key=f"rs_{idx}",
                        placeholder="es. Zambon S.p.A.",
                    )
                with col_m2:
                    anno = st.selectbox(
                        "Anno bilancio",
                        options=anni_disponibili,
                        index=1 if len(anni_disponibili) > 1 else 0,
                        key=f"anno_{idx}",
                    )
                with col_m3:
                    pagine_manuali = st.text_input(
                        "Pagine da analizzare (opzionale)",
                        key=f"pagine_{idx}",
                        placeholder="es. 2-9, 105-126",
                        help="Lascia vuoto per la scelta automatica. Compila solo se vuoi "
                             "decidere tu quali pagine del PDF inviare.",
                    )

                # Analisi pagine: una volta sola per file e opzioni (cache di sessione).
                # In memoria resta solo il PDF ridotto, non l'originale.
                cache_key = "pdf_analysis_" + hashlib.md5(
                    f"{f.name}|{f.size}|{perimetro_scelto}|{pagine_max}|{pagine_manuali.strip()}".encode()
                ).hexdigest()
                chiavi_attive.add(cache_key)

                if cache_key not in st.session_state:
                    with st.spinner(f"Analisi pagine di {f.name}..."):
                        try:
                            raw_bytes = f.getvalue()
                            surgical_bytes, info = be.prepara_bilancio(
                                raw_bytes, perimetro_scelto, pagine_max, pagine_manuali
                            )
                            del raw_bytes
                            st.session_state[cache_key] = {
                                "surgical": surgical_bytes,
                                "info": info,
                            }
                        except (be.ErroreEstrazione, ValueError) as e:
                            st.error(f"❌ {f.name}: {e}")
                            st.markdown("---")
                            continue
                        except Exception as e:
                            st.error(f"❌ Impossibile leggere {f.name}: {type(e).__name__}: {e}")
                            st.markdown("---")
                            continue

                surgical_bytes = st.session_state[cache_key]["surgical"]
                info = st.session_state[cache_key]["info"]

                # Badge risultato analisi
                mode = info.get("mode", "text")
                total = info.get("total_pages", 0)
                selected = len(info.get("selected_pages", []))
                sections = ", ".join(info.get("sections_found", [])) or "—"
                riduzione = info.get("riduzione_pct", 0)
                elenco_pagine = _intervalli([p + 1 for p in info.get("selected_pages", [])])

                if mode == "scanned":
                    st.markdown(
                        f'<div class="pdf-badge pdf-badge-scan">'
                        f'📷 PDF scansionato — invio {selected} pagine su {total} '
                        f'({info.get("criterio", "")}). Pagine: {elenco_pagine}'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                    if selected < total:
                        st.caption(
                            "Nei PDF scansionati il testo non è leggibile, quindi le pagine sono scelte "
                            "per posizione. Se conosci le pagine giuste, indicale nel campo qui sopra."
                        )
                elif mode == "manual":
                    st.markdown(
                        f'<div class="pdf-badge pdf-badge-ok">'
                        f'✅ {selected} pagine indicate a mano su {total} totali. Pagine: {elenco_pagine}'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                elif mode == "fallback":
                    st.markdown(
                        f'<div class="pdf-badge pdf-badge-warn">'
                        f'⚠️ Layout atipico — invio {selected} pagine centrali su {total} totali. '
                        f'Conviene indicare le pagine a mano.'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    if total > 300 and info.get("pagine_max", 100) < 100:
                        st.caption(
                            f"Documento molto lungo ({total} pagine): per ridurre i dati mancanti puoi "
                            "portare a 100 le pagine massime nelle Opzioni di analisi."
                        )
                    perimetro_txt = ""
                    if info.get("ha_consolidato") and info.get("ha_esercizio"):
                        perimetro_txt = f' — Il PDF contiene consolidato ed esercizio: privilegiato il {info.get("perimetro")}'
                    st.markdown(
                        f'<div class="pdf-badge pdf-badge-ok">'
                        f'✅ {selected} pagine rilevanti su {total} totali '
                        f'({riduzione}% riduzione) — Sezioni: {sections}{perimetro_txt}'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

                # Pannello di dettaglio opzionale
                with st.expander("🔍 Dettaglio analisi pagine", expanded=False):
                    st.caption(
                        f"Pagine inviate: {elenco_pagine} · "
                        f"Caratteri medi per pagina: {info.get('avg_chars', 0):.0f} · "
                        f"Pagine scansionate: {info.get('scanned_pages', 0)} · "
                        f"PDF: {info.get('mb_originale', '?')} MB → {info.get('mb_inviati', '?')} MB"
                    )
                    scores = info.get("page_scores", [])
                    categorie = info.get("page_categories", [])
                    selected_set = set(info.get("selected_pages", []))
                    if scores and any(scores):
                        sigle = {"prospetti": "PRO", "pfn": "PFN", "debito": "DEB",
                                 "ownership": "SOC", "ma": "M&A", "overview": "ATT"}
                        celle = ""
                        for i, score in enumerate(scores[:300]):
                            color = "#c8e04a" if i in selected_set else "#555"
                            sigla = sigle.get(categorie[i], "") if i < len(categorie) else ""
                            celle += (
                                f"<div style='display:inline-block;width:46px;text-align:center;"
                                f"font-size:10px;color:{color};border:1px solid {color};"
                                f"border-radius:4px;padding:3px;margin:2px;'>"
                                f"<b>{i+1}</b><br>{sigla or '·'} {score}</div>"
                            )
                        st.markdown(celle, unsafe_allow_html=True)
                        st.caption(
                            "In verde le pagine inviate. PRO = prospetti contabili, PFN = posizione "
                            "finanziaria, DEB = note sul debito, SOC = soci, M&A = operazioni, "
                            "ATT = attività."
                            + (f" Mostrate le prime 300 di {len(scores)} pagine." if len(scores) > 300 else "")
                        )

                if rs.strip():
                    bilanci_pronti.append({
                        "ragione_sociale": rs.strip(),
                        "anno": anno,
                        "pdf_bytes": surgical_bytes,
                        "filename": f.name,
                        "info": info,
                    })
                else:
                    st.warning("⬆️ Inserisci la ragione sociale per includere questo bilancio.")

                st.markdown("---")

    # Libera la memoria delle analisi di file non più caricati (o con opzioni cambiate)
    for chiave in [k for k in list(st.session_state.keys())
                   if str(k).startswith("pdf_analysis_") and k not in chiavi_attive]:
        del st.session_state[chiave]

    # ── ALTRI DOCUMENTI (invariati) ───────────────────────────────────────────

    st.markdown("### 📎 Altri documenti (opzionale)")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.caption("📈 Export Mergermarket")
        mergermarket = st.file_uploader(
            "Mergermarket", type=["xlsx", "csv", "pdf"],
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

    # Raccolta testi documenti supplementari.
    # Ogni documento viene letto una sola volta per sessione (cache): prima la
    # visura veniva rimandata a Claude a ogni clic sulla pagina.
    testi_supplementari = {}
    documenti_binari = {}

    if mergermarket:
        contenuto = mergermarket.getvalue()

        def _leggi_mergermarket(contenuto=contenuto, nome_file=mergermarket.name):
            return be.leggi_mergermarket(nome_file, contenuto)

        try:
            testi_supplementari["Mergermarket"] = _cache_sessione(
                "suppl_merger", contenuto, _leggi_mergermarket
            )
            documenti_binari[mergermarket.name] = contenuto
            st.success("✅ Mergermarket caricato")
        except Exception as e:
            st.error(f"❌ Impossibile leggere l'export Mergermarket: {e}")

    if url_azienda:
        testi_supplementari["Sito Aziendale"] = _cache_sessione(
            "suppl_url", url_azienda.strip().encode(), lambda: estrai_testo_url(url_azienda.strip())
        )
        st.success("✅ URL acquisito")

    if visura:
        contenuto = visura.getvalue()

        def _leggi_visura(contenuto=contenuto):
            # Visura con testo: si tengono le sezioni utili (sede, statuto, capitale,
            # soci e patti, amministratori, sindaci, operazioni straordinarie).
            testo = be.leggi_visura(contenuto)
            if testo:
                return testo
            # Visura scansionata: la leggono le prime 40 pagine via Claude.
            try:
                n_pagine = len(PdfReader(io.BytesIO(contenuto)).pages)
                estratto = be.costruisci_pdf_chirurgico(contenuto, list(range(min(n_pagine, 40))))
                client_temp = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
                return be.testo_da_modello_veloce(client_temp, [
                    {"type": "document", "source": {
                        "type": "base64", "media_type": "application/pdf",
                        "data": base64.b64encode(estratto).decode()
                    }},
                    {"type": "text", "text": "Trascrivi fedelmente tutte le informazioni rilevanti "
                     "della visura, senza interpretarle e senza aggiungere nulla: ragione sociale, "
                     "forma giuridica, sede, codice fiscale, data di costituzione, capitale sociale, "
                     "soci con quote e patti parasociali, forma di amministrazione esattamente come "
                     "indicata, amministratori e cariche, sindaci, oggetto sociale, data di estrazione "
                     "della visura."}
                ])
            except Exception:
                reader = PyPDF2.PdfReader(io.BytesIO(contenuto))
                return "".join(p.extract_text() or "" for p in reader.pages)

        try:
            with st.spinner("Lettura della visura..."):
                testi_supplementari["Visura Camerale"] = _cache_sessione(
                    "suppl_visura", contenuto, _leggi_visura
                )
            documenti_binari[visura.name] = contenuto
            st.success("✅ Visura caricata")
        except Exception as e:
            st.error(f"❌ Impossibile leggere la visura: {e}")

    # ── PULSANTE GENERA ───────────────────────────────────────────────────────

    st.markdown("---")

    # Abilita il bottone solo se c'è almeno un bilancio pronto
    # oppure almeno un documento supplementare
    ha_contenuto = len(bilanci_pronti) > 0 or len(testi_supplementari) > 0

    if not ha_contenuto:
        st.info("Carica almeno un bilancio o documento per procedere.")

    # Senza bilanci il nome dell'azienda va chiesto PRIMA del pulsante
    # (dentro il pulsante il campo spariva al primo clic).
    nome_solo_suppl = ""
    if not bilanci_pronti and testi_supplementari:
        nome_solo_suppl = st.text_input(
            "Nome dell'azienda", placeholder="es. Eco Eridania S.p.A.",
            key="nome_solo_suppl"
        )

    if st.button("🚀 Genera Report", disabled=not ha_contenuto):

        lingua_prompt = "in inglese" if lingua == "English" else "in italiano"
        client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

        # Reset lista report per questa nuova generazione
        st.session_state["reports_generati"] = []

        # ── Caso A: uno o più bilanci caricati → report per ciascuno ─────────
        if bilanci_pronti:
            totale = len(bilanci_pronti)
            generati = 0
            progress = st.progress(0, text="Preparazione...")

            for idx, item in enumerate(bilanci_pronti):
                nome = item["ragione_sociale"]
                anno = item["anno"]
                info = item["info"]

                if idx > 0:
                    time.sleep(5)  # pausa tra bilanci per evitare rate limit

                progress.progress(
                    idx / totale,
                    text=f"Analisi {idx+1}/{totale}: {nome} ({anno})... può richiedere alcuni minuti"
                )

                # Testo supplementare per questo report
                testo_suppl = be.componi_supplementari(testi_supplementari)

                stato = st.empty()
                try:
                    report = be.genera_report_bilancio(
                        client, item["pdf_bytes"], info, nome, anno, lingua_prompt,
                        testo_suppl=testo_suppl,
                        fonti_supplementari=list(testi_supplementari.keys()),
                        perimetro=perimetro_scelto,
                        verifica=verifica_attiva,
                        dettaglio=dettaglio_scelto,
                        on_evento=lambda msg, stato=stato: stato.info(msg),
                    )

                    # Normalizza KPI a euro interi, calcola la leva e valida
                    report = normalizza_kpi(report)
                    be.normalizza_precedente(report)
                    be.completa_leva(report)
                    report["_validation_warnings"] = valida_kpi(report) + be.controlli_fonti(report)

                    salva_report(nome, report, documenti_binari)

                    st.session_state["reports_generati"].append({
                        "nome": nome,
                        "anno": anno,
                        "report": report,
                    })
                    # Compatibilità con il codice esistente
                    st.session_state["report"] = report
                    generati += 1

                except be.ErroreEstrazione as e:
                    st.error(f"❌ {nome}: {e}")
                except Exception as e:
                    st.error(f"❌ Errore per {nome}: {type(e).__name__}: {e}")
                finally:
                    stato.empty()

            progress.progress(1.0, text="Completato!")
            if generati:
                st.success(f"✅ {generati} report su {totale} generati e salvati!")

        # ── Caso B: solo documenti supplementari ──────────────────────────────
        else:
            nome_azienda = nome_solo_suppl.strip()
            if not nome_azienda:
                st.warning("Inserisci il nome dell'azienda.")
            else:
                with st.spinner("Claude sta analizzando i documenti..."):
                    testo_completo = ""
                    for fonte, testo in testi_supplementari.items():
                        testo_completo += f"\n\n--- FONTE: {fonte} ---\n{testo[:be.LIMITI_SUPPLEMENTARI.get(fonte, 10000)]}"

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

Usa SOLO informazioni scritte nei documenti qui sopra: nessuna conoscenza esterna,
nessuna supposizione. Riporta denominazioni, quote, date e forme di amministrazione
come sono scritte. Se un dato non è disponibile scrivi N/D. Non inventare dati."""

                    try:
                        risposta, _modello, _uso = be.chiama_claude(
                            client, [{"type": "text", "text": prompt2}], 8000,
                            modelli=be.MODELLI_VELOCI,
                        )
                        report = be.estrai_json(risposta)
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
            imposta_formato(report)
            fin = report.get("dati_finanziari", {})

            st.markdown(f"## 📋 {nome_r} — {anno_r}")

            # Pannello validazione
            kpi_warnings = report.get("_validation_warnings", [])
            render_pannello_validazione(kpi_warnings)
            render_fonti_verifica(report)

            # Calcolo margine EBITDA
            r_val = fin.get("ricavi")
            e_val = fin.get("ebitda")
            margine_str = "N/D"
            if r_val and e_val and report.get("_normalizzato"):
                try:
                    margine_str = fmt_pct(float(e_val)/float(r_val)*100)
                except:
                    pass

            st.markdown(f"""
            <div class="kpi-grid">
                <div class="kpi-card"><div class="kpi-label">Ricavi</div><div class="kpi-value">{fmt_kpi(report,'ricavi')}</div></div>
                <div class="kpi-card"><div class="kpi-label">EBITDA</div><div class="kpi-value">{fmt_kpi(report,'ebitda')}</div></div>
                <div class="kpi-card"><div class="kpi-label">Margine EBITDA</div><div class="kpi-value">{margine_str}</div></div>
                <div class="kpi-card"><div class="kpi-label">Utile Netto</div><div class="kpi-value">{fmt_kpi(report,'utile_netto')}</div></div>
                <div class="kpi-card"><div class="kpi-label">Totale Attivo</div><div class="kpi-value">{fmt_kpi(report,'totale_attivo')}</div></div>
                <div class="kpi-card"><div class="kpi-label">Patrimonio Netto</div><div class="kpi-value">{fmt_kpi(report,'patrimonio_netto')}</div></div>
                <div class="kpi-card"><div class="kpi-label">Anno Rif.</div><div class="kpi-value">{fin.get('anno_riferimento','N/D')}</div></div>
            </div>
            """, unsafe_allow_html=True)


            righe_cfr = confronto_esercizi(report)
            if righe_cfr:
                with st.expander(f"📈 Confronto con l'esercizio precedente ({anno_precedente(report, anno_r)})", expanded=True):
                    st.dataframe(
                        [{"Dato": e, str(anno_r or "Corrente"): c, anno_precedente(report, anno_r): p, "Variazione": v}
                         for e, c, p, v in righe_cfr],
                        hide_index=True, use_container_width=True,
                    )

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
                        <div class="kpi-card"><div class="kpi-label">Indebitamento Totale</div><div class="kpi-value">{fmt_kpi(report,'indebitamento_totale','struttura_debito')}</div></div>
                        <div class="kpi-card"><div class="kpi-label">Debito Bancario</div><div class="kpi-value">{fmt_kpi(report,'debito_bancario','struttura_debito')}</div></div>
                        <div class="kpi-card"><div class="kpi-label">Obbligazioni</div><div class="kpi-value">{fmt_kpi(report,'obbligazioni','struttura_debito')}</div></div>
                        <div class="kpi-card"><div class="kpi-label">Debito Netto</div><div class="kpi-value">{fmt_kpi(report,'debito_netto','struttura_debito')}</div></div>
                        <div class="kpi-card"><div class="kpi-label">Leva (PFN/EBITDA)</div><div class="kpi-value">{fmt_leva(debito)}</div></div>
                    </div>
                    """, unsafe_allow_html=True)
                    if debito.get('scadenze_principali', 'N/D') != 'N/D':
                        st.markdown(f'<div class="section-box"><div class="section-title">Scadenze Principali</div><div class="section-text">{debito.get("scadenze_principali","N/D")}</div></div>', unsafe_allow_html=True)
                    if voci_note(debito.get('note')):
                        st.markdown('<div class="section-box"><div class="section-title">Note</div><div class="section-text">'
                                    + "".join(f"• {html.escape(v)}<br>" for v in voci_note(debito.get('note')))
                                    + '</div></div>', unsafe_allow_html=True)
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
                            {('<div class="ma-desc" style="margin-top:6px;font-size:13px;">' + ''.join('• ' + html.escape(str(x)) + '<br>' for x in (op.get('passaggi') or [])) + '</div>') if op.get('passaggi') else ''}
                            {('<div class="ma-tipo">Fonte: ' + html.escape(str(op.get('fonte'))) + '</div>') if op.get('fonte') else ''}
                        </div>""", unsafe_allow_html=True)
                else:
                    st.write("Nessuna operazione rilevata.")

            with st.expander("📝 Note Aggiuntive"):
                st.markdown(f'<div class="section-box"><div class="section-text">{report.get("note_aggiuntive","N/D")}</div></div>', unsafe_allow_html=True)

            # Confronto YoY se ci sono più report dello stesso soggetto
            reports_stesso_soggetto = [
                r for r in st.session_state.get("reports_generati", [])
                if r["nome"] == nome_r and r["anno"] != anno_r
            ]
            if reports_stesso_soggetto:
                yoy_warnings = []
                for r_prec in reports_stesso_soggetto:
                    yoy_warnings += confronta_anni(report, r_prec["report"])
                if yoy_warnings:
                    with st.expander("📊 Variazioni YoY anomale", expanded=True):
                        render_pannello_validazione(yoy_warnings)

            st.markdown("---")

    # ── Bottoni download (se ci sono report generati) ─────────────────────────
    if st.session_state.get("reports_generati"):
        reports_list = st.session_state["reports_generati"]
        st.markdown("### 💾 Esporta")
        col_xl, col_wd = st.columns(2)

        with col_xl:
            try:
                excel_bytes = genera_excel(reports_list)
                nome_file_xl = f"TaxiReport_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                st.download_button(
                    label="📊 Scarica Excel",
                    data=excel_bytes,
                    file_name=nome_file_xl,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Errore generazione Excel: {e}")

        with col_wd:
            try:
                word_bytes = genera_word(reports_list)
                nome_file_wd = f"TaxiReport_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                st.download_button(
                    label="📄 Scarica Word",
                    data=word_bytes,
                    file_name=nome_file_wd,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Errore generazione Word: {e}")

        st.markdown("---")

    # Compatibilità con singolo report (aperto dall'archivio o caso B)
    elif "report" in st.session_state and "reports_generati" not in st.session_state:
        report = st.session_state["report"]
        imposta_formato(report)
        # Normalizza se necessario (report vecchi pre-M2)
        if not report.get("_normalizzato"):
            report = normalizza_kpi(report)
            st.session_state["report"] = report
        fin    = report.get("dati_finanziari", {})
        debito = report.get("struttura_debito", {})

        st.markdown(f"## 📋 {report.get('nome_azienda', '')}")

        # Pannello validazione
        kpi_warnings = report.get("_validation_warnings") or valida_kpi(report)
        render_pannello_validazione(kpi_warnings)
        render_fonti_verifica(report)

        # Margine EBITDA
        r_val = fin.get("ricavi")
        e_val = fin.get("ebitda")
        margine_str = "N/D"
        if r_val and e_val:
            try:
                margine_str = fmt_pct(float(e_val)/float(r_val)*100)
            except:
                pass

        st.markdown(f"""
        <div class="kpi-grid">
            <div class="kpi-card"><div class="kpi-label">Ricavi</div><div class="kpi-value">{fmt_kpi(report,'ricavi')}</div></div>
            <div class="kpi-card"><div class="kpi-label">EBITDA</div><div class="kpi-value">{fmt_kpi(report,'ebitda')}</div></div>
            <div class="kpi-card"><div class="kpi-label">Margine EBITDA</div><div class="kpi-value">{margine_str}</div></div>
            <div class="kpi-card"><div class="kpi-label">Utile Netto</div><div class="kpi-value">{fmt_kpi(report,'utile_netto')}</div></div>
            <div class="kpi-card"><div class="kpi-label">Totale Attivo</div><div class="kpi-value">{fmt_kpi(report,'totale_attivo')}</div></div>
            <div class="kpi-card"><div class="kpi-label">Patrimonio Netto</div><div class="kpi-value">{fmt_kpi(report,'patrimonio_netto')}</div></div>
            <div class="kpi-card"><div class="kpi-label">Anno Rif.</div><div class="kpi-value">{fin.get('anno_riferimento','N/D')}</div></div>
        </div>
        """, unsafe_allow_html=True)

        with st.expander("🏢 Overview", expanded=True):
            st.markdown(f'<div class="section-box"><div class="section-title">Overview</div><div class="section-text">{report.get("overview","N/D")}</div></div>', unsafe_allow_html=True)

        with st.expander("⚙️ Core Business & Mercati"):
            st.markdown(f'<div class="section-box"><div class="section-title">Core Business</div><div class="section-text">{report.get("core_business","N/D")}</div></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="section-box"><div class="section-title">Mercati</div><div class="section-text">{report.get("mercati","N/D")}</div></div>', unsafe_allow_html=True)

        with st.expander("💰 Struttura del Debito"):
            if isinstance(debito, dict):
                st.markdown(f"""
                <div class="kpi-grid">
                    <div class="kpi-card"><div class="kpi-label">Indebitamento Totale</div><div class="kpi-value">{fmt_kpi(report,'indebitamento_totale','struttura_debito')}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Debito Bancario</div><div class="kpi-value">{fmt_kpi(report,'debito_bancario','struttura_debito')}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Obbligazioni</div><div class="kpi-value">{fmt_kpi(report,'obbligazioni','struttura_debito')}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Debito Netto</div><div class="kpi-value">{fmt_kpi(report,'debito_netto','struttura_debito')}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Leva (PFN/EBITDA)</div><div class="kpi-value">{fmt_leva(debito)}</div></div>
                </div>
                """, unsafe_allow_html=True)
                if debito.get('scadenze_principali', 'N/D') not in ('N/D', '', None):
                    st.markdown(f'<div class="section-box"><div class="section-title">Scadenze Principali</div><div class="section-text">{debito.get("scadenze_principali")}</div></div>', unsafe_allow_html=True)
                if voci_note(debito.get('note')):
                    st.markdown('<div class="section-box"><div class="section-title">Note</div><div class="section-text">'
                                + "".join(f"• {html.escape(v)}<br>" for v in voci_note(debito.get('note')))
                                + '</div></div>', unsafe_allow_html=True)
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
                for lbl, key in [("Altri Azionisti","altri_azionisti"),
                                  ("Struttura di Controllo","struttura_controllo"),
                                  ("Note","note")]:
                    val = ownership.get(key, "N/D")
                    if val not in ("N/D", "", None):
                        st.markdown(f'<div class="section-box"><div class="section-title">{lbl}</div><div class="section-text">{val}</div></div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="section-box"><div class="section-text">{ownership}</div></div>', unsafe_allow_html=True)

        with st.expander("🔀 Operazioni M&A"):
            operazioni = report.get("operazioni_ma", [])
            if operazioni and operazioni[0].get("descrizione","N/D") != "N/D":
                for op in operazioni:
                    st.markdown(f"""
                    <div class="ma-item">
                        <span class="ma-anno">{op.get('anno','')}</span>
                        <div class="ma-tipo">{op.get('tipo','')}</div>
                        <div class="ma-desc">{op.get('descrizione','')}</div>
                            {('<div class="ma-desc" style="margin-top:6px;font-size:13px;">' + ''.join('• ' + html.escape(str(x)) + '<br>' for x in (op.get('passaggi') or [])) + '</div>') if op.get('passaggi') else ''}
                            {('<div class="ma-tipo">Fonte: ' + html.escape(str(op.get('fonte'))) + '</div>') if op.get('fonte') else ''}
                    </div>""", unsafe_allow_html=True)
            else:
                st.write("Nessuna operazione rilevata.")

        with st.expander("📝 Note Aggiuntive"):
            st.markdown(f'<div class="section-box"><div class="section-text">{report.get("note_aggiuntive","N/D")}</div></div>', unsafe_allow_html=True)

        # Export anche per report singolo
        st.markdown("---")
        st.markdown("### 💾 Esporta")
        _exp_item = [{"nome": report.get("nome_azienda","Report"),
                      "anno": fin.get("anno_riferimento",""),
                      "report": report}]
        col_xl2, col_wd2 = st.columns(2)
        with col_xl2:
            try:
                st.download_button("📊 Scarica Excel", genera_excel(_exp_item),
                    file_name=f"TaxiReport_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            except Exception as e:
                st.error(f"Errore Excel: {e}")
        with col_wd2:
            try:
                st.download_button("📄 Scarica Word", genera_word(_exp_item),
                    file_name=f"TaxiReport_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True)
            except Exception as e:
                st.error(f"Errore Word: {e}")

# ══════════════════════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════════════════════
# PAGINA: ARCHIVIO
# ══════════════════════════════════════════════════════════════════════════════

elif st.session_state["pagina"] == "archivio":

    tab_archivio, tab_confronta = st.tabs(["🗂️ Archivio", "📊 Confronta aziende"])

    # ── TAB ARCHIVIO ──────────────────────────────────────────────────────────
    with tab_archivio:
        reports_arch = carica_archivio()

        if not reports_arch:
            st.info("Nessun report salvato. Genera il tuo primo report!")
        else:
            # Ricerca
            cerca = st.text_input("🔍 Cerca per nome azienda", placeholder="es. Angelini...",
                                  label_visibility="collapsed")
            if cerca:
                reports_arch = [x for x in reports_arch
                                if cerca.lower() in x.get("nome","").lower()
                                or cerca.lower() in x["report"].get("nome_azienda","").lower()]

            st.caption(f"{len(reports_arch)} report trovati")

            for item in reports_arch:
                r   = item["report"]
                fin = r.get("dati_finanziari", {})
                nome_az = r.get("nome_azienda", item.get("nome", "—"))
                anno_az = fin.get("anno_riferimento", "—")
                warnings_count = len(r.get("_validation_warnings", []))
                warn_badge = (f' <span style="color:#e65100;font-size:11px;">'
                              f'⚠️ {warnings_count} warning</span>'
                              if warnings_count else "")

                # Report con schema flessibile (caricati da docx)
                is_flex = "sezioni" in r

                if is_flex:
                    # Anteprima: mostra primi campi corti dalle prime sezioni
                    kpi_items_html = ""
                    for sez in r.get("sezioni", [])[:3]:
                        for campo in sez.get("campi", [])[:2]:
                            lbl = campo.get("label", "")
                            val = str(campo.get("valore", ""))[:40]
                            if len(val) <= 40:
                                kpi_items_html += f"""
                                <div class="archivio-kpi-item">
                                    <span class="archivio-kpi-label">{lbl}</span>
                                    <span class="archivio-kpi-val">{val}</span>
                                </div>"""
                    num_sezioni = len(r.get("sezioni", []))
                    st.markdown(f"""
                    <div class="archivio-card">
                        <div class="archivio-nome">{nome_az}{warn_badge}</div>
                        <div class="archivio-data">📅 {item["data"]} &nbsp;|&nbsp; 📄 Report caricato ({num_sezioni} sezioni)</div>
                        <div class="archivio-kpi">{kpi_items_html}</div>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="archivio-card">
                        <div class="archivio-nome">{nome_az}{warn_badge}</div>
                        <div class="archivio-data">📅 {item["data"]} &nbsp;|&nbsp; Anno bilancio: {anno_az}</div>
                        <div class="archivio-kpi">
                            <div class="archivio-kpi-item">
                                <span class="archivio-kpi-label">Ricavi</span>
                                <span class="archivio-kpi-val">{_fmt(fin.get("ricavi"))}</span>
                            </div>
                            <div class="archivio-kpi-item">
                                <span class="archivio-kpi-label">EBITDA</span>
                                <span class="archivio-kpi-val">{_fmt(fin.get("ebitda"))}</span>
                            </div>
                            <div class="archivio-kpi-item">
                                <span class="archivio-kpi-label">Utile Netto</span>
                                <span class="archivio-kpi-val">{_fmt(fin.get("utile_netto"))}</span>
                            </div>
                            <div class="archivio-kpi-item">
                                <span class="archivio-kpi-label">PN</span>
                                <span class="archivio-kpi-val">{_fmt(fin.get("patrimonio_netto"))}</span>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                col_a, col_b, col_c = st.columns([1, 1, 4])
                with col_a:
                    if st.button("📖 Apri", key=f"apri_{item['riga']}"):
                        if is_flex:
                            # Report con schema flessibile → apri in carica_esistente
                            st.session_state["report_caricato"] = r
                            st.session_state["report_caricato_nome"] = nome_az
                            st.session_state["report_caricato_file"] = ""
                            st.session_state.pop("reports_generati", None)
                            st.session_state["pagina"] = "carica_esistente"
                        else:
                            # Report con schema classico → apri in genera
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

    # ── TAB CONFRONTA ─────────────────────────────────────────────────────────
    with tab_confronta:
        reports_tutti = carica_archivio()

        if len(reports_tutti) < 2:
            st.info("Servono almeno 2 report in archivio per confrontare.")
        else:
            # Etichetta per ogni report (funziona con entrambi gli schemi)
            def _label_report(x):
                r = x["report"]
                nome = r.get("nome_azienda", x.get("nome", "—"))
                if "sezioni" in r:
                    return f"{nome} (report caricato)"
                anno = r.get("dati_finanziari", {}).get("anno_riferimento", "—")
                return f"{nome} ({anno})"

            opzioni = [_label_report(x) for x in reports_tutti]
            selezionati = st.multiselect(
                "Seleziona 2-5 aziende da confrontare",
                options=opzioni,
                max_selections=5,
            )

            # Benchmark settore (solo per report classici)
            settore_sel = st.selectbox(
                "Benchmark di settore (opzionale)",
                options=["— Nessun benchmark —"] + list(BENCHMARK_SETTORE.keys()),
            )
            bm = BENCHMARK_SETTORE.get(settore_sel) if settore_sel != "— Nessun benchmark —" else None

            if len(selezionati) < 2:
                st.caption("Seleziona almeno 2 aziende per visualizzare il confronto.")
            else:
                idx_sel = [opzioni.index(s) for s in selezionati]
                sel_reports = [reports_tutti[i] for i in idx_sel]

                # ── Funzione per appiattire qualsiasi report in sezioni unificate ──

                def _flatten_report(report):
                    """
                    Converte qualsiasi report (classico o flex) in una lista di
                    (titolo_sezione, [(label, valore), ...]) per il confronto.
                    """
                    sezioni = []

                    if "sezioni" in report:
                        # Schema flessibile (report caricato da docx)
                        for sez in report["sezioni"]:
                            titolo = sez.get("titolo_sezione", "")
                            campi = [(c["label"], str(c.get("valore", "")))
                                     for c in sez.get("campi", [])]
                            if campi:
                                sezioni.append((titolo, campi))
                    else:
                        # Schema classico (report da bilancio)
                        fin = report.get("dati_finanziari", {})
                        deb = report.get("struttura_debito", {})
                        own = report.get("ownership", {})

                        # General Information
                        info_campi = []
                        if report.get("nome_azienda"):
                            info_campi.append(("Company name", report["nome_azienda"]))
                        if report.get("overview") and report["overview"] != "N/D":
                            info_campi.append(("Overview", report["overview"]))
                        if report.get("core_business") and report["core_business"] != "N/D":
                            info_campi.append(("Core Business", report["core_business"]))
                        if report.get("mercati") and report["mercati"] != "N/D":
                            info_campi.append(("Markets", report["mercati"]))
                        if info_campi:
                            sezioni.append(("General Information", info_campi))

                        # Financial Data
                        fin_campi = []
                        for lbl, key in [
                            ("Ricavi / Revenue", "ricavi"), ("EBITDA", "ebitda"),
                            ("EBIT", "ebit"), ("Utile Netto / Net Result", "utile_netto"),
                            ("Ammortamenti", "ammortamenti"),
                            ("Totale Attivo / Total Assets", "totale_attivo"),
                            ("Patrimonio Netto / Shareholders' Equity", "patrimonio_netto"),
                            ("Cassa / Cash", "cassa"),
                        ]:
                            v = fin.get(key)
                            if v is not None:
                                fin_campi.append((lbl, _fmt(v) if report.get("_normalizzato") else str(v)))
                        if fin.get("anno_riferimento"):
                            fin_campi.append(("Anno riferimento", fin["anno_riferimento"]))
                        if fin_campi:
                            sezioni.append(("Financial Data", fin_campi))

                        # Debt Structure
                        deb_campi = []
                        for lbl, key in [
                            ("Indebitamento Totale", "indebitamento_totale"),
                            ("Debito Bancario", "debito_bancario"),
                            ("Obbligazioni", "obbligazioni"),
                            ("Debito Netto / Net Financial Position", "debito_netto"),
                            ("Leva Finanziaria", "leva_finanziaria"),
                        ]:
                            v = deb.get(key)
                            if v is not None:
                                if key == "leva_finanziaria":
                                    deb_campi.append((lbl, f"{v}x"))
                                else:
                                    deb_campi.append((lbl, _fmt(v) if report.get("_normalizzato") else str(v)))
                        if deb.get("scadenze_principali") and deb["scadenze_principali"] != "N/D":
                            deb_campi.append(("Scadenze principali", deb["scadenze_principali"]))
                        if deb.get("note") and deb["note"] != "N/D":
                            deb_campi.append(("Note debito", deb["note"]))
                        if deb_campi:
                            sezioni.append(("Debt Structure", deb_campi))

                        # Governance / Ownership
                        own_campi = []
                        if isinstance(own, dict):
                            for lbl, key in [
                                ("Azionista Principale / Main Shareholder", "azionista_principale"),
                                ("Quota", "quota_principale"),
                                ("Altri Azionisti", "altri_azionisti"),
                                ("Struttura Controllo", "struttura_controllo"),
                            ]:
                                v = own.get(key)
                                if v and v != "N/D":
                                    own_campi.append((lbl, str(v)))
                        if own_campi:
                            sezioni.append(("Governance", own_campi))

                        # M&A / Deals
                        ops = report.get("operazioni_ma", [])
                        ops_campi = []
                        for op in ops:
                            desc = op.get("descrizione", "N/D")
                            if desc and desc != "N/D":
                                anno_op = op.get("anno", "")
                                tipo_op = op.get("tipo", "")
                                label = f"{anno_op} — {tipo_op}" if anno_op else (tipo_op or "Transaction")
                                ops_campi.append((label, desc))
                        if ops_campi:
                            sezioni.append(("Deals", ops_campi))

                        # Note aggiuntive
                        if report.get("note_aggiuntive") and report["note_aggiuntive"] != "N/D":
                            sezioni.append(("Additional Notes", [("Note", report["note_aggiuntive"])]))

                    return sezioni

                # ── Costruzione indice unificato di tutte le sezioni e campi ──

                # Per ogni report, appiattisci
                flat_reports = [_flatten_report(item["report"]) for item in sel_reports]

                # Raccogli tutte le sezioni e i loro campi nell'ordine in cui appaiono
                sezioni_ordinate = []  # lista di titoli sezione in ordine
                campi_per_sezione = {}  # titolo → lista ordinata di label unici
                for flat in flat_reports:
                    for titolo, campi in flat:
                        if titolo not in campi_per_sezione:
                            sezioni_ordinate.append(titolo)
                            campi_per_sezione[titolo] = []
                        for lbl, _ in campi:
                            if lbl not in campi_per_sezione[titolo]:
                                campi_per_sezione[titolo].append(lbl)

                # Per lookup veloce: report_idx → {(sezione, label): valore}
                lookup = []
                for flat in flat_reports:
                    d = {}
                    for titolo, campi in flat:
                        for lbl, val in campi:
                            d[(titolo, lbl)] = val
                    lookup.append(d)

                # ── Rendering confronto ──────────────────────────────────────

                st.markdown("---")
                st.markdown("#### Tabella comparativa")

                for titolo_sez in sezioni_ordinate:
                    labels = campi_per_sezione[titolo_sez]
                    if not labels:
                        continue

                    st.markdown(
                        f'<div style="background:#2a2a2a;border-left:4px solid #c8e04a;'
                        f'padding:8px 14px;margin:16px 0 8px 0;border-radius:6px;">'
                        f'<span style="color:#c8e04a;font-weight:700;">{titolo_sez}</span></div>',
                        unsafe_allow_html=True
                    )

                    # Intestazione colonne
                    header_cols = st.columns([3] + [2] * len(sel_reports))
                    header_cols[0].markdown(
                        "<span style='color:#999;font-size:11px;text-transform:uppercase;'>Campo</span>",
                        unsafe_allow_html=True
                    )
                    for j, s in enumerate(selezionati):
                        header_cols[j+1].markdown(
                            f"<span style='color:#c8e04a;font-size:11px;font-weight:600;'>{s}</span>",
                            unsafe_allow_html=True
                        )
                    st.markdown('<hr style="border-color:#444;margin:2px 0 6px 0;">', unsafe_allow_html=True)

                    for lbl in labels:
                        row_cols = st.columns([3] + [2] * len(sel_reports))
                        row_cols[0].markdown(
                            f"<span style='color:#999;font-size:13px;'>{lbl}</span>",
                            unsafe_allow_html=True
                        )
                        for j in range(len(sel_reports)):
                            val = lookup[j].get((titolo_sez, lbl), "—")
                            # Tronca valori molto lunghi nel confronto
                            val_display = val if len(val) <= 120 else val[:117] + "..."
                            row_cols[j+1].markdown(
                                f"<span style='font-size:12px;color:#ddd;'>{val_display}</span>",
                                unsafe_allow_html=True
                            )
                        st.markdown('<hr style="border-color:#2a2a2a;margin:1px 0;">', unsafe_allow_html=True)

                st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# PAGINA: CARICA REPORT ESISTENTE
# ══════════════════════════════════════════════════════════════════════════════

elif st.session_state["pagina"] == "carica_esistente":

    st.markdown("### 📄 Carica Report Esistente")
    st.caption(
        "Carica un file Word (.docx) di un taxi report già compilato. "
        "Il sistema estrarrà e riprodurrà fedelmente tutte le sezioni del documento."
    )

    docx_file = st.file_uploader(
        "Seleziona un file Word (.docx) di un taxi report esistente",
        type=["docx"],
        accept_multiple_files=False,
        key="carica_report_docx",
        label_visibility="collapsed",
    )

    if docx_file:
        # ── Estrazione testo strutturato dal .docx ────────────────────────
        try:
            doc_bytes = docx_file.read()
            doc = Document(io.BytesIO(doc_bytes))

            # Estrai paragrafi
            paragrafi = [p.text for p in doc.paragraphs if p.text.strip()]

            # Estrai tabelle preservando struttura label → valore
            # Gestisce merged cells orizzontali e verticali:
            # - orizzontali: stessa cella ripetuta nella stessa riga → deduplica
            # - verticali: stessa cella ripetuta su righe consecutive → prendi solo
            #   le celle nuove dalla riga successiva
            testo_tabelle = []
            for table in doc.tables:
                contenuti_visti = set()  # traccia tutti i contenuti già emessi
                for row in table.rows:
                    # Deduplica celle orizzontali (merged)
                    celle_uniche = []
                    viste_riga = set()
                    for cell in row.cells:
                        t = cell.text.strip()
                        if t and t not in viste_riga:
                            celle_uniche.append(t)
                            viste_riga.add(t)
                    # Dalle celle uniche, prendi solo quelle mai viste prima
                    # (gestisce merged cells verticali)
                    celle_nuove = [c for c in celle_uniche if c not in contenuti_visti]
                    # Registra tutto come visto
                    for c in celle_uniche:
                        contenuti_visti.add(c)
                    # Emetti solo se ci sono celle nuove
                    if celle_nuove:
                        testo_tabelle.append(" | ".join(celle_nuove))

            testo_completo = "\n".join(paragrafi)
            if testo_tabelle:
                testo_completo += "\n\n--- TABELLE ---\n" + "\n".join(testo_tabelle)

        except Exception as e:
            testo_completo = None
            st.markdown(f"""
            <div class="section-box" style="border-left-color: #ef5350;">
                <div class="section-title" style="color: #ef5350;">
                    ❌ Errore lettura file
                </div>
                <div class="section-text">
                    Impossibile leggere il file Word: {type(e).__name__}: {e}<br>
                    Assicurati che il file sia un documento .docx valido e non corrotto.
                </div>
            </div>
            """, unsafe_allow_html=True)

        if testo_completo:
            with st.expander("📃 Anteprima testo estratto dal documento"):
                st.text(testo_completo[:5000] + ("..." if len(testo_completo) > 5000 else ""))
            st.caption(f"Estratti {len(paragrafi)} paragrafi e {len(doc.tables)} tabelle — {len(testo_completo)} caratteri totali")

            if st.button("🔍 Estrai dati dal report", use_container_width=True):
                client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

                prompt_estrazione = f"""You are a legal M&A analyst. You receive the full text extracted from a BD (Business Development) taxi report document.

Your task: extract EVERY piece of information from this document and reproduce it faithfully in a structured JSON.

--- DOCUMENT TEXT ---
{testo_completo[:20000]}
--- END ---

CRITICAL RULES:
1. DO NOT invent, summarize, or paraphrase. Copy the content exactly as written in the document.
2. DO NOT skip any section, field, or piece of information. If it's in the document, it must be in the JSON.
3. Preserve the ORIGINAL LANGUAGE of the document (English, Italian, etc.) — do not translate.
4. For each section you find, create an entry in the "sezioni" array with the exact section title and ALL its fields.
5. Each field has a "label" (the field name as it appears) and a "valore" (the full content, verbatim).
6. If a field contains multiple paragraphs or bullet points, include them all in "valore" separated by newlines.
7. NEVER repeat the same content in multiple fields. Each piece of information must appear EXACTLY ONCE. If the same text appears under multiple labels in the source (due to merged cells), include it only in the most specific/relevant field.
8. If a cell contains multiple sub-sections (e.g. "Rumored Transactions" and "Critical Issues / Outlook" in the same cell), split them into SEPARATE fields with their own label and value.

Respond ONLY with valid JSON, no backticks, no additional text.

{{
  "nome_azienda": "",
  "sezioni": [
    {{
      "titolo_sezione": "General Information",
      "campi": [
        {{"label": "Company name", "valore": "..."}},
        {{"label": "Business Description", "valore": "... full text ..."}},
        {{"label": "Headquarter", "valore": "..."}},
        {{"label": "...", "valore": "..."}}
      ]
    }},
    {{
      "titolo_sezione": "Financial Data",
      "campi": [
        {{"label": "...", "valore": "..."}},
        {{"label": "...", "valore": "..."}}
      ]
    }}
  ]
}}

Extract ALL sections you find. Common sections include (but are not limited to):
General Information, Financial Data, Deals/Transactions, Governance, Legal Advisors, etc.
If you find sections with different names, use those names exactly."""

                with st.spinner("Claude sta analizzando il documento..."):
                    for tentativo in range(3):
                        try:
                            risposta, _modello, _uso = be.chiama_claude(
                                client, [{"type": "text", "text": prompt_estrazione}], 32000
                            )
                            report = be.estrai_json(risposta)

                            nome_azienda = report.get("nome_azienda", docx_file.name)
                            st.session_state["report_caricato"] = report
                            st.session_state["report_caricato_nome"] = nome_azienda
                            st.session_state["report_caricato_file"] = docx_file.name
                            st.success("✅ Estrazione completata!")
                            break

                        except (json.JSONDecodeError, be.ErroreEstrazione) as e:
                            if tentativo < 2:
                                st.warning(f"Tentativo {tentativo+1}: risposta non valida, riprovo...")
                                time.sleep(5)
                            else:
                                st.markdown(f"""
                                <div class="section-box" style="border-left-color: #ef5350;">
                                    <div class="section-title" style="color: #ef5350;">
                                        ❌ Errore di estrazione
                                    </div>
                                    <div class="section-text">
                                        Impossibile interpretare la risposta AI dopo 3 tentativi.<br>
                                        Dettaglio: {e}
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)
                        except Exception as e:
                            if tentativo < 2:
                                st.warning(f"Tentativo {tentativo+1} fallito: {e}, riprovo...")
                                time.sleep(10)
                            else:
                                st.markdown(f"""
                                <div class="section-box" style="border-left-color: #ef5350;">
                                    <div class="section-title" style="color: #ef5350;">
                                        ❌ Errore
                                    </div>
                                    <div class="section-text">
                                        {type(e).__name__}: {e}
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)

    # ── VISUALIZZAZIONE REPORT ESTRATTO DA DOCX ───────────────────────────

    if st.session_state.get("report_caricato"):
        report = st.session_state["report_caricato"]
        nome_az = st.session_state.get("report_caricato_nome", "")
        file_origine = st.session_state.get("report_caricato_file", "")

        st.markdown("---")
        st.markdown(f"## 📋 {nome_az}")
        if file_origine:
            st.caption(f"Estratto da: {file_origine}")

        sezioni = report.get("sezioni", [])

        if not sezioni:
            st.warning("Nessuna sezione estratta dal documento.")
        else:
            # Icone per le sezioni comuni
            ICONE_SEZIONI = {
                "general information": "🏢",
                "informazioni generali": "🏢",
                "financial data": "💰",
                "dati finanziari": "💰",
                "financial data of the group": "💰",
                "deals": "🔀",
                "transactions": "🔀",
                "recent transactions": "🔀",
                "deals/transactions": "🔀",
                "governance": "👥",
                "ownership": "👥",
                "legal advisors": "⚖️",
                "current preferred legal advisor": "⚖️",
                "critical issues": "⚠️",
                "critical issues / outlook": "⚠️",
                "outlook": "📊",
                "real estate portfolio": "🏗️",
                "rumored transactions": "📰",
            }

            for idx_s, sezione in enumerate(sezioni):
                titolo = sezione.get("titolo_sezione", f"Sezione {idx_s+1}")
                campi = sezione.get("campi", [])
                icona = ICONE_SEZIONI.get(titolo.lower().strip(), "📌")

                with st.expander(f"{icona} {titolo}", expanded=(idx_s == 0)):
                    if not campi:
                        st.write("Nessun dato in questa sezione.")
                        continue

                    # Identifica campi "corti" (KPI-like) vs "lunghi" (testo)
                    campi_corti = []
                    campi_lunghi = []
                    for c in campi:
                        val = str(c.get("valore", ""))
                        if len(val) <= 80 and "\n" not in val:
                            campi_corti.append(c)
                        else:
                            campi_lunghi.append(c)

                    # Mostra campi corti come KPI cards
                    if campi_corti:
                        cards_html = ""
                        for c in campi_corti:
                            lbl = c.get("label", "")
                            val = str(c.get("valore", "N/D"))
                            cards_html += f"""
                            <div class="kpi-card">
                                <div class="kpi-label">{lbl}</div>
                                <div class="kpi-value" style="font-size:14px;">{val}</div>
                            </div>"""
                        st.markdown(f'<div class="kpi-grid">{cards_html}</div>',
                                    unsafe_allow_html=True)

                    # Mostra campi lunghi come section-box
                    for c in campi_lunghi:
                        lbl = c.get("label", "")
                        val = str(c.get("valore", ""))
                        # Converti newline in <br> per HTML
                        val_html = val.replace("\n", "<br>")
                        st.markdown(f"""
                        <div class="section-box">
                            <div class="section-title">{lbl}</div>
                            <div class="section-text">{val_html}</div>
                        </div>""", unsafe_allow_html=True)

        # ── SALVATAGGIO ───────────────────────────────────────────────────

        st.markdown("---")

        if st.button("💾 Salva su Google Sheets", use_container_width=True, key="salva_caricato"):
            salva_report(nome_az, report, {})
            st.success("✅ Report salvato nell'archivio Google Sheets!")
