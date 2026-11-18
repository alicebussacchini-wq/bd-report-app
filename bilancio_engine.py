"""
bilancio_engine.py — Motore di analisi dei bilanci per il Taxi Report Generator.

Non dipende da Streamlit: contiene solo la logica, così può essere provata
anche fuori dall'app.

Cosa fa, in ordine:
  1. legge il testo di ogni pagina del PDF con poca memoria (una pagina alla
     volta, liberandola subito);
  2. sceglie le pagine da inviare a Claude per categorie (prospetti, PFN,
     note sul debito, ownership, M&A, descrizione dell'attività), gestendo
     anche i fascicoli scansionati del Registro Imprese;
  3. costruisce il PDF ridotto;
  4. chiede a Claude i dati con la pagina e la voce di origine di ogni numero;
  5. fa fare a Claude una seconda lettura indipendente dei numeri;
  6. ricalcola nel codice ciò che è aritmetica (EBITDA civilistico, leva).
"""

import io
import json
import re
import threading
import time

from pypdf import PdfReader, PdfWriter

# ── Parametri ────────────────────────────────────────────────────────────────

# Modelli provati nell'ordine: se il primo non è disponibile per la chiave API
# si passa automaticamente al successivo.
MODELLI_ESTRAZIONE = ["claude-sonnet-5-5", "claude-sonnet-4-6"]
MODELLI_VELOCI = ["claude-haiku-5-5", "claude-haiku-4-5-20251001"]

# Comprendono anche i token di ragionamento dei modelli recenti: si paga solo l'usato.
MAX_TOKENS_ESTRAZIONE = 40000
MAX_TOKENS_VERIFICA = 24000

PAGINE_MAX_DEFAULT = 60
# Se la richiesta è troppo grande per i limiti dell'account si riprova con meno pagine.
GRADINI_PAGINE = [100, 60, 40, 25, 15]

DIMENSIONE_MAX_MB = 150          # oltre questa soglia il file viene rifiutato
SOGLIA_PAGINA_SCANSIONATA = 350  # caratteri: sotto, con un'immagine, è una scansione

KPI_FINANZIARI = [
    "ricavi", "ebitda", "ebit", "utile_netto", "ammortamenti",
    "totale_attivo", "patrimonio_netto", "cassa",
]
KPI_DEBITO = ["indebitamento_totale", "debito_bancario", "obbligazioni", "debito_netto"]
KPI_TUTTI = KPI_FINANZIARI + KPI_DEBITO

ETICHETTE_KPI = {
    "ricavi": "Ricavi", "ebitda": "EBITDA", "ebit": "EBIT",
    "utile_netto": "Utile netto", "ammortamenti": "Ammortamenti",
    "totale_attivo": "Totale attivo", "patrimonio_netto": "Patrimonio netto",
    "cassa": "Cassa", "indebitamento_totale": "Indebitamento totale",
    "debito_bancario": "Debito bancario", "obbligazioni": "Obbligazioni",
    "debito_netto": "Debito netto (PFN)", "leva_finanziaria": "Leva finanziaria",
}
ETICHETTE_KPI.update({f"precedente.{k}": f"{v} (esercizio precedente)"
                      for k, v in list(ETICHETTE_KPI.items())})

PERIMETRI = {
    "auto": "Automatico (consolidato se presente)",
    "consolidato": "Consolidato",
    "esercizio": "Esercizio / separato",
}


class RichiestaTroppoGrande(Exception):
    """La richiesta supera i limiti dell'account o del modello: servono meno pagine."""


class ErroreEstrazione(Exception):
    """Errore leggibile da mostrare all'utente."""


# ── 1. Lettura del testo pagina per pagina ───────────────────────────────────

# pdfium non è thread-safe e Streamlit serve ogni utente in un thread diverso.
_LOCK_PDFIUM = threading.Lock()


def leggi_pagine(pdf_bytes: bytes) -> list:
    """
    Restituisce una lista di dict {"testo", "caratteri", "immagine"} per pagina.

    Ogni pagina viene chiusa appena letta: la memoria resta di pochi MB anche
    con bilanci di centinaia di pagine. (La versione precedente teneva in
    memoria tutti i caratteri di tutte le pagine e con i bilanci grandi
    esauriva la memoria del server.)
    """
    try:
        return _leggi_pagine_pdfium(pdf_bytes)
    except Exception:
        return _leggi_pagine_pdfplumber(pdf_bytes)


def _leggi_pagine_pdfium(pdf_bytes: bytes) -> list:
    import pypdfium2 as pdfium

    try:
        import pypdfium2.raw as pdfium_raw
        tipo_immagine = pdfium_raw.FPDF_PAGEOBJ_IMAGE
    except Exception:
        tipo_immagine = None

    pagine = []
    with _LOCK_PDFIUM:
        pdf = pdfium.PdfDocument(pdf_bytes)
        try:
            for i in range(len(pdf)):
                page = pdf[i]
                testo = ""
                immagine = None
                try:
                    textpage = page.get_textpage()
                    try:
                        testo = textpage.get_text_range() or ""
                    finally:
                        textpage.close()
                    if tipo_immagine is not None and len(testo) < SOGLIA_PAGINA_SCANSIONATA:
                        try:
                            immagine = any(
                                True for _ in page.get_objects(
                                    filter=[tipo_immagine], max_depth=2
                                )
                            )
                        except Exception:
                            immagine = None
                finally:
                    page.close()
                if immagine is None:
                    immagine = len(testo) < SOGLIA_PAGINA_SCANSIONATA
                pagine.append({
                    "testo": testo, "caratteri": len(testo), "immagine": bool(immagine),
                })
        finally:
            pdf.close()
    return pagine


def _leggi_pagine_pdfplumber(pdf_bytes: bytes) -> list:
    """Riserva: pdfplumber, svuotando la cache di ogni pagina dopo l'uso."""
    import pdfplumber

    pagine = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            try:
                testo = page.extract_text() or ""
            except Exception:
                testo = ""
            try:
                immagine = len(testo) < SOGLIA_PAGINA_SCANSIONATA and bool(page.images)
            except Exception:
                immagine = len(testo) < SOGLIA_PAGINA_SCANSIONATA
            pagine.append({"testo": testo, "caratteri": len(testo), "immagine": immagine})
            try:
                page.flush_cache()
            except Exception:
                pass
    return pagine


# ── 2. Selezione delle pagine ────────────────────────────────────────────────

# Qualsiasi cifra isolata (anche "312" nei bilanci in migliaia), esclusi anni e date.
_RE_NUMERO = re.compile(r"(?<![\w/.,])\(?-?\d[\d.]*(?:,\d+)?\)?(?![\w/])")
_RE_ANNO = re.compile(r"^(?:19|20)\d{2}$")
_RE_CAPITOLO = re.compile(r"capitolo\s+(\d+)\s*[-–]\s*([a-zà-ù' ]{4,60})")
_RE_SPAZI = re.compile(r"\s+")

# (parola chiave, peso) per categoria. Il testo è già in minuscolo.
CATEGORIE = {
    "prospetti": [
        ("totale attivo", 12), ("totale passivo", 12),
        ("totale patrimonio netto e passiv", 12), ("totale passività e patrimonio netto", 12),
        ("differenza tra valore e costi della produzione", 12),
        ("stato patrimoniale", 6), ("situazione patrimoniale", 6),
        ("conto economico", 6), ("rendiconto finanziario", 6),
        ("prospetto delle variazioni", 4), ("valore della produzione", 5),
        ("ricavi delle vendite", 5), ("risultato operativo", 5),
        ("utile (perdita) dell'esercizio", 6), ("utile (perdita) del periodo", 6),
        ("risultato netto", 4), ("disponibilità liquide", 4),
        ("flusso finanziario", 5), ("flusso di cassa", 5),
        ("totale debiti", 8), ("patrimonio netto", 3),
        ("statement of financial position", 6), ("income statement", 6),
        ("total assets", 12), ("total equity and liabilities", 12),
        ("total liabilities and equity", 12), ("statement of cash flows", 6),
        ("cash flow statement", 6), ("statement of changes in equity", 4),
        ("operating profit", 5), ("operating result", 5),
        ("profit (loss) for the year", 6), ("core business revenue", 5),
        ("total revenue", 5), ("cash and cash equivalents", 4), ("net cash flows", 5),
        ("total equity", 4),
    ],
    "pfn": [
        ("posizione finanziaria netta", 12), ("indebitamento finanziario netto", 12),
        ("net financial position", 12), ("net financial debt", 10),
        ("indebitamento finanziario", 6), ("esma", 6), ("ebitda", 5),
        ("margine operativo lordo", 5), ("leverage", 5), ("covenant", 4),
        ("indicatori alternativi di performance", 6), ("dati di sintesi", 4),
        ("principali indicatori", 4), ("highlights", 3),
        ("net financial debt", 12), ("gross operating margin", 8), ("net debt", 6),
        ("free cash flow", 5), ("alternative performance", 6), ("credit rating", 8), ("rating agency", 6),
        ("key figures", 4),
    ],
    "debito": [
        ("debiti verso banche", 10), ("passività finanziarie", 9),
        ("debiti finanziari", 8), ("finanziament", 5),
        ("prestito obbligazionario", 9), ("prestiti obbligazionari", 9),
        ("obbligazion", 4), ("eurobond", 7), (" bond", 3), ("covenant", 7),
        ("scadenz", 4), ("oltre l'esercizio", 3), ("oltre 5 anni", 5),
        ("oltre cinque anni", 5), ("mutui", 4), ("linee di credito", 5),
        ("revolving", 5), ("passività per leasing", 6), ("tasso di interesse", 3),
        ("garanzie", 2), ("pegno", 3),
        ("financial liabilities", 9), ("bank loans", 8), ("borrowings", 5), ("bonds", 5),
        ("convertible bond", 7), ("maturit", 4), ("revolving credit", 7),
        ("credit lines", 5), ("lease liabilities", 5), ("interest rate", 3),
        ("notional", 3), ("future contractual payments", 6), ("liquidity risk", 5),
    ],
    "ownership": [
        ("compagine sociale", 10), ("azionist", 6), ("capitale sociale", 5),
        ("direzione e coordinamento", 9), ("controllante", 5), ("controllata da", 7),
        ("detenut", 3), ("soci ", 3), ("partecipazione di maggioranza", 6),
        ("patto parasociale", 6), ("azioni proprie", 3),
        ("consiglio di amministrazione", 3), ("organi sociali", 4),
        ("shareholders' agreement", 9), ("jointly control", 9),
        ("shareholding structure", 8), ("ownership structure", 8),
        ("major shareholders", 8), ("parent company", 5), ("share capital", 5),
        ("% owned", 3), ("treasury shares", 4), ("controlled by", 5),
        ("related party transactions", 5), ("other shareholders", 7),
    ],
    "ma": [
        ("acquisizion", 6), ("cession", 4), ("fusione", 6), ("scissione", 5),
        ("aggregazion", 7), ("business combination", 8), ("operazioni straordinarie", 8),
        ("fatti di rilievo", 9), ("eventi successivi", 7),
        ("dopo la chiusura dell'esercizio", 7), ("area di consolidamento", 6),
        ("perimetro di consolidamento", 6), ("closing", 4), ("signing", 4),
        ("opa ", 5), ("offerta pubblica", 6), ("conferimento", 4),
        ("acquisition", 5), ("merger", 6), ("disposal", 4), ("divestment", 5),
        ("sale of equity interest", 9), ("assets held for sale", 10),
        ("attività destinate alla vendita", 10),
        ("discontinued operations", 5), ("consolidation scope", 7),
        ("scope of consolidation", 7), ("subsequent events", 7), ("events after", 6),
        ("memorandum of understanding", 5), ("sold to", 4),
    ],
    "overview": [
        ("relazione sulla gestione", 8), ("attività del gruppo", 7),
        ("attività della società", 7), ("il gruppo opera", 8), ("la società opera", 8),
        ("settore", 2), ("mercat", 2), ("andamento della gestione", 6),
        ("profilo del gruppo", 8), ("chi siamo", 6), ("modello di business", 7),
        ("aree geografiche", 5), ("stabiliment", 3), ("dipendenti", 2),
        ("evoluzione prevedibile", 4), ("lettera agli azionisti", 4),
        ("letter to the shareholders", 9), ("business model", 7),
        ("reporting segments", 7), ("backlog", 5), ("geographic area", 9), ("by geographic area", 8),
        ("per area geografica", 14), ("ripartizione geografica", 10),
        ("by geographical", 5), ("market context", 5), ("new contracts", 5),
        ("business line", 4), ("employees", 2),
    ],
}

# Quota di pagine riservata a ogni categoria (frazione del massimo).
QUOTE = {
    "prospetti": 0.30, "pfn": 0.12, "debito": 0.22,
    "ownership": 0.08, "ma": 0.12, "overview": 0.12,
}
SOGLIA_CATEGORIA = {
    "prospetti": 14, "pfn": 12, "debito": 12, "ownership": 9, "ma": 9, "overview": 8,
}
ORDINE_PRIORITA = ["prospetti", "pfn", "debito", "ownership", "ma", "overview"]


TITOLI_PROSPETTI = (
    "stato patrimoniale", "situazione patrimoniale", "conto economico",
    "rendiconto finanziario", "prospetto delle variazioni",
    "statement of financial position", "income statement", "cash flow statement",
    "balance sheet", "statement of cash flows", "statement of changes in equity",
)


def _normalizza(testo: str) -> str:
    return _RE_SPAZI.sub(" ", testo.lower().replace("’", "'"))


def _punteggi_pagina(testo_norm: str) -> dict:
    numeri = sum(1 for n in _RE_NUMERO.findall(testo_norm) if not _RE_ANNO.match(n))
    punti = {}
    for categoria, parole in CATEGORIE.items():
        punti[categoria] = sum(peso for parola, peso in parole if parola in testo_norm)
    # I prospetti contabili sono tabelle: molte cifre. Una pagina discorsiva che
    # cita "stato patrimoniale" senza numeri non è un prospetto.
    if numeri >= 14 and punti["prospetti"] >= 6:
        punti["prospetti"] += min(numeri, 80) // 3
        # Il titolo del prospetto in testa alla pagina è il segnale più forte:
        # perdere un prospetto costa molto più che inviare una pagina in più.
        if any(titolo in testo_norm[:300] for titolo in TITOLI_PROSPETTI):
            punti["prospetti"] += 12
    elif numeri < 14:
        punti["prospetti"] = punti["prospetti"] // 3
    if numeri >= 14:
        if punti["pfn"] >= 6:
            punti["pfn"] += min(numeri, 40) // 8
        if punti["debito"] >= 6:
            punti["debito"] += min(numeri, 40) // 8
    punti["_numeri"] = numeri
    return punti


_ES_MINUSCOLO = ("bilancio d'esercizio", "bilancio di esercizio", "bilancio separato",
                 "bilancio civilistico", "statutory financial statements",
                 "separate financial statements")
_CONS_MAIUSCOLO = ("CONSOLIDATED", "CONSOLIDATO")
_ES_MAIUSCOLO = ("STATUTORY", "SEPARATE FINANCIAL", "BILANCIO D'ESERCIZIO",
                 "BILANCIO DI ESERCIZIO", "BILANCIO SEPARATO")


def _perimetro_pagine(testi_norm: list, testi_originali: list = None) -> list:
    """
    Per ogni pagina: "consolidato", "esercizio" o None, in base alle intestazioni
    e ai piè di pagina. L'ultima indicazione vista vale fino alla successiva.

    Molte relazioni annuali hanno in testa a ogni pagina una barra di navigazione
    con i nomi di tutte le sezioni, dove quella corrente è in MAIUSCOLO: in quel
    caso conta solo il maiuscolo.
    """
    esito = []
    corrente = None
    for i, t in enumerate(testi_norm):
        bordi = t[:220] + " " + t[-220:]
        testa = (testi_originali[i][:260] if testi_originali else "").replace("’", "'")
        cons_maiuscolo = any(k in testa for k in _CONS_MAIUSCOLO)
        es_maiuscolo = any(k in testa for k in _ES_MAIUSCOLO)
        cons = "consolidat" in bordi
        es = any(k in bordi for k in _ES_MINUSCOLO) or "statutory" in bordi

        if cons_maiuscolo != es_maiuscolo:
            corrente = "consolidato" if cons_maiuscolo else "esercizio"
        elif cons_maiuscolo and es_maiuscolo:
            pass                                  # ambiguo: resta l'ultima indicazione
        elif cons and es:
            # Barra di navigazione senza maiuscolo sulle sezioni di bilancio:
            # la pagina sta in un'altra sezione (es. relazione sulla gestione).
            if testi_originali:
                corrente = None
        elif cons:
            corrente = "consolidato"
        elif es:
            corrente = "esercizio"
        esito.append(corrente)
    return esito


def interpreta_pagine(testo: str, totale: int) -> list:
    """
    "1-9, 105-126, 40" -> indici 0-based ordinati e senza duplicati.
    Ignora ciò che è fuori dal documento. Solleva ValueError se il testo non è leggibile.
    """
    indici = set()
    for parte in re.split(r"[;,]", testo or ""):
        parte = parte.strip()
        if not parte:
            continue
        m = re.fullmatch(r"(\d+)\s*[-–]\s*(\d+)", parte)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a > b:
                a, b = b, a
            indici.update(range(a - 1, b))
        elif parte.isdigit():
            indici.add(int(parte) - 1)
        else:
            raise ValueError(f"Intervallo di pagine non valido: «{parte}»")
    return sorted(i for i in indici if 0 <= i < totale)


def _capitoli_registro_imprese(testi_norm: list) -> list:
    """
    I fascicoli scaricati dal Registro Imprese hanno, anche se scansionati, un
    timbro di testo "Capitolo N - TITOLO" sulla prima pagina di ogni capitolo.
    Restituisce [(titolo, inizio, fine_esclusa)].
    """
    inizi = []
    for i, t in enumerate(testi_norm):
        m = _RE_CAPITOLO.match(t[:120].strip())
        if m:
            titolo = m.group(2).split(" registro")[0].strip()
            inizi.append((titolo, i))
    capitoli = []
    for k, (titolo, inizio) in enumerate(inizi):
        fine = inizi[k + 1][1] if k + 1 < len(inizi) else len(testi_norm)
        capitoli.append((titolo, inizio, fine))
    return capitoli


def _priorita_scansione(testi_norm: list) -> tuple:
    """Ordine di priorità delle pagine per un PDF scansionato. -> (lista, descrizione)"""
    totale = len(testi_norm)
    capitoli = _capitoli_registro_imprese(testi_norm)
    if not capitoli:
        return list(range(totale)), "prime pagine del documento"

    def trova(*parole):
        for titolo, a, b in capitoli:
            if any(p in titolo for p in parole):
                return list(range(a, b))
        return []

    prospetto = trova("prospetto contabile", "prospetto", "bilancio")
    relazione = trova("relazione gestione", "relazione sulla gestione")
    nota = trova("nota integrativa", "note esplicative")

    ordine = []
    ordine += list(range(0, capitoli[0][1]))[:2]          # frontespizio e indice
    ordine += prospetto[:16]                              # SP, CE, rendiconto
    ordine += relazione[:12]                              # attività, andamento, PFN
    if nota:                                              # passivo e debiti: parte centrale
        n = len(nota)
        ordine += nota[int(n * 0.45):int(n * 0.85)]
    ordine += relazione[12:]
    ordine += nota[:6]                                    # criteri e perimetro
    ordine += prospetto[16:]
    ordine += nota
    ordine += list(range(totale))

    visti, unico = set(), []
    for i in ordine:
        if i not in visti:
            visti.add(i)
            unico.append(i)
    nomi = ", ".join(t.title() for t, _, _ in capitoli)
    return unico, f"capitoli del fascicolo ({nomi})"


PAGINE_MAX_DOCUMENTI_LUNGHI = 100   # oltre SOGLIA_DOCUMENTO_LUNGO pagine
SOGLIA_DOCUMENTO_LUNGO = 250


def analizza_pdf(pdf_bytes: bytes, perimetro: str = "auto", pagine_max: int = None,
                 pagine_manuali: str = "") -> dict:
    """
    Analizza il PDF e decide quali pagine inviare.

    Chiavi principali del risultato:
      mode             "text" | "scanned" | "manual" | "fallback"
      total_pages      pagine del documento
      priority_pages   tutte le pagine candidate, dalla più alla meno importante
      selected_pages   le prime `pagine_max` in ordine di documento (0-based)
      page_scores      punteggio complessivo per pagina (per il pannello di dettaglio)
      page_categories  categoria prevalente per pagina
      sections_found   sezioni riconosciute
      perimetro        perimetro effettivamente privilegiato
    """
    info = {
        "mode": "text", "total_pages": 0, "priority_pages": [], "selected_pages": [],
        "page_scores": [], "page_categories": [], "avg_chars": 0.0,
        "sections_found": [], "perimetro": perimetro, "scanned_pages": 0,
        "criterio": "", "ha_consolidato": False, "ha_esercizio": False,
    }

    pagine = leggi_pagine(pdf_bytes)
    totale = len(pagine)
    info["total_pages"] = totale
    if totale == 0:
        raise ErroreEstrazione("Il PDF non contiene pagine leggibili.")

    # Scelta automatica: i documenti lunghi (relazioni annuali) hanno bisogno di più pagine.
    info["pagine_max_automatico"] = not pagine_max
    if not pagine_max:
        pagine_max = (PAGINE_MAX_DOCUMENTI_LUNGHI if totale > SOGLIA_DOCUMENTO_LUNGO
                      else PAGINE_MAX_DEFAULT)
    info["pagine_max"] = pagine_max

    testi_norm = [_normalizza(p["testo"]) for p in pagine]
    info["avg_chars"] = sum(p["caratteri"] for p in pagine) / totale
    scansionate = [
        p["caratteri"] < SOGLIA_PAGINA_SCANSIONATA and p["immagine"] for p in pagine
    ]
    info["scanned_pages"] = sum(scansionate)

    # Sezioni riconosciute (anche dall'indice o dai timbri del fascicolo)
    tutto = " ".join(testi_norm)
    for parola, etichetta in [
        ("stato patrimoniale", "Stato Patrimoniale"), ("situazione patrimoniale", "Stato Patrimoniale"),
        ("prospetto contabile", "Prospetto Contabile"), ("conto economico", "Conto Economico"),
        ("rendiconto finanziario", "Rendiconto Finanziario"),
        ("posizione finanziaria netta", "PFN"), ("indebitamento finanziario netto", "PFN"),
        ("nota integrativa", "Nota Integrativa"), ("note esplicative", "Note Esplicative"),
        ("relazione sulla gestione", "Relazione sulla Gestione"),
        ("relazione gestione", "Relazione sulla Gestione"),
    ]:
        if parola in tutto and etichetta not in info["sections_found"]:
            info["sections_found"].append(etichetta)

    # ── a) pagine scelte a mano ──────────────────────────────────────────────
    if pagine_manuali and pagine_manuali.strip():
        scelte = interpreta_pagine(pagine_manuali, totale)
        if not scelte:
            raise ErroreEstrazione(
                f"Nessuna delle pagine indicate esiste nel documento (ha {totale} pagine)."
            )
        info.update({
            "mode": "manual", "priority_pages": scelte,
            "selected_pages": scelte[:max(pagine_max, len(scelte))],
            "page_scores": [0] * totale, "page_categories": [""] * totale,
            "criterio": "pagine indicate manualmente",
        })
        return info

    # ── b) documento scansionato (in tutto o in gran parte) ──────────────────
    if info["scanned_pages"] >= 0.6 * totale:
        priorita, criterio = _priorita_scansione(testi_norm)
        info.update({
            "mode": "scanned", "priority_pages": priorita,
            "selected_pages": sorted(priorita[:pagine_max]),
            "page_scores": [0] * totale, "page_categories": [""] * totale,
            "criterio": criterio,
        })
        return info

    # ── c) documento con testo: punteggio per categorie ──────────────────────
    punteggi = [_punteggi_pagina(t) for t in testi_norm]
    perimetri = _perimetro_pagine(testi_norm, [p["testo"] for p in pagine])
    info["ha_consolidato"] = perimetri.count("consolidato") >= 3
    info["ha_esercizio"] = perimetri.count("esercizio") >= 3

    voluto = perimetro
    if perimetro == "auto":
        voluto = "consolidato" if info["ha_consolidato"] else None
    info["perimetro"] = voluto or "unico"

    if voluto and info["ha_consolidato"] and info["ha_esercizio"]:
        for i, per in enumerate(perimetri):
            if per is None:
                continue
            fattore = 1.5 if per == voluto else 0.25
            for cat in ("prospetti", "pfn", "debito"):
                punteggi[i][cat] = int(punteggi[i][cat] * fattore)

    # Continuità: una tabella o una nota prosegue spesso sulla pagina successiva,
    # che da sola avrebbe poche parole chiave ("segue", senza titolo).
    base = [dict(p) for p in punteggi]
    for i in range(totale - 1):
        if scansionate[i + 1]:
            continue
        for cat in ("pfn", "debito", "ownership", "ma", "overview"):
            eredita = base[i][cat] // 2
            if eredita >= SOGLIA_CATEGORIA[cat] and eredita > punteggi[i + 1][cat]:
                punteggi[i + 1][cat] = eredita

    info["page_scores"] = [max(p[c] for c in ORDINE_PRIORITA) for p in punteggi]
    info["page_categories"] = [
        max(ORDINE_PRIORITA, key=lambda c: p[c]) if max(p[c] for c in ORDINE_PRIORITA) > 0 else ""
        for p in punteggi
    ]

    def migliori(categoria: str) -> list:
        candidati = [
            i for i in range(totale)
            if punteggi[i][categoria] >= SOGLIA_CATEGORIA[categoria] and not scansionate[i]
        ]
        return sorted(candidati, key=lambda i: (-punteggi[i][categoria], i))

    classifiche = {c: migliori(c) for c in ORDINE_PRIORITA}

    priorita, visti = [], set()

    def aggiungi(i):
        if 0 <= i < totale and i not in visti:
            visti.add(i)
            priorita.append(i)

    # Primo giro: ogni categoria riceve la sua quota, nell'ordine di importanza.
    for cat in ORDINE_PRIORITA:
        quota = max(2, int(round(QUOTE[cat] * pagine_max)))
        presi = 0
        for i in classifiche[cat]:
            if presi >= quota:
                break
            if i in visti:
                continue
            aggiungi(i)
            presi += 1
            # I prospetti proseguono sulla pagina successiva (attivo → passivo).
            if cat == "prospetti" and i + 1 < totale and punteggi[i + 1]["_numeri"] >= 14 \
                    and punteggi[i + 1]["prospetti"] >= 6 and presi < quota:
                if i + 1 not in visti:
                    aggiungi(i + 1)
                    presi += 1

    # Secondo giro: il resto, a rotazione fra le categorie, poi per punteggio.
    code = {c: [i for i in classifiche[c] if i not in visti] for c in ORDINE_PRIORITA}
    while any(code.values()):
        for cat in ORDINE_PRIORITA:
            while code[cat]:
                i = code[cat].pop(0)
                if i not in visti:
                    aggiungi(i)
                    break

    if not priorita:
        # Nessuna pagina riconosciuta: metà centrale del documento.
        info["mode"] = "fallback"
        priorita = list(range(totale // 4, max(totale // 4 + 1, 3 * totale // 4)))
        info["criterio"] = "nessuna sezione riconosciuta: pagine centrali"
    else:
        info["criterio"] = "pagine scelte per categoria"

    info["priority_pages"] = priorita
    info["selected_pages"] = sorted(priorita[:pagine_max])
    return info


# ── 3. PDF ridotto ───────────────────────────────────────────────────────────

def costruisci_pdf_chirurgico(pdf_bytes: bytes, page_indices: list) -> bytes:
    """Costruisce un PDF con le sole pagine indicate (0-based), in quell'ordine."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    for i in page_indices:
        if 0 <= i < len(reader.pages):
            writer.add_page(reader.pages[i])
    output = io.BytesIO()
    writer.write(output)
    writer.close()
    return output.getvalue()


def prepara_bilancio(pdf_bytes: bytes, perimetro: str = "auto",
                     pagine_max: int = None, pagine_manuali: str = "") -> tuple:
    """
    Punto d'ingresso: restituisce (pdf_ridotto_bytes, info).

    Il PDF ridotto contiene le pagine scelte (al massimo `pagine_max`). Se la
    richiesta risultasse troppo grande per l'account, `riduci_pdf` lo accorcia
    tenendo le pagine più importanti, senza bisogno dell'originale (che non
    viene conservato in memoria).
    """
    mb = len(pdf_bytes) / 1_000_000
    if mb > DIMENSIONE_MAX_MB:
        raise ErroreEstrazione(
            f"Il file pesa {mb:.0f} MB: il limite è {DIMENSIONE_MAX_MB} MB. "
            f"Riduci il PDF (ad esempio esportando solo le pagine del bilancio) e riprova."
        )

    info = analizza_pdf(pdf_bytes, perimetro, pagine_max, pagine_manuali)
    scelte = info["selected_pages"]
    ridotto = costruisci_pdf_chirurgico(pdf_bytes, scelte)
    info["pagine_inviate"] = len(scelte)
    info["mb_originale"] = round(mb, 1)
    info["mb_inviati"] = round(len(ridotto) / 1_000_000, 1)
    info["riduzione_pct"] = round((1 - len(ridotto) / len(pdf_bytes)) * 100, 1) if pdf_bytes else 0
    return ridotto, info


def riduci_pdf(pdf_ridotto: bytes, info: dict, nuovo_massimo: int) -> tuple:
    """
    Riduce ulteriormente un PDF già preparato tenendo le `nuovo_massimo` pagine
    più importanti. Restituisce (pdf, pagine_originali_selezionate).
    """
    attuali = info["selected_pages"]
    if len(attuali) <= nuovo_massimo:
        return pdf_ridotto, attuali
    rango = {p: k for k, p in enumerate(info.get("priority_pages", attuali))}
    tenute = sorted(sorted(attuali, key=lambda p: rango.get(p, 10 ** 6))[:nuovo_massimo])
    posizioni = [attuali.index(p) for p in tenute]
    return costruisci_pdf_chirurgico(pdf_ridotto, posizioni), tenute


# ── 4. Chiamate a Claude ─────────────────────────────────────────────────────

def _testo_risposta(messaggio) -> str:
    """Unisce i soli blocchi di testo (i modelli recenti antepongono blocchi di ragionamento)."""
    parti = []
    for blocco in getattr(messaggio, "content", []) or []:
        if getattr(blocco, "type", None) == "text":
            parti.append(getattr(blocco, "text", "") or "")
    return "".join(parti).strip()


def estrai_json(testo: str) -> dict:
    """Ricava l'oggetto JSON dalla risposta, tollerando backtick o testo attorno."""
    t = (testo or "").strip()
    if t.startswith("```"):
        t = re.sub(r"^```[a-zA-Z]*\s*", "", t)
        t = re.sub(r"\s*```\s*$", "", t)
    try:
        return json.loads(t)
    except json.JSONDecodeError:
        pass
    inizio, fine = t.find("{"), t.rfind("}")
    if inizio == -1 or fine <= inizio:
        raise ErroreEstrazione("La risposta del modello non contiene un JSON.")
    try:
        return json.loads(t[inizio:fine + 1])
    except json.JSONDecodeError as e:
        raise ErroreEstrazione(f"JSON non valido nella risposta del modello: {e}")


def _messaggio_errore(e) -> str:
    return str(getattr(e, "message", "") or e).lower()


def chiama_claude(client, content: list, max_tokens: int, modelli: list = None,
                  on_evento=None) -> tuple:
    """
    Invia la richiesta in streaming e restituisce (testo, modello_usato, uso).

    - passa al modello successivo se quello richiesto non è disponibile;
    - attende e riprova su limiti di frequenza, sovraccarico ed errori di rete;
    - solleva RichiestaTroppoGrande se la richiesta non può passare così com'è.
    """
    import anthropic

    modelli = modelli or MODELLI_ESTRAZIONE
    avvisa = on_evento or (lambda _msg: None)
    ultimo_errore = None

    for modello in modelli:
        limite_frequenza = 0
        for tentativo in range(4):
            try:
                with client.messages.stream(
                    model=modello,
                    max_tokens=max_tokens,
                    messages=[{"role": "user", "content": content}],
                ) as stream:
                    messaggio = stream.get_final_message()

                testo = _testo_risposta(messaggio)
                if getattr(messaggio, "stop_reason", None) == "max_tokens":
                    raise ErroreEstrazione(
                        "La risposta del modello è stata troncata per lunghezza. Riprova."
                    )
                if not testo:
                    raise ErroreEstrazione("Il modello non ha restituito testo.")
                return testo, modello, getattr(messaggio, "usage", None)

            except (anthropic.NotFoundError, anthropic.PermissionDeniedError) as e:
                ultimo_errore = e
                avvisa(f"Modello {modello} non disponibile, provo il successivo.")
                break

            except anthropic.RateLimitError as e:
                ultimo_errore = e
                limite_frequenza += 1
                if limite_frequenza >= 2:
                    raise RichiestaTroppoGrande(str(e))
                attesa = 65
                try:
                    attesa = min(90, max(5, int(float(e.response.headers.get("retry-after", 65)))))
                except Exception:
                    pass
                avvisa(f"Limite di richieste raggiunto: attendo {attesa} secondi.")
                time.sleep(attesa)

            except anthropic.APIStatusError as e:
                ultimo_errore = e
                stato = getattr(e, "status_code", 0)
                msg = _messaggio_errore(e)
                if stato == 413 or "too long" in msg or "too large" in msg \
                        or "too many" in msg or "maximum of" in msg or "exceed" in msg:
                    raise RichiestaTroppoGrande(str(e))
                if stato == 400 and "model" in msg:
                    avvisa(f"Modello {modello} non utilizzabile, provo il successivo.")
                    break
                if stato >= 500 or stato == 529:
                    attesa = 15 * (tentativo + 1)
                    avvisa(f"Servizio momentaneamente sovraccarico: riprovo fra {attesa} secondi.")
                    time.sleep(attesa)
                    continue
                raise ErroreEstrazione(f"Errore API ({stato}): {getattr(e, 'message', e)}")

            except (anthropic.APIConnectionError, anthropic.APITimeoutError) as e:
                ultimo_errore = e
                attesa = 10 * (tentativo + 1)
                avvisa(f"Problema di connessione: riprovo fra {attesa} secondi.")
                time.sleep(attesa)

    raise ErroreEstrazione(f"Impossibile completare la richiesta: {ultimo_errore}")


def _blocco_documento(pdf_bytes: bytes) -> dict:
    import base64
    return {
        "type": "document",
        "source": {
            "type": "base64",
            "media_type": "application/pdf",
            "data": base64.standard_b64encode(pdf_bytes).decode("ascii"),
        },
        # La seconda lettura (verifica) riusa lo stesso documento dalla cache.
        "cache_control": {"type": "ephemeral"},
    }


# ── 5. Prompt ────────────────────────────────────────────────────────────────

def _regola_perimetro(perimetro: str) -> str:
    if perimetro == "consolidato":
        return ("PERIMETRO RICHIESTO: CONSOLIDATO. Usa esclusivamente i dati del bilancio "
                "consolidato (Gruppo). Se una voce non è disponibile a livello consolidato "
                "metti null: NON sostituirla con il dato della capogruppo.")
    if perimetro == "esercizio":
        return ("PERIMETRO RICHIESTO: BILANCIO D'ESERCIZIO / SEPARATO della società. Usa "
                "esclusivamente i dati individuali della società, anche se il documento "
                "contiene anche il consolidato. Se una voce non è disponibile metti null.")
    return ("PERIMETRO: se il documento contiene sia il consolidato sia il bilancio "
            "d'esercizio, usa il CONSOLIDATO per tutte le voci. Se contiene un solo "
            "perimetro, usa quello.")


DETTAGLI = {"completo": "Completo", "sintetico": "Sintetico"}


def _regola_dettaglio(dettaglio: str) -> str:
    if dettaglio == "sintetico":
        return (
            "LIVELLO DI DETTAGLIO: SINTETICO. Il report deve stare in due pagine. Limiti: "
            "overview al massimo 120 parole; core_business e mercati 90 parole ciascuno; "
            "struttura_controllo 90 parole; note di struttura_debito 80 parole; "
            "note_aggiuntive 80 parole; ogni operazione_ma al massimo 60 parole, con al più "
            "4 voci in \"passaggi\" (solo quelle che decidono l'esito). Riporta solo ciò che "
            "serve a un avvocato per inquadrare società, debito, controllo e operazioni: "
            "niente elenchi di progetti, tecnologie o dati accessori."
        )
    return (
        "LIVELLO DI DETTAGLIO: COMPLETO. Sii esauriente ma ordinato: ogni informazione nel "
        "campo che le compete, senza ripetere la stessa informazione in campi diversi."
    )


def costruisci_prompt_estrazione(nome: str, anno: str, lingua_prompt: str, n_pagine: int,
                                 nota_estrazione: str, testo_suppl: str,
                                 perimetro: str, fonti_supplementari: list,
                                 dettaglio: str = "completo") -> str:
    elenco_fonti = ", ".join(fonti_supplementari) if fonti_supplementari else "nessuno"
    return f"""Sei un analista M&A e finance di uno studio legale internazionale.
Analizza il bilancio allegato di **{nome}** (esercizio {anno}) e, se presenti, i documenti supplementari.
{nota_estrazione}
Produci un report strutturato in JSON, con tutti i testi {lingua_prompt}.
{_regola_dettaglio(dettaglio)}

Documenti supplementari forniti: {elenco_fonti}.
{testo_suppl}

Rispondi SOLO con un oggetto JSON valido, senza backtick, senza testo aggiuntivo.

═══════════════════════════════════════════════════════════════════
REGOLE DI METODO (valgono per tutto il report)
═══════════════════════════════════════════════════════════════════

1. SOLO DOCUMENTI. Ogni numero e ogni affermazione deve essere leggibile nei
   documenti forniti. Non usare conoscenze esterne sull'azienda, nemmeno se
   la conosci. Se un'informazione non c'è, scrivi "N/D" (testi) o null (numeri).
   Niente supposizioni ("presumibilmente", "tipicamente", "riconducibile a")
   e niente caratterizzazioni non scritte nei documenti.

2. {_regola_perimetro(perimetro)}
   Non mescolare MAI i due perimetri: tutti i numeri del report devono
   appartenere allo stesso perimetro. Indica il perimetro usato nel campo
   "perimetro" ("consolidato" oppure "esercizio").

3. ESERCIZIO {anno}. I prospetti hanno due colonne (esercizio corrente e
   precedente). Leggi SEMPRE la colonna dell'esercizio {anno}. Prima di
   leggere una tabella controlla l'intestazione delle colonne: l'ordine
   può essere invertito da una tabella all'altra. Lo stesso vale per i
   testi: non riportare come {anno} fatti, oneri o componenti dell'esercizio
   precedente. Se il documento riguarda un esercizio diverso da {anno},
   usa l'esercizio del documento, indicalo in "anno_riferimento" e segnalalo
   in "note_aggiuntive".

4. UNITÀ. Guarda l'unità dichiarata in testa a OGNI tabella che leggi
   (euro, migliaia, milioni): i prospetti e le tabelle della relazione
   sulla gestione spesso usano unità diverse. Riporta tutti i numeri nella
   stessa unità, quella indicata in "unita".

5. FONTE DI OGNI NUMERO. Per ogni valore numerico non nullo compila la voce
   corrispondente in "fonti" con:
   - "p": numero d'ordine della pagina NELL'ALLEGATO (da 1 a {n_pagine}),
     non il numero stampato sulla pagina;
   - "voce": l'etichetta esatta della riga letta, come è scritta;
   - "letto": il valore esattamente come stampato, con la sua unità
     (es. "1.234.567 (migliaia di euro)");
   - "calcolo": solo se il valore è una somma o differenza di più righe,
     scrivi l'operazione con gli addendi (es. "25.300 + 4.100 + 900").
   Se non riesci a indicare pagina e voce, il valore va messo a null.

6. PAROLE E DATE ESATTE. Riporta denominazioni, forme di amministrazione,
   quote e date come scritte nei documenti. Distingui data di chiusura
   dell'esercizio, data di approvazione, data di deposito e data di
   estrazione del documento: non confonderle.

7. CATENE DI CONTROLLO. Distingui partecipazioni dirette e indirette e
   non invertire la direzione del controllo (chi controlla chi).

8. PAGINE CITATE NEI TESTI. Ogni volta che in un campo di testo citi una
   pagina del bilancio, anche nel mezzo di una frase, scrivi SOLO il
   marcatore [p.N], dove N è il numero d'ordine della pagina nell'allegato
   (da 1 a {n_pagine}); per più pagine [p.N,M] o [p.N-M]. Esempi: "come
   indicato a [p.34]", "le tabelle di [p.35,45]". Il programma converte i
   marcatori nel numero di pagina del PDF originale. Non scrivere mai
   "p. 34", "pagina 34", "dell'allegato" né il numero stampato sulla pagina.

9. NIENTE COMMENTI SUL MATERIALE NEI TESTI. Nei campi overview,
   core_business, mercati, ownership e operazioni_ma non scrivere frasi
   come "il testo si interrompe", "la visura è troncata", "pagina non
   inclusa": riporta ciò che i documenti dicono. Le lacune del materiale
   vanno solo in "dati_non_trovati" o, in una riga, in "note_aggiuntive".
   Il report è letto da avvocati che non conoscono questo strumento: non
   citare mai lo schema JSON, i nomi dei campi, il prompt, "l'estratto",
   "l'allegato", "il programma" o i valori predefiniti.

11. ESATTEZZA DEI TESTI. Copia nomi di persone e società lettera per
    lettera come sono scritti nella fonte, anche se sembrano refusi (in tal
    caso aggiungi "[sic]"). Scrivi i numeri con il punto per le migliaia e
    la virgola per i decimali (1.940.282.698; 21,19%). Prima di scrivere un
    conteggio ("28 operazioni", "tre clienti") riconta gli elementi.
    Segnala come discrepanza solo due valori che dovrebbero coincidere e non
    coincidono: valore nominale, valore contabile e pagamenti contrattuali
    dello stesso debito sono grandezze diverse, non una discrepanza.

12. OGNI COSA AL SUO POSTO, UNA VOLTA SOLA.
    - "mercati": mercati, aree geografiche, clienti, contesto. Non sede,
      addetti o unità locali.
    - "ownership": soci, controllo, patti, organi sociali. Non l'elenco
      delle controllate (al più il loro numero, in una frase).
    - Non scrivere nei testi l'esito dei tuoi controlli ("check contabile:
      verificato"), né rimandi alla numerazione dei documenti che hai
      ricevuto ("operazioni n. 7-37 dell'elenco"), né conteggi di servizio
      della visura ("Numero in carica: 5").
    - Non ripetere la stessa informazione due volte, nemmeno con parole
      diverse, nello stesso campo o in campi diversi.

13. ESERCIZIO PRECEDENTE. Compila "dati_anno_precedente" con gli stessi
    dati letti nella colonna dell'esercizio precedente degli stessi
    prospetti e delle stesse tabelle (stesse regole, stessa unità, stesso
    perimetro). Se un dato dell'esercizio precedente non è esposto nelle
    pagine fornite metti null: non ricavarlo da percentuali di variazione.
    Per i bilanci civilistici lascia "ebitda" a null e compila i quattro
    componenti, come per l'esercizio corrente.

10. DOCUMENTI SUPPLEMENTARI. La visura è già ridotta alle sezioni utili e
    l'export Mergermarket elenca tutte le operazioni: leggili per intero.
    Per ogni operazione Mergermarket riporta date di annuncio e di
    completamento, controparti, corrispettivo e, se l'operazione non è
    completata, l'ULTIMO aggiornamento disponibile (le descrizioni sono in
    ordine cronologico: l'ultima notizia è in fondo) e lo stato di ciascuna
    autorizzazione o indagine citata (autorità, esito o stato, data),
    senza ometterne nessuna. Se due notizie della
    stessa fonte si contraddicono, segnalalo invece di sceglierne una.

═══════════════════════════════════════════════════════════════════
STEP 0 — DETERMINA IL TIPO DI BILANCIO (CRITICO)
═══════════════════════════════════════════════════════════════════

Prima di estrarre qualsiasi dato, identifica il TIPO di bilancio analizzando
il documento. Popola il campo "tipo_bilancio" con UNO di questi valori:

(A) "consolidato_ifrs" — Bilancio consolidato redatto secondo IAS/IFRS.
    INDICATORI: presenza di "bilancio consolidato", "area di consolidamento",
    "patrimonio netto di terzi", "interessenze di minoranza", riferimenti a
    "IAS", "IFRS", voci come "Attività per diritti d'uso" / "Passività per
    leasing" (IFRS 16), schema ESMA 32-382-1138 per la PFN, conto economico
    complessivo, presenza di "Avviamento" con impairment test.

(B) "esercizio_ifrs" — Bilancio d'esercizio (separato) che applica IFRS.
    INDICATORI: titolo "Bilancio d'esercizio" o "Bilancio separato" +
    riferimenti a IAS/IFRS + presenza di passività per leasing IFRS 16.

(C) "esercizio_civilistico" — Bilancio d'esercizio secondo Codice Civile (OIC).
    INDICATORI: stato patrimoniale con macroclassi A/B/C/D lato attivo e
    A/B/C/D/E lato passivo (art. 2424 C.C.), conto economico ex art. 2425
    C.C., riferimenti a "OIC" / "principi contabili nazionali", voce
    "D) Debiti" con sottovoci D.1 Obbligazioni, D.4 Debiti verso banche.

(D) "consolidato_civilistico" — Bilancio consolidato redatto secondo OIC
    (D.Lgs. 127/1991): stessi schemi del caso (C), ma di Gruppo, con
    "patrimonio netto di terzi" e "area di consolidamento". Applica le
    regole di estrazione del caso (C).

═══════════════════════════════════════════════════════════════════
STEP 1 — REGOLE DI ESTRAZIONE DATI FINANZIARI
═══════════════════════════════════════════════════════════════════

REGOLE GENERALI PER TUTTI I TIPI:
- Tutti i valori in "dati_finanziari" e "struttura_debito" devono essere
  NUMERI PURI (es. 1234567), mai stringhe.
- I valori negativi (perdite, PFN negativa quando cassa > debito) devono
  avere il segno negativo (es. -5000). Nei prospetti i negativi sono spesso
  tra parentesi.
- Se un valore non è presente nel documento usa null, MAI 0 o "N/D".
  (Unica eccezione: "obbligazioni" = 0 se dal documento risulta che la
  società non ha emesso obbligazioni.)
- Non inventare e non stimare: meglio null che un numero sbagliato.

CONTO ECONOMICO — regole per tipo:

▶ Se IFRS (casi A e B):
- "ricavi": la riga dei RICAVI TOTALI del conto economico (es. "Ricavi",
  "Ricavi delle vendite e delle prestazioni", "Totale ricavi"). Se il CE
  espone più righe di ricavo con un totale, usa il totale dei ricavi, non
  una sola componente; NON includere "Altri proventi" se esposti fuori dal
  totale ricavi.
- "ebitda": "Margine operativo lordo" o "EBITDA" se esposto esplicitamente
  (nel CE o nella relazione sulla gestione). Usa il valore REPORTED, non
  quello adjusted/ricorrente; se esiste anche un valore adjusted, indicalo
  in "note_aggiuntive". Se l'EBITDA non è esposto in nessun punto, metti
  null in "ebitda" e compila "componenti_ebitda".
- "ebit": "Risultato operativo" o "EBIT".
- "ammortamenti": ammortamenti immateriali + materiali + diritti d'uso
  + svalutazioni di immobilizzazioni.
- "utile_netto": il RISULTATO NETTO CONSOLIDATO DELL'ESERCIZIO (Gruppo +
  terzi), cioè l'ultima riga del CE prima della ripartizione. Indica in
  "note_aggiuntive" la quota di pertinenza del Gruppo e quella di terzi,
  se esposte.

▶ Se civilistico (casi C e D):
- "ricavi": voce A.1 "Ricavi delle vendite e delle prestazioni".
  NON usare "Valore della produzione" (voce A totale).
- "ebit": "Differenza tra valore e costi della produzione" (A − B).
- "ammortamenti": voce B.10 "Ammortamenti e svalutazioni" (totale B.10).
- "ebitda": NON è una voce del CE civilistico. Metti null in "ebitda" e
  compila "componenti_ebitda" con le quattro righe lette dal CE: il
  calcolo (A − B) + B.10 + B.12 + B.13 viene fatto dal programma.
  Se la relazione sulla gestione espone un proprio EBITDA/MOL, riportalo
  in "ebitda_dichiarato".
- "utile_netto": "Utile (perdita) dell'esercizio" — ultima riga del CE
  (nel consolidato: utile consolidato, Gruppo + terzi).

STATO PATRIMONIALE — ATTIVO (tutti i tipi):
- "totale_attivo": riga "Totale attivo".
- "cassa":
  ▶ IFRS: "Disponibilità liquide e mezzi equivalenti".
  ▶ Civilistico: voce C.IV dell'attivo circolante "Disponibilità liquide".

STATO PATRIMONIALE — PASSIVO (tutti i tipi):
- "patrimonio_netto":
  ▶ Consolidato: patrimonio netto TOTALE (Gruppo + terzi).
  ▶ Esercizio: "Totale patrimonio netto" (voce A del passivo in civilistico).

═══════════════════════════════════════════════════════════════════
STEP 2 — STRUTTURA DEL DEBITO (SEZIONE CRITICA, LEGGI ATTENTAMENTE)
═══════════════════════════════════════════════════════════════════

Il CHECK CONTABILE deve sempre tornare:

    indebitamento_totale >= debito_bancario + obbligazioni

─────────────────────────────────────────────────────────
CIVILISTICO (casi C e D) — schema OIC / art. 2424 C.C.
─────────────────────────────────────────────────────────

- "indebitamento_totale": "TOTALE DEBITI" — voce D del Passivo (riga
  riassuntiva di TUTTI i debiti: obbligazioni, banche, altri finanziatori,
  fornitori, tributari, previdenziali, ecc.).
  NON confondere con "totale debiti finanziari".
- "debito_bancario": voce D.4 "Debiti verso banche" (entro + oltre l'esercizio).
- "obbligazioni": voci D.1 "Obbligazioni" + D.2 "Obbligazioni convertibili".
- "debito_netto": "Posizione Finanziaria Netta" o "Indebitamento Finanziario
  Netto" come esposto nella Relazione sulla Gestione o nelle note. Se non
  esposto, calcola: (D.1 + D.2 + D.3 finanziario + D.4 + D.5 finanziario)
  − C.IV Disponibilità liquide − C.III attività finanziarie liquide, e
  scrivi gli addendi in "calcolo" nella fonte.

─────────────────────────────────────────────────────────
IFRS (casi A e B) — schema ESMA
─────────────────────────────────────────────────────────

In IFRS NON esiste la "voce D" del passivo.

- "indebitamento_totale": INDEBITAMENTO FINANZIARIO LORDO =
    Passività finanziarie (non correnti + correnti) + Passività per leasing
    (non correnti + correnti).
  ERRORE DA NON COMMETTERE: NON usare "Totale Passivo − Patrimonio Netto"
  (include debiti commerciali, tributari, fondi, imposte differite).
- "debito_bancario": solo i finanziamenti bancari e istituzionali (BEI, CDP,
  SACE), quota corrente + non corrente, scoperti e debiti bancari a breve,
  al netto di eventuali commissioni capitalizzate. ESCLUSI bond e leasing.
  ATTENZIONE DOUBLE COUNTING: se il totale "Passività finanziarie" include
  il bond, sottrailo: il bond va esclusivamente in "obbligazioni".
- "obbligazioni": valore NOMINALE dei prestiti obbligazionari in essere.
- "debito_netto": Posizione Finanziaria Netta INCLUSIVA delle passività per
  leasing (schema ESMA 32-382-1138), come pubblicata dalla società. Se la
  società espone due valori (incl. ed escl. leasing), usa quello INCL.
  leasing e riporta in "note" il valore escl. leasing.

LEVA (tutti i tipi): NON calcolare la leva: la calcola il programma come
debito_netto / ebitda. Se la società dichiara ESPLICITAMENTE un proprio
leverage ratio (es. ai fini dei covenant), riportalo in "leva_dichiarata"
con la sua definizione in "leva_dichiarata_definizione"; altrimenti null.

═══════════════════════════════════════════════════════════════════
STEP 3 — CHECK prima di rispondere
═══════════════════════════════════════════════════════════════════

1. indebitamento_totale >= debito_bancario + obbligazioni.
2. patrimonio_netto < totale_attivo; cassa < totale_attivo.
3. Tutte le voci appartengono allo stesso perimetro e all'esercizio {anno}.
4. Ogni numero non nullo ha la sua fonte (pagina dell'allegato + voce).
5. Se la relazione sulla gestione e i prospetti riportano valori diversi per
   la stessa grandezza, usa il valore dei prospetti e segnala la differenza
   in "note_aggiuntive".
6. Ogni frase dei campi di testo è sostenuta da un passaggio dei documenti.
Se un check non torna e non riesci a correggerlo, annotalo in "note".

═══════════════════════════════════════════════════════════════════
SCHEMA JSON DI OUTPUT
═══════════════════════════════════════════════════════════════════

{{
  "nome_azienda": "{nome}",
  "tipo_bilancio": "",
  "perimetro": "",
  "overview": "",
  "core_business": "",
  "mercati": "",
  "dati_finanziari": {{
    "unita": "<unità dei prospetti: euro | migliaia di euro | milioni di euro>",
    "ricavi": null,
    "ebitda": null,
    "ebit": null,
    "utile_netto": null,
    "ammortamenti": null,
    "totale_attivo": null,
    "patrimonio_netto": null,
    "cassa": null,
    "anno_riferimento": "{anno}"
  }},
  "componenti_ebitda": {{
    "differenza_a_b": null,
    "ammortamenti_svalutazioni_b10": null,
    "accantonamenti_rischi_b12": null,
    "altri_accantonamenti_b13": null
  }},
  "ebitda_dichiarato": null,
  "dati_anno_precedente": {{
    "anno": "",
    "ricavi": null,
    "ebitda": null,
    "ebit": null,
    "utile_netto": null,
    "totale_attivo": null,
    "patrimonio_netto": null,
    "cassa": null,
    "indebitamento_totale": null,
    "debito_bancario": null,
    "obbligazioni": null,
    "debito_netto": null,
    "differenza_a_b": null,
    "ammortamenti_svalutazioni_b10": null,
    "accantonamenti_rischi_b12": null,
    "altri_accantonamenti_b13": null
  }},
  "struttura_debito": {{
    "indebitamento_totale": null,
    "debito_bancario": null,
    "obbligazioni": null,
    "debito_netto": null,
    "leva_dichiarata": null,
    "leva_dichiarata_definizione": "",
    "scadenze_principali": "",
    "note": ""
  }},
  "fonti": {{
    "ricavi": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "ebitda": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "ebit": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "utile_netto": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "ammortamenti": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "totale_attivo": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "patrimonio_netto": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "cassa": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "indebitamento_totale": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "debito_bancario": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "obbligazioni": {{"p": null, "voce": "", "letto": "", "calcolo": ""}},
    "debito_netto": {{"p": null, "voce": "", "letto": "", "calcolo": ""}}
  }},
  "ownership": {{
    "azionista_principale": "",
    "quota_principale": "",
    "altri_azionisti": "",
    "struttura_controllo": "",
    "note": ""
  }},
  "operazioni_ma": [
    {{"anno": "", "tipo": "", "descrizione": "", "passaggi": [], "fonte": ""}}
  ],
  "dati_non_trovati": [
    {{"campo": "", "motivo": ""}}
  ],
  "note_aggiuntive": ""
}}

═══════════════════════════════════════════════════════════════════
ISTRUZIONI PER I CAMPI DI TESTO
═══════════════════════════════════════════════════════════════════

"note" di struttura_debito — includi, separati da " | ", quando applicabili:
1. Metodologia PFN: "PFN incl. leasing (schema ESMA)" / "PFN escl. leasing" /
   "PFN ricostruita (non esposta dalla società)".
2. PFN alternativa, se dichiarata: es. "PFN escl. leasing: 961.805".
3. Covenant e garanzie rilevanti (pegni, ipoteche), se descritti.
4. Eventuali discrepanze nei check contabili.

"scadenze_principali" — profilo di scadenza del debito bancario e
obbligazionario (esclusi leasing) per anno, usando il VALORE NOMINALE da
rimborsare, in formato sintetico, es.:
"2026: €131,9M | 2027: €245,4M | 2028: €430,4M (di cui bond €350M)".
Se la società espone solo i pagamenti contrattuali (che possono includere
interessi) o i valori contabili, usa quelli ma dichiaralo ("pagamenti
contrattuali", "valore contabile"): non chiamarli nominali.
Se non disponibile nelle pagine fornite, scrivi "N/D".

"ownership" — usa SOLO soci e quote scritti nei documenti (bilancio o
visura). Per gli organi sociali prevale la visura. Le società quotate non
hanno elenco soci in visura: cerca le quote nel bilancio (tabella dell'area
di consolidamento, relazione sul governo societario, nota sulle parti
correlate) e riporta in "struttura_controllo" patti parasociali e natura
del controllo (solitario, congiunto, nessuno).

"passaggi" (in ogni operazione) — lista di frasi brevi, una per ogni
autorizzazione, indagine o tappa rilevante di un'operazione NON ancora
completata, nella forma "Autorità o tappa: stato o esito (data)". Esempio:
["UK CMA: autorizzata (04/11/2025)", "Commissione Europea: Fase II in
corso, termini sospesi (27/08/2026)"]. In "descrizione" resta solo la
sintesi dell'operazione e del suo stato; il dettaglio va in "passaggi".
Per le operazioni completate lascia la lista vuota.

"note" di struttura_debito — una informazione per voce, separate da " | ".

"operazioni_ma" — SOLO operazioni descritte nei documenti forniti. Per
ognuna compila "fonte": "Bilancio [p.N]" oppure il nome del documento
supplementare (es. "Mergermarket"). Riporta tutte le operazioni degli
ultimi 5 esercizi presenti nei documenti; per quelle più vecchie basta una
riga riassuntiva finale. Incrocia bilancio e Mergermarket: un'operazione
descritta in entrambi va riportata una sola volta, con entrambe le fonti. NON attribuire nulla a
Mergermarket se l'export Mergermarket non è tra i documenti forniti. Se
non risultano operazioni, restituisci una lista vuota [].

"dati_non_trovati" — per ogni numero lasciato a null indica il campo e il
motivo: "pagina non presente nell'estratto" oppure "dato non esposto nel
bilancio". Lista vuota se hai trovato tutto.

"overview", "core_business", "mercati" — sintesi fedele di quanto scritto
nei documenti; "N/D" se le pagine fornite non ne parlano.
"""


def costruisci_prompt_verifica(nome: str, anno: str, n_pagine: int, estratto: dict,
                               testo_suppl: str = "") -> str:
    supplementari = (
        "DOCUMENTI SUPPLEMENTARI usati dal collega (oltre al bilancio allegato):" + testo_suppl
        if testo_suppl else "Il collega non aveva documenti supplementari."
    )
    return f"""Sei un revisore. Un collega ha estratto dal bilancio allegato di **{nome}**
(esercizio {anno}) i dati riportati sotto. Il tuo compito è RILEGGERE il documento
e controllare ogni numero in modo indipendente, senza dare per buono il lavoro del collega.

DATI ESTRATTI DAL COLLEGA:
{json.dumps(estratto, ensure_ascii=False, indent=1)}

COME PROCEDERE, per ogni campo numerico:
1. Vai alla pagina indicata in "fonti" (numero d'ordine nell'allegato, da 1 a {n_pagine})
   e cerca la riga indicata in "voce".
2. Controlla: (a) che la colonna sia quella dell'esercizio {anno} e non del precedente;
   (b) che il perimetro sia quello dichiarato in "perimetro" e non l'altro;
   (c) che la riga sia quella giusta secondo la definizione del campo qui sotto;
   (d) che l'unità sia stata convertita correttamente nell'unità dichiarata in "unita";
   (e) il segno; (f) se c'è un "calcolo", rifai l'operazione.
3. Se un campo è null, cerca comunque il dato nel documento: se lo trovi, proponilo.
4. Controlla allo stesso modo "dati_anno_precedente": ogni valore deve venire dalla colonna
   dell'esercizio PRECEDENTE della stessa riga usata per il {anno}. L'errore tipico è lo
   scambio delle due colonne. Per questi campi usa il nome "precedente.<campo>"
   (es. "precedente.ricavi") e non servono pagina e voce.

DEFINIZIONI DEI CAMPI:
- ricavi: civilistico = voce A.1 (non il valore della produzione); IFRS = totale ricavi del CE.
- ebitda: valore reported esposto dalla società; null nel civilistico se non esposto
  (in quel caso si controllano i "componenti_ebitda": A−B, B.10, B.12, B.13).
- ebit: risultato operativo (civilistico: differenza A − B).
- utile_netto: risultato netto dell'esercizio (nel consolidato: Gruppo + terzi).
- ammortamenti: civilistico = totale B.10; IFRS = ammortamenti e svalutazioni di immobilizzazioni.
- patrimonio_netto: totale (nel consolidato: Gruppo + terzi).
- cassa: disponibilità liquide (e mezzi equivalenti).
- indebitamento_totale: civilistico = TOTALE DEBITI (voce D del passivo);
  IFRS = passività finanziarie + passività per leasing, correnti e non correnti
  (MAI totale passivo − patrimonio netto).
- debito_bancario: debiti verso banche (IFRS: senza bond e senza leasing).
- obbligazioni: prestiti obbligazionari (IFRS: valore nominale).
- debito_netto: PFN / indebitamento finanziario netto pubblicato (IFRS: incl. leasing).

{supplementari}

Controlla anche i TESTI (overview, core_business, mercati, ownership, operazioni_ma,
note, scadenze): confronta ogni affermazione con il bilancio allegato e con i documenti
supplementari qui sopra e segnala ciò che non trova riscontro o li contraddice. In particolare:
- date, quote, importi, denominazioni, direzione del controllo, anno di riferimento;
- nomi di persone e società: devono essere scritti lettera per lettera come nella fonte;
- conteggi ("28 operazioni", "tre clienti"): ricontali;
- formato dei numeri (punto per le migliaia, virgola per i decimali) e somme;
- etichette sbagliate (un pagamento contrattuale chiamato "nominale" o "valore contabile");
- false discrepanze: grandezze diverse presentate come valori in contrasto;
- riferimenti di pagina che non siano nella forma [p.N] (N = pagina dell'allegato);
- frasi che parlano dello strumento anziché dell'azienda (schema, campi, estratto, allegato).
Per ogni problema indica la frase ESATTA, copiata carattere per carattere dal testo del
collega, e la frase corretta che deve sostituirla.

Rispondi SOLO con un oggetto JSON valido, senza backtick:

{{
  "numeri": [
    {{
      "campo": "ricavi",
      "esito": "ok | corretto | non_verificabile",
      "valore_corretto": null,
      "p": null,
      "voce": "",
      "letto": "",
      "motivo": ""
    }}
  ],
  "testi": [
    {{"campo": "overview", "frase": "", "problema": "", "correzione": ""}}
  ]
}}

REGOLE PER LA RISPOSTA:
- Una riga in "numeri" per OGNI campo numerico (dati_finanziari, struttura_debito,
  componenti_ebitda), anche quando è corretto.
- "ok": hai ritrovato lo stesso valore. "corretto": il valore giusto è un altro
  (mettilo in "valore_corretto", nella stessa unità dichiarata in "unita", con pagina,
  voce e valore letto). "non_verificabile": la pagina necessaria non è nell'allegato.
- Usa "corretto" solo se sei certo e puoi indicare pagina e voce; nel dubbio usa
  "non_verificabile" spiegando il motivo.
- "testi": lista vuota se non hai nulla da segnalare. "frase" deve essere un passaggio
  breve (una frase o parte di frase) copiato esattamente; "correzione" è il testo che lo
  sostituisce, già completo e corretto, nella stessa lingua del report. Se la frase va
  eliminata scrivi "correzione": "". Se non sei certo della correzione, lascia
  "correzione": null e spiega il problema.
"""


# ── 6. Elaborazione del risultato ────────────────────────────────────────────

def _numero(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(" ", ""))
    except (TypeError, ValueError):
        return None


def _it(n: float) -> str:
    """Numero con il punto come separatore delle migliaia."""
    return f"{n:,.0f}".replace(",", ".")


def _sezione_di(campo: str) -> str:
    if campo.startswith("precedente."):
        return "dati_anno_precedente"
    if campo in KPI_FINANZIARI:
        return "dati_finanziari"
    if campo in KPI_DEBITO:
        return "struttura_debito"
    return "componenti_ebitda"


def applica_verifica(report: dict, verifica: dict) -> list:
    """
    Applica le correzioni proposte dalla seconda lettura e restituisce il
    registro degli esiti (lista di dict), da mostrare all'utente.
    """
    registro = []
    fonti = report.setdefault("fonti", {})
    for riga in (verifica or {}).get("numeri", []) or []:
        if not isinstance(riga, dict):
            continue
        campo = str(riga.get("campo", "")).strip()
        esito = str(riga.get("esito", "")).strip().lower()
        if not campo:
            continue
        sezione = _sezione_di(campo)
        contenitore = report.get(sezione)
        precedente = sezione == "dati_anno_precedente"
        chiave = campo.split(".", 1)[1] if precedente else campo
        if not isinstance(contenitore, dict) or chiave not in contenitore:
            continue
        if precedente:
            # Esercizio precedente: nessuna fonte per voce; si corregge il valore e lo si annota.
            if esito == "corretto":
                nuovo = _numero(riga.get("valore_corretto"))
                vecchio = _numero(contenitore.get(chiave))
                if nuovo is not None and (vecchio is None
                                          or abs(nuovo - vecchio) > 0.005 * max(abs(vecchio), 1)):
                    contenitore[chiave] = nuovo
                    registro.append({
                        "campo": campo, "esito": "corretto", "prima": vecchio, "dopo": nuovo,
                        "motivo": riga.get("motivo", ""),
                    })
            continue

        if esito == "corretto":
            nuovo = _numero(riga.get("valore_corretto"))
            vecchio = _numero(contenitore.get(campo))
            if nuovo is None or riga.get("p") in (None, "") or not riga.get("voce"):
                registro.append({
                    "campo": campo, "esito": "da_verificare",
                    "motivo": riga.get("motivo") or "correzione proposta senza fonte: non applicata",
                })
                continue
            if vecchio is not None and abs(nuovo - vecchio) <= 0.005 * max(abs(vecchio), 1):
                registro.append({"campo": campo, "esito": "ok"})
                continue
            contenitore[campo] = nuovo
            fonti[campo] = {
                "p": riga.get("p"), "voce": riga.get("voce", ""),
                "letto": riga.get("letto", ""), "calcolo": "",
            }
            registro.append({
                "campo": campo, "esito": "corretto", "prima": vecchio, "dopo": nuovo,
                "motivo": riga.get("motivo", ""),
            })
        elif esito == "non_verificabile":
            if _numero(contenitore.get(campo)) is not None:
                registro.append({
                    "campo": campo, "esito": "non_verificabile",
                    "motivo": riga.get("motivo", ""),
                })
        else:
            registro.append({"campo": campo, "esito": "ok"})

    for riga in (verifica or {}).get("testi", []) or []:
        if not (isinstance(riga, dict) and riga.get("frase")):
            continue
        frase = str(riga.get("frase"))
        correzione = riga.get("correzione")
        applicata = False
        if isinstance(correzione, str) and correzione.strip() != frase.strip():
            applicata = _sostituisci_nei_testi(report, frase, correzione)
        voce = {
            "campo": str(riga.get("campo", "testo")),
            "esito": "testo_corretto" if applicata else "testo_da_verificare",
            "frase": frase[:400],
            "motivo": str(riga.get("problema", ""))[:400],
        }
        if applicata:
            voce["correzione"] = correzione[:400]
        registro.append(voce)
    return registro


CAMPI_TESTO = ("overview", "core_business", "mercati", "ownership", "operazioni_ma",
               "note_aggiuntive", "struttura_debito")


def _sostituisci_nei_testi(report: dict, frase: str, correzione: str) -> bool:
    """
    Sostituisce `frase` con `correzione` nei campi di testo del report, solo se la
    frase compare esattamente una volta (altrimenti la correzione resta da fare a mano).
    """
    trovati = []

    def cerca(contenitore, chiave):
        valore = contenitore[chiave]
        if isinstance(valore, str):
            if frase in valore:
                trovati.append((contenitore, chiave, valore.count(frase)))
        elif isinstance(valore, dict):
            for k in list(valore.keys()):
                cerca(valore, k)
        elif isinstance(valore, list):
            for i in range(len(valore)):
                cerca(valore, i)

    for campo in CAMPI_TESTO:
        if campo in report:
            cerca(report, campo)
    if len(trovati) != 1 or trovati[0][2] != 1:
        return False
    contenitore, chiave, _n = trovati[0]
    nuovo = contenitore[chiave].replace(frase, correzione)
    contenitore[chiave] = re.sub(r"\s{2,}", " ", nuovo).strip()
    return True


def completa_calcoli(report: dict) -> list:
    """
    Aritmetica fatta dal programma (prima della normalizzazione delle unità):
    EBITDA civilistico dai componenti del conto economico.
    Restituisce eventuali note da mostrare.
    """
    note = []
    fin = report.get("dati_finanziari")
    if not isinstance(fin, dict):
        return note
    comp = report.get("componenti_ebitda") or {}
    a_b = _numero(comp.get("differenza_a_b"))
    b10 = _numero(comp.get("ammortamenti_svalutazioni_b10"))
    b12 = _numero(comp.get("accantonamenti_rischi_b12")) or 0.0
    b13 = _numero(comp.get("altri_accantonamenti_b13")) or 0.0

    civilistico = "civilistico" in str(report.get("tipo_bilancio", "")).lower()
    if a_b is not None and b10 is not None:
        calcolato = a_b + b10 + b12 + b13
        esistente = _numero(fin.get("ebitda"))
        if esistente is None or civilistico:
            if esistente is not None and abs(esistente - calcolato) > 0.01 * max(abs(calcolato), 1):
                note.append(
                    f"EBITDA indicato dal modello ({_it(esistente)}) sostituito con il valore "
                    f"ricalcolato dai componenti del conto economico ({_it(calcolato)})."
                )
            fin["ebitda"] = calcolato
            report.setdefault("fonti", {})["ebitda"] = {
                "p": (report.get("fonti", {}).get("ebit") or {}).get("p"),
                "voce": "EBITDA calcolato: (A − B) + B.10 + B.12 + B.13",
                "letto": "",
                "calcolo": f"{_it(a_b)} + {_it(b10)} + {_it(b12)} + {_it(b13)}",
            }
            report["ebitda_metodo"] = "calcolato"
            if _numero(fin.get("ebit")) is None:
                fin["ebit"] = a_b
            if _numero(fin.get("ammortamenti")) is None:
                fin["ammortamenti"] = b10
    if fin.get("ebitda") is not None and "ebitda_metodo" not in report:
        report["ebitda_metodo"] = "esposto"

    dichiarato = _numero(report.get("ebitda_dichiarato"))
    usato = _numero(fin.get("ebitda"))
    if (report.get("ebitda_metodo") == "calcolato" and dichiarato is not None and usato is not None
            and abs(dichiarato - usato) > 0.01 * max(abs(usato), 1)):
        note.append(
            f"La relazione sulla gestione dichiara un EBITDA/MOL di {_it(dichiarato)}, diverso da "
            f"quello calcolato dal conto economico ({_it(usato)}) usato nel report (unità del bilancio)."
        )
    return note


def completa_leva(report: dict) -> None:
    """
    Da chiamare DOPO la normalizzazione a euro. Calcola la leva nel codice e
    ne esplicita la definizione, così non resta ambigua (PFN/EBITDA e non PFN/PN).
    """
    fin = report.get("dati_finanziari") or {}
    deb = report.get("struttura_debito")
    if not isinstance(deb, dict):
        return
    pfn = _numero(deb.get("debito_netto"))
    ebitda = _numero(fin.get("ebitda"))
    dichiarata = _numero(deb.get("leva_dichiarata"))
    definizione = str(deb.get("leva_dichiarata_definizione") or "").strip()

    def x(valore):
        return f"{valore:.2f}".replace(".", ",") + "x"

    calcolata, descrizione = None, ""
    if pfn is not None and ebitda is not None:
        if pfn <= 0:
            descrizione = "Leva (PFN/EBITDA) non significativa: posizione di cassa netta."
        elif ebitda <= 0:
            descrizione = "Leva (PFN/EBITDA) non significativa: EBITDA nullo o negativo."
        else:
            calcolata = round(pfn / ebitda, 2)
            descrizione = f"Leva: PFN/EBITDA = {x(calcolata)}."
    else:
        descrizione = "Leva (PFN/EBITDA) non calcolabile: PFN o EBITDA non disponibili."

    if dichiarata is not None:
        descrizione += (f" Leva dichiarata dalla società: {x(dichiarata)}"
                        + (f" ({definizione.rstrip('. ')})." if definizione else "."))

    deb["leva_finanziaria"] = calcolata
    deb["leva_definizione"] = descrizione
    nota = str(deb.get("note") or "").strip()
    if nota in ("", "N/D"):
        deb["note"] = descrizione
    elif descrizione not in nota:
        deb["note"] = f"{nota} | {descrizione}"


CAMPI_PRECEDENTE = KPI_TUTTI  # stessi campi dell'esercizio corrente


def completa_precedente(report: dict) -> None:
    """
    Prima della normalizzazione: mette in ordine "dati_anno_precedente"
    (numeri puri, EBITDA civilistico dai componenti).
    """
    prec = report.get("dati_anno_precedente")
    if not isinstance(prec, dict):
        report["dati_anno_precedente"] = {}
        return
    pulito = {"anno": str(prec.get("anno") or "").strip()}
    for campo in CAMPI_PRECEDENTE:
        pulito[campo] = _numero(prec.get(campo))
    a_b = _numero(prec.get("differenza_a_b"))
    b10 = _numero(prec.get("ammortamenti_svalutazioni_b10"))
    civilistico = "civilistico" in str(report.get("tipo_bilancio", "")).lower()
    if a_b is not None and b10 is not None and (pulito["ebitda"] is None or civilistico):
        pulito["ebitda"] = (a_b + b10 + (_numero(prec.get("accantonamenti_rischi_b12")) or 0.0)
                            + (_numero(prec.get("altri_accantonamenti_b13")) or 0.0))
        if pulito["ebit"] is None:
            pulito["ebit"] = a_b
        if pulito["ammortamenti"] is None:
            pulito["ammortamenti"] = b10
    report["dati_anno_precedente"] = pulito


def normalizza_precedente(report: dict) -> None:
    """
    Dopo la normalizzazione a euro dell'esercizio corrente: porta in euro anche
    l'esercizio precedente con lo stesso fattore e scarta i confronti impossibili.
    """
    prec = report.get("dati_anno_precedente")
    fattore = (report.get("dati_finanziari") or {}).get("_fattore")
    if not isinstance(prec, dict) or not fattore or prec.get("_normalizzato"):
        return
    for campo in CAMPI_PRECEDENTE:
        if prec.get(campo) is not None:
            prec[campo] = prec[campo] * fattore
    prec["_normalizzato"] = True
    if not any(prec.get(c) is not None for c in CAMPI_PRECEDENTE):
        report["dati_anno_precedente"] = {}


_RE_FRASE = re.compile(r"(?<=[.!?])\s+(?=[A-ZÀ-Ü0-9(])")


def _senza_ripetizioni(testo: str) -> str:
    """Toglie le frasi ripetute identiche (o quasi) all'interno dello stesso testo."""
    frasi = _RE_FRASE.split(testo)
    if len(frasi) < 2:
        return testo
    viste, tenute = [], []
    for frase in frasi:
        chiave = re.sub(r"[^a-zà-ú0-9]+", " ", frase.lower()).strip()
        parole = set(chiave.split())
        doppia = False
        if len(parole) >= 6:
            for altra in viste:
                comuni = len(parole & altra)
                if comuni >= 0.85 * max(len(parole), len(altra)):
                    doppia = True
                    break
        if not doppia:
            tenute.append(frase)
            if len(parole) >= 6:
                viste.append(parole)
    return " ".join(tenute)


def pulisci_testi(report: dict) -> None:
    """Rimuove le frasi ripetute nei campi di testo."""
    def visita(valore):
        if isinstance(valore, str):
            return _senza_ripetizioni(valore)
        if isinstance(valore, list):
            return [visita(v) for v in valore]
        if isinstance(valore, dict):
            return {k: (visita(v) if k not in ("fonte",) else v) for k, v in valore.items()}
        return valore
    for campo in ("overview", "core_business", "mercati", "note_aggiuntive", "operazioni_ma", "ownership"):
        if campo in report:
            report[campo] = visita(report[campo])
    deb = report.get("struttura_debito")
    if isinstance(deb, dict):
        for campo in ("scadenze_principali",):
            if isinstance(deb.get(campo), str):
                deb[campo] = _senza_ripetizioni(deb[campo])
        if isinstance(deb.get("note"), str):
            # le note sono voci separate da " | ": si tolgono le voci doppie
            voci, viste = [], set()
            for voce in [v.strip() for v in deb["note"].split("|") if v.strip()]:
                chiave = re.sub(r"\W+", " ", voce.lower()).strip()
                if chiave not in viste:
                    viste.add(chiave)
                    voci.append(voce)
            deb["note"] = " | ".join(voci)


def rimappa_fonti(report: dict, pagine_originali: list) -> None:
    """Converte "p" (posizione nell'allegato) nel numero di pagina del PDF originale."""
    fonti = report.get("fonti")
    if not isinstance(fonti, dict):
        report["fonti"] = {}
        return
    for campo, fonte in list(fonti.items()):
        if not isinstance(fonte, dict):
            fonti.pop(campo)
            continue
        p = fonte.get("p")
        try:
            p = int(p)
        except (TypeError, ValueError):
            p = None
        if p is not None and 1 <= p <= len(pagine_originali):
            fonte["pagina"] = pagine_originali[p - 1] + 1
        else:
            fonte["pagina"] = None
        fonte.pop("p", None)


def controlli_fonti(report: dict) -> list:
    """Avvisi aggiuntivi: numeri senza fonte e dati chiave mancanti."""
    avvisi = []
    fonti = report.get("fonti") or {}
    fin = report.get("dati_finanziari") or {}
    deb = report.get("struttura_debito") or {}
    if not isinstance(deb, dict):
        deb = {}

    senza_fonte = []
    for campo in KPI_TUTTI:
        valore = (fin if campo in KPI_FINANZIARI else deb).get(campo)
        if valore is None or valore == 0:
            continue
        fonte = fonti.get(campo) or {}
        if not fonte.get("pagina") and not fonte.get("calcolo"):
            senza_fonte.append(ETICHETTE_KPI[campo])
    if senza_fonte:
        avvisi.append({
            "livello": "warning", "campo": "Valori senza pagina di origine",
            "messaggio": ", ".join(senza_fonte) + " — da controllare sul bilancio",
        })

    mancanti = [
        ETICHETTE_KPI[c] for c in ("ricavi", "ebitda", "utile_netto", "totale_attivo",
                                   "patrimonio_netto", "debito_netto")
        if (fin if c in KPI_FINANZIARI else deb).get(c) is None
    ]
    if mancanti:
        motivi = {
            str(d.get("campo", "")): str(d.get("motivo", ""))
            for d in report.get("dati_non_trovati", []) or [] if isinstance(d, dict)
        }
        dettaglio = "; ".join(f"{k}: {v}" for k, v in motivi.items() if v)[:300]
        avvisi.append({
            "livello": "warning", "campo": "Dati non trovati",
            "messaggio": ", ".join(mancanti) + (f" — {dettaglio}" if dettaglio else "")
                         + ". Se il dato è nel bilancio, indica le pagine a mano e rigenera.",
        })
    return avvisi


def _estratto_per_verifica(report: dict) -> dict:
    """La parte del report che il revisore deve controllare."""
    chiavi = ["tipo_bilancio", "perimetro", "dati_finanziari", "componenti_ebitda",
              "ebitda_dichiarato", "dati_anno_precedente", "struttura_debito", "fonti", "overview", "core_business",
              "mercati", "ownership", "operazioni_ma", "note_aggiuntive"]
    return {k: report.get(k) for k in chiavi if k in report}


def genera_report_bilancio(client, pdf_ridotto: bytes, info: dict, nome: str, anno: str,
                           lingua_prompt: str, testo_suppl: str = "",
                           fonti_supplementari: list = None, perimetro: str = "auto",
                           verifica: bool = True, on_evento=None,
                           dettaglio: str = "completo") -> dict:
    """
    Esegue estrazione, verifica e calcoli e restituisce il report (dict),
    non ancora normalizzato a euro: la normalizzazione resta in app.py.

    Chiavi tecniche aggiunte al report:
      fonti                 per ogni numero: pagina del PDF originale, voce, valore letto
      _verifica             esiti della seconda lettura
      _note_calcolo         note sui ricalcoli fatti dal programma
      _estrazione           modello usato, pagine inviate, criterio di selezione
    """
    avvisa = on_evento or (lambda _msg: None)
    pagine_originali = list(info.get("selected_pages", []))
    pdf_corrente = pdf_ridotto

    if info.get("mode") == "scanned":
        nota = ("Il bilancio è una scansione: leggi i dati dalle immagini delle pagine. "
                f"L'allegato contiene {len(pagine_originali)} pagine su "
                f"{info.get('total_pages', 0)} del documento completo.")
    else:
        nota = (f"L'allegato contiene {len(pagine_originali)} pagine selezionate su "
                f"{info.get('total_pages', 0)} del documento completo: alcune sezioni possono mancare.")

    gradini = [g for g in GRADINI_PAGINE if g < len(pagine_originali)]
    testo = modello = None

    while True:
        prompt = costruisci_prompt_estrazione(
            nome, anno, lingua_prompt, len(pagine_originali), nota, testo_suppl,
            perimetro, fonti_supplementari or [], dettaglio,
        )
        documento = _blocco_documento(pdf_corrente)
        try:
            ultimo = None
            for tentativo in range(2):
                testo, modello, _uso = chiama_claude(
                    client, [documento, {"type": "text", "text": prompt}],
                    MAX_TOKENS_ESTRAZIONE, on_evento=avvisa,
                )
                try:
                    report = estrai_json(testo)
                    break
                except ErroreEstrazione as e:
                    ultimo = e
                    avvisa("Risposta non leggibile, riprovo una volta.")
            else:
                raise ultimo
            break
        except RichiestaTroppoGrande:
            if not gradini:
                raise ErroreEstrazione(
                    "La richiesta supera i limiti dell'account API anche con poche pagine. "
                    "Riprova fra qualche minuto oppure indica a mano le pagine da analizzare."
                )
            nuovo_max = gradini.pop(0)
            avvisa(f"Richiesta troppo grande per i limiti dell'account: riprovo con {nuovo_max} pagine.")
            info_corrente = dict(info, selected_pages=pagine_originali)
            pdf_corrente, pagine_originali = riduci_pdf(pdf_corrente, info_corrente, nuovo_max)
            nota = (f"L'allegato contiene {len(pagine_originali)} pagine selezionate su "
                    f"{info.get('total_pages', 0)} del documento completo: alcune sezioni possono mancare.")
            time.sleep(20)

    if not isinstance(report, dict):
        raise ErroreEstrazione("La risposta del modello non ha la struttura attesa.")
    for sezione in ("dati_finanziari", "struttura_debito", "ownership"):
        if not isinstance(report.get(sezione), dict):
            report[sezione] = {}
    if not isinstance(report.get("operazioni_ma"), list):
        report["operazioni_ma"] = []
    report["operazioni_ma"] = [
        op for op in report["operazioni_ma"]
        if isinstance(op, dict) and str(op.get("descrizione", "")).strip() not in ("", "N/D")
    ]

    # ── Seconda lettura indipendente ─────────────────────────────────────────
    registro = []
    verifica_eseguita = False
    if verifica:
        avvisa("Seconda lettura di controllo di numeri e testi...")
        try:
            prompt_v = costruisci_prompt_verifica(
                nome, anno, len(pagine_originali), _estratto_per_verifica(report), testo_suppl
            )
            testo_v, _m, _u = chiama_claude(
                client, [documento, {"type": "text", "text": prompt_v}],
                MAX_TOKENS_VERIFICA, modelli=[modello] if modello else None, on_evento=avvisa,
            )
            registro = applica_verifica(report, estrai_json(testo_v))
            verifica_eseguita = True
        except (ErroreEstrazione, RichiestaTroppoGrande) as e:
            registro = [{
                "campo": "verifica", "esito": "non_eseguita",
                "motivo": f"Seconda lettura non riuscita: {e}",
            }]

    for op in report["operazioni_ma"]:
        passaggi = op.get("passaggi")
        op["passaggi"] = ([str(p).strip() for p in passaggi if str(p).strip()]
                          if isinstance(passaggi, list) else [])
    note_calcolo = completa_calcoli(report)
    completa_precedente(report)
    pulisci_testi(report)
    rimappa_fonti(report, pagine_originali)
    report["_verifica"] = registro
    report["_note_calcolo"] = note_calcolo
    converti_riferimenti_pagina(report, pagine_originali, inglese="ingles" in lingua_prompt)

    report["_estrazione"] = {
        "modello": modello,
        "pagine_inviate": [p + 1 for p in pagine_originali],
        "pagine_totali": info.get("total_pages", 0),
        "modalita": info.get("mode", ""),
        "criterio": info.get("criterio", ""),
        "verifica_eseguita": verifica_eseguita,
        "dettaglio": dettaglio,
    }
    report["_lingua"] = "en" if "ingles" in lingua_prompt else "it"
    return report


def testo_da_modello_veloce(client, content: list, max_tokens: int = 8000) -> str:
    """Chiamata breve (es. lettura della visura) con il modello veloce."""
    testo, _modello, _uso = chiama_claude(client, content, max_tokens, modelli=MODELLI_VELOCI)
    return testo


# ── 7. Documenti supplementari: visura e Mergermarket ────────────────────────

LIMITI_SUPPLEMENTARI = {"Visura Camerale": 30000, "Mergermarket": 45000}
LIMITE_SUPPLEMENTARE_DEFAULT = 8000

_RE_RIGA_SEZIONE = re.compile(r"^\s*(\d{1,2})\s+([A-Za-zÀ-ú][^\d]{2,80})$")
_RUMORE_VISURA = (
    "registro imprese", "archivio ufficiale della cciaa", "documento n .",
    "estratto dal registro imprese in data", "visura ordinaria", "codice fiscale 0",
    "servizio realizzato da infocamere", "camera di commercio di",
)
# Caratteri riservati a ogni sezione della visura (il resto, soprattutto le
# decine di pagine di procure, viene tagliato).
_QUOTE_VISURA = [
    ("sede", 1500), ("statuto", 6000), ("capitale", 2500), ("soci", 9000),
    ("amministratori", 4500), ("sindaci", 3000), ("titolari di altre cariche", 300),
    ("trasferimenti", 5000), ("attivit", 2500), ("sedi secondarie", 600),
    ("aggiornamento", 500),
]


def leggi_visura(pdf_bytes: bytes, max_caratteri: int = 30000) -> str:
    """
    Riduce una visura camerale (anche di centinaia di pagine) alle parti utili:
    riepilogo iniziale, sede, statuto, capitale, soci e patti parasociali,
    amministratori, sindaci, operazioni straordinarie, attività.
    Restituisce "" se il PDF non ha testo (visura scansionata).
    """
    pagine = leggi_pagine(pdf_bytes)
    if not pagine or sum(p["caratteri"] for p in pagine) < 300 * len(pagine) * 0.3:
        return ""

    data_estrazione = ""
    righe = []
    for p in pagine:
        for riga in p["testo"].replace("\r", "\n").split("\n"):
            pulita = riga.strip()
            if not pulita:
                continue
            bassa = pulita.lower()
            if not data_estrazione:
                m = re.search(r"estratto dal registro imprese in data (\d{2}/\d{2}/\d{4})", bassa)
                if m:
                    data_estrazione = m.group(1)
            if any(bassa.startswith(r) for r in _RUMORE_VISURA) or re.fullmatch(r"[.\s\d•]+", pulita) \
                    or re.fullmatch(r"\.{5,}\s*\d*", pulita) or bassa.startswith("visura ") \
                    or re.fullmatch(r"\d+", pulita):
                continue
            righe.append(pulita)

    # Indice: sequenza 1, 2, 3... dopo la riga "Indice"
    titoli, fine_indice = [], 0
    for i, riga in enumerate(righe[:400]):
        if riga.lower() == "indice":
            atteso = 1
            for j in range(i + 1, min(i + 40, len(righe))):
                m = _RE_RIGA_SEZIONE.match(righe[j])
                if m and int(m.group(1)) == atteso:
                    titoli.append((atteso, m.group(2).strip()))
                    atteso += 1
                    fine_indice = j + 1
                elif titoli:
                    break
            break

    if len(titoli) < 3:
        # Struttura non riconosciuta: inizio del documento più i passaggi sui soci.
        testo = "\n".join(righe)
        return testo[:max_caratteri]

    inizi = []
    cursore = fine_indice
    for numero, titolo in titoli:
        chiave = f"{numero} {titolo}".lower()
        for j in range(cursore, len(righe)):
            if righe[j].lower() == chiave:
                inizi.append((titolo, j))
                cursore = j + 1
                break

    blocchi = []
    intestazione = "\n".join(righe[:fine_indice - len(titoli) - 1])[:3000]
    blocchi.append("RIEPILOGO\n" + intestazione)
    for k, (titolo, inizio) in enumerate(inizi):
        fine = inizi[k + 1][1] if k + 1 < len(inizi) else len(righe)
        corpo = "\n".join(righe[inizio + 1:fine])
        quota = 2000
        for parola, q in _QUOTE_VISURA:
            if parola in titolo.lower():
                quota = q
                break
        if len(corpo) > quota:
            corpo = corpo[:quota] + f"\n[... sezione abbreviata: omessi {len(corpo) - quota} caratteri]"
        blocchi.append(f"SEZIONE {titolo.upper()}\n{corpo}")

    testa = f"Visura camerale estratta dal Registro Imprese il {data_estrazione}.\n" if data_estrazione else ""
    return (testa + "\n\n".join(blocchi))[:max_caratteri]


def _righe_tabella(nome_file: str, contenuto: bytes) -> list:
    """Righe (liste di celle) di un file .xlsx o .csv."""
    nome = nome_file.lower()
    if nome.endswith((".xlsx", ".xlsm")):
        import warnings
        import openpyxl
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(io.BytesIO(contenuto), data_only=True, read_only=True)
        righe = []
        for ws in wb.worksheets:
            for riga in ws.iter_rows(values_only=True):
                righe.append(list(riga))
        wb.close()
        return righe
    import csv
    testo = contenuto.decode("utf-8-sig", errors="replace")
    try:
        dialetto = csv.Sniffer().sniff(testo[:4000], delimiters=",;\t")
    except csv.Error:
        dialetto = csv.excel
    return [riga for riga in csv.reader(io.StringIO(testo), dialetto)]


def _cella(v) -> str:
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float):
        return f"{v:,.2f}".rstrip("0").rstrip(".")
    return str(v).strip()


def leggi_mergermarket(nome_file: str, contenuto: bytes, max_caratteri: int = 45000) -> str:
    """
    Trasforma l'export Mergermarket (.xlsx, .csv o .pdf) in un elenco di
    operazioni leggibile. Tutte le operazioni vengono elencate; se il testo
    supera il limite, le descrizioni delle operazioni più vecchie sono abbreviate.
    """
    if nome_file.lower().endswith(".pdf"):
        testo = "\n".join(p["testo"] for p in leggi_pagine(contenuto))
        return testo[:max_caratteri]

    righe = _righe_tabella(nome_file, contenuto)
    indice_intestazione = None
    for i, riga in enumerate(righe):
        celle = [_cella(c).lower() for c in riga]
        if "target" in celle and any("announced" in c for c in celle):
            indice_intestazione = i
            break
    if indice_intestazione is None:
        testo = "\n".join(" | ".join(_cella(c) for c in riga if _cella(c)) for riga in righe)
        return testo[:max_caratteri]

    intestazione = [_cella(c) for c in righe[indice_intestazione]]
    col = {nome.lower(): k for k, nome in enumerate(intestazione) if nome}

    def campo(riga, *nomi):
        for nome in nomi:
            k = col.get(nome)
            if k is not None and k < len(riga):
                valore = _cella(riga[k])
                if valore:
                    return valore
        return ""

    operazioni = []
    for riga in righe[indice_intestazione + 1:]:
        target = campo(riga, "target")
        if not target:
            continue
        testa = (
            f"Annunciata: {campo(riga, 'announced') or 'n.d.'} | "
            f"Completata: {campo(riga, 'completed') or 'non risulta completata'} | "
            f"Target: {target} | Acquirente: {campo(riga, 'acquirer', 'bidder') or 'n.d.'} | "
            f"Venditore: {campo(riga, 'divestor', 'seller', 'vendor') or 'n.d.'}"
        )
        valore = campo(riga, "deal value $m", "deal value usd(m)", "deal value")
        if valore:
            testa += f" | Valore ($m): {valore}"
        operazioni.append((testa, campo(riga, "deal description", "description")))

    quota_descrizioni = max_caratteri - sum(len(t) + 60 for t, _ in operazioni) - 300
    blocchi = []
    for n, (testa, descrizione) in enumerate(operazioni, start=1):
        if descrizione and quota_descrizioni > 0:
            if len(descrizione) > quota_descrizioni:
                descrizione = descrizione[:max(quota_descrizioni, 0)] + " [descrizione abbreviata]"
            quota_descrizioni -= len(descrizione)
        elif descrizione:
            descrizione = descrizione[:250] + " [descrizione abbreviata]"
        blocchi.append(f"OPERAZIONE {n}\n{testa}\nDescrizione: {descrizione or 'n.d.'}")

    premessa = (f"Export Mergermarket: {len(operazioni)} operazioni, dalla più recente. "
                f"L'elenco è completo; ogni descrizione riporta gli aggiornamenti in ordine cronologico.\n\n")
    return (premessa + "\n\n".join(blocchi))[:max_caratteri + 2000]


def componi_supplementari(testi: dict) -> str:
    """Testo dei documenti supplementari per il prompt, ciascuno entro il suo limite."""
    parti = []
    for fonte, testo in testi.items():
        limite = LIMITI_SUPPLEMENTARI.get(fonte, LIMITE_SUPPLEMENTARE_DEFAULT)
        corpo = testo or ""
        if len(corpo) > limite:
            corpo = corpo[:limite] + "\n[... testo abbreviato dal programma]"
        parti.append(f"\n\n--- DOCUMENTO SUPPLEMENTARE: {fonte} ---\n{corpo}")
    return "".join(parti)


# ── 8. Riferimenti di pagina nei testi ───────────────────────────────────────

_NUM = r"\d{1,3}(?!\d)"
_ELENCO = rf"{_NUM}(?:\s?[-–]\s?{_NUM})?(?:\s?(?:,|;|\be\b|\band\b)\s?{_NUM}(?:\s?[-–]\s?{_NUM})?)*"
# "p. 34", "pp. 34-36", "pagina 34", "pages 3, 5 and 9" scritti senza marcatore
_RE_RIF_NUDO = re.compile(
    rf"(?<![\w\[.])(?:pp?\.|pagg?\.|pagin[ae]|pages?)\s?({_ELENCO})(?!\s*\])", re.IGNORECASE)
_RE_RIF_MARCATORE = re.compile(r"\[\s*(?:pp?|pagg?|pagin[ae]|pages?)\.?\s*([\d\s,;\-–eand]+?)\s*\]",
                               re.IGNORECASE)
_RE_MARCATORI_VICINI = re.compile(r"\[p\.([\d,\-]+)\](?:\s*[,;]?\s*(?:e|and)?\s*\[p\.([\d,\-]+)\])+")
_RE_SUFFISSO_ALLEGATO = re.compile(r"(\[p\.[\d,\-]+\])\s*(?:dell['’]allegato|of the attachment)", re.IGNORECASE)


# "a [p.34]", "vedi [p.3]", "di [p.5]": il riferimento fa parte della frase, niente parentesi
_RE_PREPOSIZIONE = (r"(?:\b(?:a|ad|di|da|in|su|alla|alle|nella|nelle|della|delle|dalla|vedi|cfr\.?|"
                    r"at|on|in|see|of|from)\s)$")


def _posizioni(elenco: str) -> list:
    """ "3, 5-7 e 9" -> [3, 5, 6, 7, 9] """
    numeri = []
    for parte in re.split(r"[,;]|\be\b|\band\b", elenco):
        m = re.fullmatch(r"\s*(\d{1,3})\s*(?:[-–]\s*(\d{1,3}))?\s*", parte)
        if not m:
            continue
        a, b = int(m.group(1)), int(m.group(2) or m.group(1))
        if a > b:
            a, b = b, a
        numeri.extend(range(a, min(b, a + 60) + 1))
    return numeri


def _pagine_in_parole(pagine: list, inglese: bool) -> str:
    """[65, 66, 67, 69] -> "pp. 65-67 e 69 del PDF" """
    pagine = sorted(set(pagine))
    gruppi, inizio, prec = [], pagine[0], pagine[0]
    for p in pagine[1:]:
        if p == prec + 1:
            prec = p
            continue
        gruppi.append((inizio, prec))
        inizio = prec = p
    gruppi.append((inizio, prec))
    pezzi = [str(a) if a == b else f"{a}-{b}" for a, b in gruppi]
    congiunzione = " and " if inglese else " e "
    elenco = pezzi[0] if len(pezzi) == 1 else ", ".join(pezzi[:-1]) + congiunzione + pezzi[-1]
    sigla = "p." if len(pagine) == 1 else "pp."
    return f"PDF {sigla} {elenco}" if inglese else f"{sigla} {elenco} del PDF"


def _converti_testo(testo: str, pagine_originali: list, inglese: bool, parentesi: bool) -> str:
    if not re.search(r"\d", testo):
        return testo
    # 1. tutto ciò che è un riferimento di pagina diventa un marcatore [p.N,...]
    testo = _RE_RIF_MARCATORE.sub(
        lambda m: "[p." + ",".join(str(n) for n in _posizioni(m.group(1))) + "]", testo)
    def nudo(m):
        # un riferimento senza marcatore si converte solo se può essere una pagina dell'allegato
        numeri = _posizioni(m.group(1))
        prima = m.string[max(0, m.start() - 30):m.start()].lower()
        dopo = m.string[m.end():m.end() + 8].lower()
        if (not numeri or max(numeri) > len(pagine_originali)
                or re.match(r"\s*(?:di|of)\s+\d", dopo)            # "p. 3 di 347"
                or "visura" in prima or "mergermarket" in prima):  # pagina di un altro documento
            return m.group(0)
        return "[p." + ",".join(str(n) for n in numeri) + "]"
    testo = _RE_RIF_NUDO.sub(nudo, testo)
    testo = _RE_SUFFISSO_ALLEGATO.sub(r"\1", testo)
    # 2. marcatori adiacenti: uno solo
    def unisci(m):
        numeri = re.findall(r"\d+", m.group(0))
        return "[p." + ",".join(numeri) + "]"
    testo = _RE_MARCATORI_VICINI.sub(unisci, testo)
    # 3. dal numero d'ordine nell'allegato alla pagina del PDF
    def converti(m):
        originali = [pagine_originali[n - 1] + 1 for n in map(int, re.findall(r"\d+", m.group(0)))
                     if 1 <= n <= len(pagine_originali)]
        if not originali:
            return "(pagina non identificata)" if not inglese else "(page not identified)"
        parole = _pagine_in_parole(originali, inglese)
        prima = m.string[:m.start()]
        dentro_parentesi = prima.count("(") > prima.count(")")
        dopo_preposizione = re.search(_RE_PREPOSIZIONE, prima) is not None
        if not parentesi or dentro_parentesi or dopo_preposizione:
            return parole
        return f"({parole})"
    testo = re.sub(r"\[p\.[\d,\-]*\]", converti, testo)
    # 4. pulizia degli spazi attorno ai riferimenti
    testo = re.sub(r"(?<=[^\s(])\((?=(?:PDF )?pp?\. )", " (", testo)
    testo = re.sub(r"\s+([,.;:)])", r"\1", testo)
    testo = re.sub(r"\(\s+", "(", testo)
    return re.sub(r" {2,}", " ", testo).strip()


def converti_riferimenti_pagina(report: dict, pagine_originali: list, inglese: bool = False) -> None:
    """
    Nei testi del report sostituisce ogni riferimento alle pagine dell'allegato
    ("[p.12]", "p. 12", "pp. 3-5 dell'allegato") con le pagine del PDF originale,
    unendo i riferimenti vicini: "(pp. 363-364 e 395 del PDF)".
    """
    def visita(valore, parentesi=True):
        if isinstance(valore, str):
            return _converti_testo(valore, pagine_originali, inglese, parentesi)
        if isinstance(valore, list):
            return [visita(v, parentesi) for v in valore]
        if isinstance(valore, dict):
            # nel campo "fonte" il riferimento va senza parentesi: "Bilancio pp. 6-7 del PDF"
            return {k: visita(v, parentesi and k != "fonte") for k, v in valore.items()}
        return valore

    for chiave in list(report.keys()):
        if chiave in ("fonti", "_estrazione", "dati_finanziari", "componenti_ebitda"):
            continue
        report[chiave] = visita(report[chiave])
