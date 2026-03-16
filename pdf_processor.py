"""
pdf_processor.py — Estrazione intelligente pagine da bilanci italiani completi.

Strategia in due fasi:
  1. pdfplumber estrae il testo di tutte le pagine (gratis, istantaneo)
  2. Si trovano le pagine con le sezioni finanziarie tramite keyword matching
  3. Si costruisce un PDF "chirurgico" con solo quelle pagine
  4. Il PDF chirurgico viene mandato a Claude (invece del PDF completo)

Fallback per PDF scansionati: se pdfplumber trova meno di 100 caratteri
per pagina in media, assume PDF scansionato e manda tutto a Claude
con un avviso all'utente.
"""

import io
import pdfplumber
from pypdf import PdfReader, PdfWriter


# ── Keyword per trovare le sezioni rilevanti ──────────────────────────────────

# Ogni entry è (keyword_lowercase, peso)
# Le pagine vengono scorrate, si sommano i pesi delle keyword trovate,
# e si selezionano le pagine con punteggio > soglia.
FINANCIAL_KEYWORDS = [
    # Sezioni principali — peso alto
    ("stato patrimoniale",          10),
    ("conto economico",             10),
    ("rendiconto finanziario",      10),
    ("prospetto delle variazioni",   8),
    ("nota integrativa",             6),
    ("posizione finanziaria netta",  9),
    ("indebitamento finanziario",    9),
    ("pfn",                          7),

    # Voci di bilancio — peso medio
    ("ricavi",                       5),
    ("ricavi delle vendite",         6),
    ("proventi",                     4),
    ("ebitda",                       8),
    ("ebit",                         7),
    ("utile netto",                  6),
    ("perdita netta",                6),
    ("risultato netto",              6),
    ("risultato operativo",          6),
    ("ammortamenti",                 5),
    ("svalutazioni",                 4),

    # Stato patrimoniale
    ("totale attivo",                6),
    ("totale passivo",               6),
    ("patrimonio netto",             7),
    ("capitale sociale",             5),
    ("debiti finanziari",            7),
    ("disponibilità liquide",        5),
    ("cassa e mezzi",                5),

    # Indicatori e indici
    ("margine",                      4),
    ("roe",                          5),
    ("roa",                          5),
    ("leverage",                     5),
    ("debt/equity",                  5),

    # Header di tabella numerica — segnale forte
    ("(migliaia di euro)",           8),
    ("(milioni di euro)",            8),
    ("in migliaia",                  7),
    ("in milioni",                   7),
    ("esercizio chiuso",             6),
    ("31 dicembre",                  5),
    ("31/12/",                       5),
]

# Soglia minima di punteggio perché una pagina venga inclusa
SCORE_THRESHOLD = 8

# Quante pagine di contesto prendere prima e dopo ogni pagina rilevante
CONTEXT_PAGES = 1


def _score_page(text: str) -> int:
    """Calcola il punteggio finanziario di una pagina."""
    text_lower = text.lower()
    score = 0
    for keyword, weight in FINANCIAL_KEYWORDS:
        if keyword in text_lower:
            score += weight
    return score


def find_financial_pages(pdf_bytes: bytes) -> dict:
    """
    Analizza un PDF e restituisce un dizionario con:
      - 'mode': 'text' | 'scanned'
      - 'total_pages': int
      - 'selected_pages': list[int]  (0-indexed)
      - 'page_scores': list[int]
      - 'avg_chars': float
      - 'sections_found': list[str]  (nomi sezioni identificate)
    """
    result = {
        "mode": "text",
        "total_pages": 0,
        "selected_pages": [],
        "page_scores": [],
        "avg_chars": 0.0,
        "sections_found": [],
    }

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total = len(pdf.pages)
        result["total_pages"] = total

        texts = []
        scores = []
        char_counts = []

        for page in pdf.pages:
            t = page.extract_text() or ""
            texts.append(t)
            scores.append(_score_page(t))
            char_counts.append(len(t))

        result["page_scores"] = scores
        avg_chars = sum(char_counts) / total if total > 0 else 0
        result["avg_chars"] = avg_chars

        # Rilevamento PDF scansionato
        if avg_chars < 100:
            result["mode"] = "scanned"
            result["selected_pages"] = list(range(total))
            return result

        # Selezione pagine rilevanti con contesto
        relevant = set()
        for i, score in enumerate(scores):
            if score >= SCORE_THRESHOLD:
                for j in range(
                    max(0, i - CONTEXT_PAGES),
                    min(total, i + CONTEXT_PAGES + 1)
                ):
                    relevant.add(j)

        # Se non troviamo nulla (bilancio con layout atipico),
        # prendiamo le 40 pagine centrali come fallback
        if not relevant:
            start = max(0, total // 4)
            end = min(total, 3 * total // 4)
            relevant = set(range(start, end))
            result["mode"] = "fallback"

        result["selected_pages"] = sorted(relevant)

        # Identifica le sezioni trovate (per il debug panel)
        sections = []
        main_sections = [
            ("stato patrimoniale", "Stato Patrimoniale"),
            ("conto economico", "Conto Economico"),
            ("rendiconto finanziario", "Rendiconto Finanziario"),
            ("posizione finanziaria netta", "PFN"),
            ("nota integrativa", "Nota Integrativa"),
        ]
        all_text = " ".join(texts).lower()
        for kw, label in main_sections:
            if kw in all_text:
                sections.append(label)
        result["sections_found"] = sections

    return result


def build_surgical_pdf(pdf_bytes: bytes, page_indices: list) -> bytes:
    """
    Costruisce un nuovo PDF contenente solo le pagine indicate (0-indexed).
    Restituisce i bytes del PDF ridotto.
    """
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()

    for i in page_indices:
        if 0 <= i < len(reader.pages):
            writer.add_page(reader.pages[i])

    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()


def process_pdf_for_claude(pdf_bytes: bytes) -> tuple[bytes, dict]:
    """
    Entry point principale. Dato un PDF completo, restituisce:
      - pdf_to_send: bytes — il PDF da mandare a Claude (ridotto o completo)
      - info: dict — informazioni sull'analisi (per mostrare all'utente)
    """
    info = find_financial_pages(pdf_bytes)

    if info["mode"] == "scanned":
        # PDF scansionato: manda tutto, Claude legge nativamente
        return pdf_bytes, info

    # PDF con testo: costruisci il PDF chirurgico
    surgical = build_surgical_pdf(pdf_bytes, info["selected_pages"])
    info["surgical_pages"] = len(info["selected_pages"])
    info["size_reduction_pct"] = round(
        (1 - len(surgical) / len(pdf_bytes)) * 100, 1
    )

    return surgical, info
