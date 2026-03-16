"""
uploader.py — Componente Streamlit per upload multiplo di bilanci.

Gestisce:
  - Upload di 1-5 PDF contemporaneamente
  - Form metadati per ogni PDF (ragione sociale, anno, tipo bilancio)
  - Analisi automatica delle pagine con pdf_processor
  - Debug panel espandibile con l'analisi per ogni PDF
  - Stato nella session_state di Streamlit

Come usarlo in app.py:
    from uploader import render_upload_section
    pdf_queue = render_upload_section()
    # pdf_queue è una lista di dict, uno per PDF pronto al processing
"""

import base64
import streamlit as st
from pdf_processor import process_pdf_for_claude


def _tipo_bilancio_options():
    return [
        "Bilancio consolidato",
        "Bilancio separato / individuale",
        "Bilancio abbreviato",
        "Visura camerale",
        "Non so / Altro",
    ]


def _anno_options():
    """Anni dal 2018 al 2025 in ordine decrescente."""
    return [str(y) for y in range(2025, 2017, -1)]


def _render_analysis_badge(info: dict):
    """Mostra un badge colorato con il risultato dell'analisi pagine."""
    mode = info.get("mode", "unknown")
    total = info.get("total_pages", 0)
    selected = len(info.get("selected_pages", []))

    if mode == "scanned":
        st.warning(
            f"📷 **PDF scansionato** — {total} pagine totali. "
            "Claude leggerà tutto il documento (può essere più lento)."
        )
    elif mode == "fallback":
        st.warning(
            f"⚠️ **Layout atipico** — non ho trovato le sezioni standard. "
            f"Mando le {selected} pagine centrali su {total} totali."
        )
    else:
        reduction = info.get("size_reduction_pct", 0)
        st.success(
            f"✅ **{selected} pagine rilevanti** trovate su {total} totali "
            f"({reduction}% riduzione). "
            f"Sezioni: {', '.join(info.get('sections_found', ['—']))}."
        )


def _render_debug_expander(info: dict, filename: str):
    """Pannello di debug espandibile con dettaglio pagina per pagina."""
    with st.expander(f"🔍 Dettaglio analisi — {filename}", expanded=False):
        scores = info.get("page_scores", [])
        selected = set(info.get("selected_pages", []))

        if not scores:
            st.write("Nessun dato disponibile.")
            return

        st.caption(
            f"Soglia punteggio: 8 | "
            f"Caratteri medi per pagina: {info.get('avg_chars', 0):.0f}"
        )

        # Mostra solo le prime 50 pagine per non appesantire l'interfaccia
        max_show = min(50, len(scores))
        cols = st.columns(5)
        for i in range(max_show):
            col = cols[i % 5]
            score = scores[i]
            included = i in selected
            color = "#c8e04a" if included else "#444"
            bg = "#2a2a2a" if included else "#1e1e1e"
            col.markdown(
                f"""<div style="
                    background:{bg}; border:1px solid {color};
                    border-radius:6px; padding:6px 4px;
                    text-align:center; margin-bottom:6px;
                    font-size:11px; color:{color};">
                    <b>p.{i+1}</b><br>{score}pt
                </div>""",
                unsafe_allow_html=True,
            )

        if len(scores) > max_show:
            st.caption(f"... e altre {len(scores) - max_show} pagine.")


def render_upload_section() -> list:
    """
    Renderizza la sezione upload e restituisce la coda di PDF pronti
    per il processing da parte di Claude.

    Ogni elemento della lista è un dict:
    {
        "ragione_sociale": str,
        "anno": str,
        "tipo_bilancio": str,
        "pdf_bytes": bytes,        # PDF chirurgico da mandare a Claude
        "pdf_b64": str,            # base64 del PDF chirurgico
        "original_filename": str,
        "analysis_info": dict,     # output di process_pdf_for_claude
    }
    """
    st.markdown("### 📂 Carica i bilanci")
    st.caption(
        "Puoi caricare fino a 5 PDF contemporaneamente. "
        "Puoi caricare l'intero bilancio — il sistema estrarrà "
        "automaticamente solo le pagine finanziarie rilevanti."
    )

    uploaded_files = st.file_uploader(
        "Seleziona uno o più PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key="multi_pdf_uploader",
        help="Bilanci completi, visure camerali, bilanci consolidati o separati.",
    )

    if not uploaded_files:
        st.info("Carica almeno un PDF per procedere.")
        return []

    if len(uploaded_files) > 5:
        st.error("Massimo 5 PDF alla volta. Rimuovi qualche file.")
        return []

    st.markdown("---")
    st.markdown("#### Conferma i dettagli per ogni documento")

    queue = []

    for idx, f in enumerate(uploaded_files):
        st.markdown(f"**📄 {f.name}**")

        col_meta1, col_meta2, col_meta3 = st.columns([3, 1, 2])

        with col_meta1:
            ragione_sociale = st.text_input(
                "Ragione sociale",
                key=f"rs_{idx}",
                placeholder="es. Zambon S.p.A.",
            )
        with col_meta2:
            anno = st.selectbox(
                "Anno",
                options=_anno_options(),
                key=f"anno_{idx}",
            )
        with col_meta3:
            tipo = st.selectbox(
                "Tipo documento",
                options=_tipo_bilancio_options(),
                key=f"tipo_{idx}",
            )

        # Analisi pagine — eseguita una volta sola e messa in cache
        # nella session_state per non rifare ogni rerun
        cache_key = f"analysis_{f.name}_{f.size}"
        if cache_key not in st.session_state:
            with st.spinner(f"Analisi pagine di {f.name}..."):
                pdf_bytes = f.read()
                surgical_pdf, info = process_pdf_for_claude(pdf_bytes)
                st.session_state[cache_key] = {
                    "surgical_pdf": surgical_pdf,
                    "info": info,
                    "original_bytes": pdf_bytes,
                }
        else:
            # Rewind non necessario — usiamo la cache
            surgical_pdf = st.session_state[cache_key]["surgical_pdf"]
            info = st.session_state[cache_key]["info"]

        # Badge risultato analisi
        _render_analysis_badge(info)

        # Debug panel
        _render_debug_expander(info, f.name)

        # Aggiungi alla coda solo se la ragione sociale è stata inserita
        if ragione_sociale.strip():
            queue.append({
                "ragione_sociale": ragione_sociale.strip(),
                "anno": anno,
                "tipo_bilancio": tipo,
                "pdf_bytes": surgical_pdf,
                "pdf_b64": base64.b64encode(surgical_pdf).decode(),
                "original_filename": f.name,
                "analysis_info": info,
            })
        else:
            st.warning("⬆️ Inserisci la ragione sociale per includere questo documento.")

        st.markdown("---")

    # Riepilogo finale
    if queue:
        st.success(
            f"✅ **{len(queue)} documento/i pronti** per il processing. "
            f"Premi **Genera Report** per continuare."
        )
    elif uploaded_files:
        st.warning("Completa i metadati per tutti i documenti prima di procedere.")

    return queue
