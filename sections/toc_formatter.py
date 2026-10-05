"""
toc_formatter.py
"TOC Formatter" page: embeds the standalone TOC Formatter Tool (originally
its own separate Streamlit app, at github.com/akshay-click-seo/toc-formatter-
tool) as its own menu item here. The tool's own parsing/rendering logic
(toc_tool/formatter.py + toc_tool/docx_io.py, copied over unmodified apart
from one renamed internal import) is used exactly as-is -- this file only
builds the UI layer, styled to match this app's own navy/teal brand (the
same palette the generated report images use) instead of default Streamlit
widget styling, and adapted to fit as ONE PAGE inside this app's existing
single-sidebar-navigated shell (the original tool was a full standalone app
of its own, with its own st.set_page_config and its own sidebar).
"""

import os
import sys

import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from toc_tool import docx_io
from utils.page_header import render_hero, NAVY, TEAL


FORMAT_LABELS = {
    "txt": "📄 Plain Text (.txt)",
    "md": "📝 Markdown (.md)",
    "json": "📊 JSON (.json)",
}

# Scoped page CSS -- re-injected fresh every time this page renders (like
# RD Description's own <style> block), so it never bleeds into the other
# pages' look when navigating away.
_PAGE_CSS = f"""
<style>
/* bordered containers -> soft cards matching the app's navy/teal identity */
div[data-testid="stVerticalBlockBorderWrapper"] {{
    border-radius: 14px !important;
    border-color: #d8e0f2 !important;
    background: #fbfcff;
    box-shadow: 0 1px 2px rgba(11,47,122,.04), 0 6px 18px rgba(11,47,122,.05);
}}

/* segmented control ("Modo") -- navy fill on the selected pill */
div[data-testid="stButtonGroup"] button[aria-checked="true"] {{
    background-color: {NAVY} !important;
    border-color: {NAVY} !important;
    color: #fff !important;
}}
div[data-testid="stButtonGroup"] button:hover {{
    border-color: {NAVY} !important;
    color: {NAVY} !important;
}}

/* primary buttons (Convertir Todos, downloads) -> navy-to-teal gradient */
button[kind="primary"] {{
    background: linear-gradient(120deg, {NAVY} 0%, #123c9c 60%, {TEAL} 150%) !important;
    border: none !important;
}}
button[kind="primary"]:hover {{
    filter: brightness(1.08);
}}

/* metric values in the summary card */
div[data-testid="stMetricValue"] {{
    color: {NAVY};
}}

/* section card headers */
.toc-card-head {{
    display:flex;
    align-items:center;
    gap:8px;
    font-weight:700;
    font-size:15px;
    color:{NAVY};
    margin-bottom:10px;
}}
.toc-badge {{
    display:inline-block;
    font-size:11.5px;
    font-weight:700;
    letter-spacing:.03em;
    color:#fff;
    background:{TEAL};
    padding:2px 10px;
    border-radius:20px;
    margin-left:auto;
}}

.toc-preview {{
    background-color:#161a23;
    border:1px solid #2a2f3a;
    border-radius:10px;
    padding:1.2rem 1.5rem;
    font-family:'JetBrains Mono','Courier New',monospace;
    font-size:0.92rem;
    white-space:pre-wrap;
    color:#e6e6e6;
    max-height:560px;
    overflow-y:auto;
}}
</style>
"""


def _card_head(icon, text, badge=None):
    badge_html = f'<span class="toc-badge">{badge}</span>' if badge else ""
    st.markdown(
        f'<div class="toc-card-head">{icon}&nbsp; {text}{badge_html}</div>',
        unsafe_allow_html=True,
    )


def _render_single_file(fmt):
    with st.container(border=True):
        _card_head("📂", "Sube tu documento")
        uploaded = st.file_uploader(
            "Sube un archivo .docx", type=["docx"], key="toc_single_upload",
            label_visibility="collapsed",
        )

    if uploaded is None:
        st.info("👆 Sube un archivo .docx para comenzar.")
        return

    try:
        with st.spinner("Leyendo y formateando TOC..."):
            entries, rendered = docx_io.docx_to_result(uploaded, fmt=fmt)
    except Exception as exc:
        st.error(f"No se pudo procesar este archivo: {exc}")
        return

    if not entries:
        st.warning(
            "No se detectaron líneas numeradas/tipo TOC en este documento. "
            "Asegúrate de que el archivo contenga una sección de Tabla de Contenidos."
        )
        return

    col1, col2 = st.columns([3, 2])
    with col1:
        with st.container(border=True):
            _card_head("🔍", "Vista Previa", badge=f"{len(entries)} entradas")
            st.markdown(f'<div class="toc-preview">{rendered}</div>', unsafe_allow_html=True)
    with col2:
        with st.container(border=True):
            _card_head("📊", "Resumen")
            m1, m2 = st.columns(2)
            m1.metric("Entradas", len(entries))
            m2.metric("Profundidad máx.", max(e.level for e in entries))
            st.dataframe(
                [{"Nivel": e.level, "Número": e.number, "Título": e.title} for e in entries],
                width="stretch",
                height=300,
            )

    ext = {"txt": "txt", "md": "md", "json": "json"}[fmt]
    base_name = uploaded.name.rsplit(".", 1)[0]
    st.download_button(
        label=f"⬇️ Descargar {ext.upper()}",
        data=rendered,
        file_name=f"{base_name}_formatted.{ext}",
        mime="text/plain",
        width="stretch",
        type="primary",
    )


def _render_batch(fmt):
    with st.container(border=True):
        _card_head("📂", "Sube tus documentos")
        st.caption("Sube hasta ~1000 archivos DOCX. Conviértelos todos en un clic y descarga un solo ZIP.")
        uploaded_files = st.file_uploader(
            "Sube varios archivos .docx", type=["docx"], accept_multiple_files=True,
            key="toc_batch_upload", label_visibility="collapsed",
        )

    if not uploaded_files:
        st.info("👆 Sube varios archivos .docx para comenzar.")
        return

    st.write(f"**{len(uploaded_files)}** archivo(s) listo(s).")
    convert_clicked = st.button("🚀 Convertir Todos", type="primary", width="stretch")

    if not convert_clicked:
        return

    progress_bar = st.progress(0, text="Iniciando conversión por lotes...")
    status_placeholder = st.empty()

    def _on_progress(done, total, current_name):
        progress_bar.progress(done / total, text=f"Procesando {done}/{total}: {current_name}")

    results = docx_io.batch_process(uploaded_files, fmt=fmt, progress_callback=_on_progress)
    progress_bar.progress(1.0, text="¡Listo!")

    n_ok = sum(1 for r in results if not r["error"])
    n_err = len(results) - n_ok
    status_placeholder.success(f"✅ {n_ok} archivo(s) convertido(s) con éxito.")
    if n_err:
        st.error(f"⚠️ {n_err} archivo(s) fallaron:")
        for r in results:
            if r["error"]:
                st.write(f"- **{r['filename']}**: {r['error']}")

    with st.container(border=True):
        _card_head("🔍", "Vista previa de resultados")
        with st.expander("Ver cada archivo convertido", expanded=False):
            for r in results:
                if r["error"]:
                    continue
                st.markdown(f"**{r['output_name']}**")
                st.markdown(f'<div class="toc-preview">{r["rendered"]}</div>', unsafe_allow_html=True)

    zip_bytes = docx_io.results_to_zip_bytes(results)
    st.download_button(
        label="⬇️ Descargar Todo como ZIP",
        data=zip_bytes,
        file_name="toc_formatted_batch.zip",
        mime="application/zip",
        width="stretch",
        type="primary",
    )


def render_page():
    render_hero(
        "📑", "TOC Formatter", eyebrow="DOCX → TOC",
        subtitle=(
            "Sube archivos DOCX, detecta automáticamente la numeración (estilo antiguo o "
            "nuevo), y exporta una Tabla de Contenidos limpia y consistentemente indentada."
        ),
    )
    st.markdown(_PAGE_CSS, unsafe_allow_html=True)

    with st.container(border=True):
        col_mode, col_fmt = st.columns([2, 1])
        with col_mode:
            st.markdown("**Modo**")
            mode = st.segmented_control(
                "Modo", ["Single File", "Batch Processing"], default="Single File",
                required=True, key="toc_mode", label_visibility="collapsed",
            )
        with col_fmt:
            st.markdown("**Formato de salida**")
            fmt = st.selectbox(
                "Formato de salida", options=list(FORMAT_LABELS.keys()),
                format_func=lambda k: FORMAT_LABELS[k], key="toc_fmt",
                label_visibility="collapsed",
            )
        with st.expander("Estilos de numeración soportados"):
            st.markdown(
                "- Decimal moderno: `1.` `1.1.` `5.4.1.3.1.`\n"
                "- Esquema legado: `A.` `a.` `i.` `1.`\n"
                "- Se convierte automáticamente a numeración decimal limpia, siempre."
            )

    st.write("")

    if mode == "Single File":
        _render_single_file(fmt)
    else:
        _render_batch(fmt)
