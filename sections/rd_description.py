"""
rd_description.py
"RD Description" page: embeds the standalone DOCX -> HTML source conversion
tool (a self-contained HTML/JS page using Mammoth.js in the browser) as its
own menu item, exactly as provided -- no server-side reimplementation, so
its behavior (multi-file support, tabs, dark mode, the stats-line/heading
cleanup rules baked into its JS) stays identical to the standalone file.
"""

import os
from pathlib import Path

import streamlit as st

from utils.page_header import render_hero

ASSETS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
TOOL_HTML_PATH = Path(ASSETS_DIR) / "rd_description.html"


def render_page():
    render_hero(
        "📝", "RD Description", eyebrow="DOCX → HTML",
        subtitle=(
            "Convierte archivos DOCX a HTML fuente (para pegar en el sistema/CMS) -- "
            "todo se procesa en tu navegador, sin subir el archivo a ningún servidor."
        ),
    )

    if not TOOL_HTML_PATH.exists():
        st.error(f"No se encontró la herramienta en `{TOOL_HTML_PATH}`.")
        return

    # `st.iframe` reads the local HTML file directly and auto-sizes to its
    # measured content height, so the embedded tool (multi-file tabs, dark
    # mode toggle, textarea output) behaves exactly like the standalone
    # file the tool was authored/tested as.
    try:
        st.iframe(TOOL_HTML_PATH, width="stretch", height="content")
    except Exception:
        import streamlit.components.v1 as components
        components.html(TOOL_HTML_PATH.read_text(encoding="utf-8"), height=1400, scrolling=True)
