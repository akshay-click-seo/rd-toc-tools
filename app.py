"""RD Description (DOCX -> HTML) + TOC Formatter.  Run: streamlit run app.py"""
import os, sys
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

st.set_page_config(page_title="RD Description & TOC Formatter", page_icon="📝", layout="wide")

page = st.sidebar.radio("Tools", ["📝 RD Description", "📑 TOC Formatter"])

if page == "📝 RD Description":
    from sections import rd_description
    rd_description.render_page()
else:
    from sections import toc_formatter
    toc_formatter.render_page()
