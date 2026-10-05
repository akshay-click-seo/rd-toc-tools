"""
page_header.py
Shared "hero" banner used by the tool pages (RD Description, TOC Formatter)
that embed a heavier, self-contained UI of their own -- gives both a
consistent, on-brand header instead of a plain st.title/st.caption, using the
same navy-to-teal palette as the generated report images themselves
(utils.units / templates.segmentation_style.DEFAULT_PALETTE: "#0B2F7A" navy,
"#159A9C" teal), so the app's own chrome and its output images read as one
visual identity.
"""

import streamlit as st

NAVY = "#0B2F7A"
TEAL = "#159A9C"


def render_hero(icon: str, title: str, subtitle: str, eyebrow: str = ""):
    """Render a navy-to-teal gradient banner at the top of a tool page."""
    eyebrow_html = (
        f'<div style="display:inline-block;font-size:11.5px;font-weight:700;'
        f'letter-spacing:.08em;text-transform:uppercase;color:#fff;'
        f'background:rgba(255,255,255,.16);padding:4px 11px;border-radius:20px;'
        f'margin-bottom:12px;">{eyebrow}</div><br/>'
        if eyebrow else ""
    )
    st.markdown(
        f"""
        <div style="
            background:linear-gradient(135deg, {NAVY} 0%, #123c9c 55%, {TEAL} 145%);
            border-radius:16px;
            padding:26px 32px;
            margin-bottom:22px;
            box-shadow:0 8px 24px rgba(11,47,122,.20);
        ">
            {eyebrow_html}
            <div style="font-size:29px;font-weight:800;color:#fff;line-height:1.25;">
                {icon}&nbsp; {title}
            </div>
            <div style="font-size:14.5px;color:rgba(255,255,255,.92);margin-top:8px;max-width:78ch;line-height:1.55;">
                {subtitle}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
