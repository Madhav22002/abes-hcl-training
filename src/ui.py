"""
Tailwind CSS Styled UI Components and Header Cards
Utilizes supported Streamlit HTML component API (st.components.v1.html) with Tailwind CDN.
Strictly escapes HTML content to prevent XSS vulnerabilities.
"""

import html
import streamlit as st
import streamlit.components.v1 as components
from typing import List, Dict, Any

TAILWIND_CDN = '<script src="https://cdn.tailwindcss.com"></script>'

def render_app_header(title: str, subtitle: str, badge_text: str = "SYNTHETIC DEMO"):
    """Render a clean modern Tailwind-styled header banner."""
    safe_title = html.escape(title)
    safe_subtitle = html.escape(subtitle)
    safe_badge = html.escape(badge_text)

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>body {{ font-family: 'Inter', sans-serif; margin: 0; background-color: transparent; }}</style>
    </head>
    <body class="p-2">
        <div class="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 rounded-xl p-6 text-white shadow-lg border border-slate-700">
            <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                <div>
                    <div class="flex items-center gap-3">
                        <h1 class="text-2xl font-bold tracking-tight text-white">{safe_title}</h1>
                        <span class="bg-blue-500/20 text-blue-300 border border-blue-400/30 text-xs font-semibold px-2.5 py-0.5 rounded-full uppercase tracking-wider">
                            {safe_badge}
                        </span>
                    </div>
                    <p class="text-sm text-slate-300 mt-1 max-w-3xl">{safe_subtitle}</p>
                </div>
                <div class="text-right text-xs text-slate-400 border-l border-slate-700 pl-4 hidden md:block">
                    <p class="font-medium text-slate-200">AuraNova Digital Solutions</p>
                    <p>Sector 62, Noida, UP, India</p>
                    <p class="text-slate-400 mt-0.5">MCA Data Analytics Project</p>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=130)

def render_kpi_cards(cards: List[Dict[str, Any]]):
    """
    Render responsive Tailwind CSS KPI cards.
    cards: list of dicts with 'label', 'value', 'subtext', 'color' ('blue', 'green', 'amber', 'purple', 'emerald')
    """
    cards_html = []
    color_map = {
        "blue": ("border-blue-500", "text-blue-600", "bg-blue-50"),
        "green": ("border-emerald-500", "text-emerald-600", "bg-emerald-50"),
        "amber": ("border-amber-500", "text-amber-600", "bg-amber-50"),
        "purple": ("border-purple-500", "text-purple-600", "bg-purple-50"),
        "rose": ("border-rose-500", "text-rose-600", "bg-rose-50"),
        "indigo": ("border-indigo-500", "text-indigo-600", "bg-indigo-50")
    }

    for c in cards:
        safe_label = html.escape(str(c.get("label", "")))
        safe_value = html.escape(str(c.get("value", "—")))
        safe_sub = html.escape(str(c.get("subtext", "")))
        c_theme = color_map.get(c.get("color", "blue"), color_map["blue"])

        card_snippet = f"""
        <div class="bg-white rounded-lg p-4 shadow-sm border-l-4 {c_theme[0]} border-t border-r border-b border-slate-200 hover:shadow-md transition duration-150">
            <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider">{safe_label}</p>
            <p class="text-2xl font-bold {c_theme[1]} mt-1">{safe_value}</p>
            <p class="text-xs text-slate-500 mt-1">{safe_sub}</p>
        </div>
        """
        cards_html.append(card_snippet)

    cards_grid = "".join(cards_html)
    grid_cols = min(len(cards), 4)

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>body {{ font-family: 'Inter', sans-serif; margin: 0; background-color: transparent; }}</style>
    </head>
    <body class="p-1">
        <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-{grid_cols} gap-3">
            {cards_grid}
        </div>
    </body>
    </html>
    """
    height = 110 if len(cards) <= 4 else 220
    components.html(html_code, height=height)
