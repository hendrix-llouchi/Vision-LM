"""
GDSS-Maverick Hackathon
AI-Driven Image-to-IMDB Tool
Streamlit UI
"""

import os
import streamlit as st
from supabase import create_client
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "") or os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "") or os.environ.get("SUPABASE_KEY", "")
supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY) if SUPABASE_URL and SUPABASE_KEY else None

import streamlit.components.v1 as components
import json

import re
import base64
import requests
import io
import time
from pathlib import Path
from PIL import Image
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="IMDB Auto-Fill Tool",
    page_icon="🏷️",
    layout="wide"
)

# ─────────────────────────────────────────────
# STYLING
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400..900;1,400..900&family=Outfit:wght@300;400;500;600;700;800&display=swap');

    .stApp, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
        background-color: #f8fafc !important;
        color: #475569 !important;
        font-family: 'Outfit', sans-serif !important;
    }
    .main {
        background-color: transparent !important;
    }
    .block-container {
        padding-top: 2rem !important;
    }

    /* Typography */
    h1, h2, h3, h4, h5, h6, [data-testid="stHeader"] {
        font-family: 'Playfair Display', Georgia, serif !important;
        font-weight: 700 !important;
        color: #0f172a !important;
    }
    
    p, label, caption, small {
        font-family: 'Outfit', sans-serif !important;
        color: #475569 !important;
    }

    /* Sidebar elegant light theme */
    section[data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #e2e8f0 !important;
        box-shadow: 4px 0 20px rgba(0, 0, 0, 0.01) !important;
        z-index: 1000 !important; /* stay below our navbar (z-index: 99999) */
    }
    section[data-testid="stSidebar"] > div {
        background-color: #ffffff !important;
    }
    /* Adjust sidebar top padding */
    section[data-testid="stSidebar"] > div:first-child {
        padding-top: 2rem !important;
    }
    section[data-testid="stSidebar"] h3 {
        font-size: 1.1rem !important;
        color: #0f172a !important;
        margin-top: 1rem !important;
    }
    /* ── Hide Streamlit's native sidebar toggle buttons ──
       Moved off-screen (NOT display:none or opacity:0) so they remain
       fully rendered and React's .click() always works. */
    [data-testid="stSidebarCollapsedControl"] {
        position: fixed !important;
        left: -9999px !important;
        top: -9999px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] {
        position: fixed !important;
        left: -9999px !important;
        top: -9999px !important;
    }

    /* Input Fields & Widgets */
    div[data-testid="stTextInput"] [data-baseweb="input"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
        height: 42px !important;
        transition: border-color 0.15s ease, box-shadow 0.15s ease !important;
    }
    div[data-testid="stTextInput"] input {
        border: none !important;
        background-color: transparent !important;
        color: #0f172a !important;
        font-family: 'Outfit', sans-serif !important;
        height: 100% !important;
        box-shadow: none !important;
        outline: none !important;
    }
    div[data-testid="stSelectbox"] [data-baseweb="select"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        color: #0f172a !important;
        border-radius: 10px !important;
        font-family: 'Outfit', sans-serif !important;
        height: 42px !important;
        transition: border-color 0.15s ease, box-shadow 0.15s ease !important;
    }
    /* Target selectbox value wrapper to ensure it is white-bg and has dark text */
    div[data-testid="stSelectbox"] [data-baseweb="select"] > div {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }
    /* Ensure child text and spans inside the selectbox also use the dark color */
    div[data-testid="stSelectbox"] [data-baseweb="select"] * {
        color: #0f172a !important;
    }
    div[data-testid="stTextInput"] [data-baseweb="input"]:hover,
    div[data-testid="stSelectbox"] [data-baseweb="select"]:hover {
        border-color: #94a3b8 !important;
    }
    div[data-testid="stTextInput"] [data-baseweb="input"]:focus-within,
    div[data-testid="stSelectbox"] [data-baseweb="select"]:focus {
        border-color: #0f172a !important;
        box-shadow: 0 0 0 1px #0f172a !important;
    }
    /* Placeholder colors */
    div[data-testid="stTextInput"] input::placeholder {
        color: #94a3b8 !important;
        opacity: 1 !important;
    }
    /* Selectbox dropdown items and portals */
    [data-baseweb="popover"] ul,
    [data-baseweb="menu"],
    div[role="listbox"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 10px !important;
    }
    [data-baseweb="popover"] ul li,
    [data-baseweb="menu"] li,
    div[role="option"],
    li[role="option"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        padding: 0.5rem 1rem !important;
        transition: background-color 0.15s ease !important;
    }
    [data-baseweb="popover"] ul li:hover,
    [data-baseweb="menu"] li:hover,
    div[role="option"]:hover,
    li[role="option"]:hover {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
    }
    div[data-testid="stSelectbox"] [data-baseweb="select"],
    div[role="option"],
    li[role="option"],
    [data-baseweb="menu"] * {
        cursor: pointer !important;
    }
    div[data-testid="stTextInput"] label, div[data-testid="stSelectbox"] label {
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        color: #475569 !important;
        font-size: 0.85rem !important;
    }

    .nav-wrapper {
        container-type: inline-size;
        width: 100%;
    }

    /* Top Navigation Capsule */
    .nav-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: rgba(255, 255, 255, 0.55) !important;
        -webkit-backdrop-filter: blur(40px) saturate(200%) !important;
        backdrop-filter: blur(40px) saturate(200%) !important;
        border: 1px solid rgba(255, 255, 255, 0.75) !important;
        border-radius: 24px !important;
        padding: 0.6rem 1.75rem !important;
        box-shadow: 
            0 20px 40px rgba(15, 23, 42, 0.06), 
            0 5px 15px rgba(15, 23, 42, 0.04), 
            inset 0 1px 0 rgba(255, 255, 255, 0.7) !important;
        position: fixed !important;
        top: 12px !important;
        left: 50% !important;
        transform: translateX(-50%) !important;
        width: 92% !important;
        max-width: 1100px !important;
        z-index: 99999 !important;
        margin: 0 !important;
    }
    /* Responsive navbar container styling */
    @media (max-width: 850px) {
        .nav-container {
            width: 96% !important;
            padding: 0.5rem 1rem !important;
            border-radius: 20px !important;
        }
        .nav-logo {
            font-size: 1.15rem !important;
        }
    }

    /* Adjust viewport-based padding for the main content block container */
    @media (max-width: 950px) {
        .block-container {
            padding-top: 1.5rem !important;
        }
    }
    /* Hide the zero-height iframe container used for JS injection so it doesn't add blank space */
    div.element-container:has(iframe[height="0"]) {
        display: none !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    
    /* Hide Streamlit's ⋮ mobile menu button */
    [data-testid="stMainMenuButton"],
    button[data-testid="baseButton-header"] {
        display: none !important;
    }
    /* Hamburger toggle button inside the navbar capsule */
    .nav-hamburger {
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 40px !important;
        height: 40px !important;
        background: #000000 !important;
        border-radius: 12px !important;
        cursor: pointer !important;
        margin-right: 1.25rem !important;
        flex-shrink: 0 !important;
        transition: background 0.2s ease, transform 0.15s ease !important;
        user-select: none !important;
    }
    .nav-hamburger:hover {
        background: #000000 !important;
        transform: scale(1.05) !important;
    }
    /* Hide Streamlit's native header bar (Deploy btn, ⋮ menu) */
    header[data-testid="stHeader"] {
        display: none !important;
    }
    .nav-left {
        display: flex;
        align-items: center;
    }
    .nav-logo {

        font-family: 'Playfair Display', Georgia, serif !important;
        font-weight: 800 !important;
        font-size: 1.25rem !important;
        color: #0f172a !important;
    }
    .nav-center {
        display: flex;
        align-items: center;
        gap: 1.5rem;
    }
    .nav-tab {
        font-family: 'Outfit', sans-serif !important;
        font-size: 0.9rem !important;
        font-weight: 500 !important;
        color: #000000 !important;
        opacity: 0.55 !important;
        padding: 0.5rem 0.25rem !important;
        cursor: pointer;
        position: relative;
        transition: all 0.2s ease !important;
        display: inline-flex !important;
        align-items: center !important;
    }
    .nav-tab:hover {
        opacity: 0.95 !important;
    }
    .active-tab {
        color: #000000 !important;
        font-weight: 700 !important;
        opacity: 1 !important;
    }
    .active-tab::after {
        content: '';
        position: absolute;
        bottom: 0;
        left: 0;
        right: 0;
        height: 2.5px;
        background-color: #000000;
        border-radius: 2px;
    }
    .step-num {
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 20px !important;
        height: 20px !important;
        background-color: rgba(0, 0, 0, 0.08) !important;
        color: #000000 !important;
        border-radius: 50% !important;
        font-size: 0.72rem !important;
        font-weight: 600 !important;
        margin-right: 8px !important;
        transition: all 0.2s ease !important;
        opacity: 0.8 !important;
    }
    .active-num {
        background-color: #000000 !important;
        color: #ffffff !important;
        opacity: 1 !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.2) !important;
    }
    .nav-right {
        display: flex;
        align-items: center;
        gap: 1rem;
    }
    .nav-badge {
        font-family: 'Outfit', sans-serif !important;
        font-size: 0.8rem !important;
        font-weight: 600 !important;
        color: #475569 !important;
        background-color: #f1f5f9 !important;
        padding: 0.35rem 0.75rem !important;
        border-radius: 20px !important;
        display: flex;
        align-items: center;
        gap: 0.25rem;
        border: 1px solid #e2e8f0 !important;
    }
    .nav-profile {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        background: #f8fafc !important;
        border: 1px solid #e2e8f0 !important;
        padding: 0.35rem 0.85rem !important;
        border-radius: 20px !important;
    }
    .nav-profile-name {
        font-family: 'Outfit', sans-serif !important;
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #0f172a !important;
    }

    /* Cards & Containers */
    div[data-testid="stMetric"], 
    div[data-testid="stAlert"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 20px !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.02) !important;
        color: #475569 !important;
        transition: all 0.3s ease !important;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 15px 35px rgba(0, 0, 0, 0.04) !important;
    }
    div[data-testid="stMetric"] {
        padding: 1.25rem !important;
    }
    div[data-testid="stMetricLabel"] > div {
        font-family: 'Outfit', sans-serif !important;
        text-transform: uppercase !important;
        font-size: 0.75rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.05em !important;
        color: #64748b !important;
    }
    div[data-testid="stMetricValue"] > div {
        font-family: 'Playfair Display', Georgia, serif !important;
        font-size: 2.25rem !important;
        font-weight: 800 !important;
        color: #0f172a !important;
        margin-top: 0.25rem !important;
    }

    /* Expander & Bordered Containers */
    div[data-testid="stExpander"],
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 16px !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.02) !important;
        color: #475569 !important;
        transition: all 0.3s ease !important;
        overflow: hidden !important;
        margin-bottom: 1.5rem !important;
    }
    div[data-testid="stExpander"]:hover,
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {
        border-color: #cbd5e1 !important;
        box-shadow: 0 12px 35px rgba(0, 0, 0, 0.03) !important;
    }

    /* Expander Header (Modern Streamlit uses summary element) */
    div[data-testid="stExpander"] summary,
    .streamlit-expanderHeader {
        background-color: #ffffff !important;
        border-bottom: none !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        color: #0f172a !important;
        padding: 1rem 1.5rem !important;
        transition: background-color 0.2s ease !important;
    }
    /* Expander Header Hover */
    div[data-testid="stExpander"] summary:hover {
        background-color: #f8fafc !important;
    }
    /* Target the text wrapper inside summary */
    div[data-testid="stExpander"] summary > div[data-testid="stMarkdownContainer"] p {
        color: #0f172a !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
    }
    /* Target the toggle icon */
    div[data-testid="stExpander"] summary > svg[data-testid="stExpanderToggleIcon"] {
        fill: #0f172a !important;
    }

    /* Expander Content */
    div[data-testid="stExpander"] div[data-testid="stExpanderDetails"],
    .streamlit-expanderContent {
        background-color: #ffffff !important;
        border-top: 1px solid #f1f5f9 !important;
        padding: 1.5rem !important;
    }

    /* Buttons */
    div.stButton > button, div.stDownloadButton > button {
        background-color: #000000 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.65rem 1.75rem !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15) !important;
        transition: all 0.2s ease !important;
        width: 100% !important;
    }
    /* Force white on all nested label elements Streamlit injects inside buttons */
    div.stButton > button *, div.stDownloadButton > button * {
        color: #ffffff !important;
    }
    div.stButton > button:hover, div.stDownloadButton > button:hover {
        background-color: #222222 !important;
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.25) !important;
        transform: translateY(-1px) !important;
    }
    div.stButton > button:active, div.stDownloadButton > button:active {
        transform: translateY(0px) !important;
    }


    /* File Uploader dashed zone */
    div[data-testid="stFileUploader"] section {
        border: 2px dashed #cbd5e1 !important;
        background-color: #ffffff !important;
        border-radius: 20px !important;
        padding: 2rem !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.005) !important;
    }
    div[data-testid="stFileUploader"] section:hover {
        border-color: #0f172a !important;
        background-color: #f8fafc !important;
    }
    /* Fix file uploader text elements that default to white/invisible in dark theme base */
    div[data-testid="stFileUploader"] section span,
    div[data-testid="stFileUploader"] section p,
    div[data-testid="stFileUploader"] section small,
    div[data-testid="stFileUploader"] section div {
        color: #475569 !important;
    }
    /* Ensure the Upload button inside file uploader keeps its dark background and white text */
    div[data-testid="stFileUploader"] section button,
    div[data-testid="stFileUploader"] section button * {
        color: #ffffff !important;
        background-color: #000000 !important;
    }
    div[data-testid="stFileUploader"] section button:hover {
        background-color: #222222 !important;
    }

    /* Table & Dataframe */
    div[data-testid="stDataFrame"] {
        border: 1px solid #e2e8f0 !important;
        border-radius: 16px !important;
        overflow: hidden !important;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.01) !important;
    }

    /* Success banner */
    .success-banner {
        background-color: #ecfdf5 !important;
        border: 1px solid #a7f3d0 !important;
        border-radius: 12px !important;
        color: #065f46 !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        padding: 1rem !important;
        text-align: center !important;
        margin-top: 1rem;
        margin-bottom: 1rem;
    }

    /* Camera Input Widget styling */
    div[data-testid="stCameraInput"] {
        border: 1px solid #cbd5e1 !important;
        border-radius: 20px !important;
        overflow: hidden !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.02) !important;
        background-color: #ffffff !important;
        margin-bottom: 1rem !important;
    }
    div[data-testid="stCameraInput"] button,
    div[data-testid="stCameraInput"] button * {
        color: #ffffff !important;
        background-color: #000000 !important;
        border-radius: 10px !important;
    }
    div[data-testid="stCameraInput"] button:hover {
        background-color: #222222 !important;
    }

    /* Compact Snapshot Gallery Grid */
    .snap-card-title {
        font-family: 'Outfit', sans-serif !important;
        font-size: 0.8rem !important;
        font-weight: 600 !important;
        color: #475569 !important;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        margin-top: 0.25rem !important;
        margin-bottom: 0.5rem !important;
    }

    /* Navbar clickable view links */
    .nav-view-link {
        font-family: 'Outfit', sans-serif !important;
        font-size: 0.88rem !important;
        font-weight: 600 !important;
        color: #000000 !important;
        opacity: 0.5 !important;
        padding: 0.45rem 0.85rem !important;
        border-radius: 10px !important;
        cursor: pointer !important;
        text-decoration: none !important;
        transition: all 0.2s ease !important;
        display: inline-flex !important;
        align-items: center !important;
        gap: 0.35rem !important;
    }
    .nav-view-link:hover {
        opacity: 0.85 !important;
        background-color: rgba(0,0,0,0.05) !important;
    }
    .nav-view-link.active-view {
        opacity: 1 !important;
        background-color: #000000 !important;
        color: #ffffff !important;
    }
    .nav-divider {
        width: 1px;
        height: 20px;
        background: #e2e8f0;
        margin: 0 0.5rem;
    }
    /* Item Master Database table card */
    .db-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.02);
        margin-bottom: 1.5rem;
    }
    .db-stat-chip {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        background: #f1f5f9;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 0.3rem 0.75rem;
        font-family: 'Outfit', sans-serif;
        font-size: 0.8rem;
        font-weight: 600;
        color: #475569;
        margin-right: 0.5rem;
        margin-bottom: 0.5rem;
    }

    /* Stepper card in Pipeline Page */
    .stepper-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 0.85rem 1.5rem;
        box-shadow: 0 4px 20px rgba(15, 23, 42, 0.015);
        margin-bottom: 2rem;
        width: 100%;
    }
    .stepper-container {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 1.25rem;
        flex-wrap: wrap;
        width: 100%;
    }
    .page-tab {
        font-family: 'Outfit', sans-serif !important;
        font-size: 0.95rem !important;
        font-weight: 500 !important;
        color: #475569 !important;
        opacity: 0.55 !important;
        padding: 0.4rem 0.8rem !important;
        display: inline-flex !important;
        align-items: center !important;
        border-radius: 10px;
        transition: all 0.2s ease !important;
    }
    .page-tab.active-tab {
        color: #0f172a !important;
        font-weight: 700 !important;
        opacity: 1 !important;
        background-color: #f1f5f9;
        border: 1px solid #cbd5e1;
    }
    .stepper-arrow {
        color: #cbd5e1;
        font-size: 1rem;
        font-weight: 500;
        user-select: none;
    }

    /* Responsive Navbar Text wrappers */
    .nav-text-mobile {
        display: none !important;
    }
    .nav-left-group {
        display: flex;
        align-items: center;
    }
    
    /* Shorten database label on tablets/phones */
    @media (max-width: 768px) {
        .nav-text-full {
            display: none !important;
        }
        .nav-text-mobile {
            display: inline !important;
        }
    }
    
    /* Hide divider, center logo, and structure 3-column grid on mobile/small screens */
    @media (max-width: 600px) {
        .nav-left-group {
            display: contents !important;
        }
        .nav-divider {
            display: none !important;
        }
        .nav-container {
            width: 96% !important;
            padding: 0.45rem 0.75rem !important;
            border-radius: 16px !important;
            display: grid !important;
            grid-template-columns: 1fr auto 1fr !important;
            align-items: center !important;
        }
        .nav-logo-container {
            order: 2 !important;
            grid-column: 2 !important;
            justify-self: center !important;
            display: flex !important;
            align-items: center !important;
        }
        .nav-left {
            order: 1 !important;
            grid-column: 1 !important;
            justify-self: start !important;
            display: flex !important;
            align-items: center !important;
            gap: 0.35rem !important;
        }
        .nav-right {
            order: 3 !important;
            grid-column: 3 !important;
            justify-self: end !important;
            display: flex !important;
            align-items: center !important;
        }
        .nav-logo {
            font-size: 1.1rem !important;
            margin-right: 0 !important;
        }
        .nav-badge {
            font-size: 0.75rem !important;
            padding: 0.25rem 0.5rem !important;
        }
        .nav-view-link {
            font-size: 0.82rem !important;
            padding: 0.35rem 0.65rem !important;
            gap: 0.25rem !important;
        }
    }
    
    /* Icon-only mode for very small screens (phones in portrait) and pill badge styling */
    @media (max-width: 480px) {
        .nav-text-mobile, .nav-text-full {
            display: none !important;
        }
        .nav-view-link {
            padding: 0 !important; /* clear padding to avoid centering offset */
            border-radius: 50% !important;
            width: 36px !important;
            height: 36px !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
        }
        .nav-left {
            gap: 0.35rem !important;
        }
        .nav-logo {
            font-size: 1rem !important;
        }
        .nav-badge {
            font-size: 0.72rem !important; /* Keep font compact but visible */
            padding: 0.25rem 0.5rem !important;
            border-radius: 12px !important; /* Keep pill shape */
            width: auto !important;
            height: auto !important;
            display: inline-flex !important;
            align-items: center !important;
            gap: 0.2rem !important;
        }
        .nav-badge span {
            margin-right: 2px !important; /* Restore dot margin */
        }
    }
    
    /* Stepper Mobile Collapsing Rules */
    @media (max-width: 600px) {
        /* Hide text labels of inactive steps */
        .page-tab:not(.active-tab) .step-label {
            display: none !important;
        }
        .stepper-container {
            gap: 0.5rem !important;
        }
        .page-tab {
            padding: 0.35rem 0.55rem !important;
        }
        .stepper-arrow {
            font-size: 0.85rem !important;
        }
        .step-num {
            margin-right: 0 !important; /* center number since label is hidden */
        }
        .page-tab.active-tab .step-num {
            margin-right: 6px !important; /* restore margin for active step */
        }
    }

</style>
""", unsafe_allow_html=True)


def render_navbar(active_step, model_name="Groq API", current_view="pipeline", key_suffix="main"):
    compact_model_name = model_name.replace(" API", "")
    
    st.markdown("""
    <style>
        div[data-testid="column"] button {
            border-radius: 12px !important;
            font-family: 'Outfit', sans-serif !important;
            font-weight: 600 !important;
            height: 38px !important;
            transition: all 0.2s ease !important;
        }
    </style>
    """, unsafe_allow_html=True)
    
    col_logo, col_pipe, col_db, col_badge = st.columns([1.5, 1.2, 1.8, 1.5])
    
    with col_logo:
        st.markdown("""
        <div style="display: flex; align-items: center; gap: 8px; height: 38px; margin-top: 2px;">
            <div style="
                background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                font-family: 'Outfit', sans-serif; 
                font-weight: 900; 
                font-size: 1.5rem;
                letter-spacing: -0.03em;
                line-height: 1;
            ">VisionLM</div>
            <div style="font-size: 0.65rem; background: #f1f5f9; color: #64748b; padding: 2px 6px; border-radius: 99px; font-weight: 700; letter-spacing: 0.05em; height: 16px; display: flex; align-items: center; justify-content: center;">BETA</div>
        </div>
        """, unsafe_allow_html=True)
        
    with col_pipe:
        btn_type = "primary" if current_view == "pipeline" else "secondary"
        if st.button("🏭 Pipeline", key=f"nav_pipe_{key_suffix}", use_container_width=True, type=btn_type):
            st.session_state.current_view = "pipeline"
            st.rerun()
            
    with col_db:
        btn_type = "primary" if current_view == "database" else "secondary"
        if st.button("📊 Item Master Database", key=f"nav_db_{key_suffix}", use_container_width=True, type=btn_type):
            st.session_state.current_view = "database"
            st.rerun()
            
    with col_badge:
        st.markdown(f"""
        <div style="display: flex; justify-content: flex-end; align-items: center; height: 38px;">
            <span style="
                display: inline-flex;
                align-items: center;
                background: white;
                border: 1px solid #e2e8f0;
                padding: 0.4rem 0.85rem;
                border-radius: 9999px;
                font-family: 'Inter', sans-serif;
                font-size: 0.85rem;
                font-weight: 600;
                color: #475569;
                box-shadow: 0 1px 2px rgba(0,0,0,0.02);
            ">
                <span style='color:#10b981; margin-right:8px; font-size: 0.6rem; animation: pulse 2s infinite;'>●</span>{compact_model_name}
            </span>
        </div>
        <style>
            @keyframes pulse {{
                0% {{ opacity: 1; }}
                50% {{ opacity: 0.4; }}
                100% {{ opacity: 1; }}
            }}
        </style>
        """, unsafe_allow_html=True)
        
    st.markdown("<hr style='margin-top: 0.5rem; margin-bottom: 1.5rem;'/>", unsafe_allow_html=True)


def render_pipeline_stepper(active_step):
    steps = [
        ("Upload", 1),
        ("Extract", 2),
        ("Aggregate", 3),
        ("Validate", 4),
        ("Export", 5)
    ]
    tabs_html = ""
    for i, (label, step_num) in enumerate(steps):
        if step_num == active_step:
            tabs_html += f'<div class="page-tab active-tab"><span class="step-num active-num">{step_num}</span><span class="step-label">{label}</span></div>'
        else:
            tabs_html += f'<div class="page-tab"><span class="step-num">{step_num}</span><span class="step-label">{label}</span></div>'
        if i < len(steps) - 1:
            tabs_html += '<span class="stepper-arrow">→</span>'

    stepper_html = f"""
    <div class="stepper-card">
        <div class="stepper-container">
            {tabs_html}
        </div>
    </div>
    """
    st.markdown(stepper_html, unsafe_allow_html=True)



# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────
IMDB_COLS = [
    "ITEM NAME", "BARCODE", "MANUFACTURER", "BRAND", "WEIGHT",
    "PACKAGING TYPE", "COUNTRY", "VARIANT", "TYPE",
    "FRAGRANCE FLAVOR", "PROMOTION", "ADDONS", "TAGLINE",
]


SYSTEM_PROMPT = """You are a product data extraction engine for a retail catalog system.
YOUR ONLY OUTPUT IS A SINGLE VALID JSON OBJECT. NO PREAMBLE. NO EXPLANATION. NO MARKDOWN. NO CODE BLOCKS. JUST THE RAW JSON.

Extract the following 13 fields:
- ITEM_NAME: full descriptive product name for catalog e.g. "KNORR CHICKEN STOCK CUBE 20G"
- BARCODE: numeric digits only, no dashes or spaces
- MANUFACTURER: full company name that manufactures the product
- BRAND: brand name exactly as shown on package e.g. "KNORR", "SUNLIGHT"
- WEIGHT: net weight or volume with unit, no spaces e.g. "80G", "500ML", "1.5KG"
- PACKAGING_TYPE: one of SACHET, BOTTLE, TUB, CAN, BOX, GLASS JAR, CARTON, POUCH, TUBE, PACKET, BAG
- COUNTRY: country of manufacture as printed on package e.g. "MALAYSIA", "THAILAND"
- VARIANT: product variant if present e.g. "ORIGINAL", "LOW FAT", "REDUCED SALT", "PREMIUM"
- TYPE: short product category e.g. "MARGARINE", "MAYONNAISE", "SEASONING", "DETERGENT"
- FRAGRANCE_FLAVOR: scent or flavor descriptor if present e.g. "RICH", "ORIGINAL", "LEMON", "VANILLA"
- PROMOTION: any on-pack promotion text e.g. "BUY 2 FREE 1", "20% EXTRA FREE", "NEW LOWER PRICE"
- ADDONS: additional included items or features e.g. "SPOON INCLUDED", "FREE RECIPE BOOK"
- TAGLINE: short promotional tagline or slogan e.g. "TASTE THE DIFFERENCE", "LOVED BY FAMILIES"

Rules:
- UPPERCASE for all string values
- Use "" for any field not visible or not applicable
- NEVER use null
- BARCODE must be digits only
- WEIGHT must have no space between number and unit"""

TRIGGER = 'Extract the product data from this image and return only the JSON object. Use "" for missing fields. Never use null.'

# ─────────────────────────────────────────────
# SUPABASE HELPERS
# ─────────────────────────────────────────────
def get_supabase_client():
    url = st.session_state.get("supabase_url", SUPABASE_URL)
    key = st.session_state.get("supabase_key", SUPABASE_KEY)
    if url == SUPABASE_URL and key == SUPABASE_KEY:
        return supabase_client
    if url and key:
        try:
            return create_client(url, key)
        except Exception:
            return None
    return None


def supabase_upsert_product(imdb_row):
    """
    Upsert a product into the Supabase imdb_products table.
    """
    client = get_supabase_client()
    if not client:
        return False, "Supabase credentials not configured."
    
    item_name = imdb_row.get("ITEM NAME", "").strip()
    if not item_name:
        return False, "No item name provided."
        
    db_fields = {
        "item_name":        item_name,
        "barcode":          imdb_row.get("BARCODE", "").strip(),
        "manufacturer":     imdb_row.get("MANUFACTURER", "").strip(),
        "brand":            imdb_row.get("BRAND", "").strip(),
        "weight":           imdb_row.get("WEIGHT", "").strip(),
        "packaging_type":   imdb_row.get("PACKAGING TYPE", "").strip(),
        "country":          imdb_row.get("COUNTRY", "").strip(),
        "variant":          imdb_row.get("VARIANT", "").strip(),
        "type":             imdb_row.get("TYPE", "").strip(),
        "fragrance_flavor": imdb_row.get("FRAGRANCE FLAVOR", "").strip(),
        "promotion":        imdb_row.get("PROMOTION", "").strip(),
        "addons":           imdb_row.get("ADDONS", "").strip(),
        "tagline":          imdb_row.get("TAGLINE", "").strip(),
    }
    
    try:
        resp = client.table("imdb_products").select("*").eq("item_name", item_name).execute()
        existing_records = resp.data
        
        if existing_records and len(existing_records) > 0:
            existing = existing_records[0]
            scan_count = int(existing.get("scan_count") or 1) + 1
            client.table("imdb_products").update({"scan_count": scan_count}).eq("id", existing["id"]).execute()
            return True, f"Updated (scan #{scan_count})"
        else:
            db_fields["scan_count"] = 1
            client.table("imdb_products").insert(db_fields).execute()
            return True, "Inserted as new record"
    except Exception as e:
        return False, str(e)


def supabase_get_products(search_query=""):
    """
    Fetch all products from Supabase imdb_products, optionally filtered by brand or item_name.
    """
    client = get_supabase_client()
    if not client:
        return [], "Supabase credentials not configured."
    try:
        query = client.table("imdb_products").select("*").order("last_updated", desc=True)
        if search_query:
            query = query.or_(f"brand.ilike.*{search_query}*,item_name.ilike.*{search_query}*")
        resp = query.execute()
        return resp.data, None
    except Exception as e:
        return [], str(e)


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
class NamedBytesIO(io.BytesIO):
    """File-like wrapper for captured image bytes with a custom name and size."""
    def __init__(self, content, name):
        super().__init__(content)
        self.name = name
        self.size = len(content)


def preprocess_image(image_file, max_size=(1024, 1024)):
    """Resize and base64-encode an image. Use smaller max_size for local models."""
    if hasattr(image_file, "seek"):
        image_file.seek(0)
    img = Image.open(image_file)
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail(max_size, Image.LANCZOS)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=85)
    buffer.seek(0)
    return base64.standard_b64encode(buffer.read()).decode("utf-8")


GROQ_API_KEY       = st.secrets.get("GROQ_API_KEY", "") or os.environ.get("GROQ_API_KEY", "")
OPENROUTER_API_KEY = st.secrets.get("OPENROUTER_API_KEY", "") or os.environ.get("OPENROUTER_API_KEY", "")
GROQ_DELAY_SEC     = 6                       # Minimum seconds between Groq/OpenRouter requests


def extract_via_groq(b64_image, api_key, model="llama-3.2-11b-vision-preview", max_retries=5):
    """
    Send image to Groq chat-completions endpoint (llama-3.2-11b-vision-preview or llama-3.2-90b-vision-preview).
    Exponential backoff on 429 errors: 5s, 10s, 20s, 40s, 80s.
    """
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": model,
        "max_tokens": 1000,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_image}"}},
                {"type": "text", "text": TRIGGER}
            ]}
        ]
    }
    for attempt in range(max_retries):
        resp = requests.post(url, json=payload, headers=headers, timeout=90)
        if resp.status_code == 429:
            wait = 5 * (2 ** attempt)
            st.warning(f"⏳ Groq rate limit hit. Waiting {wait}s before retry (attempt {attempt + 1}/{max_retries})...")
            time.sleep(wait)
            continue
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"].strip()
        content = re.sub(r"^```json\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
        return json.loads(content)
    raise Exception(f"Groq API returned 429 after {max_retries} retries.")


def extract_via_openrouter(b64_image, api_key, model="google/gemma-4-26b-a4b-it:free", max_retries=5):
    """
    Send image to OpenRouter chat-completions endpoint (google/gemma-4-26b-a4b-it:free).
    Exponential backoff on 429 errors: 30s, 60s, 120s, 240s, 480s.
    """
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": model,
        "max_tokens": 1000,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_image}"}},
                {"type": "text", "text": TRIGGER}
            ]}
        ]
    }
    for attempt in range(max_retries):
        resp = requests.post(url, json=payload, headers=headers, timeout=90)
        if resp.status_code == 429:
            wait = 30 * (2 ** attempt)
            st.warning(f"⏳ OpenRouter rate limit hit. Waiting {wait}s before retry (attempt {attempt + 1}/{max_retries})...")
            time.sleep(wait)
            continue
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"].strip()
        content = re.sub(r"^```json\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
        return json.loads(content)
    raise Exception(f"OpenRouter API returned 429 after {max_retries} retries.")


def aggregate(records):
    merged = {}
    fields = [
        "ITEM_NAME", "BARCODE", "MANUFACTURER", "BRAND", "WEIGHT",
        "PACKAGING_TYPE", "COUNTRY", "VARIANT", "TYPE",
        "FRAGRANCE_FLAVOR", "PROMOTION", "ADDONS", "TAGLINE",
    ]
    for col in fields:
        values = [r.get(col, "").strip() for r in records if r.get(col, "").strip()]
        if not values:
            merged[col] = ""
        else:
            count = Counter(values)
            max_count = count.most_common(1)[0][1]
            candidates = [v for v, c in count.items() if c == max_count]
            merged[col] = max(candidates, key=len)
    return merged


def validate(record):
    record["BARCODE"] = re.sub(r"[^\d]", "", record.get("BARCODE", ""))
    weight = record.get("WEIGHT", "").upper()
    weight = re.sub(r"(\d)\s+(G|KG|ML|L|MG)\b", r"\1\2", weight)
    record["WEIGHT"] = weight
    for field in ["ITEM_NAME", "MANUFACTURER", "BRAND", "PACKAGING_TYPE",
                  "COUNTRY", "VARIANT", "TYPE", "FRAGRANCE_FLAVOR",
                  "PROMOTION", "ADDONS", "TAGLINE"]:
        record[field] = record.get(field, "").upper().strip()
    return record


def to_imdb(record):
    return {
        "ITEM NAME":        record.get("ITEM_NAME", ""),
        "BARCODE":          record.get("BARCODE", ""),
        "MANUFACTURER":     record.get("MANUFACTURER", ""),
        "BRAND":            record.get("BRAND", ""),
        "WEIGHT":           record.get("WEIGHT", ""),
        "PACKAGING TYPE":   record.get("PACKAGING_TYPE", ""),
        "COUNTRY":          record.get("COUNTRY", ""),
        "VARIANT":          record.get("VARIANT", ""),
        "TYPE":             record.get("TYPE", ""),
        "FRAGRANCE FLAVOR": record.get("FRAGRANCE_FLAVOR", ""),
        "PROMOTION":        record.get("PROMOTION", ""),
        "ADDONS":           record.get("ADDONS", ""),
        "TAGLINE":          record.get("TAGLINE", ""),
    }


def build_excel(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "IMDB"
    thin  = Side(style='thin', color='000000')
    thick = Side(style='medium', color='000000')
    ws.append(IMDB_COLS)
    for col_idx in range(1, len(IMDB_COLS)+1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(name='Calibri', bold=True, size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(left=thick, right=thick, top=thick, bottom=thick)
    ws.row_dimensions[1].height = 20
    for row_idx, row in enumerate(rows, 2):
        for col_idx, col in enumerate(IMDB_COLS, 1):
            val = row.get(col, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val if val else None)
            cell.font = Font(name='Calibri', size=10)
            cell.alignment = Alignment(horizontal='left', vertical='center')
            cell.border = Border(left=Side(style='thin', color='000000'),
                                 right=Side(style='thin', color='000000'),
                                 top=Side(style='thin', color='000000'),
                                 bottom=Side(style='thin', color='000000'))
        ws.row_dimensions[row_idx].height = 15
    col_widths = [32, 16, 32, 18, 12, 18, 14, 18, 20, 20, 22, 22, 35]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "results" not in st.session_state:
    st.session_state.results = []
if "edited_df" not in st.session_state:
    st.session_state.edited_df = None
if "groq_api_key" not in st.session_state:
    st.session_state.groq_api_key = GROQ_API_KEY
if "openrouter_api_key" not in st.session_state:
    st.session_state.openrouter_api_key = OPENROUTER_API_KEY
if "engine" not in st.session_state:
    st.session_state.engine = "OpenRouter API"
if "camera_snapshots" not in st.session_state:
    st.session_state.camera_snapshots = []
if "last_added_photo" not in st.session_state:
    st.session_state.last_added_photo = None
if "supabase_url" not in st.session_state:
    st.session_state.supabase_url = os.environ.get("SUPABASE_URL", "")
if "supabase_key" not in st.session_state:
    st.session_state.supabase_key = os.environ.get("SUPABASE_KEY", "")


# ─────────────────────────────────────────────
# VIEW ROUTING (Pipeline vs Item Master Database)
# ─────────────────────────────────────────────
# Navigation handles URL updates directly via javascript History API (popstate)

if "current_view" not in st.session_state:
    st.session_state.current_view = "pipeline"
current_view = st.session_state.current_view

# ─────────────────────────────────────────────
# STATE DETECTION FOR PIPELINE TABS
# ─────────────────────────────────────────────
active_step = 1
uploaded_files_in_state = st.session_state.get("uploader", [])
camera_snapshots_in_state = st.session_state.get("camera_snapshots", [])
if uploaded_files_in_state or camera_snapshots_in_state:
    active_step = 2
    if st.session_state.get("results"):
        active_step = 5

# ─────────────────────────────────────────────
# HEADER (NAVBAR)
# ─────────────────────────────────────────────
navbar_placeholder = st.empty()
with navbar_placeholder.container():
    render_navbar(active_step, st.session_state.engine, current_view, key_suffix="main")

# ─────────────────────────────────────────────
# ENGINE CONFIG (MOVED FROM SIDEBAR)
# ─────────────────────────────────────────────
if current_view == "pipeline":
    with st.expander("⚙️ Model Configuration", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.engine = st.selectbox(
                "Choose API Provider",
                ["Groq API", "OpenRouter API"],
                index=0 if st.session_state.get("engine", "Groq API") == "Groq API" else 1
            )
        with col2:
            if st.session_state.engine == "Groq API":
                st.session_state.groq_api_key = st.text_input(
                    "Groq API Key",
                    type="password",
                    placeholder="gsk_...",
                    value=st.session_state.get("groq_api_key", "")
                )
            elif st.session_state.engine == "OpenRouter API":
                st.session_state.openrouter_api_key = st.text_input(
                    "OpenRouter API Key",
                    type="password",
                    placeholder="sk-or-v1-...",
                    value=st.session_state.get("openrouter_api_key", "")
                )
        
        # Model Selection Sub-Row
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            if st.session_state.engine == "Groq API":
                groq_models = ["llama-3.2-11b-vision-preview", "llama-3.2-90b-vision-preview"]
                current_groq = st.session_state.get("groq_model", "llama-3.2-11b-vision-preview")
                idx = groq_models.index(current_groq) if current_groq in groq_models else 0
                st.session_state.groq_model = st.selectbox(
                    "Groq Vision Model (100% Free)",
                    groq_models,
                    index=idx,
                    help="Free-tier vision models on GroqCloud."
                )
            elif st.session_state.engine == "OpenRouter API":
                or_models = [
                    "google/gemma-4-26b-a4b-it:free",
                    "qwen/qwen-2.5-vl-72b-instruct:free",
                    "meta-llama/llama-3.2-11b-vision-instruct:free"
                ]
                current_or = st.session_state.get("openrouter_model", "google/gemma-4-26b-a4b-it:free")
                idx = or_models.index(current_or) if current_or in or_models else 0
                st.session_state.openrouter_model = st.selectbox(
                    "OpenRouter Free Vision Model",
                    or_models,
                    index=idx,
                    help="Free vision models on OpenRouter."
                )
        with col_m2:
            st.caption("ℹ️ **Free Tier Note:** Both `llama-3.2-11b-vision-preview` and `llama-3.2-90b-vision-preview` on Groq are free for developers with up to 30 RPM rate limits.")
        st.markdown("---")
        col3, col4 = st.columns(2)
        with col3:
            st.session_state.supabase_url = st.text_input(
                "Supabase URL",
                placeholder="https://xxxx.supabase.co",
                value=st.session_state.get("supabase_url", "")
            )
        with col4:
            st.session_state.supabase_key = st.text_input(
                "Supabase API Key",
                type="password",
                placeholder="your-supabase-anon-key",
                value=st.session_state.get("supabase_key", "")
            )

# Resolve the active api_key from the correct session state slot
if st.session_state.engine == "Groq API":
    api_key = st.session_state.get("groq_api_key", "")
else:
    api_key = st.session_state.get("openrouter_api_key", "")

# ─────────────────────────────────────────────
# ═══════════════════════════════════════════════
# VIEW: ITEM MASTER DATABASE
# ═══════════════════════════════════════════════
if current_view == "database":
    st.markdown("<h2 style='font-family: \"Playfair Display\", Georgia, serif; font-weight: 800; font-size: 2.2rem; color: #0f172a; margin-top: 1rem; margin-bottom: 0.25rem;'>📊 Item Master Database</h2>", unsafe_allow_html=True)
    st.markdown("<p style='font-family: \"Outfit\", sans-serif; color: #64748b; font-size: 1.05rem; margin-bottom: 1.5rem;'>Browse and manage all previously extracted products stored in Supabase.</p>", unsafe_allow_html=True)

    if st.button("🔄 Refresh Data", type="primary"):
        st.rerun()

    # Resolve client
    supabase_client_inst = get_supabase_client() or supabase_client
    
    if not supabase_client_inst:
        st.error("❌ Supabase client is not initialized. Please configure SUPABASE_URL and SUPABASE_KEY.")
    else:
        try:
            supabase_client = supabase_client_inst
            response = supabase_client.table("imdb_products").select("*").order("created_at", desc=True).execute()
            df = pd.DataFrame(response.data)
            st.dataframe(df)
        except Exception as e:
            st.error(f"❌ Error fetching data: {e}")

    st.stop()   # Do not render pipeline below

# ═══════════════════════════════════════════════
# VIEW: PIPELINE (default)
# ═══════════════════════════════════════════════
st.markdown("<h2 style='font-family: \"Playfair Display\", Georgia, serif; font-weight: 800; font-size: 2.2rem; color: #0f172a; margin-top: 1rem; margin-bottom: 0.25rem;'>🏷️ IMDB Auto-Fill</h2>", unsafe_allow_html=True)
st.markdown("<p style='font-family: \"Outfit\", sans-serif; color: #64748b; font-size: 1.05rem; margin-bottom: 1.5rem;'>Welcome back! Streamline retail master data cataloging with AI-driven extraction.</p>", unsafe_allow_html=True)

# Render the page-level progress stepper
stepper_placeholder = st.empty()
with stepper_placeholder.container():
    render_pipeline_stepper(active_step)

st.divider()

# ─────────────────────────────────────────────
# STEP 1 — ACQUIRE PRODUCT IMAGES
# ─────────────────────────────────────────────
st.markdown("### Step 1 — Acquire Product Images")
st.caption("Provide product images by uploading files or capturing snapshots using your device camera.")

tab_upload, tab_camera = st.tabs(["📁 Upload Files", "📸 Capture Snapshots"])

with tab_upload:
    uploaded_files = st.file_uploader(
        "Drop images here or click to browse",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="uploader"
    )

with tab_camera:
    col_pid, col_clear = st.columns([2, 1])
    with col_pid:
        camera_product_id = st.text_input(
            "Product ID / Prefix",
            value="PRODUCT1",
            help="Specify a prefix (e.g. S12345) to group your camera snapshots."
        ).strip().upper()
    with col_clear:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        if st.button("🗑️ Clear Snapshots", use_container_width=True):
            st.session_state.camera_snapshots = []
            st.success("Snapshot queue cleared!")
            st.rerun()

    camera_photo = st.camera_input("Take a snapshot of the product or product parts", key="camera_widget")

    if camera_photo is None:
        st.session_state.last_added_photo = None

    if camera_photo is not None:
        if not camera_product_id:
            st.warning("⚠️ Please specify a Product ID / Prefix before saving.")
        else:
            photo_bytes = camera_photo.getvalue()
            # Check if this photo was already added to prevent repeat clicks
            is_already_added = (st.session_state.last_added_photo is not None and 
                                st.session_state.last_added_photo == photo_bytes)
            
            if is_already_added:
                st.success("✅ Snapshot successfully added to queue! Tap 'Clear photo' below to capture the next angle/part.")
            else:
                if st.button(f"➕ Add Snapshot to {camera_product_id} Queue", use_container_width=True):
                    timestamp = int(time.time())
                    filename = f"{camera_product_id}_snap_{timestamp}.jpg"
                    
                    # Wrap photo bytes in NamedBytesIO
                    snap_file = NamedBytesIO(photo_bytes, filename)
                    st.session_state.camera_snapshots.append(snap_file)
                    st.session_state.last_added_photo = photo_bytes
                    st.toast(f"✅ Saved `{filename}`!")
                    st.rerun()

    # Display camera snapshot queue grid if any exist
    if st.session_state.camera_snapshots:
        st.markdown("<h4 style='margin-top: 1.5rem;'>📸 Captured Snapshots Queue</h4>", unsafe_allow_html=True)
        snap_cols = st.columns(2)
        for idx, snap in enumerate(st.session_state.camera_snapshots):
            with snap_cols[idx % 2]:
                with st.container(border=True):
                    st.image(snap.getvalue(), use_container_width=True)
                    st.markdown(f"<p class='snap-card-title' style='text-align: center;'>{snap.name}</p>", unsafe_allow_html=True)
                    if st.button("🗑️ Remove", key=f"del_snap_{idx}", use_container_width=True):
                        st.session_state.camera_snapshots.pop(idx)
                        st.rerun()

# Merge all files from both uploader and camera
all_files = []
if uploaded_files:
    all_files.extend(uploaded_files)
if st.session_state.camera_snapshots:
    all_files.extend(st.session_state.camera_snapshots)

if all_files:
    # Group by product ID (filename prefix)
    groups = defaultdict(list)
    for f in all_files:
        product_id = f.name.split("_")[0]
        groups[product_id].append(f)

    col1, col2, col3 = st.columns(3)
    col1.metric("Total images", len(all_files))
    col2.metric("Products detected", len(groups))
    col3.metric("Avg images/product", f"{len(all_files)/len(groups):.1f}" if groups else "0.0")

    # Show image previews
    with st.expander("Preview product images", expanded=False):
        cols = st.columns(5)
        for i, f in enumerate(all_files[:10]):
            with cols[i % 5]:
                if hasattr(f, "seek"):
                    f.seek(0)
                if hasattr(f, "getvalue"):
                    st.image(f.getvalue(), caption=f.name, use_container_width=True)
                else:
                    st.image(f, caption=f.name, use_container_width=True)

    st.divider()

    # ─────────────────────────────────────────────
    # STEP 2 — EXTRACT
    # ─────────────────────────────────────────────
    st.markdown("### Step 2 — Extract & Process")

    checkpoint_path = Path("./checkpoint.json")
    checkpoint_records = []
    done_ids = set()
    if checkpoint_path.exists():
        try:
            with open(checkpoint_path, "r", encoding="utf-8") as f:
                checkpoint_records = json.load(f)
            done_ids = {r["_product_id"] for r in checkpoint_records if "_product_id" in r}
            if done_ids:
                col_info, col_clear = st.columns([3, 1])
                with col_info:
                    st.info(f"ℹ️ **Found checkpoint on disk:** {len(done_ids)} product(s) already processed.")
                with col_clear:
                    if st.button("🗑️ Clear Checkpoint", use_container_width=True):
                        checkpoint_path.unlink(missing_ok=True)
                        st.rerun()
        except Exception as e:
            st.warning(f"⚠️ Could not read checkpoint file: {e}")

    if st.button("🚀 Run Pipeline", use_container_width=True):
        pipeline_success = True
        # Create a dictionary to hold results by product ID to preserve the original order of products
        results_by_product = {}
        for product_id in groups:
            if product_id in done_ids:
                matched_record = next(r for r in checkpoint_records if r.get("_product_id") == product_id)
                results_by_product[product_id] = to_imdb(matched_record)

        # Build a single flat queue of images to process across all non-checkpointed products
        queue = []
        for product_id, files in groups.items():
            if product_id in done_ids:
                continue
            for f in files:
                queue.append((product_id, f))

        total_images = len(queue)
        successful_extractions = 0
        failed_extractions = 0

        if total_images > 0:
            progress = st.progress(0)
            status = st.empty()
            
            extracted_by_product = defaultdict(list)
            processed_count_by_product = defaultdict(int)

            for idx, (product_id, f) in enumerate(queue):
                # Update stepper to "Extract" (active_step = 2)
                with stepper_placeholder.container():
                    render_pipeline_stepper(2)
                
                status.markdown(f"**Processing Image {idx + 1}/{total_images}:** `{f.name}` (Product `{product_id}`)...")
                
                timer_slot = st.empty()
                start = time.time()
                try:
                    b64 = preprocess_image(f, max_size=(1024, 1024))
                    # Read keys directly from session_state (persists across reruns)
                    engine = st.session_state.get("engine", "OpenRouter API")
                    if engine == "Groq API":
                        _api_key = st.session_state.get("groq_api_key", "")
                        if not _api_key:
                            raise ValueError("Groq API key is missing. Please enter it in Model Configuration.")
                        _model = st.session_state.get("groq_model", "llama-3.2-11b-vision-preview")
                        record = extract_via_groq(b64, _api_key, model=_model)
                    elif engine == "OpenRouter API":
                        _api_key = st.session_state.get("openrouter_api_key", "")
                        if not _api_key:
                            raise ValueError("OpenRouter API key is missing. Please enter it in Model Configuration.")
                        _model = st.session_state.get("openrouter_model", "google/gemma-4-26b-a4b-it:free")
                        record = extract_via_openrouter(b64, _api_key, model=_model)
                    else:
                        raise ValueError("No valid engine selected.")

                    elapsed = time.time() - start
                    timer_slot.success(f"✅ `{f.name}` done in {elapsed:.1f}s")
                    extracted_by_product[product_id].append(record)
                    successful_extractions += 1
                except Exception as e:
                    timer_slot.warning(f"⚠️ Could not extract from `{f.name}`: {e}")
                    failed_extractions += 1

                # Enforce minimum gap between API calls — only sleep what's left of the window
                if idx < total_images - 1:
                    elapsed_total = time.time() - start
                    remaining = GROQ_DELAY_SEC - elapsed_total
                    if remaining > 0:
                        status.markdown(f"⏳ Pacing... waiting {remaining:.1f}s before next call")
                        time.sleep(remaining)

                processed_count_by_product[product_id] += 1
                if processed_count_by_product[product_id] == len(groups[product_id]):
                    # All images for this product have been processed, let's aggregate and validate!
                    if extracted_by_product[product_id]:
                        with stepper_placeholder.container():
                            render_pipeline_stepper(3)
                        merged = aggregate(extracted_by_product[product_id])
                        
                        with stepper_placeholder.container():
                            render_pipeline_stepper(4)
                        merged = validate(merged)
                        
                        merged["_product_id"] = product_id
                        checkpoint_records.append(merged)
                        try:
                            with open(checkpoint_path, "w", encoding="utf-8") as f_out:
                                json.dump(checkpoint_records, f_out, ensure_ascii=False, indent=2)
                        except Exception as e:
                            st.warning(f"⚠️ Could not save checkpoint: {e}")
                            
                        results_by_product[product_id] = to_imdb(merged)

                progress.progress((idx + 1) / total_images)

            status.empty()
            progress.empty()

            if successful_extractions == 0:
                st.error("❌ Pipeline could not complete — extraction failed for all images. Please check your API key or try again in 60 seconds.")
                pipeline_success = False
            elif failed_extractions > 0:
                st.warning("⚠️ Some images could not be extracted. Proceeding with available data.")

        if pipeline_success:
            # Reassemble the final results in the exact original order of the products
            results = [results_by_product[pid] for pid in groups if pid in results_by_product]
            
            st.session_state.results = results
            st.session_state.edited_df = pd.DataFrame(results, columns=IMDB_COLS)

            # Clean up checkpoint after successful full completion
            if checkpoint_path.exists():
                try:
                    checkpoint_path.unlink()
                except Exception as e:
                    st.warning(f"⚠️ Could not delete checkpoint file: {e}")

            st.markdown('<div class="success-banner">✅ Pipeline complete!</div>', unsafe_allow_html=True)

            # ── Auto-sync to Supabase ─────────────────────────────────────
            # Only attempt Supabase sync if at least one product was successfully extracted.
            if results:
                url = st.session_state.get("supabase_url", SUPABASE_URL)
                key = st.session_state.get("supabase_key", SUPABASE_KEY)
                if url and key:
                    try:
                        client = get_supabase_client()
                        if client:
                            sync_success = True
                            for imdb_row in results:
                                success, msg = supabase_upsert_product(imdb_row)
                                if not success:
                                    sync_success = False
                                    break
                            if sync_success:
                                st.success("✅ Synced to Item Master Database")
                            else:
                                st.warning("⚠️ Extraction complete but could not sync to database. Your export is still available.")
                        else:
                            st.warning("⚠️ Extraction complete but could not sync to database. Your export is still available.")
                    except Exception as e:
                        st.warning("⚠️ Extraction complete but could not sync to database. Your export is still available.")

            # Update stepper to "Export/Preview" (active_step = 5)
            with stepper_placeholder.container():
                render_pipeline_stepper(5)

# ─────────────────────────────────────────────
# STEP 3 — PREVIEW & EDIT
# ─────────────────────────────────────────────
if st.session_state.results:
    st.divider()
    st.markdown("### Step 3 — Preview & Edit")
    st.caption("Click any cell to edit before exporting.")

    edited = st.data_editor(
        st.session_state.edited_df,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True
    )
    st.session_state.edited_df = edited

    # Stats
    df = edited
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total products", len(df))
    col2.metric("With barcode", int(df["BARCODE"].astype(bool).sum()))
    col3.metric("With manufacturer", int(df["MANUFACTURER"].astype(bool).sum()))
    col4.metric("With country", int(df["COUNTRY"].astype(bool).sum()))

    st.divider()

    # ─────────────────────────────────────────────
    # STEP 4 — EXPORT
    # ─────────────────────────────────────────────
    st.markdown("### Step 4 — Export")
    col1, col2 = st.columns(2)

    with col1:
        excel_buffer = build_excel(edited.to_dict("records"))
        st.download_button(
            label="📥 Download as Excel (.xlsx)",
            data=excel_buffer,
            file_name="IMDB_predictions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        csv = edited.to_csv(index=False)
        st.download_button(
            label="📥 Download as CSV (.csv)",
            data=csv,
            file_name="IMDB_predictions.csv",
            mime="text/csv",
            use_container_width=True
        )
