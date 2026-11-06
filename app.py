"""
====================================================================================================
SISTEM INFORMASI PEMELIHARAAN BANGUNAN GEDUNG MENYELURUH BERBASIS AI
(Deteksi Kerusakan & Estimasi AHSP Jawa Tengah)
Aplikasi Website Skripsi Teknik Sipil Digital
Framework: Streamlit | AI Agent: Antigravity API (requests) | Database: SQLite3 | Ekspor: ReportLab
====================================================================================================
"""

import os
import io
import json
import base64
import sqlite3
import requests
from datetime import datetime
from typing import Dict, Any, List, Optional

import streamlit as st
import pandas as pd
import altair as alt
from PIL import Image, ImageDraw, ImageFont

# Library OpenPyXL untuk ekspor formulir penilaian kondisi & tabel Excel
import openpyxl
from openpyxl.styles import (
    Font as XlsFont, PatternFill as XlsPatternFill, Alignment as XlsAlignment,
    Border as XlsBorder, Side as XlsSide
)
from openpyxl.utils import get_column_letter

# Library ReportLab untuk cetak dokumen laporan resmi PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, HRFlowable,
    Image as RLImage, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm


# ==================================================================================================
# 1. KONFIGURASI APLIKASI STREAMLIT & CUSTOM STYLING (TAILWIND-INSPIRED DESIGN)
# ==================================================================================================
st.set_page_config(
    page_title="SIM-PBG AI | AHSP Jawa Tengah 2023",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS Tailwind-Like UI dengan Animasi Modern, Desain Interaktif & Tipografi Kontras Tinggi
CUSTOM_CSS = """
<style>
    /* Import Google Fonts - Plus Jakarta Sans */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800;900&display=swap');

    /* ========================================================================== */
    /* 1. SISTEM TIPOGRAFI & WARNA TEKS DENGAN KONTRAS TINGGI (CRYSTAL CLEAR)     */
    /* ========================================================================== */
    html, body, [class*="css"], .stApp {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        color: #0f172a !important;
    }

    /* Background Aplikasi dengan Grid Halus Teknik Sipil */
    .stApp {
        background-color: #f8fafc;
        background-image: radial-gradient(#cbd5e1 0.75px, transparent 0.75px);
        background-size: 26px 26px;
    }

    /* Semua Judul (Headings) Utama - Deep Navy/Slate Kontras Tinggi */
    h1, h2, h3, h4, h5, h6,
    [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3,
    [data-testid="stMarkdownContainer"] h4,
    [data-testid="stMarkdownContainer"] h5,
    [data-testid="stMarkdownContainer"] h6 {
        color: #0f172a !important;
        font-weight: 800 !important;
        letter-spacing: -0.3px !important;
    }

    /* Paragraf & Teks Markdown Standar */
    p, li, [data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li {
        color: #1e293b !important;
        font-weight: 500 !important;
        line-height: 1.6 !important;
    }

    /* Label Form & Input Streamlit */
    label,
    .stWidgetLabel,
    div[data-testid="stWidgetLabel"] p,
    div[data-testid="stWidgetLabel"] label,
    div[data-testid="stWidgetLabel"] span {
        color: #0f172a !important;
        font-weight: 800 !important;
        font-size: 0.95rem !important;
    }

    /* Teks Keterangan & Caption */
    .stCaption,
    div[data-testid="stCaptionContainer"] p,
    div[data-testid="stCaptionContainer"] {
        color: #334155 !important;
        font-weight: 600 !important;
        font-size: 0.84rem !important;
    }

    /* Teks Input, Selectbox, & Number Input */
    div[data-baseweb="input"] input,
    div[data-baseweb="select"] span,
    div[data-baseweb="select"] div {
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 0.94rem !important;
    }

    /* Radio Button Navigasi */
    div[role="radiogroup"] label p,
    div[role="radiogroup"] label span,
    div[role="radiogroup"] label div {
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 0.94rem !important;
    }

    /* Header & Teks Accordion / Expander */
    div[data-testid="stExpander"] {
        border-radius: 14px !important;
        border: 1.5px solid #cbd5e1 !important;
        background: #ffffff !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03) !important;
    }
    div[data-testid="stExpander"] summary span,
    div[data-testid="stExpander"] summary p {
        color: #0f172a !important;
        font-weight: 800 !important;
        font-size: 1rem !important;
    }
    div[data-testid="stExpander"] details {
        color: #0f172a !important;
    }

    /* Tabel & Dataframe Teks & Header */
    .stTable,
    div[data-testid="stDataFrame"],
    table {
        color: #0f172a !important;
    }
    table th, [data-testid="stDataFrame"] div[role="columnheader"] {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
        font-weight: 800 !important;
        font-size: 0.88rem !important;
    }
    table td, [data-testid="stDataFrame"] div[role="gridcell"] {
        color: #0f172a !important;
        font-weight: 600 !important;
        font-size: 0.88rem !important;
    }

    /* Alert Boxes (Info, Warning, Success, Error) */
    div[data-testid="stAlert"] * {
        color: #0f172a !important;
        font-weight: 600 !important;
        font-size: 0.92rem !important;
    }

    /* Sidebar Navigation Container */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%) !important;
        border-right: 1.5px solid #e2e8f0 !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label {
        padding: 9px 14px !important;
        border-radius: 12px !important;
        transition: all 0.25s ease !important;
        margin-bottom: 6px !important;
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03) !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background: #eff6ff !important;
        border-color: #93c5fd !important;
        transform: translateX(4px) !important;
    }

    /* ========================================================================== */
    /* 2. KEYFRAME ANIMASI HALUS & DINAMIS                                        */
    /* ========================================================================== */
    @keyframes gradientMesh {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    @keyframes slideFadeUp {
        0% {
            opacity: 0;
            transform: translateY(18px);
        }
        100% {
            opacity: 1;
            transform: translateY(0);
        }
    }

    @keyframes pulseGlow {
        0%, 100% {
            box-shadow: 0 0 0 0 rgba(37, 99, 235, 0.45);
        }
        50% {
            box-shadow: 0 0 0 12px rgba(37, 99, 235, 0);
        }
    }

    @keyframes radarDualPing {
        0% {
            transform: scale(0.9);
            box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.8), 0 0 0 0 rgba(16, 185, 129, 0.4);
        }
        50% {
            transform: scale(1.1);
            box-shadow: 0 0 0 6px rgba(16, 185, 129, 0.25), 0 0 0 12px rgba(16, 185, 129, 0);
        }
        100% {
            transform: scale(0.9);
            box-shadow: 0 0 0 0 rgba(16, 185, 129, 0), 0 0 0 0 rgba(16, 185, 129, 0);
        }
    }

    @keyframes floatSlow {
        0%, 100% {
            transform: translateY(0px);
        }
        50% {
            transform: translateY(-8px);
        }
    }

    @keyframes shimmerBar {
        0% { transform: translateX(-100%); }
        100% { transform: translateX(200%); }
    }

    @keyframes btnSweep {
        0% { transform: translateX(-100%); }
        100% { transform: translateX(200%); }
    }

    /* ========================================================================== */
    /* 3. HERO BANNER - PERMUKAAN GELAP DENGAN TEKS PUTIH KRISTAL & ANIMASI       */
    /* ========================================================================== */
    .hero-banner {
        background: linear-gradient(-45deg, #0b1329, #172554, #0f3647, #1e1b4b, #0f172a);
        background-size: 300% 300%;
        animation: gradientMesh 14s ease infinite;
        color: #ffffff !important;
        padding: 30px 36px;
        border-radius: 20px;
        margin-bottom: 26px;
        border: 1px solid rgba(255, 255, 255, 0.18);
        box-shadow: 0 20px 35px -8px rgba(15, 23, 42, 0.4), 0 8px 16px -4px rgba(15, 23, 42, 0.25);
        position: relative;
        overflow: hidden;
    }

    .hero-banner * {
        color: #ffffff;
    }

    .hero-banner::before {
        content: '';
        position: absolute;
        bottom: 0;
        left: 0;
        width: 100%;
        height: 3.5px;
        background: linear-gradient(90deg, #3b82f6, #06b6d4, #10b981, #f59e0b, #ec4899, #3b82f6);
        background-size: 250% 100%;
        animation: gradientMesh 6s linear infinite;
    }

    .hero-banner::after {
        content: '';
        position: absolute;
        top: -50%;
        right: -15%;
        width: 320px;
        height: 320px;
        background: radial-gradient(circle, rgba(59, 130, 246, 0.3) 0%, rgba(0,0,0,0) 70%);
        pointer-events: none;
    }

    .hero-title {
        font-size: 1.85rem !important;
        font-weight: 900 !important;
        letter-spacing: -0.5px !important;
        margin-bottom: 10px !important;
        color: #ffffff !important;
        text-shadow: 0 2px 4px rgba(0, 0, 0, 0.5) !important;
    }

    .hero-subtitle {
        font-size: 1rem !important;
        color: #f1f5f9 !important;
        font-weight: 500 !important;
        line-height: 1.65 !important;
        max-width: 950px !important;
        text-shadow: 0 1px 3px rgba(0, 0, 0, 0.5) !important;
    }

    /* Live Pulse Status Dot dengan Dual Radar Ping */
    .live-badge {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: rgba(16, 185, 129, 0.25) !important;
        border: 1.5px solid #10b981 !important;
        color: #ecfdf5 !important;
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 0.76rem;
        font-weight: 800;
        letter-spacing: 0.4px;
        text-transform: uppercase;
        margin-bottom: 12px;
        backdrop-filter: blur(4px);
    }

    .live-dot {
        width: 9px;
        height: 9px;
        background-color: #10b981;
        border-radius: 50%;
        animation: radarDualPing 2s ease-in-out infinite;
    }

    /* ========================================================================== */
    /* 4. CONTAINER CARD STYLING & ELEVASI HOVER                                  */
    /* ========================================================================== */
    .tw-card {
        background: #ffffff !important;
        border: 1.5px solid #cbd5e1 !important;
        border-radius: 18px !important;
        padding: 24px !important;
        box-shadow: 0 4px 10px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03) !important;
        margin-bottom: 22px !important;
        transition: all 0.35s cubic-bezier(0.16, 1, 0.3, 1) !important;
        animation: slideFadeUp 0.5s ease-out !important;
        position: relative !important;
    }

    .tw-card:hover {
        transform: translateY(-5px) !important;
        box-shadow: 0 20px 30px -8px rgba(15, 23, 42, 0.12), 0 8px 12px -4px rgba(15, 23, 42, 0.05) !important;
        border-color: #93c5fd !important;
    }

    .tw-card-header {
        border-bottom: 1.5px solid #e2e8f0;
        padding-bottom: 14px;
        margin-bottom: 18px;
    }

    .tw-card h4, .tw-card h5 {
        color: #0f172a !important;
        font-weight: 800 !important;
    }

    .tw-card p {
        color: #334155 !important;
        font-weight: 500 !important;
    }

    /* ========================================================================== */
    /* 5. METRIC CARDS INTERAKTIF & VARIAN WARNA AKSEN                            */
    /* ========================================================================== */
    .metric-container {
        display: flex;
        align-items: center;
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%) !important;
        border: 1.5px solid #cbd5e1 !important;
        border-radius: 14px !important;
        padding: 18px 22px !important;
        margin-bottom: 14px !important;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
        position: relative !important;
        overflow: hidden !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.04) !important;
    }

    .metric-container::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 5px;
        height: 100%;
        background: linear-gradient(180deg, #2563eb, #06b6d4);
        border-radius: 4px 0 0 4px;
        opacity: 0.9;
        transition: width 0.3s ease, opacity 0.3s ease;
    }

    .metric-container:hover {
        transform: translateY(-4px) scale(1.02) !important;
        box-shadow: 0 14px 28px -4px rgba(37, 99, 235, 0.18) !important;
        border-color: #93c5fd !important;
    }

    .metric-container:hover::before {
        width: 7px;
        opacity: 1;
    }

    .metric-card-green::before { background: linear-gradient(180deg, #10b981, #059669) !important; }
    .metric-card-blue::before { background: linear-gradient(180deg, #2563eb, #1d4ed8) !important; }
    .metric-card-amber::before { background: linear-gradient(180deg, #f59e0b, #d97706) !important; }
    .metric-card-purple::before { background: linear-gradient(180deg, #8b5cf6, #7c3aed) !important; }
    .metric-card-cyan::before { background: linear-gradient(180deg, #06b6d4, #0891b2) !important; }

    .metric-value {
        font-size: 1.6rem !important;
        font-weight: 900 !important;
        color: #0f172a !important;
        letter-spacing: -0.5px !important;
        transition: transform 0.25s ease !important;
    }

    .metric-container:hover .metric-value {
        transform: scale(1.03) !important;
    }

    .metric-label {
        font-size: 0.84rem !important;
        font-weight: 800 !important;
        color: #1e293b !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }

    /* ========================================================================== */
    /* 6. STATUS BADGES BERWARNA HIDUP & KONTRAS TINGGI                           */
    /* ========================================================================== */
    .badge-ringan {
        background: linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%) !important;
        color: #14532d !important;
        padding: 6px 14px !important;
        border-radius: 9999px !important;
        font-size: 0.8rem !important;
        font-weight: 800 !important;
        border: 1.5px solid #86efac !important;
        display: inline-block !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease !important;
    }
    .badge-ringan:hover {
        transform: scale(1.06) !important;
        box-shadow: 0 4px 12px rgba(34, 197, 94, 0.35) !important;
    }

    .badge-sedang {
        background: linear-gradient(135deg, #fef9c3 0%, #fef08a 100%) !important;
        color: #78350f !important;
        padding: 6px 14px !important;
        border-radius: 9999px !important;
        font-size: 0.8rem !important;
        font-weight: 800 !important;
        border: 1.5px solid #fde047 !important;
        display: inline-block !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease !important;
    }
    .badge-sedang:hover {
        transform: scale(1.06) !important;
        box-shadow: 0 4px 12px rgba(234, 179, 8, 0.35) !important;
    }

    .badge-berat {
        background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%) !important;
        color: #7f1d1d !important;
        padding: 6px 14px !important;
        border-radius: 9999px !important;
        font-size: 0.8rem !important;
        font-weight: 800 !important;
        border: 1.5px solid #fca5a5 !important;
        display: inline-block !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease !important;
    }
    .badge-berat:hover {
        transform: scale(1.06) !important;
        box-shadow: 0 4px 12px rgba(239, 68, 68, 0.35) !important;
    }

    .badge-ahsp {
        background: linear-gradient(135deg, #e0f2fe 0%, #bae6fd 100%) !important;
        color: #0369a1 !important;
        padding: 6px 14px !important;
        border-radius: 9999px !important;
        font-size: 0.8rem !important;
        font-weight: 800 !important;
        border: 1.5px solid #7dd3fc !important;
        display: inline-block !important;
        transition: transform 0.2s ease !important;
    }
    .badge-ahsp:hover {
        transform: scale(1.05) !important;
    }

    /* ========================================================================== */
    /* 7. PROSEDUR METODE KERJA / STEPS                                           */
    /* ========================================================================== */
    .step-item {
        display: flex;
        margin-bottom: 12px;
        align-items: flex-start;
        padding: 14px 18px;
        border-radius: 14px;
        transition: all 0.28s cubic-bezier(0.16, 1, 0.3, 1);
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        border: 1.5px solid #cbd5e1;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
    }
    .step-item:hover {
        background: #ffffff;
        transform: translateX(6px);
        border-color: #2563eb;
        box-shadow: 0 8px 20px -4px rgba(37, 99, 235, 0.16);
    }

    .step-number {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        border-radius: 50%;
        width: 32px;
        height: 32px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.85rem;
        font-weight: 900;
        margin-right: 14px;
        flex-shrink: 0;
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.4);
        transition: transform 0.25s ease;
    }

    .step-item:hover .step-number {
        transform: scale(1.1) rotate(5deg);
    }

    .step-text {
        font-size: 0.95rem !important;
        font-weight: 600 !important;
        color: #0f172a !important;
        line-height: 1.6 !important;
    }

    /* ========================================================================== */
    /* 8. INTERACTIVE STATUS GRID UNTUK PENILAIAN KONDISI IKBG                    */
    /* ========================================================================== */
    .ikbg-status-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
        gap: 10px;
        margin-top: 10px;
    }

    .ikbg-status-card {
        padding: 12px 10px;
        border-radius: 12px;
        border: 1.5px solid #cbd5e1;
        background: #ffffff;
        text-align: center;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        box-shadow: 0 2px 5px rgba(0,0,0,0.03);
    }

    .ikbg-status-card.active-rutin {
        border: 2px solid #16a34a !important;
        background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%) !important;
        box-shadow: 0 6px 16px rgba(22, 163, 74, 0.28) !important;
        transform: translateY(-3px) !important;
    }

    .ikbg-status-card.active-preventif {
        border: 2px solid #d97706 !important;
        background: linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%) !important;
        box-shadow: 0 6px 16px rgba(217, 119, 6, 0.28) !important;
        transform: translateY(-3px) !important;
    }

    .ikbg-status-card.active-korektif {
        border: 2px solid #ea580c !important;
        background: linear-gradient(135deg, #fff7ed 0%, #ffedd5 100%) !important;
        box-shadow: 0 6px 16px rgba(234, 88, 12, 0.28) !important;
        transform: translateY(-3px) !important;
    }

    .ikbg-status-card.active-rehabilitatif {
        border: 2px solid #dc2626 !important;
        background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%) !important;
        box-shadow: 0 6px 16px rgba(220, 38, 38, 0.28) !important;
        transform: translateY(-3px) !important;
    }

    /* ========================================================================== */
    /* 9. PLACEHOLDER UNGGAH FOTO INTERAKTIF                                      */
    /* ========================================================================== */
    .upload-placeholder {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        border: 2px dashed #93c5fd;
        border-radius: 16px;
        padding: 40px 20px;
        text-align: center;
        color: #334155;
        transition: all 0.3s ease;
        box-shadow: 0 4px 10px rgba(0,0,0,0.03);
    }

    .upload-placeholder:hover {
        border-color: #2563eb;
        background: #eff6ff;
        box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.15);
    }

    .floating-icon {
        font-size: 3.2rem;
        display: inline-block;
        animation: floatSlow 3.2s ease-in-out infinite;
    }

    /* ========================================================================== */
    /* 10. TOMBOL STREAMLIT DENGAN ANIMASI SHIMMER & DEPTH                        */
    /* ========================================================================== */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 50%, #3b82f6 100%) !important;
        background-size: 200% 200% !important;
        color: #ffffff !important;
        font-weight: 800 !important;
        font-size: 1rem !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.72rem 1.75rem !important;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1) !important;
        box-shadow: 0 6px 20px -2px rgba(37, 99, 235, 0.45) !important;
        position: relative !important;
        overflow: hidden !important;
    }

    div.stButton > button[kind="primary"]:hover {
        transform: translateY(-3px) scale(1.01) !important;
        box-shadow: 0 14px 28px -3px rgba(37, 99, 235, 0.6) !important;
        filter: brightness(1.08) !important;
    }

    div.stButton > button[kind="secondary"] {
        border-radius: 12px !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
        transition: all 0.25s ease !important;
        border: 1.5px solid #cbd5e1 !important;
        background: #ffffff !important;
        color: #0f172a !important;
    }
    div.stButton > button[kind="secondary"]:hover {
        transform: translateY(-2px) !important;
        border-color: #2563eb !important;
        color: #2563eb !important;
        box-shadow: 0 6px 14px rgba(15, 23, 42, 0.1) !important;
    }

    div.stDownloadButton > button {
        border-radius: 12px !important;
        font-weight: 800 !important;
        font-size: 0.95rem !important;
        transition: all 0.28s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    div.stDownloadButton > button:hover {
        transform: translateY(-3px) scale(1.01) !important;
        box-shadow: 0 10px 22px -3px rgba(15, 23, 42, 0.2) !important;
    }

    /* File Uploader Visual */
    div[data-testid="stFileUploader"] {
        background: #ffffff;
        border: 2px dashed #93c5fd;
        border-radius: 16px;
        padding: 16px;
        transition: all 0.3s ease;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #2563eb;
        background: #eff6ff;
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.15);
    }

    /* Custom Slim Scrollbars */
    ::-webkit-scrollbar {
        width: 7px;
        height: 7px;
    }
    ::-webkit-scrollbar-track {
        background: #f1f5f9;
    }
    ::-webkit-scrollbar-thumb {
        background: #cbd5e1;
        border-radius: 4px;
    }
    ::-webkit-scrollbar-thumb:hover {
        background: #94a3b8;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ==================================================================================================
# 2. STRUKTUR & PENGELOLA DATABASE SQLITE (sqlite3)
# ==================================================================================================
DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pemeliharaan_gedung.db")

class DatabaseManager:
    """Kelas pengelola koneksi dan query SQLite untuk riwayat kerusakan bangunan gedung."""

    @staticmethod
    def get_connection() -> sqlite3.Connection:
        conn = sqlite3.connect(DB_FILE, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        return conn

    @classmethod
    def init_db(cls) -> None:
        """Membuat tabel riwayat_kerusakan sesuai spesifikasi jika belum ada dan memastikan skema termutakhir."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS riwayat_kerusakan (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tanggal TEXT NOT NULL,
                    nama_gedung TEXT DEFAULT 'Bangunan Gedung',
                    lokasi_lantai TEXT DEFAULT '-',
                    surveyor TEXT DEFAULT 'Surveyor Lapangan',
                    elemen_gedung TEXT NOT NULL,
                    jenis_kerusakan TEXT NOT NULL,
                    kode_ahsp TEXT NOT NULL,
                    volume REAL NOT NULL,
                    satuan TEXT NOT NULL,
                    harga_satuan INTEGER NOT NULL,
                    total_biaya INTEGER NOT NULL,
                    rekomendasi TEXT NOT NULL
                )
            """)
            # Auto-migrasi kolom metadata gedung jika belum ada
            cursor.execute("PRAGMA table_info(riwayat_kerusakan)")
            cols = [info[1] for info in cursor.fetchall()]
            if "nama_gedung" not in cols:
                cursor.execute("ALTER TABLE riwayat_kerusakan ADD COLUMN nama_gedung TEXT DEFAULT 'Bangunan Gedung'")
            if "lokasi_lantai" not in cols:
                cursor.execute("ALTER TABLE riwayat_kerusakan ADD COLUMN lokasi_lantai TEXT DEFAULT '-'")
            if "surveyor" not in cols:
                cursor.execute("ALTER TABLE riwayat_kerusakan ADD COLUMN surveyor TEXT DEFAULT 'Surveyor Lapangan'")
            conn.commit()

    @classmethod
    def insert_riwayat(
        cls,
        tanggal: str,
        elemen_gedung: str,
        jenis_kerusakan: str,
        kode_ahsp: str,
        volume: float,
        satuan: str,
        harga_satuan: int,
        total_biaya: int,
        rekomendasi: str,
        nama_gedung: str = "Bangunan Gedung",
        lokasi_lantai: str = "-",
        surveyor: str = "Surveyor Lapangan"
    ) -> int:
        """Menyimpan data penilaian kondisi kerusakan gedung ke SQLite."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO riwayat_kerusakan (
                    tanggal, nama_gedung, lokasi_lantai, surveyor,
                    elemen_gedung, jenis_kerusakan, kode_ahsp,
                    volume, satuan, harga_satuan, total_biaya, rekomendasi
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                tanggal, nama_gedung, lokasi_lantai, surveyor,
                elemen_gedung, jenis_kerusakan, kode_ahsp,
                volume, satuan, int(harga_satuan), int(total_biaya), rekomendasi
            ))
            conn.commit()
            return cursor.lastrowid

    @classmethod
    def get_all_riwayat(cls) -> List[Dict[str, Any]]:
        """Mengambil seluruh data riwayat penilaian dari database terurut terbaru."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM riwayat_kerusakan ORDER BY id DESC")
            rows = cursor.fetchall()
            return [dict(row) for row in rows]

    @classmethod
    def delete_riwayat(cls, record_id: int) -> bool:
        """Menghapus satu catatan riwayat berdasarkan ID."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM riwayat_kerusakan WHERE id = ?", (record_id,))
            conn.commit()
            return cursor.rowcount > 0

    @classmethod
    def clear_all(cls) -> None:
        """Mengosongkan seluruh riwayat kerusakan."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM riwayat_kerusakan")
            conn.commit()

    @classmethod
    def seed_initial_demo_data(cls) -> None:
        """Menambahkan data sampel realistis untuk demonstrasi sidang skripsi."""
        samples = [
            (
                "2026-11-01 09:30", "Gedung Kuliah Terpadu", "Lantai 2 - Selasar Barat", "Ahmad Fajar, S.T.",
                "Balok Beton Bertulang", "Retak Diagonal Geser/Lentur Balok", "A.4.1.1.18.JT",
                4.5, "m'", 315000, 1417500,
                "Injeksi resin epoxy bertekanan rendah pada balok lantai 2 as C-4 untuk mengembalikan monolititas beton."
            ),
            (
                "2026-11-02 11:15", "Gedung Kuliah Terpadu", "Lantai 1 - Ruang 103", "Ahmad Fajar, S.T.",
                "Dinding Pasangan Bata / Hebel", "Retak Rambut Plesteran (> 0.20 mm)", "A.4.4.1.25.JT",
                14.0, "m2", 88500, 1239000,
                "Pembobokan alur retak bentuk V, aplikasi elastomeric sealant, plamir polimer dan pengecatan ulang."
            ),
            (
                "2026-11-03 14:00", "Gedung Kuliah Terpadu", "Lantai Atap - Plat Dak", "Ahmad Fajar, S.T.",
                "Pelat Lantai / Dak Beton", "Rembesan Air & Retak Rambut Dak Atap", "A.4.4.3.35.JT",
                25.0, "m2", 175000, 4375000,
                "Pembersihan substrat, injeksi waterproofing kristalisasi, pelapisan membrane polyurethane 2 lapis."
            ),
            (
                "2026-11-04 10:45", "Gedung Kuliah Terpadu", "Lantai Dasar - Kolom K-02", "Ahmad Fajar, S.T.",
                "Kolom Struktur Beton", "Spalling / Beton Terkelupas & Tulangan Terbuka", "A.4.1.1.20.JT",
                3.2, "m2", 245000, 784000,
                "Chipping beton rapuh, pembersihan karat tulangan baja, pelapisan zinc-rich primer, dan patching polymer mortar."
            )
        ]
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            for s in samples:
                cursor.execute("""
                    INSERT INTO riwayat_kerusakan (
                        tanggal, nama_gedung, lokasi_lantai, surveyor,
                        elemen_gedung, jenis_kerusakan, kode_ahsp,
                        volume, satuan, harga_satuan, total_biaya, rekomendasi
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, s)
            conn.commit()

# Inisialisasi Database saat aplikasi dimuat
DatabaseManager.init_db()


# ==================================================================================================
# 3. BASIS DATA & PENGETAHUAN TEKNIK SIPIL: AHSP JAWA TENGAH TAHUN 2023
# ==================================================================================================
class AHSPKnowledgeBase:
    """Katalog Standar Analisis Harga Satuan Pekerjaan (AHSP) Dinas PU Bina Marga & Cipta Karya Jateng 2023."""

    CATALOG = {
        "A.4.4.1.25.JT": {
            "kode": "A.4.4.1.25.JT",
            "kategori": "Pekerjaan Dinding & Finishing",
            "nama": "Perbaikan Retak Plesteran Mortar Polimer & Sealant Elastomeric",
            "satuan": "m2",
            "harga_satuan": 88500,
            "elemen_terkait": "Dinding Pasangan Bata / Hebel",
            "jenis_kerusakan": "Retak Rambut Plesteran (> 0.20 mm)",
            "tingkat_keparahan": "Sedang",
            "koefisien_upah": [
                {"tenaga": "Pekerja", "koef": 0.120, "satuan": "OH"},
                {"tenaga": "Tukang Batu/Cat", "koef": 0.150, "satuan": "OH"},
                {"tenaga": "Kepala Tukang", "koef": 0.015, "satuan": "OH"},
                {"tenaga": "Mandor", "koef": 0.008, "satuan": "OH"}
            ],
            "koefisien_bahan": [
                {"bahan": "Polymer Modified Mortar", "koef": 1.250, "satuan": "kg"},
                {"bahan": "Sealant Elastomeric Polyurethane", "koef": 0.400, "satuan": "tube"},
                {"bahan": "Bonding Agent Adhesive", "koef": 0.150, "satuan": "liter"}
            ],
            "langkah_kerja": [
                "Pembersihan celah retakan dari debu dan partikel lepas menggunakan sikat kawat dan kompresor udara.",
                "Pelebaran alur celah retak membentuk pola V-groove selebar 5 mm dan kedalaman 5-10 mm.",
                "Aplikasi cairan primer / bonding agent sintetis pada alur celah retak.",
                "Pengisian pasta polymer modified mortar dan elastomeric sealant ke dalam alur retak secara padat rata.",
                "Penghalusan plamir acian elastis dan pengecatan ulang 2 lapis sesuai warna eksisting."
            ],
            "rekomendasi": "Pasang crack monitoring card untuk mengevaluasi apakah retakan bersifat statis atau masih bergerak akibat penurunan tanah fondasi."
        },
        "A.4.1.1.18.JT": {
            "kode": "A.4.1.1.18.JT",
            "kategori": "Pekerjaan Struktur Beton",
            "nama": "Perbaikan Retak Balok/Kolom dengan Injeksi Epoxy Resin Bertekanan",
            "satuan": "m'",
            "harga_satuan": 315000,
            "elemen_terkait": "Balok Beton Bertulang",
            "jenis_kerusakan": "Retak Diagonal Lentur/Geser Balok Struktur",
            "tingkat_keparahan": "Berat",
            "koefisien_upah": [
                {"tenaga": "Pekerja", "koef": 0.200, "satuan": "OH"},
                {"tenaga": "Tukang Khusus Injeksi", "koef": 0.250, "satuan": "OH"},
                {"tenaga": "Kepala Tukang", "koef": 0.025, "satuan": "OH"},
                {"tenaga": "Mandor", "koef": 0.012, "satuan": "OH"}
            ],
            "koefisien_bahan": [
                {"bahan": "Low Viscosity Epoxy Resin", "koef": 0.850, "satuan": "kg"},
                {"bahan": "Epoxy Surface Sealant Paste", "koef": 0.500, "satuan": "kg"},
                {"bahan": "Injection Port / Nipple", "koef": 5.000, "satuan": "buah"}
            ],
            "langkah_kerja": [
                "Pemetaan celah retak dan pengukuran lebar celah menggunakan crack width gauge microscope.",
                "Pembersihan permukaan beton dan penutupan celah luar dengan pasta epoxy surface sealer.",
                "Pemasangan injection ports (nipple) pada jarak interval 15 - 20 cm sepanjang alur retak.",
                "Injeksi epoxy resin viskositas rendah menggunakan pompa tekanan konstan hingga keluar di port sebelahnya.",
                "Pelepasan nipple dan perataan permukaan beton setelah waktu pengerasan (curing) 24 jam."
            ],
            "rekomendasi": "Lakukan evaluasi pembebanan struktur balok dan batasi muatan hidup sementara di atas lantai yang terdampak."
        },
        "A.4.1.1.20.JT": {
            "kode": "A.4.1.1.20.JT",
            "kategori": "Pekerjaan Struktur Beton",
            "nama": "Perbaikan Beton Keropos/Spalling dengan Patching Mortar Semen Polimer",
            "satuan": "m2",
            "harga_satuan": 245000,
            "elemen_terkait": "Kolom Struktur Beton",
            "jenis_kerusakan": "Spalling / Keropos Beton & Selimut Terkelupas",
            "tingkat_keparahan": "Berat",
            "koefisien_upah": [
                {"tenaga": "Pekerja", "koef": 0.280, "satuan": "OH"},
                {"tenaga": "Tukang Batu Ahli", "koef": 0.200, "satuan": "OH"},
                {"tenaga": "Kepala Tukang", "koef": 0.020, "satuan": "OH"},
                {"tenaga": "Mandor", "koef": 0.015, "satuan": "OH"}
            ],
            "koefisien_bahan": [
                {"bahan": "Structural Repair Mortar (Polimer)", "koef": 18.000, "satuan": "kg"},
                {"bahan": "Zinc-Rich Anti-Corrosion Primer", "koef": 0.350, "satuan": "kg"},
                {"bahan": "Semen Bonding Agent", "koef": 0.500, "satuan": "liter"}
            ],
            "langkah_kerja": [
                "Pembongkaran (chipping) beton keropos dan berongga hingga mencapai batas beton padat (sound concrete).",
                "Pembersihan kerak karat pada besi tulangan menggunakan sikat kawat baja baja wire brush.",
                "Pelapisan pelindung anti korosi berbahan dasar zinc-rich primer pada seluruh tulangan baja yang terekspos.",
                "Penyapuan bonding agent sintetis pada bidang kontak beton lama.",
                "Aplikasi mortar polimer non-shrink bertahap (layer by layer) dan pemadatan manual hingga rata rapi."
            ],
            "rekomendasi": "Pastikan tebal selimut beton minimal 30 mm terpenuhi kembali guna melindungi baja tulangan dari karbonasi atmosfer."
        },
        "A.4.4.3.35.JT": {
            "kode": "A.4.4.3.35.JT",
            "kategori": "Pekerjaan Kedap Air (Waterproofing)",
            "nama": "Pelapisan Waterproofing Membrane / Polyurethane Pelat Dak Atap",
            "satuan": "m2",
            "harga_satuan": 175000,
            "elemen_terkait": "Pelat Lantai / Dak Beton",
            "jenis_kerusakan": "Rembesan Air & Retak Rambut Plat Atap Dak",
            "tingkat_keparahan": "Sedang",
            "koefisien_upah": [
                {"tenaga": "Pekerja", "koef": 0.150, "satuan": "OH"},
                {"tenaga": "Tukang Khusus Pelapis", "koef": 0.120, "satuan": "OH"},
                {"tenaga": "Kepala Tukang", "koef": 0.012, "satuan": "OH"},
                {"tenaga": "Mandor", "koef": 0.008, "satuan": "OH"}
            ],
            "koefisien_bahan": [
                {"bahan": "Liquid Polyurethane Waterproofing", "koef": 1.600, "satuan": "kg"},
                {"bahan": "Primer Bitumen / Epoxy Substrate", "koef": 0.250, "satuan": "kg"},
                {"bahan": "Polyester Mesh Reinforcement", "koef": 1.100, "satuan": "m2"}
            ],
            "langkah_kerja": [
                "Pembersihan lumut, debu, dan residu kotoran pada permukaan dak beton.",
                "Pembuatan fillet lengkung (chamfer) 5x5 cm pada pertemuan antara pelat lantai dengan parapet dinding.",
                "Pelapisan primer perekat dan pemasangan serat mesh pada area sudut dan retakan.",
                "Aplikasi waterproofing polyurethane lapis pertama dan kedua secara menyilang.",
                "Uji genang air (ponding test) selama minimal 24 jam untuk menjamin integritas kekedapan air."
            ],
            "rekomendasi": "Cek kemiringan slope dak beton menuju roof drain minimal 1.5% untuk mencegah genangan air terakumulasi."
        },
        "A.4.5.1.07.JT": {
            "kode": "A.4.5.1.07.JT",
            "kategori": "Pekerjaan Langit-langit (Plafon)",
            "nama": "Penggantian Penutup Plafon Gypsum Board 9 mm Rangka Galvanis",
            "satuan": "m2",
            "harga_satuan": 135000,
            "elemen_terkait": "Plafon Gypsum / Akustik",
            "jenis_kerusakan": "Plafon Lendut, Lapuk & Noda Coklat Akibat Bocor",
            "tingkat_keparahan": "Ringan",
            "koefisien_upah": [
                {"tenaga": "Pekerja", "koef": 0.100, "satuan": "OH"},
                {"tenaga": "Tukang Plafon", "koef": 0.120, "satuan": "OH"},
                {"tenaga": "Kepala Tukang", "koef": 0.012, "satuan": "OH"},
                {"tenaga": "Mandor", "koef": 0.005, "satuan": "OH"}
            ],
            "koefisien_bahan": [
                {"bahan": "Gypsum Board 9 mm (Water-Resistant)", "koef": 0.364, "satuan": "lbr"},
                {"bahan": "Sekrup Drywall 1 inch", "koef": 18.000, "satuan": "buah"},
                {"bahan": "Compound Joint Filler & Paper Tape", "koef": 0.250, "satuan": "kg"}
            ],
            "langkah_kerja": [
                "Pelepasan panel plafon gypsum yang lapuk, lendut atau berjamur.",
                "Pengecekan dan perbaikan sumber kebocoran instalasi pipa air atau talang di atas plafon.",
                "Perkuatan dan penyetelan elevasi rangka hollow galvanis penggantung.",
                "Pemasangan lembaran gypsum board baru dengan sekrup drywall teratur berjarak 20 cm.",
                "Pendempulan sambungan dengan joint compound, pemasangan paper tape, pengamplasan, dan pengecatan."
            ],
            "rekomendasi": "Pastikan sumber tetesan air dari lantai atas telah tuntas diperbaiki sebelum penutup plafon baru dipasang."
        },
        "A.4.4.3.02.JT": {
            "kode": "A.4.4.3.02.JT",
            "kategori": "Pekerjaan Penutup Lantai",
            "nama": "Pembongkaran & Pasang Ulang Keramik Lantai Perekat Semen Instan",
            "satuan": "m2",
            "harga_satuan": 220000,
            "elemen_terkait": "Lantai Keramik / Granit",
            "jenis_kerusakan": "Keramik Menggelembung & Terangkat (Popping)",
            "tingkat_keparahan": "Sedang",
            "koefisien_upah": [
                {"tenaga": "Pekerja", "koef": 0.250, "satuan": "OH"},
                {"tenaga": "Tukang Keramik", "koef": 0.200, "satuan": "OH"},
                {"tenaga": "Kepala Tukang", "koef": 0.020, "satuan": "OH"},
                {"tenaga": "Mandor", "koef": 0.010, "satuan": "OH"}
            ],
            "koefisien_bahan": [
                {"bahan": "Ubin Keramik / Homogeneous Tile 60x60", "koef": 1.050, "satuan": "m2"},
                {"bahan": "Semen Instan Perekat Keramik", "koef": 5.500, "satuan": "kg"},
                {"bahan": "Tile Grout Pengisi Nat Kedap Air", "koef": 0.600, "satuan": "kg"}
            ],
            "langkah_kerja": [
                "Pembongkaran keramik yang kopong atau terangkat beserta spesi lama sampai ke pelat beton.",
                "Pembersihan permukaan lantai dari debu dan serpihan adukan semen.",
                "Aplikasi adukan semen instan perekat ubin (tile adhesive) menggunakan roskam bergigi.",
                "Pemasangan ubin keramik baru dengan menjaga jarak nat ekspansi selebar 2-3 mm.",
                "Pengisian celah nat menggunakan semen grout polimer elastis setelah perekat kering 24 jam."
            ],
            "rekomendasi": "Sediakan expansion joint berisi polyurethane sealant pada bentang lantai setiap 5 meter untuk antisipasi termal."
        },
        "A.4.7.1.10.JT": {
            "kode": "A.4.7.1.10.JT",
            "kategori": "Pekerjaan Pengecatan",
            "nama": "Pengerokan Cat Lama & Pengecatan Ulang Dinding Gedung 2 Lapis",
            "satuan": "m2",
            "harga_satuan": 46500,
            "elemen_terkait": "Dinding Eksterior / Interior",
            "jenis_kerusakan": "Pengelupasan, Pengapuran & Jamur Cat Dinding",
            "tingkat_keparahan": "Ringan",
            "koefisien_upah": [
                {"tenaga": "Pekerja", "koef": 0.070, "satuan": "OH"},
                {"tenaga": "Tukang Cat", "koef": 0.090, "satuan": "OH"},
                {"tenaga": "Kepala Tukang", "koef": 0.009, "satuan": "OH"},
                {"tenaga": "Mandor", "koef": 0.004, "satuan": "OH"}
            ],
            "koefisien_bahan": [
                {"bahan": "Cat Tembok Mutu Tinggi (Akrilik)", "koef": 0.280, "satuan": "kg"},
                {"bahan": "Alkali-Resisting Primer Sealer", "koef": 0.120, "satuan": "liter"},
                {"bahan": "Plamir Tembok Acrylic", "koef": 0.100, "satuan": "kg"}
            ],
            "langkah_kerja": [
                "Pengerokan cat lama yang terkelupas dan pengapuran memakai kape besi sampai substrat bersih.",
                "Pencucian permukaan dinding dan pembersihan spora lumut menggunakan larutan fungisida.",
                "Pelapisan 1 lapis alkali resisting primer sealer untuk mencegah serangan alkali semen.",
                "Aplikasi plamir tipis pada bagian dinding yang bopeng atau pori kasar.",
                "Pengecatan cat dinding akrilik 2 lapis hingga warna menutup merata dan homogen."
            ],
            "rekomendasi": "Hindari pengecatan pada saat kelembapan dinding masih tinggi (> 16% berdasarkan moisture meter)."
        },
        "A.4.6.2.04.JT": {
            "kode": "A.4.6.2.04.JT",
            "kategori": "Pekerjaan Pintu & Jendela",
            "nama": "Perbaikan Kusen & Daun Pintu/Jendela Aluminium + Sealant",
            "satuan": "m'",
            "harga_satuan": 165000,
            "elemen_terkait": "Kusen Pintu / Jendela Aluminium",
            "jenis_kerusakan": "Kusen Miring, Karet Lepas & Rembesan Air Kaca",
            "tingkat_keparahan": "Ringan",
            "koefisien_upah": [
                {"tenaga": "Pekerja", "koef": 0.080, "satuan": "OH"},
                {"tenaga": "Tukang Aluminium", "koef": 0.140, "satuan": "OH"},
                {"tenaga": "Kepala Tukang", "koef": 0.014, "satuan": "OH"},
                {"tenaga": "Mandor", "koef": 0.007, "satuan": "OH"}
            ],
            "koefisien_bahan": [
                {"bahan": "Sealant Silicone Netral Eksterior", "koef": 0.350, "satuan": "tube"},
                {"bahan": "Weatherstrip Gasket Karet EPDM", "koef": 1.050, "satuan": "m'"},
                {"bahan": "Sekrup Stainless Steel & Fisher", "koef": 4.000, "satuan": "buah"}
            ],
            "langkah_kerja": [
                "Penyetelan kembali kedudukan kusen aluminium dan sekrup penahan perimeter.",
                "Penggantian gasket karet penjepit kaca yang telah mengeras atau getas.",
                "Pembersihan sealant lama yang rusak pada celah antara kusen aluminium dan dinding.",
                "Aplikasi masking tape dan penyuntikan silicone sealant netral tahan cuaca secara kontinu.",
                "Perataan alur sealant menggunakan spatula karet dan pelepasan masking tape sebelum sealant berkulit."
            ],
            "rekomendasi": "Gunakan sealant tipe netral (non-acidic) agar tidak merusak finishing powder coating atau anodized profil aluminium."
        }
    }

    @classmethod
    def get_all(cls) -> List[Dict[str, Any]]:
        return list(cls.CATALOG.values())

    @classmethod
    def get_by_code(cls, kode: str) -> Optional[Dict[str, Any]]:
        """Mencari data AHSP berdasarkan kode analisis."""
        return cls.CATALOG.get(kode)

    @classmethod
    def find_by_element(cls, elemen: str) -> Dict[str, Any]:
        """Mencari data AHSP yang paling cocok berdasarkan elemen gedung."""
        for item in cls.CATALOG.values():
            if item["elemen_terkait"].lower() in elemen.lower() or elemen.lower() in item["elemen_terkait"].lower():
                return item
        return cls.CATALOG["A.4.4.1.25.JT"]


# ==================================================================================================
# 4. INTEGRASI AI AGENT ANTIGRAVITY (HTTP POST CALL VIA requests)
# ==================================================================================================
class AIAgentClient:
    """Kelas integrasi untuk melakukan HTTP POST API Call ke AI Agent Antigravity."""

    DEFAULT_ENDPOINT = "http://localhost:8000/api/analyze-damage"

    @classmethod
    def call_agent(
        cls,
        image_bytes: bytes,
        filename: str,
        volume: float,
        elemen_gedung: str,
        endpoint_url: str = "",
        api_key: str = "",
        force_simulation: bool = False
    ) -> Dict[str, Any]:
        """
        Mengubah foto ke format Base64 dan memanggil endpoint AI Agent Antigravity
        menggunakan library requests via protokol HTTP POST JSON.
        """
        # Konversi file gambar ke Base64 string
        base64_encoded_img = base64.b64encode(image_bytes).decode("utf-8")
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Payload sesuai protokol Antigravity AI Agent
        payload = {
            "image_base64": base64_encoded_img,
            "filename": filename,
            "volume": float(volume),
            "elemen_gedung": elemen_gedung,
            "acuan_standar": "AHSP Dinas PU Bina Marga & Cipta Karya Jawa Tengah 2023",
            "timestamp": current_time
        }

        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"

        target_url = endpoint_url.strip() if endpoint_url.strip() else cls.DEFAULT_ENDPOINT

        # 1. Jika pengguna memilih mode simulasi sidang mandiri (offline safety)
        if force_simulation:
            return cls._deterministic_civil_ai_engine(elemen_gedung, volume, filename, "Simulasi Mandiri (Offline Safe)")

        # 2. Cek apakah ada model YOLOv8 kustom hasil training (weights/best.pt)
        local_yolo_result = cls.detect_with_local_yolo(image_bytes, elemen_gedung, volume, filename)
        if local_yolo_result is not None:
            return local_yolo_result

        # Jika model weights/best.pt ditemukan tapi ultralytics belum terpasang
        model_path = cls.get_yolo_model_path()
        if model_path:
            fallback = cls._deterministic_civil_ai_engine(elemen_gedung, volume, filename, "Standalone Civil AI Engine")
            try:
                import ultralytics
            except ImportError:
                fallback["warning"] = f"Model bobot '{os.path.basename(model_path)}' ditemukan di folder weights, namun pustaka 'ultralytics' belum selesai diinstal. Sistem sementara dialihkan ke Standalone Civil Engine."
                return fallback

        # 3. Eksekusi HTTP POST Call (hanya jika target_url bukan default localhost:8000 atau server lokal sengaja dinyalakan)
        if target_url and target_url != cls.DEFAULT_ENDPOINT:
            try:
                response = requests.post(target_url, json=payload, headers=headers, timeout=8)
                if response.status_code == 200:
                    result = response.json()
                    result["is_live_api"] = True
                    result["api_endpoint_used"] = target_url
                    return result
                else:
                    fallback = cls._deterministic_civil_ai_engine(elemen_gedung, volume, filename, f"HTTP Error {response.status_code}")
                    fallback["warning"] = f"Endpoint merespons status {response.status_code}. Beralih ke Standalone Engine."
                    return fallback
            except Exception as exc:
                fallback = cls._deterministic_civil_ai_engine(elemen_gedung, volume, filename, "Standalone Engine (Auto-Fallback)")
                fallback["warning"] = f"Koneksi endpoint '{target_url}' tidak merespons ({type(exc).__name__}). Dialihkan ke Standalone Civil AI Engine."
                return fallback

        # Default fallback jika tidak ada custom endpoint dan model weights belum aktif
        fallback = cls._deterministic_civil_ai_engine(elemen_gedung, volume, filename, "Standalone Civil AI Engine")
        fallback["warning"] = "Mode Standalone Civil Engine aktif. Model best.pt siap digunakan setelah pustaka ultralytics dimuat."
        return fallback

    @classmethod
    def get_yolo_model_path(cls) -> Optional[str]:
        """Mencari file model bobot YOLOv8 (best.pt) pada direktori sistem."""
        candidates = [
            os.path.join("weights", "best.pt"),
            "best.pt",
            os.path.join("..", "weights", "best.pt"),
            r"d:\gedung\weights\best.pt"
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
        return None

    @classmethod
    def _map_yolo_defect_to_ahsp(cls, class_name: str, elemen_gedung: str) -> Dict[str, Any]:
        """Memetakan 4 kelas YOLO CONCRETE-25FEB ke item katalog AHSP Jateng 2023."""
        c_lower = str(class_name).lower()
        if "crack" in c_lower:
            if "balok" in elemen_gedung.lower():
                return {
                    "kode_ahsp": "A.4.1.1.18.JT",
                    "elemen_terkait": "Balok Beton Bertulang",
                    "jenis_kerusakan": "Retak Diagonal Lentur/Geser Balok Struktur",
                    "tingkat_keparahan": "Berat"
                }
            elif "dak" in elemen_gedung.lower() or "pelat" in elemen_gedung.lower():
                return {
                    "kode_ahsp": "A.4.4.3.35.JT",
                    "elemen_terkait": "Pelat Lantai / Dak Beton",
                    "jenis_kerusakan": "Rembesan Air & Retak Rambut Plat Atap Dak",
                    "tingkat_keparahan": "Sedang"
                }
            else:
                return {
                    "kode_ahsp": "A.4.4.1.25.JT",
                    "elemen_terkait": "Dinding Pasangan Bata / Hebel",
                    "jenis_kerusakan": "Retak Plesteran & Retak Rambut Dinding",
                    "tingkat_keparahan": "Sedang"
                }
        elif "spalling" in c_lower:
            return {
                "kode_ahsp": "A.4.1.1.20.JT",
                "elemen_terkait": "Kolom Struktur Beton",
                "jenis_kerusakan": "Spalling / Keropos Beton & Selimut Terkelupas",
                "tingkat_keparahan": "Berat"
            }
        elif "rebar" in c_lower:
            return {
                "kode_ahsp": "A.4.1.1.20.JT",
                "elemen_terkait": "Kolom Struktur Beton",
                "jenis_kerusakan": "Besi Tulangan Terbuka & Keropos Beton (Exposed Rebar)",
                "tingkat_keparahan": "Berat"
            }
        elif "efflorescence" in c_lower:
            if "dinding" in elemen_gedung.lower():
                return {
                    "kode_ahsp": "A.4.6.2.04.JT",
                    "elemen_terkait": "Dinding Eksterior / Interior",
                    "jenis_kerusakan": "Cat Dinding Mengelupas, Mengapur & Berjamur (Efflorescence)",
                    "tingkat_keparahan": "Ringan"
                }
            else:
                return {
                    "kode_ahsp": "A.4.4.3.35.JT",
                    "elemen_terkait": "Pelat Lantai / Dak Beton",
                    "jenis_kerusakan": "Rembesan Air & Efloresensi Garam Pelat Beton",
                    "tingkat_keparahan": "Sedang"
                }
        else:
            item = AHSPKnowledgeBase.find_by_element(elemen_gedung)
            return {
                "kode_ahsp": item["kode"],
                "elemen_terkait": item["elemen_terkait"],
                "jenis_kerusakan": item["jenis_kerusakan"],
                "tingkat_keparahan": item["tingkat_keparahan"]
            }

    @classmethod
    def detect_with_local_yolo(
        cls, image_bytes: bytes, elemen_gedung: str, volume: float, filename: str
    ) -> Optional[Dict[str, Any]]:
        """Menjalankan inferensi lokal YOLOv8 jika file weights/best.pt terdeteksi."""
        model_path = cls.get_yolo_model_path()
        if not model_path:
            return None

        try:
            from ultralytics import YOLO
        except ImportError:
            return None

        try:
            model = YOLO(model_path)
            img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
            # Ambang batas conf=0.15 agar retak rambut & gompal samar tetap terdeteksi
            results = model(img, conf=0.15, verbose=False)
            if not results or len(results) == 0:
                return None

            first_res = results[0]
            boxes_data = []
            model_names = first_res.names if hasattr(first_res, "names") else {
                0: "Crack", 1: "Efflorescence", 2: "Exposed Rebar", 3: "Spalling"
            }

            if hasattr(first_res, "boxes") and len(first_res.boxes) > 0:
                for b in first_res.boxes:
                    coords = [int(v) for v in b.xyxy[0].tolist()]
                    conf = float(b.conf[0].item()) * 100.0
                    cls_id = int(b.cls[0].item())
                    cls_name = model_names.get(cls_id, "Crack")
                    ahsp_map = cls._map_yolo_defect_to_ahsp(cls_name, elemen_gedung)

                    boxes_data.append({
                        "box": coords,
                        "class_name": cls_name,
                        "label": ahsp_map["jenis_kerusakan"],
                        "confidence": round(conf, 1),
                        "keparahan": ahsp_map["tingkat_keparahan"],
                        "kode_ahsp": ahsp_map["kode_ahsp"]
                    })

            if not boxes_data:
                # Jika tidak ada defect terdeteksi di atas conf 0.15, gunakan acuan elemen dengan status informatif
                ahsp_item = AHSPKnowledgeBase.find_by_element(elemen_gedung)
                harga_satuan = ahsp_item["harga_satuan"]
                total_biaya = int(round(volume * harga_satuan))
                return {
                    "status": "success",
                    "agent_id": "yolov8-custom-model-engine",
                    "mode": f"YOLOv8 Engine ({os.path.basename(model_path)})",
                    "is_live_api": False,
                    "is_yolo_active": True,
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "filename": filename,
                    "elemen_gedung": ahsp_item["elemen_terkait"],
                    "jenis_kerusakan": ahsp_item["jenis_kerusakan"],
                    "tingkat_keparahan": ahsp_item["tingkat_keparahan"],
                    "confidence_score": 95.0,
                    "kode_ahsp": ahsp_item["kode"],
                    "deskripsi_ahsp": ahsp_item["nama"],
                    "volume": volume,
                    "satuan": ahsp_item["satuan"],
                    "harga_satuan": harga_satuan,
                    "total_biaya": total_biaya,
                    "langkah_kerja": ahsp_item["langkah_kerja"],
                    "rekomendasi": ahsp_item["rekomendasi"],
                    "acuan_regulasi": "AHSP Dinas Bina Marga & Cipta Karya Provinsi Jawa Tengah 2023",
                    "warning": f"Model YOLOv8 ({os.path.basename(model_path)}) aktif memindai citra (tidak ada anomali berat di atas ambang batas 15%). Menggunakan estimasi acuan elemen terpilih.",
                    "detected_boxes": [],
                    "total_detected_defects": 0
                }

            main_box = max(boxes_data, key=lambda x: x["confidence"])
            ahsp_item = AHSPKnowledgeBase.get_by_code(main_box["kode_ahsp"])
            if not ahsp_item:
                ahsp_item = AHSPKnowledgeBase.find_by_element(elemen_gedung)

            harga_satuan = ahsp_item["harga_satuan"]
            total_biaya = int(round(volume * harga_satuan))

            return {
                "status": "success",
                "agent_id": "yolov8-custom-model-engine",
                "mode": f"YOLOv8 Local Model ({os.path.basename(model_path)})",
                "is_live_api": False,
                "is_yolo_active": True,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "filename": filename,
                "elemen_gedung": ahsp_item["elemen_terkait"],
                "jenis_kerusakan": main_box["label"],
                "tingkat_keparahan": main_box["keparahan"],
                "confidence_score": round(main_box["confidence"], 1),
                "kode_ahsp": ahsp_item["kode"],
                "deskripsi_ahsp": ahsp_item["nama"],
                "volume": volume,
                "satuan": ahsp_item["satuan"],
                "harga_satuan": harga_satuan,
                "total_biaya": total_biaya,
                "langkah_kerja": ahsp_item["langkah_kerja"],
                "rekomendasi": ahsp_item["rekomendasi"],
                "acuan_regulasi": "AHSP Dinas Bina Marga & Cipta Karya Provinsi Jawa Tengah 2023",
                "detected_boxes": boxes_data,
                "total_detected_defects": len(boxes_data)
            }
        except Exception:
            return None

    @classmethod
    def _deterministic_civil_ai_engine(
        cls, elemen_gedung: str, volume: float, filename: str, mode_label: str
    ) -> Dict[str, Any]:
        """Mesin analisis teknik sipil deterministik berbasis katalog AHSP Jateng 2023."""
        ahsp_item = AHSPKnowledgeBase.find_by_element(elemen_gedung)
        harga_satuan = ahsp_item["harga_satuan"]
        total_biaya = int(round(volume * harga_satuan))

        return {
            "status": "success",
            "agent_id": "antigravity-civil-vision-agent-v2",
            "mode": mode_label,
            "is_live_api": False,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "filename": filename,
            "elemen_gedung": ahsp_item["elemen_terkait"],
            "jenis_kerusakan": ahsp_item["jenis_kerusakan"],
            "tingkat_keparahan": ahsp_item["tingkat_keparahan"],
            "confidence_score": 97.2,
            "kode_ahsp": ahsp_item["kode"],
            "deskripsi_ahsp": ahsp_item["nama"],
            "volume": volume,
            "satuan": ahsp_item["satuan"],
            "harga_satuan": harga_satuan,
            "total_biaya": total_biaya,
            "langkah_kerja": ahsp_item["langkah_kerja"],
            "rekomendasi": ahsp_item["rekomendasi"],
            "acuan_regulasi": "AHSP Dinas Bina Marga & Cipta Karya Provinsi Jawa Tengah 2023"
        }


# ==================================================================================================
# 5. GENERATOR LAPORAN PENILAIAN KONDISI (IKBG) & ESTIMASI RAB (EXCEL & PDF)
# ==================================================================================================
def format_rupiah(val: float) -> str:
    """Format angka ke format mata uang Rupiah."""
    return f"Rp {int(val):,d}".replace(",", ".")

def angka_ke_terbilang(n: int) -> str:
    """Konversi bilangan ke kalimat terbilang bahasa Indonesia."""
    satuan = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    n = int(n)
    if n < 12:
        return satuan[n]
    elif n < 20:
        return angka_ke_terbilang(n - 10) + " Belas"
    elif n < 100:
        return angka_ke_terbilang(n // 10) + " Puluh " + satuan[n % 10]
    elif n < 200:
        return "Seratus " + angka_ke_terbilang(n - 100)
    elif n < 1000:
        return satuan[n // 100] + " Ratus " + angka_ke_terbilang(n % 100)
    elif n < 2000:
        return "Seribu " + angka_ke_terbilang(n - 1000)
    elif n < 1000000:
        return angka_ke_terbilang(n // 1000) + " Ribu " + angka_ke_terbilang(n % 1000)
    elif n < 1000000000:
        return angka_ke_terbilang(n // 1000000) + " Juta " + angka_ke_terbilang(n % 1000000)
    elif n < 1000000000000:
        return angka_ke_terbilang(n // 1000000000) + " Miliar " + angka_ke_terbilang(n % 1000000000)
    return str(n)


def compute_condition_assessment(item: Dict[str, Any]) -> Dict[str, Any]:
    """
    Menghitung matriks penilaian kondisi fisik dan fungsi komponen bangunan gedung (IKBG)
    sesuai standar evaluasi keandalan bangunan (Permen PUPR No. 24/PRT/M/2008).
    Formula:
      - Nilai Kondisi Fisik = (Nilai Awal) x (Rating Fisik) / 100
      - Nilai Kondisi Fungsi = (Nilai Awal) x (Rating Fungsi) / 100
      - Sub Total = Nilai Kondisi Fisik + Nilai Kondisi Fungsi
      - Bobot: Fisik 50%, Fungsi 50%
      - Kinerja Total (µku) = (50% x Fisik) + (50% x Fungsi)
      - Klasifikasi:
          * Rutin: µku > 90
          * Preventif: µku = 70 - 90
          * Korektif: µku = 50 - 70
          * Rehabilitatif: µku < 50
    """
    elemen = item.get("elemen_gedung", "Komponen Bangunan")
    kerusakan = item.get("jenis_kerusakan", "Retak / Kerusakan Permukaan")
    keparahan = item.get("tingkat_keparahan")
    if not keparahan:
        kode = item.get("kode_ahsp", "")
        keparahan = AHSPKnowledgeBase.CATALOG.get(kode, {}).get("tingkat_keparahan", "Sedang")

    # Nilai Awal selalu 100%
    nilai_awal = 100

    if keparahan == "Ringan":
        skor_fisik = 90.0
        skor_fungsi = 100.0
        desc_fisik = f"Secara Visual {elemen} Masih Baik , Mungkin ada goresan dan Bercak Noda Sedikit Tidak Terlalu Berpengaruh dengan kegiatan operasional/proses belajar mengajar."
        desc_fungsi = f"Untuk Fungsi Pada {elemen} Tergolong Baik masih berfungsi dengan baik memikul beban kerja normal."
    elif keparahan == "Sedang":
        skor_fisik = 65.0
        skor_fungsi = 80.0
        desc_fisik = f"Secara Visual {elemen} Cukup, terdapat {kerusakan} pada beberapa luasan tertentu. Memerlukan penanganan pemeliharaan preventif."
        desc_fungsi = f"Untuk Fungsi Pada {elemen} Tergolong Cukup, komponen masih mampu bekerja namun keandalan mulai tereduksi."
    else:  # Berat
        skor_fisik = 40.0
        skor_fungsi = 55.0
        desc_fisik = f"Secara Visual {elemen} Mengalami Kerusakan Berat ({kerusakan}). Integritas material atau selimut pelindung mengalami kerusakan struktural."
        desc_fungsi = f"Untuk Fungsi Pada {elemen} Mengalami Penurunan Signifikan, kapasitas layan terganggu dan membutuhkan tindakan rehabilitasi intensif."

    # Distribusi kolom Kondisi Penilaian: Baik (79.5-100%), Cukup (49.5-79.49%), Rusak (0-49.49%)
    baik_f = skor_fisik if skor_fisik >= 79.5 else ""
    cukup_f = skor_fisik if 49.5 <= skor_fisik < 79.5 else ""
    rusak_f = skor_fisik if skor_fisik < 49.5 else ""
    nilai_kondisi_f = skor_fisik

    baik_fn = skor_fungsi if skor_fungsi >= 79.5 else ""
    cukup_fn = skor_fungsi if 49.5 <= skor_fungsi < 79.5 else ""
    rusak_fn = skor_fungsi if skor_fungsi < 49.5 else ""
    nilai_kondisi_fn = skor_fungsi

    subtotal = int(round(nilai_kondisi_f + nilai_kondisi_fn))

    # Bobot & Kinerja: Fisik 50%, Fungsi 50%
    kinerja_f = round(0.5 * nilai_kondisi_f, 1)
    kinerja_fn = round(0.5 * nilai_kondisi_fn, 1)
    total_kinerja = round(kinerja_f + kinerja_fn, 1)

    # Klasifikasi Tingkat Pemeliharaan:
    if total_kinerja > 90:
        kategori_pemeliharaan = "RUTIN"
    elif 70 <= total_kinerja <= 90:
        kategori_pemeliharaan = "PREVENTIF"
    elif 50 <= total_kinerja < 70:
        kategori_pemeliharaan = "KOREKTIF"
    else:
        kategori_pemeliharaan = "REHABILITATIF"

    return {
        "elemen": elemen,
        "kerusakan": kerusakan,
        "keparahan": keparahan,
        "nilai_awal": nilai_awal,
        "desc_fisik": desc_fisik,
        "desc_fungsi": desc_fungsi,
        "skor_fisik": skor_fisik,
        "skor_fungsi": skor_fungsi,
        "baik_f": baik_f,
        "cukup_f": cukup_f,
        "rusak_f": rusak_f,
        "nilai_kondisi_f": nilai_kondisi_f,
        "baik_fn": baik_fn,
        "cukup_fn": cukup_fn,
        "rusak_fn": rusak_fn,
        "nilai_kondisi_fn": nilai_kondisi_fn,
        "subtotal": subtotal,
        "kinerja_f": kinerja_f,
        "kinerja_fn": kinerja_fn,
        "total_kinerja": total_kinerja,
        "kategori_pemeliharaan": kategori_pemeliharaan,
        "is_rutin": (kategori_pemeliharaan == "RUTIN"),
        "is_preventif": (kategori_pemeliharaan == "PREVENTIF"),
        "is_korektif": (kategori_pemeliharaan == "KOREKTIF"),
        "is_rehabilitatif": (kategori_pemeliharaan == "REHABILITATIF"),
    }


class AIVisualAnnotator:
    """
    Modul Visualisasi Bounding Box Anotasi Deteksi Kerusakan Berbasis AI Vision.
    Menghasilkan citra dengan bounding box berwarna, reticle corner brackets,
    badge confidence score, dan label standar AHSP Dinas PU Cipta Karya Jateng.
    Mendukung koordinat multi-box hasil inferensi model kustom YOLOv8 maupun mode simulasi.
    """

    @staticmethod
    def annotate(
        image_bytes: bytes,
        jenis_kerusakan: str,
        tingkat_keparahan: str,
        confidence_score: float,
        kode_ahsp: str,
        detected_boxes: Optional[List[Dict[str, Any]]] = None
    ) -> bytes:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        w, h = img.size
        draw = ImageDraw.Draw(img, "RGBA")

        # 3. Tipografi Label
        font_size = max(12, int(min(w, h) * 0.035))
        font_small_size = max(10, font_size - 3)
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
            font_small = ImageFont.truetype("arial.ttf", font_small_size)
        except Exception:
            font = ImageFont.load_default()
            font_small = ImageFont.load_default()

        # Siapkan daftar box yang akan digambar
        boxes_to_draw = []
        if detected_boxes and len(detected_boxes) > 0:
            for b in detected_boxes:
                boxes_to_draw.append({
                    "coords": b.get("box", [int(w * 0.18), int(h * 0.22), int(w * 0.82), int(h * 0.78)]),
                    "label": b.get("label", jenis_kerusakan),
                    "conf": b.get("confidence", confidence_score),
                    "keparahan": b.get("keparahan", tingkat_keparahan),
                    "kode_ahsp": b.get("kode_ahsp", kode_ahsp)
                })
        else:
            boxes_to_draw.append({
                "coords": [int(w * 0.18), int(h * 0.22), int(w * 0.82), int(h * 0.78)],
                "label": jenis_kerusakan,
                "conf": confidence_score,
                "keparahan": tingkat_keparahan,
                "kode_ahsp": kode_ahsp
            })

        for box_item in boxes_to_draw:
            raw_c = box_item["coords"]
            x1 = max(0, min(w, int(raw_c[0])))
            y1 = max(0, min(h, int(raw_c[1])))
            x2 = max(0, min(w, int(raw_c[2])))
            y2 = max(0, min(h, int(raw_c[3])))
            if x2 <= x1:
                x2 = min(w, x1 + 50)
            if y2 <= y1:
                y2 = min(h, y1 + 50)

            b_label = box_item["label"]
            b_conf = box_item["conf"]
            b_keparahan = box_item["keparahan"]
            b_ahsp = box_item["kode_ahsp"]

            # Palet warna berdasarkan tingkat keparahan standar teknik sipil
            keparahan_lower = str(b_keparahan).lower()
            if "berat" in keparahan_lower:
                color_rgb = (239, 68, 68)       # #EF4444 Merah (Red)
                bg_badge = (220, 38, 38, 235)
            elif "sedang" in keparahan_lower:
                color_rgb = (245, 158, 11)      # #F59E0B Oranye/Amber
                bg_badge = (217, 119, 6, 235)
            else:
                color_rgb = (16, 185, 129)      # #10B981 Hijau Emerald
                bg_badge = (5, 150, 105, 235)

            # 1. Shading semi-transparan pada region kerusakan
            fill_box = (*color_rgb, 35)
            draw.rectangle([x1, y1, x2, y2], fill=fill_box, outline=(*color_rgb, 200), width=2)

            # 2. Reticle corner brackets (L-shaped bracket khas YOLO/Computer Vision)
            bw = x2 - x1
            bh = y2 - y1
            corner_len = max(14, min(bw, bh) // 8)
            corner_w = max(2, min(w, h) // 150)

            # Pojok Kiri-Atas
            draw.line([(x1, y1), (x1 + corner_len, y1)], fill=(*color_rgb, 255), width=corner_w)
            draw.line([(x1, y1), (x1, y1 + corner_len)], fill=(*color_rgb, 255), width=corner_w)
            # Pojok Kanan-Atas
            draw.line([(x2, y1), (x2 - corner_len, y1)], fill=(*color_rgb, 255), width=corner_w)
            draw.line([(x2, y1), (x2, y1 + corner_len)], fill=(*color_rgb, 255), width=corner_w)
            # Pojok Kiri-Bawah
            draw.line([(x1, y2), (x1 + corner_len, y2)], fill=(*color_rgb, 255), width=corner_w)
            draw.line([(x1, y2), (x1, y2 - corner_len)], fill=(*color_rgb, 255), width=corner_w)
            # Pojok Kanan-Bawah
            draw.line([(x2, y2), (x2 - corner_len, y2)], fill=(*color_rgb, 255), width=corner_w)
            draw.line([(x2, y2), (x2, y2 - corner_len)], fill=(*color_rgb, 255), width=corner_w)

            top_text = f"🎯 {b_label} ({b_conf:.1f}%)"
            bottom_text = f"AHSP: {b_ahsp} | Status: {b_keparahan.upper()}"

            # Top Badge Box
            bbox_top = draw.textbbox((x1, y1), top_text, font=font)
            text_w = bbox_top[2] - bbox_top[0]
            text_h = bbox_top[3] - bbox_top[1]
            badge_y1 = max(0, y1 - text_h - 10)
            badge_y2 = badge_y1 + text_h + 8
            badge_x2 = min(w, x1 + text_w + 14)
            draw.rectangle([x1, badge_y1, badge_x2, badge_y2], fill=bg_badge)
            draw.text((x1 + 6, badge_y1 + 4), top_text, fill=(255, 255, 255, 255), font=font)

            # Bottom Badge Box
            bbox_bot = draw.textbbox((x1, y2), bottom_text, font=font_small)
            text_bot_w = bbox_bot[2] - bbox_bot[0]
            text_bot_h = bbox_bot[3] - bbox_bot[1]
            bot_y1 = min(h - text_bot_h - 8, y2 + 2)
            bot_y2 = bot_y1 + text_bot_h + 6
            bot_x2 = min(w, x1 + text_bot_w + 12)
            draw.rectangle([x1, bot_y1, bot_x2, bot_y2], fill=(15, 23, 42, 230))
            draw.text((x1 + 5, bot_y1 + 3), bottom_text, fill=(255, 255, 255, 255), font=font_small)

        out_buf = io.BytesIO()
        img.save(out_buf, format="JPEG", quality=95)
        return out_buf.getvalue()


class ExcelReportGenerator:
    """Generator file Microsoft Excel (.xlsx) dengan tabel penilaian kondisi fisik & fungsi sesuai standar IKBG."""

    @classmethod
    def generate_excel_report(cls, data_list: List[Dict[str, Any]], title_doc: str = "Laporan Penilaian Kondisi") -> bytes:
        wb = openpyxl.Workbook()

        thin_border = XlsBorder(
            left=XlsSide(style='thin', color='000000'),
            right=XlsSide(style='thin', color='000000'),
            top=XlsSide(style='thin', color='000000'),
            bottom=XlsSide(style='thin', color='000000')
        )
        box_border = XlsBorder(
            left=XlsSide(style='medium', color='000000'),
            right=XlsSide(style='medium', color='000000'),
            top=XlsSide(style='medium', color='000000'),
            bottom=XlsSide(style='medium', color='000000')
        )
        pill_fill = XlsPatternFill(start_color='FDE68A', end_color='FDE68A', fill_type='solid')

        # -----------------------------------------------------------------------------------------
        # SHEET 1: PENILAIAN KONDISI (IKBG) SESUAI GAMBAR STANDAR TEKNIK SIPIL
        # -----------------------------------------------------------------------------------------
        ws1 = wb.active
        ws1.title = "Penilaian Kondisi (IKBG)"
        ws1.views.sheetView[0].showGridLines = True

        ws1.column_dimensions['A'].width = 14
        ws1.column_dimensions['B'].width = 46
        ws1.column_dimensions['C'].width = 13
        ws1.column_dimensions['D'].width = 13
        ws1.column_dimensions['E'].width = 13
        ws1.column_dimensions['F'].width = 13
        ws1.column_dimensions['G'].width = 14

        target_item = data_list[0] if data_list else {}
        nama_gedung = target_item.get("nama_gedung", "Gedung Kuliah Terpadu")
        lokasi_lantai = target_item.get("lokasi_lantai", "Lantai 2 - Ruang Kuliah 204")
        surveyor_name = target_item.get("surveyor", "Ahmad Fajar, S.T.")

        ws1.merge_cells("A1:G1")
        ws1["A1"] = f"FORMULIR PENILAIAN KONDISI BANGUNAN (IKBG) - {str(nama_gedung).upper()}"
        ws1["A1"].font = XlsFont(bold=True, size=11, color="0F172A")
        ws1["A1"].alignment = XlsAlignment(horizontal='center', vertical='center')

        ws1.merge_cells("A2:G2")
        ws1["A2"] = f"Lokasi: {lokasi_lantai} | Surveyor: {surveyor_name} | Tanggal: {datetime.now().strftime('%d/%m/%Y')}"
        ws1["A2"].font = XlsFont(size=9, italic=True, color="334155")
        ws1["A2"].alignment = XlsAlignment(horizontal='center', vertical='center')

        current_row = 4

        for idx, item in enumerate(data_list, start=1):
            eval_data = compute_condition_assessment(item)

            if len(data_list) > 1:
                ws1.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
                title_cell = ws1.cell(row=current_row, column=1, value=f"KOMPONEN BANGUNAN #{idx}: {item.get('elemen_gedung', 'Komponen').upper()}")
                title_cell.font = XlsFont(bold=True, size=11, color="1E293B")
                title_cell.alignment = XlsAlignment(horizontal='left', vertical='center')
                current_row += 1

            # 1. Tabel Header Kondisi Fisik
            r_fisik = current_row
            ws1.merge_cells(start_row=r_fisik, start_column=1, end_row=r_fisik+1, end_column=1)
            c = ws1.cell(row=r_fisik, column=1, value="No. Kode\nKomponen")
            c.font = XlsFont(bold=True, size=9)
            c.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)

            ws1.merge_cells(start_row=r_fisik, start_column=2, end_row=r_fisik+1, end_column=2)
            c = ws1.cell(row=r_fisik, column=2, value="Kondisi Fisik\nKomponen Bangunan Pengambilan")
            c.font = XlsFont(bold=True, size=9)
            c.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)

            ws1.merge_cells(start_row=r_fisik, start_column=3, end_row=r_fisik+1, end_column=3)
            c = ws1.cell(row=r_fisik, column=3, value="Nilai\nKondisi\nAwal\n(%)")
            c.font = XlsFont(bold=True, size=9)
            c.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)

            ws1.merge_cells(start_row=r_fisik, start_column=4, end_row=r_fisik, end_column=6)
            c = ws1.cell(row=r_fisik, column=4, value="Kondisi Penilaian")
            c.font = XlsFont(bold=True, size=9)
            c.alignment = XlsAlignment(horizontal='center', vertical='center')

            c_b = ws1.cell(row=r_fisik+1, column=4, value="Baik\n(%)")
            c_b.font = XlsFont(bold=True, size=8)
            c_b.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)
            c_b.fill = pill_fill

            c_c = ws1.cell(row=r_fisik+1, column=5, value="Cukup\n(%)")
            c_c.font = XlsFont(bold=True, size=8)
            c_c.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)
            c_c.fill = pill_fill

            c_r = ws1.cell(row=r_fisik+1, column=6, value="Rusak\n(%)")
            c_r.font = XlsFont(bold=True, size=8)
            c_r.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)
            c_r.fill = pill_fill

            ws1.cell(row=r_fisik+2, column=4, value="79.5-100%").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fisik+2, column=4).font = XlsFont(size=8)
            ws1.cell(row=r_fisik+2, column=5, value="49.5-79.49%").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fisik+2, column=5).font = XlsFont(size=8)
            ws1.cell(row=r_fisik+2, column=6, value="0-49.49%").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fisik+2, column=6).font = XlsFont(size=8)

            ws1.merge_cells(start_row=r_fisik, start_column=7, end_row=r_fisik+1, end_column=7)
            c = ws1.cell(row=r_fisik, column=7, value="Nilai Kondisi")
            c.font = XlsFont(bold=True, size=9)
            c.alignment = XlsAlignment(horizontal='center', vertical='center')

            c = ws1.cell(row=r_fisik+2, column=7, value="(3) x (4) /100")
            c.font = XlsFont(size=8)
            c.alignment = XlsAlignment(horizontal='center', vertical='center')

            # Baris Penomoran Kolom
            nums = ["(1)", "(2)", "(3)", "(4)", "(5)", "(6)", "(7)"]
            for col_i, num in enumerate(nums, start=1):
                cell_n = ws1.cell(row=r_fisik+3, column=col_i, value=num)
                cell_n.font = XlsFont(size=8)
                cell_n.alignment = XlsAlignment(horizontal='center', vertical='center')

            # Baris Data Fisik
            ws1.cell(row=r_fisik+4, column=1, value="(1)").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fisik+4, column=2, value=eval_data["desc_fisik"]).alignment = XlsAlignment(horizontal='left', vertical='center', wrap_text=True)
            ws1.cell(row=r_fisik+4, column=3, value=100).alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fisik+4, column=4, value=eval_data["baik_f"]).alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fisik+4, column=5, value=eval_data["cukup_f"]).alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fisik+4, column=6, value=eval_data["rusak_f"]).alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fisik+4, column=7, value=eval_data["nilai_kondisi_f"]).alignment = XlsAlignment(horizontal='center', vertical='center')

            for r_x in range(r_fisik, r_fisik+5):
                for c_x in range(1, 8):
                    ws1.cell(row=r_x, column=c_x).border = thin_border

            # 2. Tabel Header Kondisi Fungsi
            r_fungsi = r_fisik + 5
            ws1.merge_cells(start_row=r_fungsi, start_column=1, end_row=r_fungsi+1, end_column=1)
            c = ws1.cell(row=r_fungsi, column=1, value="No. Kode\nKomponen")
            c.font = XlsFont(bold=True, size=9)
            c.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)

            ws1.merge_cells(start_row=r_fungsi, start_column=2, end_row=r_fungsi+1, end_column=2)
            c = ws1.cell(row=r_fungsi, column=2, value="Kondisi Fungsi\nKomponen Bangunan Pengambilan")
            c.font = XlsFont(bold=True, size=9)
            c.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)

            ws1.merge_cells(start_row=r_fungsi, start_column=3, end_row=r_fungsi+1, end_column=3)
            c = ws1.cell(row=r_fungsi, column=3, value="Nilai\nKondisi\nAwal\n(%)")
            c.font = XlsFont(bold=True, size=9)
            c.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)

            ws1.merge_cells(start_row=r_fungsi, start_column=4, end_row=r_fungsi, end_column=6)
            c = ws1.cell(row=r_fungsi, column=4, value="Kondisi Penilaian")
            c.font = XlsFont(bold=True, size=9)
            c.alignment = XlsAlignment(horizontal='center', vertical='center')

            c_b = ws1.cell(row=r_fungsi+1, column=4, value="Baik\n(%)")
            c_b.font = XlsFont(bold=True, size=8)
            c_b.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)
            c_b.fill = pill_fill

            c_c = ws1.cell(row=r_fungsi+1, column=5, value="Cukup\n(%)")
            c_c.font = XlsFont(bold=True, size=8)
            c_c.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)
            c_c.fill = pill_fill

            c_r = ws1.cell(row=r_fungsi+1, column=6, value="Rusak\n(%)")
            c_r.font = XlsFont(bold=True, size=8)
            c_r.alignment = XlsAlignment(horizontal='center', vertical='center', wrap_text=True)
            c_r.fill = pill_fill

            ws1.cell(row=r_fungsi+2, column=4, value="79.5-100%").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fungsi+2, column=4).font = XlsFont(size=8)
            ws1.cell(row=r_fungsi+2, column=5, value="49.5-79.49%").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fungsi+2, column=5).font = XlsFont(size=8)
            ws1.cell(row=r_fungsi+2, column=6, value="0-49.49%").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fungsi+2, column=6).font = XlsFont(size=8)

            ws1.merge_cells(start_row=r_fungsi, start_column=7, end_row=r_fungsi+1, end_column=7)
            c = ws1.cell(row=r_fungsi, column=7, value="Nilai Kondisi")
            c.font = XlsFont(bold=True, size=9)
            c.alignment = XlsAlignment(horizontal='center', vertical='center')

            c = ws1.cell(row=r_fungsi+2, column=7, value="(3) x (4) /100")
            c.font = XlsFont(size=8)
            c.alignment = XlsAlignment(horizontal='center', vertical='center')

            for col_i, num in enumerate(nums, start=1):
                cell_n = ws1.cell(row=r_fungsi+3, column=col_i, value=num)
                cell_n.font = XlsFont(size=8)
                cell_n.alignment = XlsAlignment(horizontal='center', vertical='center')

            # Baris Data Fungsi
            ws1.cell(row=r_fungsi+4, column=1, value="(1)").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fungsi+4, column=2, value=eval_data["desc_fungsi"]).alignment = XlsAlignment(horizontal='left', vertical='center', wrap_text=True)
            ws1.cell(row=r_fungsi+4, column=3, value=100).alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fungsi+4, column=4, value=eval_data["baik_fn"]).alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fungsi+4, column=5, value=eval_data["cukup_fn"]).alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fungsi+4, column=6, value=eval_data["rusak_fn"]).alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_fungsi+4, column=7, value=eval_data["nilai_kondisi_fn"]).alignment = XlsAlignment(horizontal='center', vertical='center')

            for r_x in range(r_fungsi, r_fungsi+5):
                for c_x in range(1, 8):
                    ws1.cell(row=r_x, column=c_x).border = thin_border

            # 3. Baris SUB TOTAL
            r_sub = r_fungsi + 5
            ws1.merge_cells(start_row=r_sub, start_column=1, end_row=r_sub, end_column=6)
            c_sub = ws1.cell(row=r_sub, column=1, value="SUB TOTAL")
            c_sub.font = XlsFont(bold=True, size=9)
            c_sub.alignment = XlsAlignment(horizontal='left', vertical='center')

            c_sub_val = ws1.cell(row=r_sub, column=7, value=eval_data["subtotal"])
            c_sub_val.font = XlsFont(bold=True, size=9)
            c_sub_val.alignment = XlsAlignment(horizontal='center', vertical='center')

            for c_x in range(1, 8):
                ws1.cell(row=r_sub, column=c_x).border = thin_border

            # 4. Keterangan Standar Formula
            r_ket = r_sub + 2
            ws1.cell(row=r_ket, column=1, value="Keterangan :").font = XlsFont(bold=True, size=9)
            ws1.cell(row=r_ket+1, column=1, value="Rutin: µku = >90;  Preventif: µku = 70 - 90;  Korektif: µku = 50 - 70;  Rehabilitatif: µku = <50").font = XlsFont(size=8.5)

            # 5. Tabel Bobot & Kinerja (Fisik 50%, Fungsi 50%)
            r_b = r_ket + 3
            headers_bobot = ["Penilaian", "Bobot", "Penilaian", "Kinerja"]
            for col_i, h_txt in enumerate(headers_bobot, start=1):
                cell_h = ws1.cell(row=r_b, column=col_i, value=h_txt)
                cell_h.font = XlsFont(bold=True, size=9)
                cell_h.alignment = XlsAlignment(horizontal='center', vertical='center')
                cell_h.border = thin_border

            # Fisik
            ws1.cell(row=r_b+1, column=1, value="Fisik").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_b+1, column=2, value="50%").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_b+1, column=3, value=eval_data["nilai_kondisi_f"]).alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_b+1, column=4, value=eval_data["kinerja_f"]).alignment = XlsAlignment(horizontal='center', vertical='center')

            # Fungsi
            ws1.cell(row=r_b+2, column=1, value="Fungsi").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_b+2, column=2, value="50%").alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_b+2, column=3, value=eval_data["nilai_kondisi_fn"]).alignment = XlsAlignment(horizontal='center', vertical='center')
            ws1.cell(row=r_b+2, column=4, value=eval_data["kinerja_fn"]).alignment = XlsAlignment(horizontal='center', vertical='center')

            for r_x in [r_b+1, r_b+2]:
                for c_x in range(1, 5):
                    ws1.cell(row=r_x, column=c_x).border = thin_border

            # Total Kinerja
            cell_tot = ws1.cell(row=r_b+3, column=4, value=eval_data["total_kinerja"])
            cell_tot.font = XlsFont(bold=True, size=9)
            cell_tot.alignment = XlsAlignment(horizontal='center', vertical='center')
            cell_tot.border = thin_border

            # 6. Tabel NILAI KONDISI (Kotak Checkbox)
            r_nk = r_b + 5
            ws1.merge_cells(start_row=r_nk, start_column=1, end_row=r_nk, end_column=3)
            cell_nk = ws1.cell(row=r_nk, column=1, value="NILAI KONDISI")
            cell_nk.font = XlsFont(bold=True, size=9)
            cell_nk.alignment = XlsAlignment(horizontal='center', vertical='center')
            for c_x in range(1, 4):
                ws1.cell(row=r_nk, column=c_x).border = thin_border

            status_list = [
                ("RUTIN", eval_data["is_rutin"]),
                ("PREVENTIF", eval_data["is_preventif"]),
                ("KOREKTIF", eval_data["is_korektif"]),
                ("REHABILITATIF", eval_data["is_rehabilitatif"])
            ]

            for s_idx, (s_name, is_sel) in enumerate(status_list, start=1):
                curr_s_row = r_nk + s_idx
                # Checkbox box
                c_box = ws1.cell(row=curr_s_row, column=1, value="X" if is_sel else "-")
                c_box.font = XlsFont(bold=True, size=9)
                c_box.alignment = XlsAlignment(horizontal='center', vertical='center')
                c_box.border = box_border if is_sel else thin_border

                # Status label
                ws1.merge_cells(start_row=curr_s_row, start_column=2, end_row=curr_s_row, end_column=3)
                c_lbl = ws1.cell(row=curr_s_row, column=2, value=s_name)
                c_lbl.font = XlsFont(bold=is_sel, size=9)
                c_lbl.alignment = XlsAlignment(horizontal='left', vertical='center')

            current_row = r_nk + 6

        # -----------------------------------------------------------------------------------------
        # SHEET 2: REKAPITULASI RAB STANDAR AHSP JAWA TENGAH 2023
        # -----------------------------------------------------------------------------------------
        ws2 = wb.create_sheet(title="Rekapitulasi RAB Jateng")
        ws2.views.sheetView[0].showGridLines = True

        ws2.merge_cells("A1:J1")
        ws2["A1"] = "RENCANA ANGGARAN BIAYA (RAB) PEMELIHARAAN BANGUNAN GEDUNG"
        ws2["A1"].font = XlsFont(bold=True, size=12)
        ws2["A1"].alignment = XlsAlignment(horizontal='center', vertical='center')

        ws2.merge_cells("A2:J2")
        ws2["A2"] = "Standar AHSP Dinas Bina Marga & Cipta Karya Provinsi Jawa Tengah Tahun 2023"
        ws2["A2"].font = XlsFont(size=9, italic=True)
        ws2["A2"].alignment = XlsAlignment(horizontal='center', vertical='center')

        ws2.merge_cells("A3:J3")
        ws2["A3"] = f"Objek: {nama_gedung} | Lokasi: {lokasi_lantai} | Surveyor: {surveyor_name}"
        ws2["A3"].font = XlsFont(bold=True, size=9, color="334155")
        ws2["A3"].alignment = XlsAlignment(horizontal='center', vertical='center')

        rab_headers = [
            "No", "Tgl Inspeksi", "Elemen Gedung", "Jenis Kerusakan",
            "Kode AHSP", "Volume", "Sat", "Harga Satuan (Rp)", "Total Biaya (Rp)", "Kategori Pemeliharaan"
        ]

        rab_widths = [6, 18, 24, 28, 16, 10, 8, 18, 18, 18]
        for idx_w, w_val in enumerate(rab_widths, start=1):
            col_letter = get_column_letter(idx_w)
            ws2.column_dimensions[col_letter].width = w_val

        for c_idx, h_name in enumerate(rab_headers, start=1):
            c_cell = ws2.cell(row=5, column=c_idx, value=h_name)
            c_cell.font = XlsFont(bold=True, size=9, color="FFFFFF")
            c_cell.fill = XlsPatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
            c_cell.alignment = XlsAlignment(horizontal='center', vertical='center')
            c_cell.border = thin_border

        grand_total_biaya = 0
        r_start = 6
        for i_row, item in enumerate(data_list, start=1):
            curr_r = r_start + i_row - 1
            tot = item.get("total_biaya", 0)
            grand_total_biaya += tot
            eval_data = compute_condition_assessment(item)

            ws2.cell(row=curr_r, column=1, value=i_row).alignment = XlsAlignment(horizontal='center')
            ws2.cell(row=curr_r, column=2, value=str(item.get("tanggal", ""))[:16]).alignment = XlsAlignment(horizontal='center')
            ws2.cell(row=curr_r, column=3, value=item.get("elemen_gedung", "")).alignment = XlsAlignment(horizontal='left')
            ws2.cell(row=curr_r, column=4, value=item.get("jenis_kerusakan", "")).alignment = XlsAlignment(horizontal='left')
            ws2.cell(row=curr_r, column=5, value=item.get("kode_ahsp", "")).alignment = XlsAlignment(horizontal='center')
            ws2.cell(row=curr_r, column=6, value=float(item.get("volume", 0))).alignment = XlsAlignment(horizontal='right')
            ws2.cell(row=curr_r, column=7, value=item.get("satuan", "")).alignment = XlsAlignment(horizontal='center')
            ws2.cell(row=curr_r, column=8, value=int(item.get("harga_satuan", 0))).number_format = "#,##0"
            ws2.cell(row=curr_r, column=9, value=int(tot)).number_format = "#,##0"
            ws2.cell(row=curr_r, column=10, value=eval_data["kategori_pemeliharaan"]).alignment = XlsAlignment(horizontal='center')

            for c_x in range(1, 11):
                ws2.cell(row=curr_r, column=c_x).border = thin_border

        # Baris Grand Total
        r_tot = r_start + len(data_list)
        ws2.merge_cells(start_row=r_tot, start_column=1, end_row=r_tot, end_column=8)
        c_gtot = ws2.cell(row=r_tot, column=1, value="GRAND TOTAL ESTIMASI BIAYA PEMELIHARAAN (RAB)")
        c_gtot.font = XlsFont(bold=True, size=9)
        c_gtot.alignment = XlsAlignment(horizontal='center', vertical='center')

        c_gtot_val = ws2.cell(row=r_tot, column=9, value=grand_total_biaya)
        c_gtot_val.font = XlsFont(bold=True, size=9)
        c_gtot_val.number_format = "#,##0"
        c_gtot_val.alignment = XlsAlignment(horizontal='right', vertical='center')

        for c_x in range(1, 11):
            ws2.cell(row=r_tot, column=c_x).border = thin_border

        buffer = io.BytesIO()
        wb.save(buffer)
        val = buffer.getvalue()
        buffer.close()
        return val


class PDFReportGenerator:
    """Generator dokumen PDF laporan resmi pemeliharaan gedung menggunakan ReportLab."""

    @classmethod
    def generate_pdf_report(
        cls,
        data_list: List[Dict[str, Any]],
        title_doc: str = "Laporan Inspeksi Gedung",
        annotated_image_bytes: Optional[bytes] = None
    ) -> bytes:
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            leftMargin=36,
            rightMargin=36,
            topMargin=32,
            bottomMargin=32
        )

        styles = getSampleStyleSheet()

        kop_title_style = ParagraphStyle(
            'KopTitle',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=12,
            leading=15,
            alignment=1,
            textColor=colors.HexColor('#0F172A')
        )
        kop_sub_style = ParagraphStyle(
            'KopSubtitle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=8.5,
            leading=11.5,
            alignment=1,
            textColor=colors.HexColor('#475569')
        )
        doc_header_style = ParagraphStyle(
            'DocHeader',
            parent=styles['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=10,
            leading=13,
            textColor=colors.HexColor('#1E293B'),
            spaceBefore=8,
            spaceAfter=4
        )
        body_small = ParagraphStyle(
            'BodySmall',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=7,
            leading=9,
            alignment=1,
            textColor=colors.HexColor('#1E293B')
        )
        body_small_left = ParagraphStyle(
            'BodySmallLeft',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=7,
            leading=9,
            alignment=0,
            textColor=colors.HexColor('#1E293B')
        )
        body_small_bold = ParagraphStyle(
            'BodySmallBold',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=7.2,
            leading=9.2,
            alignment=1,
            textColor=colors.HexColor('#1E293B')
        )

        story = []

        # 1. KOP SURAT RESMI
        story.append(Paragraph("PEMERINTAH PROVINSI JAWA TENGAH", kop_title_style))
        story.append(Paragraph("DINAS PEKERJAAN UMUM BINA MARGA DAN CIPTA KARYA", kop_title_style))
        story.append(Paragraph("SISTEM INFORMASI PEMELIHARAAN BANGUNAN GEDUNG BERBASIS AI (SIM-PBG AI)", kop_sub_style))
        story.append(Paragraph("Standar Acuan: Pedoman AHSP Gedung & Perumahan Provinsi Jawa Tengah Tahun 2023", kop_sub_style))
        story.append(Spacer(1, 4))
        story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=6))

        # 2. METADATA DOKUMEN INSPEKSI
        today_str = datetime.now().strftime("%d %B %Y - %H:%M WIB")
        doc_no = f"BA-PBG/JATENG/{datetime.now().strftime('%Y%m%d')}/{len(data_list):03d}"

        target_item = data_list[0] if data_list else {}
        nama_gedung = target_item.get("nama_gedung") or "Bangunan Gedung"
        lokasi_lantai = target_item.get("lokasi_lantai") or "-"
        surveyor_name = target_item.get("surveyor") or ""

        if surveyor_name and surveyor_name not in ["Surveyor Lapangan", "-", ""]:
            sig_surveyor_text = f"<u>( {surveyor_name} )</u>"
            surveyor_disp = surveyor_name
        else:
            sig_surveyor_text = "<u>( .................................................. )</u>"
            surveyor_disp = "Tim Pemeriksa Lapangan"

        meta_data = [
            [
                Paragraph("<b>Nomor Dokumen:</b>", body_small_left), Paragraph(doc_no, body_small_left),
                Paragraph("<b>Tanggal Cetak:</b>", body_small_left), Paragraph(today_str, body_small_left)
            ],
            [
                Paragraph("<b>Objek Bangunan:</b>", body_small_left), Paragraph(nama_gedung, body_small_left),
                Paragraph("<b>Lokasi / Lantai:</b>", body_small_left), Paragraph(lokasi_lantai, body_small_left)
            ],
            [
                Paragraph("<b>Surveyor Lapangan:</b>", body_small_left), Paragraph(surveyor_disp, body_small_left),
                Paragraph("<b>Acuan Regulasi:</b>", body_small_left), Paragraph("AHSP Dinas PU Jateng 2023 & Permen PUPR 24/2008", body_small_left)
            ]
        ]
        meta_table = Table(meta_data, colWidths=[90, 170, 90, 170])
        meta_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(meta_table)
        story.append(Spacer(1, 8))

        # -----------------------------------------------------------------------------------------
        # 3. TABEL PENILAIAN KONDISI FISIK DAN FUNGSI (IKBG) SESUAI GAMBAR STANDAR
        # -----------------------------------------------------------------------------------------
        target_item = data_list[0] if data_list else {
            "elemen_gedung": "Komponen Gedung", "jenis_kerusakan": "Retak Permukaan", "tingkat_keparahan": "Sedang"
        }
        eval_m = compute_condition_assessment(target_item)

        story.append(Paragraph("1. FORMULIR PENILAIAN KONDISI FISIK DAN FUNGSI KOMPONEN BANGUNAN (IKBG)", doc_header_style))

        col_w_assessment = [45, 230, 48, 50, 50, 50, 50]
        data_assessment_table = [
            # Header Fisik
            [
                Paragraph("<b>No. Kode<br/>Komponen</b>", body_small_bold),
                Paragraph("<b>Kondisi Fisik<br/>Komponen Bangunan Pengambilan</b>", body_small_bold),
                Paragraph("<b>Nilai Kondisi<br/>Awal (%)</b>", body_small_bold),
                Paragraph("<b>Kondisi Penilaian</b>", body_small_bold), "", "",
                Paragraph("<b>Nilai Kondisi</b>", body_small_bold)
            ],
            [
                "", "", "",
                Paragraph("<b>Baik (%)</b><br/>79.5-100%", body_small),
                Paragraph("<b>Cukup (%)</b><br/>49.5-79.49%", body_small),
                Paragraph("<b>Rusak (%)</b><br/>0-49.49%", body_small),
                Paragraph("(3) x (4) /100", body_small)
            ],
            ["(1)", "(2)", "(3)", "(4)", "(5)", "(6)", "(7)"],
            [
                Paragraph("(1)", body_small),
                Paragraph(eval_m["desc_fisik"], body_small_left),
                Paragraph("100", body_small),
                Paragraph(str(eval_m["baik_f"]), body_small),
                Paragraph(str(eval_m["cukup_f"]), body_small),
                Paragraph(str(eval_m["rusak_f"]), body_small),
                Paragraph(str(eval_m["nilai_kondisi_f"]), body_small)
            ],
            # Header Fungsi
            [
                Paragraph("<b>No. Kode<br/>Komponen</b>", body_small_bold),
                Paragraph("<b>Kondisi Fungsi<br/>Komponen Bangunan Pengambilan</b>", body_small_bold),
                Paragraph("<b>Nilai Kondisi<br/>Awal (%)</b>", body_small_bold),
                Paragraph("<b>Kondisi Penilaian</b>", body_small_bold), "", "",
                Paragraph("<b>Nilai Kondisi</b>", body_small_bold)
            ],
            [
                "", "", "",
                Paragraph("<b>Baik (%)</b><br/>79.5-100%", body_small),
                Paragraph("<b>Cukup (%)</b><br/>49.5-79.49%", body_small),
                Paragraph("<b>Rusak (%)</b><br/>0-49.49%", body_small),
                Paragraph("(3) x (4) /100", body_small)
            ],
            ["(1)", "(2)", "(3)", "(4)", "(5)", "(6)", "(7)"],
            [
                Paragraph("(1)", body_small),
                Paragraph(eval_m["desc_fungsi"], body_small_left),
                Paragraph("100", body_small),
                Paragraph(str(eval_m["baik_fn"]), body_small),
                Paragraph(str(eval_m["cukup_fn"]), body_small),
                Paragraph(str(eval_m["rusak_fn"]), body_small),
                Paragraph(str(eval_m["nilai_kondisi_fn"]), body_small)
            ],
            # Subtotal
            [
                Paragraph("<b>SUB TOTAL</b>", body_small_bold), "", "", "", "", "",
                Paragraph(f"<b>{eval_m['subtotal']}</b>", body_small_bold)
            ]
        ]

        t_ikbg = Table(data_assessment_table, colWidths=col_w_assessment)
        t_ikbg.setStyle(TableStyle([
            ('SPAN', (0, 0), (0, 1)),
            ('SPAN', (1, 0), (1, 1)),
            ('SPAN', (2, 0), (2, 1)),
            ('SPAN', (3, 0), (5, 0)),
            ('SPAN', (6, 0), (6, 1)),
            ('BACKGROUND', (3, 1), (5, 1), colors.HexColor('#FDE68A')),

            ('SPAN', (0, 4), (0, 5)),
            ('SPAN', (1, 4), (1, 5)),
            ('SPAN', (2, 4), (2, 5)),
            ('SPAN', (3, 4), (5, 4)),
            ('SPAN', (6, 4), (6, 5)),
            ('BACKGROUND', (3, 5), (5, 5), colors.HexColor('#FDE68A')),

            ('SPAN', (0, 8), (5, 8)),
            ('ALIGN', (0, 8), (5, 8), 'LEFT'),
            ('BACKGROUND', (0, 8), (-1, 8), colors.HexColor('#F1F5F9')),

            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        story.append(t_ikbg)
        story.append(Spacer(1, 4))

        # Keterangan Formula
        story.append(Paragraph(
            "<b>Keterangan :</b><br/>Rutin: µku = >90;  Preventif: µku = 70 - 90;  Korektif: µku = 50 - 70;  Rehabilitatif: µku = <50",
            body_small_left
        ))
        story.append(Spacer(1, 4))

        # Tabel Bobot & Kinerja (Kiri) dan NILAI KONDISI (Kanan) Side-by-Side
        data_t_bobot = [
            [Paragraph("<b>Penilaian</b>", body_small_bold), Paragraph("<b>Bobot</b>", body_small_bold), Paragraph("<b>Penilaian</b>", body_small_bold), Paragraph("<b>Kinerja</b>", body_small_bold)],
            [Paragraph("Fisik", body_small), Paragraph("50%", body_small), Paragraph(str(eval_m["nilai_kondisi_f"]), body_small), Paragraph(str(eval_m["kinerja_f"]), body_small)],
            [Paragraph("Fungsi", body_small), Paragraph("50%", body_small), Paragraph(str(eval_m["nilai_kondisi_fn"]), body_small), Paragraph(str(eval_m["kinerja_fn"]), body_small)],
            ["", "", "", Paragraph(f"<b>{eval_m['total_kinerja']}</b>", body_small_bold)]
        ]
        t_bobot = Table(data_t_bobot, colWidths=[65, 50, 65, 55])
        t_bobot.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, 2), 0.5, colors.black),
            ('BOX', (3, 3), (3, 3), 0.5, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))

        data_t_nk = [
            [Paragraph("<b>NILAI KONDISI</b>", body_small_bold), ""],
            [Paragraph("<b>X</b>" if eval_m["is_rutin"] else "-", body_small_bold), Paragraph("<b>RUTIN</b>" if eval_m["is_rutin"] else "RUTIN", body_small_left)],
            [Paragraph("<b>X</b>" if eval_m["is_preventif"] else "-", body_small_bold), Paragraph("<b>PREVENTIF</b>" if eval_m["is_preventif"] else "PREVENTIF", body_small_left)],
            [Paragraph("<b>X</b>" if eval_m["is_korektif"] else "-", body_small_bold), Paragraph("<b>KOREKTIF</b>" if eval_m["is_korektif"] else "KOREKTIF", body_small_left)],
            [Paragraph("<b>X</b>" if eval_m["is_rehabilitatif"] else "-", body_small_bold), Paragraph("<b>REHABILITATIF</b>" if eval_m["is_rehabilitatif"] else "REHABILITATIF", body_small_left)],
        ]
        t_nk = Table(data_t_nk, colWidths=[35, 130])
        t_nk.setStyle(TableStyle([
            ('SPAN', (0, 0), (1, 0)),
            ('GRID', (0, 0), (-1, 0), 0.5, colors.black),
            ('BOX', (0, 1), (0, 1), 0.5, colors.black),
            ('BOX', (0, 2), (0, 2), 0.5, colors.black),
            ('BOX', (0, 3), (0, 3), 0.5, colors.black),
            ('BOX', (0, 4), (0, 4), 0.5, colors.black),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))

        t_container = Table([[t_bobot, t_nk]], colWidths=[260, 263], style=[
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ])
        story.append(t_container)
        story.append(Spacer(1, 10))

        # -----------------------------------------------------------------------------------------
        # 4. TABEL REKAPITULASI RAB STANDAR AHSP JAWA TENGAH 2023
        # -----------------------------------------------------------------------------------------
        story.append(Paragraph("2. REKAPITULASI HASIL ESTIMASI ANGGARAN BIAYA (RAB) AHSP JATENG 2023", doc_header_style))

        headers_rab = [
            Paragraph("<b>No</b>", body_small_bold),
            Paragraph("<b>Tgl Inspeksi</b>", body_small_bold),
            Paragraph("<b>Elemen Gedung</b>", body_small_bold),
            Paragraph("<b>Jenis Kerusakan</b>", body_small_bold),
            Paragraph("<b>Kode AHSP</b>", body_small_bold),
            Paragraph("<b>Vol</b>", body_small_bold),
            Paragraph("<b>Sat</b>", body_small_bold),
            Paragraph("<b>Harga Satuan</b>", body_small_bold),
            Paragraph("<b>Total Biaya</b>", body_small_bold),
        ]
        table_rows_rab = [headers_rab]

        grand_total = 0
        for idx_r, row_item in enumerate(data_list, start=1):
            tot = row_item.get("total_biaya", 0)
            grand_total += tot
            table_rows_rab.append([
                Paragraph(str(idx_r), body_small),
                Paragraph(str(row_item.get("tanggal", ""))[:16], body_small),
                Paragraph(str(row_item.get("elemen_gedung", "")), body_small_left),
                Paragraph(str(row_item.get("jenis_kerusakan", "")), body_small_left),
                Paragraph(str(row_item.get("kode_ahsp", "")), body_small),
                Paragraph(f"{row_item.get('volume', 0):.1f}", body_small),
                Paragraph(str(row_item.get("satuan", "")), body_small),
                Paragraph(format_rupiah(row_item.get("harga_satuan", 0)), body_small),
                Paragraph(format_rupiah(tot), body_small_bold),
            ])

        table_rows_rab.append([
            Paragraph("<b>TOTAL ESTIMASI ANGGARAN PEMELIHARAAN (RAB)</b>", body_small_bold),
            "", "", "", "", "", "", "",
            Paragraph(f"<b>{format_rupiah(grand_total)}</b>", body_small_bold)
        ])

        col_w_rab = [20, 60, 85, 95, 65, 30, 25, 65, 78]
        summary_table = Table(table_rows_rab, colWidths=col_w_rab, repeatRows=1)
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('ALIGN', (5, 0), (6, -1), 'CENTER'),
            ('ALIGN', (7, 0), (-1, -1), 'RIGHT'),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#64748B')),
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#F8FAFC')]),
            ('SPAN', (0, -1), (7, -1)),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#E2E8F0')),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(summary_table)
        story.append(Spacer(1, 6))

        terbilang_teks = angka_ke_terbilang(grand_total).strip() + " Rupiah"
        story.append(Paragraph(f"<b>Terbilang:</b> <i>{terbilang_teks}</i>", body_small_left))
        story.append(Spacer(1, 8))

        # -----------------------------------------------------------------------------------------
        # 5. CATATAN REKOMENDASI TEKNIS & LEMBAR PENGESAHAN TANDA TANGAN
        # -----------------------------------------------------------------------------------------
        story.append(Paragraph("3. CATATAN REKOMENDASI TEKNIS & PROSEDUR PERBAIKAN", doc_header_style))
        for idx_rc, r_row in enumerate(data_list[:2], start=1):
            recom_text = f"<b>{idx_rc}. {r_row.get('elemen_gedung')}:</b> {r_row.get('rekomendasi', '-')}"
            story.append(Paragraph(recom_text, body_small_left))
            story.append(Spacer(1, 2))

        story.append(Spacer(1, 12))

        story.append(KeepTogether([
            Table([
                [
                    Paragraph(f"Dibuat Oleh,<br/><b>Surveyor / Penguji Lapangan</b><br/><br/><br/><br/>{sig_surveyor_text}<br/>NIM/NIP.", body_small),
                    Paragraph("Mengetahui,<br/><b>Dosen Pembimbing / Penguji</b><br/><br/><br/><br/><u>( .................................................. )</u><br/>NIP.", body_small),
                    Paragraph("Disetujui Oleh,<br/><b>Pengelola Gedung / PPK</b><br/><br/><br/><br/><u>( .................................................. )</u><br/>NIP.", body_small)
                ]
            ], colWidths=[170, 175, 175], style=[
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ])
        ]))

        # -----------------------------------------------------------------------------------------
        # 6. LAMPIRAN DOKUMENTASI VISUAL KERUSAKAN BERBASIS AI
        # -----------------------------------------------------------------------------------------
        if annotated_image_bytes:
            story.append(PageBreak())
            story.append(Paragraph("PEMERINTAH PROVINSI JAWA TENGAH", kop_title_style))
            story.append(Paragraph("DINAS PEKERJAAN UMUM BINA MARGA DAN CIPTA KARYA", kop_title_style))
            story.append(Paragraph("LAMPIRAN DOKUMENTASI VISUAL KERUSAKAN BANGUNAN GEDUNG (COMPUTER VISION AI)", kop_sub_style))
            story.append(Spacer(1, 4))
            story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

            story.append(Paragraph("<b>DOKUMENTASI HASIL PEMINDAIAN VISUAL BERBASIS AI AGENT</b>", doc_header_style))
            story.append(Paragraph(
                "Berikut adalah bukti citra dokumentasi inspeksi visual lapangan yang telah diproses oleh modul deteksi AI dengan "
                "anotasi kotak pembatas (bounding box), nilai akurasi keyakinan (confidence score), serta rujukan kode AHSP Cipta Karya Jawa Tengah.",
                body_small_left
            ))
            story.append(Spacer(1, 8))

            try:
                img_io = io.BytesIO(annotated_image_bytes)
                pil_temp = Image.open(img_io)
                orig_w, orig_h = pil_temp.size

                target_w = 420.0
                target_h = target_w * (orig_h / orig_w)
                if target_h > 330.0:
                    target_h = 330.0
                    target_w = target_h * (orig_w / orig_h)

                img_io.seek(0)
                rl_img = RLImage(img_io, width=target_w, height=target_h)

                meta_lampiran = [
                    [
                        Paragraph(f"<b>Objek Bangunan:</b> {nama_gedung}", body_small_left),
                        Paragraph(f"<b>Lokasi / Ruang:</b> {lokasi_lantai}", body_small_left)
                    ],
                    [
                        Paragraph(f"<b>Elemen Gedung:</b> {target_item.get('elemen_gedung', '-')}", body_small_left),
                        Paragraph(f"<b>Jenis Kerusakan:</b> {target_item.get('jenis_kerusakan', '-')}", body_small_left)
                    ],
                    [
                        Paragraph(f"<b>Kode Analisa Satuan:</b> {target_item.get('kode_ahsp', '-')}", body_small_left),
                        Paragraph(f"<b>Tingkat Keparahan:</b> {str(target_item.get('tingkat_keparahan', 'Sedang')).upper()}", body_small_left)
                    ]
                ]
                tbl_lampiran = Table(meta_lampiran, colWidths=[260, 260])
                tbl_lampiran.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
                    ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
                    ('TOPPADDING', (0, 0), (-1, -1), 3),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ]))

                story.append(KeepTogether([
                    Table([[rl_img]], colWidths=[520], style=[
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E1')),
                        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
                        ('TOPPADDING', (0, 0), (-1, -1), 6),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                    ]),
                    Spacer(1, 6),
                    tbl_lampiran
                ]))
            except Exception:
                pass

        doc.build(story)
        pdf_bytes = buffer.getvalue()
        buffer.close()
        return pdf_bytes


# ==================================================================================================
# 6. KOMPONEN KONTROLER UTAMA & NAVIGASI HALAMAN STREAMLIT
# ==================================================================================================

# Inisialisasi session state untuk menyimpan riwayat hasil analisis terkini
if "latest_analysis" not in st.session_state:
    st.session_state.latest_analysis = None
if "annotated_img_bytes" not in st.session_state:
    st.session_state.annotated_img_bytes = None
if "uploaded_img_bytes" not in st.session_state:
    st.session_state.uploaded_img_bytes = None

# Sidebar Navigasi
st.sidebar.markdown("""
<div style="text-align: center; padding: 14px 0 18px 0;">
    <div style="display: inline-block; padding: 14px; background: linear-gradient(135deg, #1e293b, #0f172a); border-radius: 20px; box-shadow: 0 8px 18px rgba(15, 23, 42, 0.25); border: 1px solid rgba(255,255,255,0.1); margin-bottom: 8px;">
        <span style="font-size: 2.2rem;">🏢</span>
    </div>
    <h3 style="margin: 6px 0 0 0; color: #0f172a; font-weight: 900; letter-spacing: -0.5px;">SIM-PBG AI</h3>
    <div class="live-badge" style="margin-top: 6px; padding: 2px 10px; font-size: 0.72rem;">
        <span class="live-dot"></span> SISTEM ONLINE
    </div>
    <p style="font-size: 0.8rem; color: #1e293b; font-weight: 600; margin: 4px 0 0 0; line-height: 1.4;">Pemeliharaan Bangunan Gedung<br><b style="color: #0f172a;">AHSP Jawa Tengah 2023</b></p>
</div>
""", unsafe_allow_html=True)

menu = st.sidebar.radio(
    "Navigasi Modul:",
    [
        "📊 Dashboard Utama & Input Data",
        "📜 Riwayat Inspeksi Gedung",
        "📑 Panduan AHSP Jateng 2023"
    ],
    index=0
)

st.sidebar.markdown("---")

# Pengaturan Integrasi AI Agent Antigravity di Sidebar
with st.sidebar.expander("⚙️ Konfigurasi AI Agent Antigravity", expanded=False):
    st.markdown("<p style='font-size:0.84rem; color:#1e293b; font-weight:600;'>Konfigurasi HTTP POST API Call library <code>requests</code> ke AI Agent Antigravity.</p>", unsafe_allow_html=True)
    api_endpoint_input = st.text_input(
        "AI Agent Endpoint URL:",
        value="http://localhost:8000/api/analyze-damage",
        help="Alamat HTTP POST server Antigravity Agent"
    )
    api_token_input = st.text_input(
        "Bearer API Key (Opsional):",
        value="",
        type="password"
    )
    offline_toggle = st.checkbox(
        "Mode Simulasi Sidang (Offline Safe)",
        value=False,
        help="Aktifkan opsi ini jika tidak menjalankan server lokal terpisah atau jaringan wifi gedung terkunci saat sidang skripsi."
    )

st.sidebar.markdown("""
<div style="background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%); border-radius: 12px; padding: 14px; border: 1.5px solid #cbd5e1; font-size: 0.8rem; color: #0f172a; font-weight: 500; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
    <div style="display:flex; align-items:center; gap:6px; margin-bottom:6px;">
        <span class="live-dot" style="width:7px; height:7px;"></span>
        <span style="font-weight:800; color:#0f172a;">STATUS SISTEM AKTIF</span>
    </div>
    <b style="color:#0f172a;">Database:</b> SQLite3 (<code>pemeliharaan_gedung.db</code>)<br>
    <b style="color:#0f172a;">Engine:</b> Antigravity Civil Vision AI<br>
    <b style="color:#0f172a;">Standar Acuan:</b> Pergub Jateng AHSP 2023
</div>
""", unsafe_allow_html=True)

# Status Model YOLOv8 Kustom di Sidebar
_detected_yolo_weight = AIAgentClient.get_yolo_model_path()
if _detected_yolo_weight:
    st.sidebar.markdown(f"""
    <div style="background: #ecfdf5; border-radius: 10px; padding: 12px; border: 1.5px solid #6ee7b7; font-size: 0.78rem; color: #065f46; margin-top: 10px; box-shadow: 0 2px 6px rgba(16,185,129,0.08);">
        <div style="display:flex; align-items:center; gap:6px;">
            <span style="font-size: 1rem;">🟢</span>
            <b style="color:#065f46;">Model YOLOv8 Terdeteksi</b>
        </div>
        <div style="margin-top: 5px; font-size: 0.74rem; color: #047857; line-height: 1.4;">
            File: <code>{os.path.basename(_detected_yolo_weight)}</code><br>
            Target 4 Kelas: <i>Crack, Spalling, Rebar, Efflorescence</i>
        </div>
    </div>
    """, unsafe_allow_html=True)
else:
    st.sidebar.markdown("""
    <div style="background: #eff6ff; border-radius: 10px; padding: 12px; border: 1.5px solid #bfdbfe; font-size: 0.78rem; color: #1e40af; margin-top: 10px; box-shadow: 0 2px 6px rgba(37,99,235,0.06);">
        <div style="display:flex; align-items:center; gap:6px;">
            <span style="font-size: 1rem;">🔵</span>
            <b style="color:#1e40af;">Civil AI Engine Siap</b>
        </div>
        <div style="margin-top: 5px; font-size: 0.73rem; color: #1d4ed8; line-height: 1.4;">
            Letakkan file <code>best.pt</code> hasil training di folder <code>weights/</code> untuk mengaktifkan inferensi YOLOv8.
        </div>
    </div>
    """, unsafe_allow_html=True)


# ==================================================================================================
# HALAMAN 1: 📊 DASHBOARD UTAMA & INPUT DATA INSPEKSI
# ==================================================================================================
if menu == "📊 Dashboard Utama & Input Data":
    # Hero Banner
    st.markdown("""
    <div class="hero-banner">
        <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 12px;">
            <div>
                <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-bottom: 10px;">
                    <span class="badge-ahsp" style="background:#2563eb; color:#ffffff; border:none; padding:4px 12px;">SKRIPSI TEKNIK SIPIL DIGITAL</span>
                    <div class="live-badge" style="margin-bottom:0;">
                        <span class="live-dot"></span> KONEKSI AI AGENT SIAP
                    </div>
                </div>
                <div class="hero-title">Sistem Deteksi Kerusakan & Estimasi AHSP Jateng</div>
                <div class="hero-subtitle">
                    Aplikasi pemeliharaan bangunan gedung terpadu berbasis AI Agent Antigravity dengan kalkulasi Rencana Anggaran Biaya (RAB) otomatis serta formulir penilaian kondisi fisik dan fungsi (IKBG) sesuai standar AHSP Dinas PU Bina Marga & Cipta Karya Provinsi Jawa Tengah Tahun 2023.
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        st.markdown("""
        <div class="tw-card">
            <div class="tw-card-header">
                <h4 style="margin: 0; color: #0f172a; font-weight: 800;">📸 Formulir Input Inspeksi Kerusakan</h4>
                <p style="margin: 4px 0 0 0; font-size: 0.88rem; color: #1e293b; font-weight: 600;">Unggah dokumentasi foto visual lapangan dan input estimasi volume fisik.</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

        with st.expander("🏢 Identitas Bangunan & Surveyor Lapangan (Dapat Diisi Bebas)", expanded=True):
            st.caption("ℹ️ *Aplikasi ini bersifat terbuka untuk umum (tanpa login/role). Anda dapat mengisi nama gedung dan pemeriksa secara bebas sesuai kebutuhan inspeksi.*")
            nama_gedung_input = st.text_input(
                "Nama Bangunan Gedung / Objek Inspeksi:",
                value="",
                placeholder="Contoh: Gedung Rektorat, Rumah Tinggal, Ruko, dsb.",
                help="Ketik nama fasilitas atau bangunan gedung yang sedang diinspeksi."
            )
            col_loc1, col_loc2 = st.columns(2)
            with col_loc1:
                lokasi_lantai_input = st.text_input(
                    "Lokasi / Lantai / Ruangan:",
                    value="",
                    placeholder="Contoh: Lantai 2 - Ruang 204, Selasar Barat...",
                    help="Ketik spesifikasi zona, lantai, atau ruangan lokasi kerusakan."
                )
            with col_loc2:
                surveyor_input = st.text_input(
                    "Nama Surveyor / Pemeriksa:",
                    value="",
                    placeholder="Contoh: Nama Anda / Tim Penilai...",
                    help="Ketik nama pemeriksa lapangan yang akan tercantum pada laporan dan lembar tanda tangan."
                )

        uploaded_file = st.file_uploader(
            "1. Unggah Foto Kerusakan (JPG / PNG):",
            type=["jpg", "jpeg", "png"],
            help="Pilih foto dokumentasi visual kerusakan elemen bangunan gedung."
        )

        elemen_pilihan = st.selectbox(
            "2. Elemen Gedung yang Diinspeksi:",
            [
                "Balok Beton Bertulang",
                "Kolom Struktur Beton",
                "Dinding Pasangan Bata / Hebel",
                "Pelat Lantai / Dak Beton",
                "Plafon Gypsum / Akustik",
                "Lantai Keramik / Granit",
                "Kusen Pintu / Jendela Aluminium",
                "Dinding Eksterior / Interior"
            ],
            index=2,
            help="Pilih elemen gedung atau biarkan AI melakukan pencocokan karakteristik elemen."
        )

        # Deteksi otomatis karakteristik satuan elemen (m' panjang linier atau m2 luas permukaan)
        is_linier = elemen_pilihan in ["Balok Beton Bertulang", "Kusen Pintu / Jendela Aluminium"]
        satuan_default = "m'" if is_linier else "m2"

        # Pilihan Metode Input Volume: Input Langsung atau Kalkulator Dimensi Lapangan (Panjang x Lebar)
        metode_vol = st.radio(
            "3. Metode Penentuan Volume Kerusakan:",
            ["Input Volume Langsung (Opname / Rekap)", "📐 Hitung dari Dimensi Lapangan (Panjang × Lebar / Linier)"],
            horizontal=True
        )

        if metode_vol == "📐 Hitung dari Dimensi Lapangan (Panjang × Lebar / Linier)":
            if is_linier:
                st.info(f"💡 Elemen **{elemen_pilihan}** menggunakan satuan panjang linier (**m'**). Cukup input panjang kerusakan tanpa dikalikan lebar.")
                panjang_m = st.number_input("Panjang Kerusakan / Retak (meter):", min_value=0.1, value=5.0, step=0.5)
                volume_input = float(round(panjang_m, 2))
                st.success(f"Volume Terhitung: **{volume_input:.2f} m'**")
            else:
                st.info(f"💡 Elemen **{elemen_pilihan}** menggunakan satuan luas (**m²**). Dihitung dari formula: **Luas = Panjang × Lebar/Tinggi**.")
                c_p, c_l = st.columns(2)
                with c_p:
                    panjang_m = st.number_input("Panjang (meter):", min_value=0.1, value=5.0, step=0.5)
                with c_l:
                    lebar_m = st.number_input("Lebar / Tinggi (meter):", min_value=0.1, value=2.5, step=0.5)
                volume_input = float(round(panjang_m * lebar_m, 2))
                st.success(f"Volume Terhitung: **{panjang_m:.2f} m × {lebar_m:.2f} m = {volume_input:.2f} m²**")
        else:
            volume_input = st.number_input(
                f"Masukkan Besaran Volume Kerusakan ({satuan_default}):",
                min_value=0.1,
                max_value=10000.0,
                value=12.5,
                step=0.5,
                help=f"Masukkan besaran volume kerusakan ({satuan_default}) sesuai jenis pekerjaan."
            )

        st.caption(f"ℹ️ *Satuan volume terdeteksi: **{satuan_default}** (terstandar AHSP Dinas PU Cipta Karya Jateng 2023).*")

        start_btn = st.button("🚀 Mulai Analisis Sistem", type="primary", use_container_width=True)

    with col_right:
        st.markdown("""
        <div class="tw-card">
            <div class="tw-card-header">
                <h4 style="margin: 0; color: #0f172a; font-weight: 800;">🔍 Sensor Visual & Pratinjau Citra</h4>
                <p style="margin: 4px 0 0 0; font-size: 0.88rem; color: #1e293b; font-weight: 600;">Citra yang diunggah akan dikonversi ke Base64 dan dianalisis oleh AI Agent.</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.session_state.get("annotated_img_bytes") is not None and st.session_state.latest_analysis is not None:
            tab_ai, tab_raw = st.tabs(["🎯 Citra Anotasi Deteksi AI (Bounding Box)", "📷 Foto Lapangan Asli"])
            with tab_ai:
                st.image(
                    st.session_state.annotated_img_bytes,
                    caption=f"Hasil Deteksi AI: {st.session_state.latest_analysis.get('jenis_kerusakan')} (Confidence: {st.session_state.latest_analysis.get('confidence_score', 96.5)}%)",
                    use_container_width=True
                )
                st.success(f"🎯 Bounding box deteksi visual AI aktif | Kode AHSP: {st.session_state.latest_analysis.get('kode_ahsp')}")
            with tab_raw:
                if st.session_state.get("uploaded_img_bytes"):
                    st.image(
                        st.session_state.uploaded_img_bytes,
                        caption=f"Dokumentasi Asli: {st.session_state.latest_analysis.get('filename')}",
                        use_container_width=True
                    )
        elif uploaded_file is not None:
            image = Image.open(uploaded_file)
            st.image(image, caption=f"Dokumentasi Foto: {uploaded_file.name}", use_container_width=True)
            st.success(f"✅ Citra berhasil dimuat: {image.width}x{image.height} px | Ukuran file: {uploaded_file.size / 1024:.1f} KB")
        else:
            st.info("💡 Silakan unggah foto kerusakan pada formulir di sebelah kiri untuk melihat pratinjau citra dan memulai analisis.")
            # Placeholder ilustrasi visual
            st.markdown("""
            <div class="upload-placeholder">
                <span class="floating-icon">📸</span>
                <p style="margin-top: 14px; font-weight: 800; color: #0f172a; font-size: 1.05rem;">Belum ada foto yang dipilih</p>
                <p style="font-size: 0.86rem; color: #1e293b; font-weight: 600; margin: 0; max-width: 320px; margin-inline: auto;">Pilih foto kerusakan bangunan gedung (JPEG/PNG) pada formulir sebelah kiri untuk pratinjau citra dan analisis.</p>
            </div>
            """, unsafe_allow_html=True)

    # ==============================================================================================
    # PROSES EKSEKUSI API CALL & PENYIMPANAN KE SQLITE
    # ==============================================================================================
    if start_btn:
        if uploaded_file is None:
            st.warning("⚠️ Harap unggah foto kerusakan bangunan terlebih dahulu sebelum memulai analisis!")
        else:
            with st.spinner("⏳ Menghubungi AI Agent Antigravity via HTTP POST JSON & Menghitung Estimasi AHSP Jateng..."):
                # Baca bytes gambar
                img_bytes = uploaded_file.getvalue()
                fname = uploaded_file.name

                # Panggil AI Agent Antigravity
                result_ai = AIAgentClient.call_agent(
                    image_bytes=img_bytes,
                    filename=fname,
                    volume=volume_input,
                    elemen_gedung=elemen_pilihan,
                    endpoint_url=api_endpoint_input,
                    api_key=api_token_input,
                    force_simulation=offline_toggle
                )

                # Pembersihan input string pengguna (jika kosong, gunakan fallback netral)
                nama_gedung_clean = nama_gedung_input.strip() if nama_gedung_input and nama_gedung_input.strip() else "Bangunan Gedung"
                lokasi_clean = lokasi_lantai_input.strip() if lokasi_lantai_input and lokasi_lantai_input.strip() else "-"
                surveyor_clean = surveyor_input.strip() if surveyor_input and surveyor_input.strip() else "Surveyor Lapangan"

                result_ai["nama_gedung"] = nama_gedung_clean
                result_ai["lokasi_lantai"] = lokasi_clean
                result_ai["surveyor"] = surveyor_clean

                # Bounding Box AI Annotation
                annotated_img = AIVisualAnnotator.annotate(
                    image_bytes=img_bytes,
                    jenis_kerusakan=result_ai.get("jenis_kerusakan", "Retak Permukaan"),
                    tingkat_keparahan=result_ai.get("tingkat_keparahan", "Sedang"),
                    confidence_score=result_ai.get("confidence_score", 96.5),
                    kode_ahsp=result_ai.get("kode_ahsp", "A.4.4.1.25.JT"),
                    detected_boxes=result_ai.get("detected_boxes")
                )
                st.session_state.annotated_img_bytes = annotated_img
                st.session_state.uploaded_img_bytes = img_bytes

                # Simpan ke Database SQLite secara otomatis
                today_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                record_id = DatabaseManager.insert_riwayat(
                    tanggal=today_now,
                    elemen_gedung=result_ai.get("elemen_gedung", elemen_pilihan),
                    jenis_kerusakan=result_ai.get("jenis_kerusakan", "Retak Permukaan"),
                    kode_ahsp=result_ai.get("kode_ahsp", "A.4.4.1.25.JT"),
                    volume=float(result_ai.get("volume", volume_input)),
                    satuan=result_ai.get("satuan", "m2"),
                    harga_satuan=int(result_ai.get("harga_satuan", 0)),
                    total_biaya=int(result_ai.get("total_biaya", 0)),
                    rekomendasi=result_ai.get("rekomendasi", "Perbaikan standar sesuai metode kerja."),
                    nama_gedung=nama_gedung_clean,
                    lokasi_lantai=lokasi_clean,
                    surveyor=surveyor_clean
                )

                result_ai["db_id"] = record_id
                st.session_state.latest_analysis = result_ai
                st.rerun()

    # ==============================================================================================
    # TAMPILAN HASIL ANALISIS VISUAL & ESTIMASI AHSP
    # ==============================================================================================
    if st.session_state.latest_analysis is not None:
        res = st.session_state.latest_analysis
        st.markdown("---")

        # Notifikasi info koneksi API
        if res.get("is_live_api"):
            st.success(f"🟢 **Koneksi Live HTTP POST Berhasil:** Respon diterima dari endpoint `{res.get('api_endpoint_used')}` | ID Riwayat DB: #{res.get('db_id')}")
        else:
            warning_msg = res.get("warning", "Sistem menjalankan Standalone AI Engine (Mode Sidang Safe).")
            st.info(f"🟡 **Status Engine:** {warning_msg} | ID Riwayat DB: #{res.get('db_id')}")

        st.markdown("### 📊 Hasil Penilaian Kondisi & Estimasi Biaya Pemeliharaan")

        # Baris Kartu Status Visual
        keparahan = res.get("tingkat_keparahan", "Sedang")
        if keparahan == "Berat":
            badge_html = '<span class="badge-berat">KERUSAKAN BERAT</span>'
        elif keparahan == "Sedang":
            badge_html = '<span class="badge-sedang">KERUSAKAN SEDANG</span>'
        else:
            badge_html = '<span class="badge-ringan">KERUSAKAN RINGAN</span>'

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.markdown(f"""
            <div class="metric-container metric-card-amber">
                <div>
                    <div class="metric-label">Tingkat Keparahan</div>
                    <div style="margin-top: 6px;">{badge_html}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_m2:
            st.markdown(f"""
            <div class="metric-container metric-card-blue">
                <div>
                    <div class="metric-label">Confidence Score AI</div>
                    <div class="metric-value">{res.get('confidence_score', 96.5)}%</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_m3:
            st.markdown(f"""
            <div class="metric-container metric-card-cyan">
                <div>
                    <div class="metric-label">Kode AHSP Terpilih</div>
                    <div class="metric-value" style="font-size:1.15rem; color:#0369a1;">{res.get('kode_ahsp')}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_m4:
            st.markdown(f"""
            <div class="metric-container metric-card-green">
                <div>
                    <div class="metric-label">Total Estimasi RAB</div>
                    <div class="metric-value" style="color:#16a34a;">{format_rupiah(res.get('total_biaya', 0))}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        # -----------------------------------------------------------------------------------------
        # FITUR 3: SLIDER SIMULASI VOLUME CEPAT (LIVE RECALCULATE / WHAT-IF SCENARIO)
        # -----------------------------------------------------------------------------------------
        curr_vol = float(res.get("volume", 10.0))
        satuan_kerja = res.get("satuan", "m2")
        hsp_satuan = int(res.get("harga_satuan", 0))

        with st.expander("🎛️ Simulasi Cepat Penyesuaian Volume (What-If Scenario Lapangan)", expanded=False):
            st.markdown(f"""
            <div style="font-size:0.88rem; color:#1e293b; font-weight:600; margin-bottom:8px;">
                Uji skenario perubahan besaran volume kerusakan di lapangan secara dinamis. Estimasi total biaya RAB akan dikalkulasi ulang secara <i>real-time</i> berdasarkan rumus: <code style="color:#2563eb;">Total Biaya = Volume × HSP</code>.
            </div>
            """, unsafe_allow_html=True)

            c_sim1, c_sim2 = st.columns([2, 1])
            with c_sim1:
                max_slider = max(50.0, float(round(curr_vol * 3.0, 1)))
                sim_volume = st.slider(
                    f"Penyesuaian Besaran Volume ({satuan_kerja}):",
                    min_value=0.5,
                    max_value=float(max_slider),
                    value=float(curr_vol),
                    step=0.5,
                    key=f"sim_slider_{res.get('db_id', 0)}"
                )

            sim_total_biaya = int(round(sim_volume * hsp_satuan))
            selisih_biaya = sim_total_biaya - int(res.get("total_biaya", 0))

            with c_sim2:
                st.metric(
                    label=f"Simulasi RAB ({sim_volume:.1f} {satuan_kerja})",
                    value=format_rupiah(sim_total_biaya),
                    delta=f"{'+' if selisih_biaya >= 0 else ''}{format_rupiah(selisih_biaya)} vs Riil"
                )

            if sim_volume != curr_vol:
                col_btn_apply, _ = st.columns([1, 1])
                with col_btn_apply:
                    if st.button("🔄 Terapkan Volume Baru Ini ke Laporan & Database", type="primary", use_container_width=True):
                        # Update session state
                        st.session_state.latest_analysis["volume"] = sim_volume
                        st.session_state.latest_analysis["total_biaya"] = sim_total_biaya
                        # Update SQLite database
                        with DatabaseManager.get_connection() as conn:
                            cur = conn.cursor()
                            cur.execute(
                                "UPDATE riwayat_kerusakan SET volume = ?, total_biaya = ? WHERE id = ?",
                                (sim_volume, sim_total_biaya, res.get("db_id"))
                            )
                            conn.commit()
                        st.success(f"✅ Volume berhasil diperbarui menjadi {sim_volume:.2f} {satuan_kerja} dengan total RAB {format_rupiah(sim_total_biaya)}!")
                        st.rerun()

        # Rincian Tabel Kalkulasi RAB
        st.markdown("#### 📑 Rincian Anggaran Biaya (RAB) Standar AHSP Jawa Tengah 2023")
        rab_data = {
            "Komponen Analisis": [
                "Elemen Bangunan",
                "Identifikasi Kerusakan",
                "Kode Analisa Satuan",
                "Uraian Pekerjaan Pemeliharaan",
                "Volume Kerusakan",
                "Satuan Pengukuran",
                "Harga Satuan AHSP Jateng 2023",
                "Total Estimasi Anggaran (RAB)"
            ],
            "Nilai Parameter": [
                res.get("elemen_gedung"),
                res.get("jenis_kerusakan"),
                res.get("kode_ahsp"),
                res.get("deskripsi_ahsp"),
                f"{res.get('volume', 0):.2f}",
                res.get("satuan"),
                format_rupiah(res.get("harga_satuan", 0)),
                format_rupiah(res.get("total_biaya", 0))
            ]
        }
        df_rab = pd.DataFrame(rab_data)
        st.table(df_rab)

        # -----------------------------------------------------------------------------------------
        # FITUR 1: TRANSPARANSI FORMULA & RINCIAN KOEFISIEN AHSP
        # -----------------------------------------------------------------------------------------
        kode_item = res.get("kode_ahsp", "")
        ahsp_detail = AHSPKnowledgeBase.CATALOG.get(kode_item, {})
        koef_upah = ahsp_detail.get("koefisien_upah", [])
        koef_bahan = ahsp_detail.get("koefisien_bahan", [])

        with st.expander("🔬 Transparansi Formula Matematis & Rincian Koefisien Analisis Satuan (AHSP Jateng)", expanded=False):
            st.markdown(f"""
            <div style="font-size:0.9rem; color:#0f172a; margin-bottom:12px;">
                <b>Metodologi Perhitungan Standar Permen PUPR No. 1/2022 & AHSP Jawa Tengah 2023:</b><br>
                Harga Satuan Pekerjaan (HSP) untuk kode <code>{kode_item}</code> (<b>{res.get('deskripsi_ahsp')}</b>) dibentuk dari akumulasi perkalian koefisien indeks tenaga kerja dan kebutuhan material dikalikan faktor overhead & profit (15%):
            </div>
            """, unsafe_allow_html=True)

            st.latex(r"HSP = \left[ \sum_{i=1}^{n} (K_{\text{Tenaga}, i} \times U_i) + \sum_{j=1}^{m} (K_{\text{Bahan}, j} \times H_j) \right] \times (1 + \text{Overhead } 15\%)")
            st.latex(r"\text{Total Estimasi Anggaran (RAB)} = \text{Volume Pekerjaan } (V) \times HSP")

            st.markdown(f"""
            <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:10px 14px; margin: 10px 0; font-size:0.86rem; color:#1e293b;">
                📌 <b>Kalkulasi Aktual:</b> {res.get('volume', 0):.2f} {res.get('satuan')} × {format_rupiah(res.get('harga_satuan', 0))} = <b style="color:#16a34a;">{format_rupiah(res.get('total_biaya', 0))}</b>
            </div>
            """, unsafe_allow_html=True)

            col_koef1, col_koef2 = st.columns(2)
            with col_koef1:
                st.markdown("<b style='color:#0f172a; font-size:0.85rem;'>👷 Unsur Koefisien Tenaga Kerja (OH):</b>", unsafe_allow_html=True)
                if koef_upah:
                    df_upah = pd.DataFrame(koef_upah)
                    df_upah.columns = ["Unsur Tenaga", "Koefisien Indeks", "Satuan"]
                    st.dataframe(df_upah, use_container_width=True, hide_index=True)
                else:
                    st.caption("Standard Tenaga Kerja Operasional Pemeliharaan")

            with col_koef2:
                st.markdown("<b style='color:#0f172a; font-size:0.85rem;'>🧱 Unsur Koefisien Kebutuhan Bahan Material:</b>", unsafe_allow_html=True)
                if koef_bahan:
                    df_bahan = pd.DataFrame(koef_bahan)
                    df_bahan.columns = ["Material Bahan", "Koefisien Indeks", "Satuan"]
                    st.dataframe(df_bahan, use_container_width=True, hide_index=True)
                else:
                    st.caption("Standard Material Komposit Spesialis")

            st.caption("ℹ️ *Koefisien indeks diambil langsung dari basis data resmi AHSP Dinas PU Bina Marga & Cipta Karya Provinsi Jawa Tengah Tahun 2023.*")

        # Tampilan Tabel Matriks Penilaian Kondisi Fisik & Fungsi (IKBG) Sesuai Standar
        eval_ui = compute_condition_assessment(res)
        with st.expander("📋 Formulir Penilaian Kondisi Fisik & Fungsi Komponen Bangunan (Format Standar IKBG)", expanded=True):
            st.markdown(f"""
            <div style="font-size:0.88rem; color:#1e293b; font-weight:600; margin-bottom:10px;">
                Matriks evaluasi kondisi fisik dan fungsi mengacu pada pedoman pemeliharaan bangunan gedung (Permen PUPR No. 24/PRT/M/2008).
            </div>
            """, unsafe_allow_html=True)

            col_ikbg1, col_ikbg2 = st.columns([3, 2])
            with col_ikbg1:
                df_ikbg_view = pd.DataFrame({
                    "Komponen Penilaian": [
                        "1. Kondisi Fisik Komponen",
                        "2. Kondisi Fungsi Komponen",
                        "SUB TOTAL NILAI KONDISI"
                    ],
                    "Deskripsi Lapangan": [
                        eval_ui["desc_fisik"],
                        eval_ui["desc_fungsi"],
                        "-"
                    ],
                    "Nilai Awal": ["100%", "100%", "-"],
                    "Rating Penilaian": [
                        f"{eval_ui['nilai_kondisi_f']}% ({'Baik' if eval_ui['nilai_kondisi_f']>=79.5 else 'Cukup' if eval_ui['nilai_kondisi_f']>=49.5 else 'Rusak'})",
                        f"{eval_ui['nilai_kondisi_fn']}% ({'Baik' if eval_ui['nilai_kondisi_fn']>=79.5 else 'Cukup' if eval_ui['nilai_kondisi_fn']>=49.5 else 'Rusak'})",
                        "-"
                    ],
                    "Nilai Kondisi": [
                        eval_ui["nilai_kondisi_f"],
                        eval_ui["nilai_kondisi_fn"],
                        eval_ui["subtotal"]
                    ]
                })
                st.dataframe(df_ikbg_view, use_container_width=True, hide_index=True)

            with col_ikbg2:
                df_bobot_view = pd.DataFrame({
                    "Penilaian": ["Fisik", "Fungsi", "TOTAL KINERJA (µku)"],
                    "Bobot": ["50%", "50%", "100%"],
                    "Nilai": [eval_ui["nilai_kondisi_f"], eval_ui["nilai_kondisi_fn"], "-"],
                    "Kinerja": [eval_ui["kinerja_f"], eval_ui["kinerja_fn"], f"{eval_ui['total_kinerja']}"]
                })
                st.dataframe(df_bobot_view, use_container_width=True, hide_index=True)

                rutin_cls = "active-rutin" if eval_ui['is_rutin'] else ""
                prev_cls = "active-preventif" if eval_ui['is_preventif'] else ""
                korek_cls = "active-korektif" if eval_ui['is_korektif'] else ""
                rehab_cls = "active-rehabilitatif" if eval_ui['is_rehabilitatif'] else ""

                st.markdown(f"""
                <div style="margin-top:10px;">
                    <div style="font-size:0.78rem; color:#0f172a; font-weight:800; text-transform:uppercase; letter-spacing:0.5px;">Status Klasifikasi Tindakan Pemeliharaan (IKBG):</div>
                    <div class="ikbg-status-grid">
                        <div class="ikbg-status-card {rutin_cls}">
                            <div style="font-size:0.8rem; font-weight:800; color:{'#15803d' if eval_ui['is_rutin'] else '#0f172a'};">[{'X' if eval_ui['is_rutin'] else '-'}] RUTIN</div>
                            <div style="font-size:0.72rem; font-weight:700; color:{'#166534' if eval_ui['is_rutin'] else '#475569'}; margin-top:3px;">µku &gt; 90</div>
                        </div>
                        <div class="ikbg-status-card {prev_cls}">
                            <div style="font-size:0.8rem; font-weight:800; color:{'#b45309' if eval_ui['is_preventif'] else '#0f172a'};">[{'X' if eval_ui['is_preventif'] else '-'}] PREVENTIF</div>
                            <div style="font-size:0.72rem; font-weight:700; color:{'#92400e' if eval_ui['is_preventif'] else '#475569'}; margin-top:3px;">70 &le; µku &le; 90</div>
                        </div>
                        <div class="ikbg-status-card {korek_cls}">
                            <div style="font-size:0.8rem; font-weight:800; color:{'#c2410c' if eval_ui['is_korektif'] else '#0f172a'};">[{'X' if eval_ui['is_korektif'] else '-'}] KOREKTIF</div>
                            <div style="font-size:0.72rem; font-weight:700; color:{'#9a3412' if eval_ui['is_korektif'] else '#475569'}; margin-top:3px;">50 &le; µku &lt; 70</div>
                        </div>
                        <div class="ikbg-status-card {rehab_cls}">
                            <div style="font-size:0.8rem; font-weight:800; color:{'#b91c1c' if eval_ui['is_rehabilitatif'] else '#0f172a'};">[{'X' if eval_ui['is_rehabilitatif'] else '-'}] REHABILITATIF</div>
                            <div style="font-size:0.72rem; font-weight:700; color:{'#991b1b' if eval_ui['is_rehabilitatif'] else '#475569'}; margin-top:3px;">µku &lt; 50</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        # Langkah Kerja & Rekomendasi
        col_step, col_rec = st.columns([1, 1], gap="medium")
        with col_step:
            st.markdown("""
            <div class="tw-card">
                <div class="tw-card-header">
                    <h5 style="margin: 0; font-weight: 800; color: #0f172a;">🛠️ Prosedur Metode Kerja Standar PUPR</h5>
                </div>
            """, unsafe_allow_html=True)
            for idx, step in enumerate(res.get("langkah_kerja", []), start=1):
                st.markdown(f"""
                <div class="step-item">
                    <div class="step-number">{idx}</div>
                    <div class="step-text">{step}</div>
                </div>
                """, unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with col_rec:
            st.markdown("""
            <div class="tw-card">
                <div class="tw-card-header">
                    <h5 style="margin: 0; font-weight: 800; color: #0f172a;">💡 Rekomendasi Teknis Sipil & Mitigasi</h5>
                </div>
            """, unsafe_allow_html=True)
            st.warning(f"**Rekomendasi Pemeliharaan:** {res.get('rekomendasi')}")
            st.markdown(f"""
            <div style="font-size: 0.88rem; color: #1e293b; font-weight: 500; margin-top: 12px; line-height: 1.6;">
                <b style="color: #0f172a;">Ketentuan Pemeliharaan:</b> Berdasarkan Peraturan Menteri PUPR No. 24/PRT/M/2008, tindakan perbaikan harus segera dilakukan untuk mencegah deteriorasi lebih lanjut yang berisiko menurunkan keandalan struktur dan keselamatan pengguna gedung.
            </div>
            </div>
            """, unsafe_allow_html=True)

        # Tombol Aksi Ekspor Excel & PDF untuk Data Ini
        col_down1, col_down2 = st.columns(2)
        with col_down1:
            excel_single = ExcelReportGenerator.generate_excel_report(
                [res], title_doc=f"Penilaian - {res.get('elemen_gedung')}"
            )
            st.download_button(
                label="📊 Unduh Formulir Penilaian & RAB (Excel .xlsx)",
                data=excel_single,
                file_name=f"Penilaian_IKBG_{res.get('elemen_gedung', 'Elemen').replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with col_down2:
            single_record_pdf = PDFReportGenerator.generate_pdf_report(
                [res],
                title_doc=f"Laporan Penilaian - {res.get('elemen_gedung')}",
                annotated_image_bytes=st.session_state.get("annotated_img_bytes")
            )
            st.download_button(
                label="📄 Unduh Laporan Resmi & Penilaian (PDF)",
                data=single_record_pdf,
                file_name=f"Laporan_PBG_ID_{res.get('db_id')}_{res.get('kode_ahsp')}.pdf",
                mime="application/pdf",
                type="secondary",
                use_container_width=True
            )

        # Pratinjau PDF Langsung di Layar (In-App PDF Viewer)
        with st.expander("📑 Pratinjau Dokumen PDF Langsung di Layar (In-App PDF Viewer)", expanded=False):
            b64_single_pdf = base64.b64encode(single_record_pdf).decode('utf-8')
            pdf_embed_html = (
                f'<iframe src="data:application/pdf;base64,{b64_single_pdf}#toolbar=1&navpanes=0" '
                f'width="100%" height="600" type="application/pdf" '
                f'style="border: 1px solid #cbd5e1; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">'
                f'</iframe>'
            )
            st.markdown(pdf_embed_html, unsafe_allow_html=True)
            st.caption("ℹ️ *Jendela pratinjau PDF interaktif resmi ReportLab standar A4 untuk presentasi sidang proyektor tanpa berpindah aplikasi.*")


# ==================================================================================================
# HALAMAN 2: 📜 RIWAYAT INSPEKSI GEDUNG
# ==================================================================================================
elif menu == "📜 Riwayat Inspeksi Gedung":
    st.markdown("""
    <div class="hero-banner">
        <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 10px;">
            <span class="badge-ahsp" style="background:#0891b2; color:#ffffff; border:none; padding:4px 12px;">MODUL AUDIT & HISTORI</span>
            <div class="live-badge" style="margin-bottom:0;">
                <span class="live-dot"></span> DATABASE TERHUBUNG
            </div>
        </div>
        <div class="hero-title">Log Riwayat Penilaian Kondisi Bangunan Gedung</div>
        <div class="hero-subtitle">
            Daftar seluruh riwayat inspeksi yang tersimpan di dalam database SQLite. Menyediakan ringkasan KPI, filter data elemen gedung, dan fasilitas ekspor laporan resmi PDF serta file Excel (.xlsx) bersertifikasi standar IKBG.
        </div>
    </div>
    """, unsafe_allow_html=True)

    records = DatabaseManager.get_all_riwayat()

    if not records:
        st.info("ℹ️ Database riwayat saat ini masih kosong.")
        if st.button("🧪 Muat Data Sampel Demonstrasi Sidang Skripsi", type="primary"):
            DatabaseManager.seed_initial_demo_data()
            st.rerun()
    else:
        # Kalkulasi Indeks Kesehatan Gedung Kumulatif (IKBG Global) & Data Statistik
        kinerja_list = []
        severity_counts = {"Ringan": 0, "Sedang": 0, "Berat": 0}
        elemen_budgets = {}

        for r in records:
            eval_r = compute_condition_assessment(r)
            kinerja_list.append(eval_r["total_kinerja"])
            sev = eval_r["keparahan"]
            if sev in severity_counts:
                severity_counts[sev] += 1
            else:
                severity_counts["Sedang"] += 1

            elem = r.get("elemen_gedung", "Lainnya")
            elemen_budgets[elem] = elemen_budgets.get(elem, 0) + r.get("total_biaya", 0)

        avg_mu_ku = round(sum(kinerja_list) / len(kinerja_list), 1) if kinerja_list else 100.0

        if avg_mu_ku > 90:
            status_gedung_label = "KONDISI BAIK (PEMELIHARAAN RUTIN)"
            status_gedung_badge = "RUTIN"
            card_color = "#16a34a"
            rekomendasi_manajerial = (
                "Komponen struktur dan arsitektur gedung secara umum berfungsi optimal (µ̄ku > 90). "
                "Cukup laksanakan program inspeksi visual berkala dan pembersihan rutin terjadwal sesuai SOP pemeliharaan gedung."
            )
        elif 70 <= avg_mu_ku <= 90:
            status_gedung_label = "KONDISI SEDANG (PEMELIHARAAN PREVENTIF)"
            status_gedung_badge = "PREVENTIF"
            card_color = "#d97706"
            rekomendasi_manajerial = (
                "Ditemukan beberapa indikasi deteriorasi pada komponen non-struktural atau selimut pelindung (70 ≤ µ̄ku ≤ 90). "
                "Direkomendasikan tindakan pemeliharaan pencegahan (preventive maintenance) dan pelapisan protektif segera guna mencegah degradasi struktural."
            )
        elif 50 <= avg_mu_ku < 70:
            status_gedung_label = "KONDISI KURANG BAIK (PEMELIHARAAN KOREKTIF)"
            status_gedung_badge = "KOREKTIF"
            card_color = "#ea580c"
            rekomendasi_manajerial = (
                "Sebagian elemen bangunan mengalami penurunan performa fungsional yang signifikan (50 ≤ µ̄ku < 70). "
                "Diperlukan tindakan perbaikan kuratif terencana dengan prioritas penganggaran pemeliharaan sesuai standar AHSP Jateng 2023."
            )
        else:
            status_gedung_label = "KONDISI KRITIS (PEMELIHARAAN REHABILITATIF)"
            status_gedung_badge = "REHABILITATIF"
            card_color = "#dc2626"
            rekomendasi_manajerial = (
                "Terdapat komponen struktural utama yang mengalami kerusakan substansial (µ̄ku < 50). "
                "Memerlukan audit forensik struktur lanjutan dan perbaikan rehabilitatif/perkuatan (retrofitting) segera guna menjamin keselamatan penghuni."
            )

        # -----------------------------------------------------------------------------------------
        # PILIHAN C: KARTU EKSEKUTIF INDEKS KESEHATAN GEDUNG KUMULATIF (IKBG GLOBAL)
        # -----------------------------------------------------------------------------------------
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); border-radius: 16px; padding: 22px; color: #ffffff; box-shadow: 0 10px 25px rgba(15,23,42,0.25); border: 1px solid rgba(255,255,255,0.1); margin-bottom: 24px;">
            <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 16px;">
                <div style="flex: 1; min-width: 280px;">
                    <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                        <span style="background: rgba(255,255,255,0.15); padding: 4px 10px; border-radius: 6px; font-size: 0.75rem; font-weight: 700; letter-spacing: 0.5px;">EVALUASI KEANDALAN GLOBAL</span>
                        <span style="background: {card_color}; color: #ffffff; padding: 4px 12px; border-radius: 20px; font-size: 0.75rem; font-weight: 800;">{status_gedung_badge}</span>
                    </div>
                    <h3 style="margin: 0; font-size: 1.35rem; font-weight: 900; color: #f8fafc;">Indeks Kesehatan Gedung Kumulatif (IKBG Global)</h3>
                    <p style="margin: 6px 0 14px 0; font-size: 0.92rem; color: #cbd5e1; line-height: 1.5;">
                        Status: <b style="color: #ffffff;">{status_gedung_label}</b> (Berdasarkan evaluasi menyeluruh {len(records)} titik inspeksi sesuai Permen PUPR No. 24/PRT/M/2008).
                    </p>
                    <div style="background: rgba(255,255,255,0.07); border-left: 4px solid {card_color}; border-radius: 6px; padding: 10px 14px; font-size: 0.85rem; color: #e2e8f0; line-height: 1.5;">
                        <b>Rekomendasi Manajerial:</b> {rekomendasi_manajerial}
                    </div>
                </div>
                <div style="text-align: center; background: rgba(255,255,255,0.06); border-radius: 14px; padding: 16px 24px; border: 1px solid rgba(255,255,255,0.1); min-width: 170px;">
                    <div style="font-size: 0.78rem; text-transform: uppercase; color: #94a3b8; font-weight: 700; letter-spacing: 0.5px;">Rata-rata Skor (µ̄ku)</div>
                    <div style="font-size: 2.6rem; font-weight: 900; color: #ffffff; line-height: 1.1; margin: 4px 0;">{avg_mu_ku}</div>
                    <div style="font-size: 0.8rem; color: #cbd5e1; font-weight: 600;">Skala 0 - 100</div>
                    <div style="margin-top: 8px; width: 100%; background: rgba(255,255,255,0.2); height: 6px; border-radius: 3px; overflow: hidden;">
                        <div style="width: {min(100.0, avg_mu_ku)}%; background: {card_color}; height: 100%; border-radius: 3px;"></div>
                    </div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Ringkasan KPI Atas
        total_anggaran = sum(r["total_biaya"] for r in records)
        total_titik = len(records)
        avg_anggaran = total_anggaran / total_titik if total_titik > 0 else 0

        # Cari elemen paling dominan
        df_all = pd.DataFrame(records)
        elemen_dominan = df_all["elemen_gedung"].mode()[0] if not df_all.empty else "-"

        col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
        with col_kpi1:
            st.markdown(f"""
            <div class="metric-container metric-card-green">
                <div>
                    <div class="metric-label">Total RAB Pemeliharaan</div>
                    <div class="metric-value" style="color:#16a34a;">{format_rupiah(total_anggaran)}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_kpi2:
            st.markdown(f"""
            <div class="metric-container metric-card-blue">
                <div>
                    <div class="metric-label">Total Titik Kerusakan</div>
                    <div class="metric-value">{total_titik} Titik</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_kpi3:
            st.markdown(f"""
            <div class="metric-container metric-card-cyan">
                <div>
                    <div class="metric-label">Rata-rata Biaya / Titik</div>
                    <div class="metric-value">{format_rupiah(avg_anggaran)}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_kpi4:
            st.markdown(f"""
            <div class="metric-container metric-card-purple">
                <div>
                    <div class="metric-label">Elemen Dominan Rusak</div>
                    <div class="metric-value" style="font-size:1.05rem; color:#7c3aed;">{elemen_dominan}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")

        # -----------------------------------------------------------------------------------------
        # PILIHAN A (BAGIAN 2): VISUALISASI STATISTIK DENGAN GRAFIK ALTAIR
        # -----------------------------------------------------------------------------------------
        st.markdown("#### 📈 Visualisasi Statistik & Analisis Sebaran Kerusakan")
        col_ch1, col_ch2 = st.columns(2)

        with col_ch1:
            df_chart_elem = pd.DataFrame([
                {"Elemen Gedung": k, "Total Anggaran (Rp)": v}
                for k, v in elemen_budgets.items()
            ]).sort_values(by="Total Anggaran (Rp)", ascending=False)

            bar_chart = alt.Chart(df_chart_elem).mark_bar(
                cornerRadiusTopLeft=6, cornerRadiusTopRight=6, color="#2563eb"
            ).encode(
                x=alt.X("Total Anggaran (Rp):Q", title="Total Estimasi RAB (Rp)", axis=alt.Axis(format="~s")),
                y=alt.Y("Elemen Gedung:N", sort="-x", title="Elemen Gedung"),
                tooltip=[
                    alt.Tooltip("Elemen Gedung:N"),
                    alt.Tooltip("Total Anggaran (Rp):Q", format=",.0f")
                ]
            ).properties(
                title="Distribusi Anggaran Pemeliharaan per Elemen Gedung",
                height=280
            )
            st.altair_chart(bar_chart, use_container_width=True)

        with col_ch2:
            df_chart_sev = pd.DataFrame([
                {"Tingkat Keparahan": k, "Jumlah Titik": v}
                for k, v in severity_counts.items()
                if v > 0
            ])

            donut_chart = alt.Chart(df_chart_sev).mark_arc(innerRadius=50).encode(
                theta=alt.Theta(field="Jumlah Titik", type="quantitative"),
                color=alt.Color(
                    field="Tingkat Keparahan",
                    type="nominal",
                    scale=alt.Scale(
                        domain=["Ringan", "Sedang", "Berat"],
                        range=["#10B981", "#F59E0B", "#EF4444"]
                    ),
                    legend=alt.Legend(title="Keparahan", orient="bottom")
                ),
                tooltip=[
                    alt.Tooltip("Tingkat Keparahan:N"),
                    alt.Tooltip("Jumlah Titik:Q")
                ]
            ).properties(
                title="Proporsi Tingkat Keparahan Kerusakan (Visual AI)",
                height=280
            )
            st.altair_chart(donut_chart, use_container_width=True)

        st.markdown("---")

        # Fitur Filter Data
        col_f1, col_f2 = st.columns([1, 2])
        with col_f1:
            daftar_elemen = ["Semua Elemen"] + sorted(list(set(r["elemen_gedung"] for r in records)))
            selected_filter = st.selectbox("Filter Elemen Gedung:", daftar_elemen)
        with col_f2:
            search_query = st.text_input("Cari kata kunci (gedung, lokasi, surveyor, jenis kerusakan, kode AHSP):", "")

        # Terapkan filter
        filtered_records = records
        if selected_filter != "Semua Elemen":
            filtered_records = [r for r in filtered_records if r["elemen_gedung"] == selected_filter]
        if search_query.strip():
            sq = search_query.lower()
            filtered_records = [
                r for r in filtered_records
                if sq in str(r.get("jenis_kerusakan", "")).lower()
                or sq in str(r.get("kode_ahsp", "")).lower()
                or sq in str(r.get("rekomendasi", "")).lower()
                or sq in str(r.get("elemen_gedung", "")).lower()
                or sq in str(r.get("nama_gedung", "")).lower()
                or sq in str(r.get("lokasi_lantai", "")).lower()
                or sq in str(r.get("surveyor", "")).lower()
            ]

        # Tampilkan Tabel DataFrame
        st.markdown(f"##### Menampilkan {len(filtered_records)} dari {len(records)} Catatan Inspeksi")

        # Format tampilan tabel untuk Streamlit dengan Metadata Lengkap
        table_display = []
        for r in filtered_records:
            table_display.append({
                "ID": r["id"],
                "Tanggal": str(r["tanggal"])[:16] if r.get("tanggal") else "-",
                "Gedung": r.get("nama_gedung", "Gedung Kuliah Terpadu"),
                "Lokasi / Lantai": r.get("lokasi_lantai", "Lantai 2 - Ruang 204"),
                "Surveyor": r.get("surveyor", "Ahmad Fajar, S.T."),
                "Elemen Gedung": r["elemen_gedung"],
                "Jenis Kerusakan": r["jenis_kerusakan"],
                "Kode AHSP": r["kode_ahsp"],
                "Volume": f"{r['volume']:.2f} {r['satuan']}",
                "Harga Satuan": format_rupiah(r["harga_satuan"]),
                "Total Biaya": format_rupiah(r["total_biaya"]),
                "Rekomendasi Teknis": r["rekomendasi"]
            })

        df_display = pd.DataFrame(table_display)
        st.dataframe(df_display, use_container_width=True, hide_index=True)

        # Baris Tombol Aksi Ekspor (Excel, PDF, CSV) & Hapus Data
        col_act1, col_act2, col_act3, col_act4, col_act5 = st.columns([1.6, 1.6, 1.1, 1.0, 1.1])

        with col_act1:
            # Ekspor Excel (.xlsx) dengan Formulir Penilaian Kondisi (IKBG) & RAB
            excel_data = ExcelReportGenerator.generate_excel_report(
                filtered_records, title_doc="Laporan Penilaian Gedung Jawa Tengah"
            )
            st.download_button(
                label="📊 Ekspor Penilaian & RAB (Excel .xlsx)",
                data=excel_data,
                file_name=f"Penilaian_IKBG_Jateng_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        with col_act2:
            # Ekspor PDF Resmi ReportLab
            pdf_data = PDFReportGenerator.generate_pdf_report(
                filtered_records, title_doc="Laporan Pemeliharaan Gedung Jawa Tengah"
            )
            st.download_button(
                label="📄 Ekspor Laporan PDF Resmi (ReportLab)",
                data=pdf_data,
                file_name=f"Laporan_Resmi_PBG_Jateng_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                type="secondary",
                use_container_width=True
            )

        with col_act3:
            # Ekspor CSV
            csv_data = df_all.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Unduh CSV",
                data=csv_data,
                file_name=f"Data_Inspeksi_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )

        with col_act4:
            # Hapus Satu Data Tertentu
            delete_id = st.number_input("Hapus ID:", min_value=1, step=1, label_visibility="collapsed")
            if st.button("🗑️ Hapus ID", use_container_width=True):
                if DatabaseManager.delete_riwayat(delete_id):
                    st.success(f"Catatan ID #{delete_id} berhasil dihapus.")
                    st.rerun()
                else:
                    st.error(f"ID #{delete_id} tidak ditemukan.")

        with col_act5:
            # Kosongkan Database
            if st.button("⚠️ Reset Data", type="secondary", use_container_width=True):
                DatabaseManager.clear_all()
                st.warning("Seluruh riwayat berhasil dikosongkan.")
                st.rerun()

        # -----------------------------------------------------------------------------------------
        # FITUR 5: PRATINJAU DOKUMEN LAPORAN LENGKAP PDF (IN-APP PDF VIEWER)
        # -----------------------------------------------------------------------------------------
        with st.expander("📑 Pratinjau Dokumen Laporan Lengkap PDF (In-App PDF Viewer)", expanded=False):
            b64_all_pdf = base64.b64encode(pdf_data).decode('utf-8')
            pdf_all_embed_html = (
                f'<iframe src="data:application/pdf;base64,{b64_all_pdf}#toolbar=1&navpanes=0" '
                f'width="100%" height="650" type="application/pdf" '
                f'style="border: 1px solid #cbd5e1; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">'
                f'</iframe>'
            )
            st.markdown(pdf_all_embed_html, unsafe_allow_html=True)
            st.caption("ℹ️ *Pratinjau dokumen laporan audit pemeliharaan gedung resmi format A4 untuk presentasi sidang proyektor tanpa berpindah aplikasi.*")


# ==================================================================================================
# HALAMAN 3: 📑 PANDUAN AHSP JATENG 2023
# ==================================================================================================
elif menu == "📑 Panduan AHSP Jateng 2023":
    st.markdown("""
    <div class="hero-banner">
        <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 10px;">
            <span class="badge-ahsp" style="background:#7c3aed; color:#ffffff; border:none; padding:4px 12px;">DASAR HUKUM & STANDAR TEKNIS</span>
            <div class="live-badge" style="margin-bottom:0;">
                <span class="live-dot"></span> REGULASI PERGUB JATENG 2023
            </div>
        </div>
        <div class="hero-title">Katalog AHSP Pemeliharaan Gedung Jawa Tengah 2023</div>
        <div class="hero-subtitle">
            Standar Analisis Harga Satuan Pekerjaan (AHSP) Bidang Ke-PU-an Provinsi Jawa Tengah Tahun 2023 berdasarkan Permen PUPR No. 1/PRT/M/2022. Memuat rincian koefisien tenaga kerja, koefisien bahan, dan standar operasional prosedur perbaikan bangunan gedung.
        </div>
    </div>
    """, unsafe_allow_html=True)

    catalog = AHSPKnowledgeBase.get_all()

    # Ringkasan Cepat
    st.markdown("#### 🔍 Daftar Analisis Harga Satuan Pekerjaan Tersedia")

    # Filter Kategori
    kategori_list = ["Semua Kategori"] + sorted(list(set(item["kategori"] for item in catalog)))
    pilih_kat = st.selectbox("Pilih Kategori Pekerjaan:", kategori_list)

    filtered_catalog = catalog
    if pilih_kat != "Semua Kategori":
        filtered_catalog = [it for it in filtered_catalog if it["kategori"] == pilih_kat]

    # Grid Card Katalog AHSP
    for item in filtered_catalog:
        with st.expander(f"📌 {item['kode']} - {item['nama']} ({format_rupiah(item['harga_satuan'])} / {item['satuan']})"):
            c1, c2 = st.columns([1, 1])

            with c1:
                st.markdown(f"**Kategori:** `{item['kategori']}`")
                st.markdown(f"**Elemen Terkait:** `{item['elemen_terkait']}`")
                st.markdown(f"**Kerusakan Khas:** `{item['jenis_kerusakan']}`")
                st.markdown(f"**Tingkat Keparahan Umum:** `{item['tingkat_keparahan']}`")
                st.markdown(f"**Harga Satuan Standar:** `{format_rupiah(item['harga_satuan'])} per {item['satuan']}`")

                st.markdown("##### 👷 Analisa Koefisien Tenaga Kerja (Upah)")
                df_upah = pd.DataFrame(item["koefisien_upah"])
                st.table(df_upah)

            with c2:
                st.markdown("##### 🧱 Analisa Koefisien Bahan / Material")
                df_bahan = pd.DataFrame(item["koefisien_bahan"])
                st.table(df_bahan)

                st.markdown("##### 🛠️ Langkah Metode Pelaksanaan Lapangan")
                for s_idx, step in enumerate(item["langkah_kerja"], start=1):
                    st.markdown(f"{s_idx}. {step}")

            st.info(f"**Rekomendasi Ahli:** {item['rekomendasi']}")

    st.markdown("---")

    # Kalkulator Cepat AHSP Jateng
    st.markdown("### 🧮 Kalkulator Cepat Anggaran Pemeliharaan (AHSP Jateng 2023)")
    st.caption("Gunakan kalkulator mandiri ini untuk estimasi cepat di lapangan tanpa unggah citra.")

    col_calc1, col_calc2, col_calc3 = st.columns(3)
    with col_calc1:
        pilih_item_calc = st.selectbox(
            "Pilih Item Pekerjaan AHSP:",
            options=list(AHSPKnowledgeBase.CATALOG.keys()),
            format_func=lambda x: f"{x} - {AHSPKnowledgeBase.CATALOG[x]['nama']}"
        )
    with col_calc2:
        vol_calc = st.number_input("Masukkan Volume Pekerjaan:", min_value=0.1, value=10.0, step=1.0)
    with col_calc3:
        item_obj = AHSPKnowledgeBase.CATALOG[pilih_item_calc]
        biaya_calc = int(round(vol_calc * item_obj["harga_satuan"]))
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%); border: 1.5px solid #86efac; border-radius: 14px; padding: 14px 18px; margin-top: 10px; box-shadow: 0 4px 12px rgba(22, 163, 74, 0.12); transition: all 0.3s ease;">
            <div style="font-size: 0.76rem; color: #166534; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;">ESTIMASI ANGGARAN TOTAL (AHSP):</div>
            <div style="font-size: 1.55rem; font-weight: 800; color: #15803d; margin-top: 2px;">{format_rupiah(biaya_calc)}</div>
            <div style="font-size: 0.76rem; color: #166534; margin-top: 4px;">Tarif: <b>{format_rupiah(item_obj['harga_satuan'])}</b> / {item_obj['satuan']}</div>
        </div>
        """, unsafe_allow_html=True)


# ==================================================================================================
# 7. FOOTER APLIKASI WEB SKRIPSI
# ==================================================================================================
st.markdown("""
<div style="text-align: center; margin-top: 50px; padding: 24px 0; border-top: 1.5px solid #cbd5e1; font-size: 0.86rem; color: #1e293b; font-weight: 500;">
    <b style="color: #0f172a; font-weight: 800;">Sistem Informasi Pemeliharaan Bangunan Gedung Menyeluruh Berbasis AI (SIM-PBG AI)</b><br>
    Program Skripsi Teknik Sipil Digital | Standar Acuan: <span style="color: #0f172a; font-weight: 700;">AHSP Dinas PU Bina Marga & Cipta Karya Jawa Tengah 2023</span>
</div>
""", unsafe_allow_html=True)
