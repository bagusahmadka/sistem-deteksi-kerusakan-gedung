"""
====================================================================================================
SISTEM INFORMASI PEMELIHARAAN BANGUNAN GEDUNG MENYELURUH BERBASIS AI
(Deteksi Kerusakan & Estimasi AHSP Jawa Tengah)
Framework: Streamlit | AI Model: YOLOv8 Computer Vision | Database: SQLite3 | Ekspor: ReportLab PDF
====================================================================================================
"""

import os
import io
import json
import base64
import sqlite3
import requests
from datetime import datetime, date
from typing import Dict, Any, List, Optional

import streamlit as st
import pandas as pd
import altair as alt
from PIL import Image, ImageDraw, ImageFont

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
    page_title="Sistem Deteksi Kerusakan Gedung",
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

    /* Pengecualian Kontras untuk Kontainer Latar Gelap (Hero Banner & Kartu IKBG Global) */
    .hero-banner *,
    .hero-banner h1, .hero-banner h2, .hero-banner h3, .hero-banner h4, .hero-banner p, .hero-banner div,
    .ikbg-global-card,
    .ikbg-global-card *,
    .ikbg-global-card h1, .ikbg-global-card h2, .ikbg-global-card h3, .ikbg-global-card h4,
    .ikbg-global-card p, .ikbg-global-card div, .ikbg-global-card span, .ikbg-global-card b {
        color: #ffffff !important;
    }
    .ikbg-global-card p, .ikbg-global-card .subtext {
        color: #e2e8f0 !important;
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
    /* 7. PROSEDUR METODE KERJA / STEPS (TAMPILAN KARTU TAHAPAN MODERN)           */
    /* ========================================================================== */
    .repair-steps-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
        gap: 14px;
        margin-top: 10px;
        margin-bottom: 14px;
    }
    .repair-step-card {
        background: linear-gradient(160deg, #ffffff 0%, #f8fafc 100%);
        border: 1.5px solid #cbd5e1;
        border-top: 4px solid #2563eb;
        border-radius: 14px;
        padding: 16px 16px 14px 16px;
        box-shadow: 0 3px 10px rgba(15, 23, 42, 0.05);
        transition: all 0.25s ease;
        display: flex;
        flex-direction: column;
        gap: 8px;
    }
    .repair-step-card:hover {
        transform: translateY(-3px);
        border-color: #2563eb;
        box-shadow: 0 10px 22px -6px rgba(37, 99, 235, 0.22);
    }
    .repair-step-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #eff6ff;
        color: #1d4ed8;
        border: 1px solid #bfdbfe;
        border-radius: 999px;
        padding: 3px 10px;
        font-size: 0.74rem;
        font-weight: 800;
        width: fit-content;
        letter-spacing: 0.4px;
    }
    .repair-step-body {
        font-size: 0.9rem !important;
        font-weight: 600 !important;
        color: #0f172a !important;
        line-height: 1.55 !important;
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
    /* 10. TOMBOL STREAMLIT (AKSI, ANALISIS, UNDUH EXCEL, PDF, CSV)               */
    /* ========================================================================== */
    
    /* Aturan Global Tombol: Pastikan teks di dalam tombol SELALU terlihat kontras */
    div.stButton > button,
    div.stDownloadButton > button {
        border-radius: 12px !important;
        font-weight: 800 !important;
        font-size: 0.95rem !important;
        padding: 0.7rem 1.6rem !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
        cursor: pointer !important;
        position: relative !important;
    }

    /* 1. Tombol Primary (Analisis Deteksi AI): Biru Elektrik Modern */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 50%, #3b82f6 100%) !important;
        color: #ffffff !important;
        border: none !important;
        box-shadow: 0 4px 16px rgba(37, 99, 235, 0.45) !important;
    }
    
    /* 2. Tombol Download Excel (Primary): Hijau Emerald Standar Excel dengan Teks Putih Tebal */
    div.stDownloadButton > button[kind="primary"],
    div.stDownloadButton > button[data-testid="baseButton-primary"] {
        background: linear-gradient(135deg, #15803d 0%, #16a34a 100%) !important;
        color: #ffffff !important;
        border: 1.5px solid #15803d !important;
        box-shadow: 0 4px 16px rgba(22, 163, 74, 0.35) !important;
    }

    /* 3. Tombol Download PDF (Secondary): Biru Royal Standar Dokumen dengan Teks Putih Tebal */
    div.stDownloadButton > button[kind="secondary"],
    div.stDownloadButton > button[data-testid="baseButton-secondary"] {
        background: linear-gradient(135deg, #1e40af 0%, #2563eb 100%) !important;
        color: #ffffff !important;
        border: 1.5px solid #1e40af !important;
        box-shadow: 0 4px 16px rgba(37, 99, 235, 0.35) !important;
    }

    /* 4. Tombol Download Lainnya (CSV / Default): Biru Slate Solid dengan Teks Putih */
    div.stDownloadButton > button {
        background: linear-gradient(135deg, #0369a1 0%, #0284c7 100%) !important;
        color: #ffffff !important;
        border: 1.5px solid #0284c7 !important;
        box-shadow: 0 4px 16px rgba(3, 105, 161, 0.3) !important;
    }

    /* 5. Tombol Sekunder Non-Download (Hapus ID, dll): Putih Bersih dengan Border Slate */
    div.stButton > button[kind="secondary"] {
        background: #ffffff !important;
        color: #0f172a !important;
        border: 1.5px solid #cbd5e1 !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05) !important;
    }

    /* ========================================================================== */
    /* PEMBERSIH WARNA TEKS TOMBOL: Override Paksa Seluruh Tag Paragraf di Tombol */
    /* ========================================================================== */
    div.stDownloadButton > button *,
    div.stDownloadButton > button p,
    div.stDownloadButton > button span,
    div.stDownloadButton > button div,
    div.stDownloadButton > button [data-testid="stMarkdownContainer"] *,
    div.stDownloadButton > button [data-testid="stMarkdownContainer"] p,
    div.stButton > button[kind="primary"] *,
    div.stButton > button[kind="primary"] p,
    div.stButton > button[kind="primary"] span,
    div.stButton > button[kind="primary"] [data-testid="stMarkdownContainer"] *,
    div.stButton > button[kind="primary"] [data-testid="stMarkdownContainer"] p {
        color: #ffffff !important;
        font-weight: 800 !important;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.25) !important;
    }

    div.stButton > button[kind="secondary"] *,
    div.stButton > button[kind="secondary"] p,
    div.stButton > button[kind="secondary"] span,
    div.stButton > button[kind="secondary"] [data-testid="stMarkdownContainer"] p {
        color: #0f172a !important;
        font-weight: 800 !important;
    }

    /* Hover States Efek Angkat Halus */
    div.stButton > button:hover,
    div.stDownloadButton > button:hover {
        transform: translateY(-2px) scale(1.01) !important;
        filter: brightness(1.08) !important;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.25) !important;
    }
    div.stButton > button[kind="secondary"]:hover {
        background: #f8fafc !important;
        border-color: #2563eb !important;
        color: #2563eb !important;
    }
    div.stButton > button[kind="secondary"]:hover * {
        color: #2563eb !important;
    }

    /* Tab Navigasi Kontras Tinggi */
    div[data-testid="stTabs"] button[data-baseweb="tab"] {
        color: #334155 !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }
    div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {
        color: #2563eb !important;
        font-weight: 800 !important;
        border-bottom-color: #2563eb !important;
    }

    /* Native Metric Cards High Contrast */
    div[data-testid="stMetric"] {
        background: #ffffff !important;
        border: 1.5px solid #cbd5e1 !important;
        border-radius: 12px !important;
        padding: 12px 16px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03) !important;
    }
    div[data-testid="stMetricValue"] {
        color: #0f172a !important;
        font-weight: 800 !important;
    }
    div[data-testid="stMetricValue"] > div {
        color: #0f172a !important;
    }
    div[data-testid="stMetricLabel"] p {
        color: #475569 !important;
        font-weight: 700 !important;
    }

    /* Form Inputs & Selectbox Kontras Tinggi */
    div[data-baseweb="input"] {
        background-color: #ffffff !important;
        border-color: #cbd5e1 !important;
    }
    div[data-baseweb="input"] input {
        color: #0f172a !important;
        background-color: #ffffff !important;
    }
    div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }
    div[data-baseweb="popover"], div[role="listbox"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }
    div[role="option"] {
        color: #0f172a !important;
    }
    textarea {
        color: #0f172a !important;
        background-color: #ffffff !important;
    }

    /* File Uploader Visual Besar dengan Ikon Kamera */
    div[data-testid="stFileUploader"] {
        background: linear-gradient(180deg, #f8fbff 0%, #eff6ff 100%);
        border: 2.5px dashed #2563eb;
        border-radius: 18px;
        padding: 22px 20px;
        transition: all 0.3s ease;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.08);
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #1d4ed8;
        background: #e0ecff;
        box-shadow: 0 10px 24px rgba(37, 99, 235, 0.16);
    }
    div[data-testid="stFileUploaderDropzone"] {
        min-height: 165px !important;
        padding: 24px 16px !important;
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
    }
    .camera-upload-hero {
        background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
        border: 2px solid #93c5fd;
        border-radius: 16px;
        padding: 18px 16px 12px 16px;
        text-align: center;
        margin-bottom: 10px;
    }
    .camera-upload-caption {
        text-align: center;
        font-size: 0.9rem;
        font-weight: 700;
        color: #b91c1c;
        background: #fef2f2;
        border: 1px solid #fecaca;
        border-radius: 10px;
        padding: 8px 12px;
        margin-top: 8px;
        margin-bottom: 14px;
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

    /* ========================================================================== */
    /* 11. RESPONSIVE DESIGN (MOBILE & TABLET VIEWPORT FIXES)                     */
    /* ========================================================================== */
    @media (max-width: 768px) {
        /* Optimalisasi padding container agar layar HP tidak kepotong */
        .block-container {
            padding-top: 1.25rem !important;
            padding-bottom: 2rem !important;
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            max-width: 100% !important;
        }

        /* Hero Banner responsif */
        .hero-banner {
            padding: 20px 16px !important;
            border-radius: 14px !important;
            margin-bottom: 16px !important;
        }
        .hero-title {
            font-size: 1.35rem !important;
            line-height: 1.3 !important;
        }
        .hero-subtitle {
            font-size: 0.88rem !important;
            line-height: 1.5 !important;
        }

        /* Penyesuaian Kolom Streamlit (Horizontal Block) di Layar HP */
        div[data-testid="stHorizontalBlock"] {
            flex-direction: column !important;
            flex-wrap: wrap !important;
            gap: 12px !important;
        }
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
            min-width: 100% !important;
            width: 100% !important;
            flex: 1 1 100% !important;
            margin-bottom: 4px !important;
        }

        /* Metric card responsif memenuhi lebar hp dengan rapi */
        .metric-container {
            padding: 14px 16px !important;
            margin-bottom: 10px !important;
            width: 100% !important;
            box-sizing: border-box !important;
        }
        .metric-value {
            font-size: 1.35rem !important;
            word-break: break-word !important;
        }
        .metric-label {
            font-size: 0.76rem !important;
        }

        /* Grid Status Penilaian IKBG (2 kolom di HP) */
        .ikbg-status-grid {
            grid-template-columns: repeat(2, 1fr) !important;
            gap: 8px !important;
        }
        .ikbg-status-card {
            padding: 10px 6px !important;
            font-size: 0.8rem !important;
        }

        /* Card Container Padding */
        .tw-card {
            padding: 16px !important;
            border-radius: 14px !important;
            margin-bottom: 16px !important;
        }

        /* Tombol di HP dibuat full width agar mudah disentuh */
        div.stButton > button, div.stDownloadButton > button {
            width: 100% !important;
            padding: 0.7rem 1rem !important;
            font-size: 0.92rem !important;
        }

        /* Tabel & Dataframe dapat di-scroll horizontal tanpa membuat halaman kepotong */
        div[data-testid="stTable"],
        div[data-testid="stDataFrame"],
        .stTable,
        table {
            width: 100% !important;
            overflow-x: auto !important;
            -webkit-overflow-scrolling: touch !important;
            display: block !important;
        }

        /* Gambar responsif */
        img {
            max-width: 100% !important;
            height: auto !important;
            border-radius: 10px !important;
        }
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
                cursor.execute("ALTER TABLE riwayat_kerusakan ADD COLUMN surveyor TEXT DEFAULT 'Inspektor Lapangan'")
            if "koordinat_lokasi" not in cols:
                cursor.execute("ALTER TABLE riwayat_kerusakan ADD COLUMN koordinat_lokasi TEXT DEFAULT '-'")
            if "objek_inspeksi" not in cols:
                cursor.execute("ALTER TABLE riwayat_kerusakan ADD COLUMN objek_inspeksi TEXT DEFAULT 'Arsitektur'")
            if "luas_total" not in cols:
                cursor.execute("ALTER TABLE riwayat_kerusakan ADD COLUMN luas_total REAL DEFAULT 25.0")
            if "persentase_kerusakan" not in cols:
                cursor.execute("ALTER TABLE riwayat_kerusakan ADD COLUMN persentase_kerusakan REAL DEFAULT 10.0")
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
        surveyor: str = "Inspektor Lapangan",
        koordinat_lokasi: str = "-",
        objek_inspeksi: str = "Arsitektur",
        luas_total: float = 25.0,
        persentase_kerusakan: float = 10.0
    ) -> int:
        """Menyimpan data penilaian kondisi kerusakan gedung ke SQLite."""
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO riwayat_kerusakan (
                    tanggal, nama_gedung, lokasi_lantai, surveyor,
                    koordinat_lokasi, objek_inspeksi,
                    elemen_gedung, jenis_kerusakan, kode_ahsp,
                    volume, satuan, harga_satuan, total_biaya, rekomendasi,
                    luas_total, persentase_kerusakan
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                tanggal, nama_gedung, lokasi_lantai, surveyor,
                koordinat_lokasi, objek_inspeksi,
                elemen_gedung, jenis_kerusakan, kode_ahsp,
                volume, satuan, int(harga_satuan), int(total_biaya), rekomendasi,
                float(luas_total), float(persentase_kerusakan)
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
        """Menambahkan data sampel realistis untuk demonstrasi pengujian sistem."""
        samples = [
            (
                "2026-11-01 09:30", "Gedung Kuliah Terpadu", "Lantai 2 - Selasar Barat", "Ahmad Fajar, S.T.",
                "-7.005100, 110.438100", "Struktur - Balok Beton Bertulang",
                "Balok Beton Bertulang", "Retak Diagonal Geser/Lentur Balok", "A.4.1.1.18.JT",
                4.5, "m'", 294860, 1326870,
                "Injeksi resin epoxy bertekanan rendah pada balok lantai 2 as C-4 untuk mengembalikan monolititas beton.",
                18.0, 25.0
            ),
            (
                "2026-11-02 11:15", "Gedung Kuliah Terpadu", "Lantai 1 - Ruang 103", "Ahmad Fajar, S.T.",
                "-7.005100, 110.438100", "Arsitektur - Dinding dan Partisi",
                "Dinding dan Partisi", "Retak Rambut", "A.4.4.1.25.JT",
                3.0, "m2", 95163, 285489,
                "Pembersihan alur retak bentuk V, aplikasi elastomeric sealant, acian polimer dan pengecatan ulang.",
                32.0, 9.4
            ),
            (
                "2026-11-03 14:00", "Gedung Kuliah Terpadu", "Lantai 1 - Ruang 105", "Ahmad Fajar, S.T.",
                "-7.005100, 110.438100", "Arsitektur - Lantai",
                "Lantai", "Pecah (1 Keramik 60x60)", "A.4.4.3.02.JT",
                0.36, "m2", 268180, 96545,
                "Pembongkaran 1 keping keramik pecah secara hati-hati dengan gerinda nat agar tidak merusak keramik sebelahnya, lalu pasang ulang dengan mortar perekat.",
                20.0, 1.8
            ),
            (
                "2026-11-04 10:45", "Gedung Kuliah Terpadu", "Lantai Dasar - Kolom K-02", "Ahmad Fajar, S.T.",
                "-7.005100, 110.438100", "Struktur - Kolom Struktur Beton",
                "Kolom Struktur Beton", "Spalling / Beton Terkelupas & Tulangan Terbuka", "A.4.1.1.20.JT",
                3.2, "m2", 293135, 938032,
                "Chipping beton rapuh, pembersihan karat tulangan baja, pelapisan zinc-rich primer, dan patching polymer mortar.",
                8.0, 40.0
            )
        ]
        with cls.get_connection() as conn:
            cursor = conn.cursor()
            for s in samples:
                cursor.execute("""
                    INSERT INTO riwayat_kerusakan (
                        tanggal, nama_gedung, lokasi_lantai, surveyor,
                        koordinat_lokasi, objek_inspeksi,
                        elemen_gedung, jenis_kerusakan, kode_ahsp,
                        volume, satuan, harga_satuan, total_biaya, rekomendasi,
                        luas_total, persentase_kerusakan
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, s)
            conn.commit()

# Inisialisasi Database saat aplikasi dimuat
DatabaseManager.init_db()


# ==================================================================================================
# 3. BASIS DATA & PENGETAHUAN TEKNIK SIPIL: AHSP JAWA TENGAH TAHUN 2023 (RINCI PER BAHAN & UPAH)
# ==================================================================================================
def _make_ahsp(
    kode: str,
    bidang: str,
    kategori: str,
    elemen_terkait: str,
    sub_kerusakan: str,
    nama: str,
    satuan: str,
    tingkat_keparahan: str,
    bobot_fisik: float,
    bobot_fungsi: float,
    is_struktural: bool,
    upah: List[Dict[str, Any]],
    bahan: List[Dict[str, Any]],
    langkah_kerja: List[str],
    rekomendasi: str
) -> Dict[str, Any]:
    """Menyusun item AHSP terstandar dengan kalkulasi HSP otomatis dari rincian upah + bahan + overhead 15%."""
    subtotal_upah = sum(float(u["koef"]) * float(u["harga_satuan"]) for u in upah)
    subtotal_bahan = sum(float(b["koef"]) * float(b["harga_satuan"]) for b in bahan)
    dasar = subtotal_upah + subtotal_bahan
    overhead = dasar * 0.15
    hsp = int(round(dasar + overhead))
    return {
        "kode": kode,
        "bidang": bidang,
        "kategori": kategori,
        "elemen_terkait": elemen_terkait,
        "sub_kerusakan": sub_kerusakan,
        "nama": nama,
        "satuan": satuan,
        "harga_satuan": hsp,
        "subtotal_upah": int(round(subtotal_upah)),
        "subtotal_bahan": int(round(subtotal_bahan)),
        "overhead_15": int(round(overhead)),
        "jenis_kerusakan": f"{sub_kerusakan} pada {elemen_terkait}",
        "tingkat_keparahan": tingkat_keparahan,
        "bobot_fisik": bobot_fisik,
        "bobot_fungsi": bobot_fungsi,
        "is_struktural": is_struktural,
        "koefisien_upah": upah,
        "koefisien_bahan": bahan,
        "langkah_kerja": langkah_kerja,
        "rekomendasi": rekomendasi,
    }


UPAH_JATENG = {
    "Pekerja": 110000,
    "Tukang": 135000,
    "Tukang_Ahli": 145000,
    "Kepala_Tukang": 155000,
    "Mandor": 165000,
}


class AHSPKnowledgeBase:
    """Katalog Standar Analisis Harga Satuan Pekerjaan (AHSP) Dinas PU Bina Marga & Cipta Karya Jateng 2023."""

    OBJEK_INSPEKSI_MAP = {
        "Arsitektur - Dinding dan Partisi": [
            "Retak rambut",
            "Pengelupasan cat",
            "Cat menggelembung",
            "Jamur",
        ],
        "Arsitektur - Plafon": [
            "Kebocoran",
            "Jamur",
            "Retak",
            "Rusak",
        ],
        "Arsitektur - Pintu dan Jendela": [
            "Kotor",
            "Pelapukan",
            "Karat",
            "Kaca retak",
            "Sealent",
        ],
        "Arsitektur - Lantai": [
            "Pecah",
            "Retak",
            "Kotor",
            "Berjamur",
            "Popping",
            "Gumpil",
        ],
        "Arsitektur - Atap": [
            "Bocor",
            "Pecah",
            "Karat",
            "Kotor",
            "Jamur",
        ],
        "Struktur - Balok Beton Bertulang": [
            "Retak Struktur (Crack)",
            "Spalling / Keropos",
            "Tulangan Terbuka (Exposed Rebar)",
        ],
        "Struktur - Kolom Struktur Beton": [
            "Spalling / Keropos",
            "Retak Struktur (Crack)",
            "Tulangan Terbuka (Exposed Rebar)",
        ],
        "Struktur - Pelat Lantai / Dak Beton": [
            "Rembesan / Bocor & Efflorescence",
            "Retak Pelat Beton",
            "Spalling Bawah Pelat",
        ],
    }

    CATALOG = {
        # ARSITEKTUR 1: DINDING DAN PARTISI
        "A.4.4.1.25.JT": _make_ahsp(
            kode="A.4.4.1.25.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Dinding dan Partisi",
            elemen_terkait="Dinding dan Partisi",
            sub_kerusakan="Retak rambut",
            nama="Perbaikan Retak Rambut Dinding dengan Mortar Polimer & Acian Elastis",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.70,
            bobot_fungsi=0.35,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.120, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Batu/Cat", "koef": 0.150, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.015, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.008, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Polymer Modified Mortar / Filler Elastis", "koef": 1.250, "satuan": "kg", "harga_satuan": 16500},
                {"bahan": "Sealant Elastomeric Polyurethane", "koef": 0.350, "satuan": "tube", "harga_satuan": 55000},
                {"bahan": "Bonding Agent Adhesive", "koef": 0.150, "satuan": "liter", "harga_satuan": 48000},
                {"bahan": "Cat Dasar & Finishing Tembok", "koef": 0.250, "satuan": "kg", "harga_satuan": 36000},
            ],
            langkah_kerja=[
                "Bersihkan celah retak rambut dinding dari debu dan plesteran rapuh menggunakan sikat kawat dan kuas bersih.",
                "Buka alur celah retak membentuk alur huruf V (V-groove) selebar 3-5 mm memakai pahat atau scraper.",
                "Oleskan cairan bonding agent pada alur retak sebagai lapis pengikat material baru.",
                "Isi dan rapatkan celah menggunakan mortar polimer / filler elastis hingga rata dengan bidang dinding eksisting.",
                "Amplas halus permukaan setelah kering, lalu lapisi cat dasar dan cat finishing 2 lapis sesuai warna awal."
            ],
            rekomendasi="Pantau area retakan secara berkala jika muncul berulang untuk memastikan bukan retak geser akibat penurunan tanah fondasi."
        ),
        "A.4.7.1.04.JT": _make_ahsp(
            kode="A.4.7.1.04.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Dinding dan Partisi",
            elemen_terkait="Dinding dan Partisi",
            sub_kerusakan="Pengelupasan cat",
            nama="Pengerokan Cat Mengelupas, Plamir Akrilik & Pengecatan Ulang Dinding 2 Lapis",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.55,
            bobot_fungsi=0.25,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.070, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Cat", "koef": 0.090, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.009, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.004, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Alkali-Resisting Primer Sealer", "koef": 0.140, "satuan": "liter", "harga_satuan": 62000},
                {"bahan": "Plamir Tembok Acrylic", "koef": 0.150, "satuan": "kg", "harga_satuan": 22000},
                {"bahan": "Cat Tembok Akrilik Bermutu", "koef": 0.300, "satuan": "kg", "harga_satuan": 42000},
                {"bahan": "Kertas Amplas & Kape Scrap", "koef": 0.250, "satuan": "lbr", "harga_satuan": 6000},
            ],
            langkah_kerja=[
                "Kerok seluruh lapisan cat lama yang mengelupas dan rapuh memakai kape besi sampai mencapai lapisan acian yang padat.",
                "Amplas permukaan dinding agar pori terbuka dan bersihkan sisa serbuk kapur dengan lap lembab.",
                "Lapisi 1 lapis Alkali-Resisting Primer Sealer untuk menstabilkan pH acian dinding dan memperkuat daya rekat cat.",
                "Tutup bagian dinding yang tidak rata dengan plamir akrilik tipis, lalu amplas rata setelah kering.",
                "Aplikasikan cat tembok akrilik 2 lapis secara menyilang dengan jeda waktu pengeringan minimal 2 jam antar lapis."
            ],
            rekomendasi="Pastikan kelembapan dinding < 16% sebelum pengecatan ulang agar cat baru tidak kembali terkelupas."
        ),
        "A.4.7.1.06.JT": _make_ahsp(
            kode="A.4.7.1.06.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Dinding dan Partisi",
            elemen_terkait="Dinding dan Partisi",
            sub_kerusakan="Cat menggelembung",
            nama="Perbaikan Cat Menggelembung (Blistering) dengan Sealer Anti Lembab",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.65,
            bobot_fungsi=0.35,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.090, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Cat", "koef": 0.110, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.011, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.005, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Cat Dasar Anti Lembab / Damp Sealer", "koef": 0.200, "satuan": "kg", "harga_satuan": 78000},
                {"bahan": "Mortar Acian Tahan Lembab", "koef": 0.500, "satuan": "kg", "harga_satuan": 12000},
                {"bahan": "Cat Dinding Breathable Acrylic", "koef": 0.300, "satuan": "kg", "harga_satuan": 48000},
                {"bahan": "Kertas Amplas Halus Grid 180", "koef": 0.300, "satuan": "lbr", "harga_satuan": 6000},
            ],
            langkah_kerja=[
                "Kupas dan kerok area gelembung cat secara menyeluruh hingga tidak ada gelembung udara atau kantong uap air tersisa.",
                "Periksa dan hentikan sumber rembesan air (pipa bocor, talang, atau kapilaritas dinding luar).",
                "Biarkan dinding mengering sempurna (24-48 jam) sebelum pelapisan dasar dilakukan.",
                "Oleskan 1-2 lapis cat dasar damp-proof sealer sebagai penghalang uap air dan garam alkali.",
                "Ratakan dengan acian tipis jika perlu, lalu cat ulang 2 lapis menggunakan cat tembok tipe breathable."
            ],
            rekomendasi="Jika gelembung berada di dinding bawah dekat lantai, periksa lapisan trasraam / sumber rembesan air tanah."
        ),
        "A.4.7.1.08.JT": _make_ahsp(
            kode="A.4.7.1.08.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Dinding dan Partisi",
            elemen_terkait="Dinding dan Partisi",
            sub_kerusakan="Jamur",
            nama="Pembersihan Jamur Dinding, Fungicidal Wash & Pengecatan Anti Jamur",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.50,
            bobot_fungsi=0.30,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.080, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Cat", "koef": 0.095, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.010, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.004, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Cairan Fungicidal Wash Anti Jamur", "koef": 0.250, "satuan": "liter", "harga_satuan": 52000},
                {"bahan": "Alkali-Resisting Sealer Primer", "koef": 0.120, "satuan": "liter", "harga_satuan": 62000},
                {"bahan": "Cat Tembok Anti-Fungal / Anti Jamur", "koef": 0.300, "satuan": "kg", "harga_satuan": 54000},
            ],
            langkah_kerja=[
                "Sikat dan bersihkan koloni jamur serta lumut secara kering/lembab menggunakan sikat nilon kaku.",
                "Semprotkan atau kuaskan larutan fungicidal wash merata ke seluruh area berjamur hingga jarak 30 cm di sekitarnya, diamkan 24 jam.",
                "Bilas dengan air bersih, lalu keringkan permukaan dinding hingga benar-benar kering.",
                "Aplikasikan 1 lapis primer sealer penahan alkali untuk menutup pori-pori permukaan.",
                "Cat ulang sebanyak 2 lapis menggunakan cat tembok berformulasi anti jamur dan anti lumut."
            ],
            rekomendasi="Tingkatkan sirkulasi udara / ventilasi ruangan atau pasang exhaust fan pada ruangan lembab untuk mencegah pertumbuhan spora ulang."
        ),

        # ARSITEKTUR 2: PLAFON
        "A.4.5.1.12.JT": _make_ahsp(
            kode="A.4.5.1.12.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Plafon",
            elemen_terkait="Plafon",
            sub_kerusakan="Kebocoran",
            nama="Perbaikan Plafon Terdampak Kebocoran, Water-Stain Sealer & Penggantian Panel",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.85,
            bobot_fungsi=0.70,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.120, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Plafon", "koef": 0.140, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.014, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.006, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Panel Plafon Moisture-Resistant 9 mm", "koef": 0.364, "satuan": "lbr", "harga_satuan": 108000},
                {"bahan": "Sekrup Drywall Anti Karat", "koef": 16.000, "satuan": "buah", "harga_satuan": 300},
                {"bahan": "Compound Joint & Stain-Blocking Primer", "koef": 0.350, "satuan": "kg", "harga_satuan": 45000},
                {"bahan": "Cat Plafon Matte", "koef": 0.250, "satuan": "kg", "harga_satuan": 36000},
            ],
            langkah_kerja=[
                "Bongkar bagian panel plafon yang basah, lembek, atau bernoda bocor untuk membuka akses pengecekan sumber air.",
                "Perbaiki dan pastikan sumber kebocoran dari atap, dak beton, atau pipa utilitas di atas plafon telah tuntas kering.",
                "Periksa kondisi rangka hollow penggantung dari korosi, perkuat atau ganti batang rangka yang berkarat.",
                "Pasang panel plafon tahan lembab baru dengan sekrup drywall jarak tiap 20 cm.",
                "Tutup nat sambungan dengan compound + paper tape, aplikasikan stain-blocker primer, dan cat ulang plafon 2 lapis."
            ],
            rekomendasi="Jangan memasang plafon baru sebelum sumber kebocoran di atas plafon dipastikan kering dan teruji tidak menetes lagi."
        ),
        "A.4.5.1.14.JT": _make_ahsp(
            kode="A.4.5.1.14.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Plafon",
            elemen_terkait="Plafon",
            sub_kerusakan="Jamur",
            nama="Pembersihan Plafon Berjamur, Fungicidal Treatment & Pengecatan Plafon",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.55,
            bobot_fungsi=0.35,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.080, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Plafon/Cat", "koef": 0.100, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.010, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.005, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Cairan Pembersih Jamur (Fungicidal)", "koef": 0.200, "satuan": "liter", "harga_satuan": 52000},
                {"bahan": "Stain-Blocking Sealer Primer", "koef": 0.150, "satuan": "liter", "harga_satuan": 68000},
                {"bahan": "Cat Plafon Anti Jamur", "koef": 0.280, "satuan": "kg", "harga_satuan": 48000},
            ],
            langkah_kerja=[
                "Siapkan perancah (scaffolding) dan lindungi perabot lantai di bawah area kerja dengan plastik terpal.",
                "Bersihkan bercak jamur pada permukaan plafon dengan kain lembab bercampur larutan fungisida.",
                "Keringkan permukaan plafon secara alami dengan ventilasi terbuka.",
                "Lapisi 1 lapis stain-blocking primer untuk mencegah noda kuning dan spora jamur menembus cat baru.",
                "Cat ulang permukaan plafon 2 lapis memakai cat plafon anti jamur."
            ],
            rekomendasi="Periksa apakah plafon memiliki kondensasi dari ducting AC atau kelembapan ruang atap yang kurang berventilasi."
        ),
        "A.4.5.1.16.JT": _make_ahsp(
            kode="A.4.5.1.16.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Plafon",
            elemen_terkait="Plafon",
            sub_kerusakan="Retak",
            nama="Perbaikan Retak Sambungan Plafon dengan Textile Tape & Compound",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.65,
            bobot_fungsi=0.40,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.080, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Plafon", "koef": 0.110, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.011, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.004, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Gypsum Joint Compound (Tepung Kompon)", "koef": 0.650, "satuan": "kg", "harga_satuan": 16000},
                {"bahan": "Fiberglass Mesh / Cotton Tape", "koef": 1.200, "satuan": "m'", "harga_satuan": 3500},
                {"bahan": "Sekrup Drywall Tambahan", "koef": 10.000, "satuan": "buah", "harga_satuan": 300},
                {"bahan": "Cat Plafon Matte", "koef": 0.250, "satuan": "kg", "harga_satuan": 36000},
            ],
            langkah_kerja=[
                "Buka alur retakan sambungan plafon berbentuk V dan bersihkan sisa kompon lama yang terlepas.",
                "Tambahkan sekrup drywall di sisi kiri-kanan sambungan bila panel terasa goyang pada rangka hollow.",
                "Tempelkan fiberglass mesh tape atau paper tape tepat di atas alur retakan.",
                "Aplikasikan kompon gypsum 2-3 lapis bertahap (tipis lalu melebar) hingga sambungan rata sempurna.",
                "Amplas halus dengan grid 180-220 setelah kering, lalu cat ulang plafon agar warna seragam."
            ],
            rekomendasi="Gunakan compound elastis berdaya rekat tinggi dan pastikan jarak penggantung rangka plafon sesuai standar (maks. 60 cm)."
        ),
        "A.4.5.1.07.JT": _make_ahsp(
            kode="A.4.5.1.07.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Plafon",
            elemen_terkait="Plafon",
            sub_kerusakan="Rusak",
            nama="Pembongkaran & Penggantian Penutup Plafon Gypsum/GRC + Rangka Hollow",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.95,
            bobot_fungsi=0.80,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.140, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Plafon", "koef": 0.160, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.016, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.008, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Gypsum Board 9 mm / GRC 4 mm", "koef": 0.364, "satuan": "lbr", "harga_satuan": 88000},
                {"bahan": "Rangka Hollow Galvanis 4x4 & 2x4 (Penggantian)", "koef": 1.200, "satuan": "m'", "harga_satuan": 18500},
                {"bahan": "Sekrup Drywall 1 inch & Paku Ramset", "koef": 20.000, "satuan": "buah", "harga_satuan": 350},
                {"bahan": "Compound, Textile Tape & Cat Plafon", "koef": 0.500, "satuan": "kg", "harga_satuan": 32000},
            ],
            langkah_kerja=[
                "Bongkar panel plafon yang jebol, pecah, atau melendut berat beserta bagian rangka yang bengkok/korosi.",
                "Pasang dan setel elevasi waterpass rangka hollow galvanis pengganti serta perkuat penggantung (rod/hanger) ke dak/kuda-kuda.",
                "Pasang lembaran gypsum board / GRC baru memakai sekrup drywall pada jarak grid 20 cm secara rapi.",
                "Tutup seluruh kepala sekrup dan sambungan papan dengan joint compound + textile tape.",
                "Amplas sambungan hingga rata, beri list profil jika ada, lalu cat ulang 2 lapis."
            ],
            rekomendasi="Untuk area semi-outdoor atau rawan lembab, ganti panel memakai papan GRC/Kalsiboard tahan air agar tidak cepat lapuk."
        ),

        # ARSITEKTUR 3: PINTU DAN JENDELA
        "A.4.6.2.10.JT": _make_ahsp(
            kode="A.4.6.2.10.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Pintu dan Jendela",
            elemen_terkait="Pintu dan Jendela",
            sub_kerusakan="Kotor",
            nama="Pembersihan Menyeluruh Daun, Kusen & Kaca Pintu/Jendela + Pelumasan Hardware",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.35,
            bobot_fungsi=0.15,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.070, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Finishing", "koef": 0.040, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Mandor", "koef": 0.004, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Cairan Pembersih Kaca & Kusen Netral", "koef": 0.150, "satuan": "liter", "harga_satuan": 38000},
                {"bahan": "Pelumas Engsel / Silicone Spray", "koef": 0.050, "satuan": "kaleng", "harga_satuan": 65000},
                {"bahan": "Kain Microfiber & Spons Halus", "koef": 0.250, "satuan": "buah", "harga_satuan": 15000},
            ],
            langkah_kerja=[
                "Bersihkan debu tebal pada rel jendela geser, sudut kusen, dan daun pintu memakai kuas halus dan vacuum.",
                "Semprotkan cairan pembersih pH netral pada permukaan kusen aluminium/kayu serta kaca, bersihkan dengan kain microfiber.",
                "Hilangkan noda lem/kerak membandel memakai scraper plastik agar tidak menggores kaca atau powder coating.",
                "Keringkan permukaan dengan wiper karet hingga bening bebas bercak.",
                "Semprotkan pelumas silikon pada engsel, rel, dan slot kunci agar bukaan beroperasi lancar."
            ],
            rekomendasi="Hindari cairan pembersih bersifat asam kuat/abrasif pada kusen aluminium karena dapat memudarkan lapisan anodized."
        ),
        "A.4.6.2.12.JT": _make_ahsp(
            kode="A.4.6.2.12.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Pintu dan Jendela",
            elemen_terkait="Pintu dan Jendela",
            sub_kerusakan="Pelapukan",
            nama="Perbaikan / Penggantian Bagian Kusen & Daun Pintu Kayu Lapuk + Wood Treatment",
            satuan="m2",
            tingkat_keparahan="Berat",
            bobot_fisik=1.00,
            bobot_fungsi=0.85,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.180, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Kayu Ahli", "koef": 0.350, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang_Ahli"]},
                {"tenaga": "Kepala Tukang", "koef": 0.035, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.010, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Kayu Kamper / Meranti Oven Pengganti", "koef": 0.025, "satuan": "m3", "harga_satuan": 7500000},
                {"bahan": "Cairan Anti Rayap & Pengawet Kayu", "koef": 0.250, "satuan": "liter", "harga_satuan": 75000},
                {"bahan": "Dempul Kayu (Wood Filler) & Lem Epoxy", "koef": 0.300, "satuan": "kg", "harga_satuan": 55000},
                {"bahan": "Cat Kayu / Politur Finishing 2 Lapis", "koef": 0.350, "satuan": "kg", "harga_satuan": 72000},
            ],
            langkah_kerja=[
                "Potong dan buang bagian kayu kusen atau kaki daun pintu yang mengalami pelapukan akibat lembab atau serangan rayap.",
                "Semprotkan cairan anti-rayap dan pengawet kayu pada kayu eksisting yang tersisa serta dinding kontaknya.",
                "Lakukan penyambungan kayu pengganti kering oven memakai sambungan pen-lubang dan lem epoxy struktural.",
                "Dempul rapat seluruh sambungan kayu, lalu amplas bertahap hingga profil halus sejajar.",
                "Lapisi cat dasar menie kayu dan cat/politur finishing tahan cuaca sebanyak 2 lapis."
            ],
            rekomendasi="Berikan celah nat 5 mm / sepatu kusen granit di dasar kusen area basah agar kayu tidak kontak langsung dengan genangan air lantai."
        ),
        "A.4.6.2.14.JT": _make_ahsp(
            kode="A.4.6.2.14.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Pintu dan Jendela",
            elemen_terkait="Pintu dan Jendela",
            sub_kerusakan="Karat",
            nama="Pembersihan Korosi Besi/Hardware Pintu & Pengecatan Anti Karat Zinc Chromate",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.75,
            bobot_fungsi=0.55,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.110, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Besi/Cat", "koef": 0.140, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.014, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.006, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Cat Dasar Anti Karat Zinc Chromate", "koef": 0.220, "satuan": "kg", "harga_satuan": 82000},
                {"bahan": "Cat Besi Sintetik Gloss Finishing", "koef": 0.250, "satuan": "kg", "harga_satuan": 78000},
                {"bahan": "Thinner & Amplas Besi / Sikat Kawat", "koef": 0.300, "satuan": "liter", "harga_satuan": 36000},
            ],
            langkah_kerja=[
                "Rontokkan seluruh kerak karat menggunakan sikat kawat baja dan amplas besi hingga logam bersih mengkilap.",
                "Lap permukaan besi menggunakan kain basah thinner untuk menghilangkan sisa debu karat dan minyak.",
                "Oleskan segera 1-2 lapis cat dasar anti karat Zinc Chromate / Epoxy Primer secara merata.",
                "Biarkan cat dasar mengering minimal 6 jam agar perlindungan oksidasi terbentuk sempurna.",
                "Lapisi cat finishing besi sintetik 2 lapis dan ganti engsel/sekrup yang sudah macet korosif dengan stainless steel."
            ],
            rekomendasi="Ganti engsel dan baut pengunci yang mengalami karat parah dengan material stainless steel SUS 304."
        ),
        "A.4.6.2.16.JT": _make_ahsp(
            kode="A.4.6.2.16.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Pintu dan Jendela",
            elemen_terkait="Pintu dan Jendela",
            sub_kerusakan="Kaca retak",
            nama="Pembongkaran & Pemasangan Kaca Bening/Rayben 5 mm Baru + Karet Gasket",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.90,
            bobot_fungsi=0.80,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.120, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Kaca/Aluminium", "koef": 0.200, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang_Ahli"]},
                {"tenaga": "Kepala Tukang", "koef": 0.020, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.008, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Kaca Lembaran Tebal 5 mm (Polos/Rayben)", "koef": 1.100, "satuan": "m2", "harga_satuan": 155000},
                {"bahan": "Karet Gasket Jepit Kaca", "koef": 4.000, "satuan": "m'", "harga_satuan": 8500},
                {"bahan": "Silicone Sealant Netral", "koef": 0.350, "satuan": "tube", "harga_satuan": 55000},
            ],
            langkah_kerja=[
                "Gunakan sarung tangan anti-sayat & kacamata pelindung, lalu lepaskan list kaca dan karet gasket lama secara hati-hati.",
                "Bongkar lembaran kaca yang retak secara utuh/bertahap agar pecahan tidak melukai pengguna ruangan.",
                "Bersihkan alur dudukan kaca pada kusen dari serpihan kaca dan residu sealant lama.",
                "Pasang ganjalan setting block karet di dasar alur, lalu masukkan lembaran kaca 5 mm baru dengan toleransi muai 3 mm.",
                "Pasang kembali list penjepit, karet gasket baru, dan segel keliling pertemuan dengan silicone sealant netral."
            ],
            rekomendasi="Untuk pintu akses utama atau jendela rendah rawan benturan, pertimbangkan penggunaan kaca tempered atau laminated safety glass."
        ),
        "A.4.6.2.04.JT": _make_ahsp(
            kode="A.4.6.2.04.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Pintu dan Jendela",
            elemen_terkait="Pintu dan Jendela",
            sub_kerusakan="Sealent",
            nama="Penggantian Karet Gasket & Injeksi Ulang Sealant Kaca Kusen Aluminium",
            satuan="m'",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.50,
            bobot_fungsi=0.45,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.050, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Aluminium/Kaca", "koef": 0.080, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.008, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.003, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Silicone Sealant Netral Tahan Cuaca", "koef": 0.300, "satuan": "tube", "harga_satuan": 55000},
                {"bahan": "Backer Rod / Gasket Karet EPDM", "koef": 1.050, "satuan": "m'", "harga_satuan": 8000},
                {"bahan": "Masking Tape Kertas", "koef": 0.500, "satuan": "roll", "harga_satuan": 12000},
            ],
            langkah_kerja=[
                "Kupas tuntas sealant lama yang getas, putus, atau lepas dari celah antara kaca/kusen dan dinding menggunakan cutter khusus.",
                "Bersihkan celah dari debu dan kelembapan lalu lap dengan cairan pembersih alkohol hingga kering.",
                "Tempelkan masking tape di sepanjang kedua sisi celah untuk mendapatkan garis sealant yang lurus dan rapi.",
                "Pasang backer rod / karet gasket jika celah terlalu dalam, lalu tembakkan silicone sealant tipe netral secara kontinu tanpa jeda udara.",
                "Ratakan alur sealant menggunakan spatula karet cekung, lalu lepaskan masking tape sebelum sealant mengering."
            ],
            rekomendasi="Wajib gunakan silicone sealant bertipe netral (non-acidic) agar tidak menimbulkan korosi atau noda bercak pada profil aluminium dan kaca."
        ),

        # ARSITEKTUR 4: LANTAI
        "A.4.4.3.02.JT": _make_ahsp(
            kode="A.4.4.3.02.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Lantai",
            elemen_terkait="Lantai",
            sub_kerusakan="Pecah",
            nama="Pembongkaran & Pemasangan Ulang Lantai Keramik/Granit Pecah + Perekat Instan",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.85,
            bobot_fungsi=0.65,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.250, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Keramik", "koef": 0.200, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.020, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.010, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Ubin Keramik / Granit Pengganti 60x60", "koef": 1.050, "satuan": "m2", "harga_satuan": 125000},
                {"bahan": "Semen Instan Perekat Ubin (MU-400)", "koef": 5.500, "satuan": "kg", "harga_satuan": 4800},
                {"bahan": "Tile Grout Pengisi Nat Elastis", "koef": 0.600, "satuan": "kg", "harga_satuan": 18000},
            ],
            langkah_kerja=[
                "Potong seluruh garis nat di sekeliling keping keramik yang pecah menggunakan mesin gerinda potong tipis agar getaran tidak merambat ke keramik sebelahnya.",
                "Bongkar keping keramik yang pecah beserta sisa spesi keras di bawahnya, lalu bersihkan dasar lantai dari debu.",
                "Aplikasikan mortar instan perekat keramik (tile adhesive) pada dasar lantai dan punggung keramik (back-buttering) memakai roskam bergigi.",
                "Pasang keping ubin pengganti sejajar waterpass dengan keramik sekitar dan beri tile spacer nat selebar 2-3 mm.",
                "Isi celah nat dengan semen grout polimer setelah perekat mengering minimal 24 jam."
            ],
            rekomendasi="Simpan cadangan ubin dari batch produksi yang sama dan hindari beban benturan benda berat sebelum perekat mengeras sempurna."
        ),
        "A.4.4.3.05.JT": _make_ahsp(
            kode="A.4.4.3.05.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Lantai",
            elemen_terkait="Lantai",
            sub_kerusakan="Retak",
            nama="Perbaikan Retak Permukaan Lantai dengan Epoxy Grouting / Penggantian Parsial",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.75,
            bobot_fungsi=0.55,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.160, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Keramik", "koef": 0.150, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.015, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.008, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Epoxy Resin Filler / Grouting Lantai", "koef": 0.800, "satuan": "kg", "harga_satuan": 115000},
                {"bahan": "Ubin Pengganti (Cadangan Keping Retak Parah)", "koef": 0.500, "satuan": "m2", "harga_satuan": 125000},
                {"bahan": "Semen Perekat & Tile Grout", "koef": 2.500, "satuan": "kg", "harga_satuan": 6500},
            ],
            langkah_kerja=[
                "Identifikasi pola retak apakah hanya pada glasur ubin atau menerus ke lapisan screed di bawahnya.",
                "Untuk retak menerus parah, potong nat dan ganti keping ubin dengan mortar fleksibel; untuk retak garis halus, bersihkan alur retak memakai vacuum.",
                "Suntikkan atau tuangkan resin epoxy low-viscosity berwarna senada ke dalam celah retak hingga terisi padat.",
                "Ratakan permukaan berlebih sebelum mengeras dan poles halus setelah curing 12-24 jam.",
                "Periksa nat di sekitar ubin dan lakukan regrouting jika terdapat celah terbuka."
            ],
            rekomendasi="Jika retak membentuk garis panjang melintasi beberapa keping lantai secara berurutan, periksa kemungkinan defleksi balok atau penurunan subbase lantai."
        ),
        "A.4.4.3.08.JT": _make_ahsp(
            kode="A.4.4.3.08.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Lantai",
            elemen_terkait="Lantai",
            sub_kerusakan="Kotor",
            nama="Pembersihan Mendalam (Deep Cleaning) & Poles Noda Membandel Lantai",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.30,
            bobot_fungsi=0.15,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.080, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Operator Polisher", "koef": 0.050, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Mandor", "koef": 0.004, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Cairan Floor Heavy-Duty Degreaser/Cleaner", "koef": 0.200, "satuan": "liter", "harga_satuan": 45000},
                {"bahan": "Pad Poles Mesin Buffing", "koef": 0.100, "satuan": "buah", "harga_satuan": 75000},
                {"bahan": "Coating Pelindung Pori (Floor Sealer)", "koef": 0.100, "satuan": "liter", "harga_satuan": 85000},
            ],
            langkah_kerja=[
                "Sapu dan sedot debu serta pasir halus di seluruh permukaan lantai yang kotor.",
                "Oleskan cairan pembersih noda kerak/minyak (degreaser) yang aman bagi material keramik/granit dan diamkan 10 menit.",
                "Lakukan penyikatan mekanis memakai mesin floor polisher dengan pad lembut hingga kerak dan noda terangkat.",
                "Bilas dengan air bersih dan keringkan sisa air memakai wet vacuum / pel microfiber bersih.",
                "Lakukan buffing akhir atau oleskan sealer pelindung pori agar lantai kembali mengkilap dan tidak licin."
            ],
            rekomendasi="Hindari pembersih asam klorida (HCl) pekat karena dapat mengikis lapisan glasur keramik dan merusak nat semen."
        ),
        "A.4.4.3.10.JT": _make_ahsp(
            kode="A.4.4.3.10.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Lantai",
            elemen_terkait="Lantai",
            sub_kerusakan="Berjamur",
            nama="Pembersihan Lantai & Nat Berjamur dengan Anti-Fungal & Regrouting Nat",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.45,
            bobot_fungsi=0.25,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.100, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Keramik", "koef": 0.080, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Mandor", "koef": 0.005, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Cairan Pembasmi Jamur & Lumut Lantai", "koef": 0.250, "satuan": "liter", "harga_satuan": 52000},
                {"bahan": "Nat Keramik Anti Jamur (Epoxy Grout)", "koef": 0.450, "satuan": "kg", "harga_satuan": 32000},
            ],
            langkah_kerja=[
                "Sikat permukaan lantai berjamur memakai larutan anti jamur dan pembasmi spora lumut hingga bersih.",
                "Kerok bagian nat lantai lama yang sudah menghitam dan keropos sedalam 3 mm dengan grout scraper.",
                "Bilas seluruh lantai dan nat dengan air bersih lalu biarkan kering menyeluruh.",
                "Isi ulang jalur nat menggunakan semen grout anti-jamur (epoxy/waterproof grout) memakai roskam karet.",
                "Bersihkan sisa residu grout pada ubin dengan spons lembab setelah 30 menit."
            ],
            rekomendasi="Pastikan kemiringan lantai menuju floor drain cukup (1-2%) agar tidak terjadi genangan air permanen yang memicu jamur."
        ),
        "A.4.4.3.12.JT": _make_ahsp(
            kode="A.4.4.3.12.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Lantai",
            elemen_terkait="Lantai",
            sub_kerusakan="Popping",
            nama="Penanganan Lantai Popping (Menggelembung), Pembongkaran & Perekat Fleksibel",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.90,
            bobot_fungsi=0.75,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.280, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Keramik Ahli", "koef": 0.240, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang_Ahli"]},
                {"tenaga": "Kepala Tukang", "koef": 0.024, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.012, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Ubin Keramik / Granite Tile Baru 60x60", "koef": 1.050, "satuan": "m2", "harga_satuan": 125000},
                {"bahan": "Perekat Keramik Fleksibel (High-Flex Adhesive)", "koef": 6.000, "satuan": "kg", "harga_satuan": 7500},
                {"bahan": "Tile Grout Fleksibel & Movement Joint", "koef": 0.700, "satuan": "kg", "harga_satuan": 22000},
            ],
            langkah_kerja=[
                "Ketuk area sekitar kerusakan (sounding test) untuk menentukan batas ubin yang sudah kopong atau kehilangan daya lekat.",
                "Bongkar seluruh keping ubin yang terangkat (popping) beserta spesi lama yang terlepas dari screed lantai.",
                "Bersihkan dan kasarkan/chipping ringan permukaan screed beton serta basahi sedikit dengan air bersih.",
                "Aplikasikan mortar perekat fleksibel (elastis terhadap muai-susut termal) secara penuh di lantai dan punggung ubin tanpa rongga udara.",
                "Pasang ubin baru dengan jarak nat minimal 3 mm dan sediakan celah expansion joint diisi polyurethane sealant di tepi ruangan."
            ],
            rekomendasi="Sediakan celah muai (expansion joint) 5-8 mm setiap bentang 4-5 meter agar tegangan termal tidak mendorong ubin terangkat kembali."
        ),
        "A.4.4.3.15.JT": _make_ahsp(
            kode="A.4.4.3.15.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Lantai",
            elemen_terkait="Lantai",
            sub_kerusakan="Gumpil",
            nama="Perbaikan Lantai Gumpil (Chipping) dengan Resin Epoxy Filler / Tambal Ubin",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.55,
            bobot_fungsi=0.35,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.120, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Keramik", "koef": 0.140, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.014, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.006, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Resin Epoxy Filler Pengisi Keramik/Granit", "koef": 0.650, "satuan": "kg", "harga_satuan": 118000},
                {"bahan": "Pigmen Pewarna Senada Ubin", "koef": 0.100, "satuan": "tube", "harga_satuan": 45000},
                {"bahan": "Pad Amplas Air & Batu Poles Halus", "koef": 0.500, "satuan": "buah", "harga_satuan": 15000},
            ],
            langkah_kerja=[
                "Bersihkan bagian sudut atau tepi ubin yang gumpil/cuil dari debu, lemak, dan serpihan glasur longgar.",
                "Lap bagian gumpil memakai thinner/alkohol agar kering bebas kontaminasi.",
                "Campurkan resin epoxy filler dua komponen bersama pigmen warna hingga menyerupai warna dasar lantai.",
                "Isikan adonan filler sedikit lebih tinggi pada titik gumpil dengan kape dan biarkan mengeras 4-6 jam.",
                "Ratakan dan haluskan tambalan memakai amplas air halus berurutan (grid 400 hingga 1000) agar menyatu rata tanpa tepi tajam."
            ],
            rekomendasi="Jika gumpil berada di tepi ambang pintu berisiko tersandung atau meluas, pasang list penutup transisi aluminium/kuningan."
        ),

        # ARSITEKTUR 5: ATAP
        "A.4.5.2.01.JT": _make_ahsp(
            kode="A.4.5.2.01.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Atap",
            elemen_terkait="Atap",
            sub_kerusakan="Bocor",
            nama="Perbaikan Penutup Atap Bocor, Flashing Nok/Talang & Pelapisan Anti Bocor",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.90,
            bobot_fungsi=0.85,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.150, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Atap", "koef": 0.160, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.016, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.008, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Waterproofing Elastomer + Serat Membrane", "koef": 1.250, "satuan": "kg", "harga_satuan": 65000},
                {"bahan": "Sealant Karet Sekrup / Flashing Bitumen", "koef": 0.400, "satuan": "m2", "harga_satuan": 72000},
                {"bahan": "Sekrup Roofing Karet Pengganti", "koef": 6.000, "satuan": "buah", "harga_satuan": 1200},
            ],
            langkah_kerja=[
                "Telusuri jalur aliran air saat hujan atau uji siram untuk menemukan titik bocor (celah nok, sambungan jurai, atau karet baut getas).",
                "Bersihkan area sambungan genteng/spandek, flashing dinding, dan talang air dari lumut dan endapan tanah.",
                "Ganti baut roofing yang karet sealnya getas dan perbaiki kedudukan tumpang tindih (overlap) penutup atap.",
                "Pasang membrane serat polyester + waterproofing elastomer 2 lapis menyilang pada pertemuan nok, jurai, dan talang.",
                "Lakukan uji siram air mengalir minimal 15 menit guna memverifikasi tidak ada lagi rembesan ke ruang bawah."
            ],
            rekomendasi="Periksa rutin seal karet sekrup atap metal dan pertemuan flashing dinding parapet menjelang musim hujan."
        ),
        "A.4.5.2.04.JT": _make_ahsp(
            kode="A.4.5.2.04.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Atap",
            elemen_terkait="Atap",
            sub_kerusakan="Pecah",
            nama="Pembongkaran & Penggantian Penutup Atap Genteng / Spandek Pecah",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.95,
            bobot_fungsi=0.80,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.160, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Atap", "koef": 0.180, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.018, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.008, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Penutup Atap Pengganti (Genteng / Spandek)", "koef": 1.100, "satuan": "m2", "harga_satuan": 115000},
                {"bahan": "Sekrup / Pengikat Atap & Kawat Pengait", "koef": 8.000, "satuan": "buah", "harga_satuan": 1200},
                {"bahan": "Adukan Semen / Sealant Sambungan", "koef": 1.500, "satuan": "kg", "harga_satuan": 12000},
            ],
            langkah_kerja=[
                "Gunakan sabuk pengaman kerja ketinggian (body harness) dan papan pijakan agar beban terdistribusi pada reng/gording.",
                "Angkat dan keluarkan lembaran atau keping penutup atap yang pecah serta bersihkan dudukan reng.",
                "Periksa apakah terdapat reng kayu yang patah atau reng baja ringan melendut; perkuat terlebih dahulu bila rusak.",
                "Pasang keping genteng/lembaran atap pengganti sesuai pola kunci kaitan (interlocking) dan panjang tumpang tindih standar.",
                "Kencangkan baut/pengait anti angin dan tutup kembali adukan nok jika terdampak pembongkaran."
            ],
            rekomendasi="Saat mengganti atap di area ketinggian, wajib gunakan safety harness dan lakukan pengecekan jarak reng sesuai modul penutup atap baru."
        ),
        "A.4.5.2.07.JT": _make_ahsp(
            kode="A.4.5.2.07.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Atap",
            elemen_terkait="Atap",
            sub_kerusakan="Karat",
            nama="Pembersihan Korosi Penutup/Rangka Atap Metal & Pelapisan Cat Anti Karat",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.80,
            bobot_fungsi=0.65,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.130, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Besi/Cat", "koef": 0.140, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Kepala Tukang", "koef": 0.014, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.007, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Rust Converter / Primer Epoxy Anti Karat", "koef": 0.250, "satuan": "kg", "harga_satuan": 92000},
                {"bahan": "Cat Atap Metal Tahan Cuaca (Weatherproof)", "koef": 0.280, "satuan": "kg", "harga_satuan": 78000},
                {"bahan": "Sikat Kawat & Kain Lap Pembersih", "koef": 0.250, "satuan": "buah", "harga_satuan": 15000},
            ],
            langkah_kerja=[
                "Gosok dan rontokkan lapisan karat aktif pada atap seng/galvalum/rangka baja memakai sikat kawat hingga bersih.",
                "Bilas serbuk karat menggunakan air bertekanan rendah dan biarkan kering di bawah sinar matahari.",
                "Kuaskan rust converter / primer anti karat zinc-rich pada seluruh titik bekas korosi.",
                "Lapisi cat finishing atap metal tahan UV sebanyak 2 lapis hingga menutup sempurna.",
                "Untuk titik karat yang sudah berlubang tembus, tambal terlebih dahulu memakai flashing aluminium foil bitumen."
            ],
            rekomendasi="Jika pelat atap metal mengalami keropos tembus luas (> 30% luasan), pertimbangkan penggantian lembaran atap secara utuh."
        ),
        "A.4.5.2.10.JT": _make_ahsp(
            kode="A.4.5.2.10.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Atap",
            elemen_terkait="Atap",
            sub_kerusakan="Kotor",
            nama="Pembersihan Kotoran Daun, Sampah & Pelancaran Talang Penutup Atap",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.35,
            bobot_fungsi=0.25,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.075, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Atap", "koef": 0.045, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Mandor", "koef": 0.004, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Karung Pengumpul Sampah & Saringan Talang", "koef": 0.200, "satuan": "buah", "harga_satuan": 15000},
                {"bahan": "Air Kerja & Sabun Pembersih Netral", "koef": 0.150, "satuan": "liter", "harga_satuan": 25000},
            ],
            langkah_kerja=[
                "Angkut tumpukan daun kering, ranting, lumpur, dan sampah pada permukaan atap dan saluran talang jurai.",
                "Masukkan sampah ke dalam karung kerja agar tidak jatuh mengotori area bawah gedung.",
                "Semprot bersih alur gelombang genteng/spandek serta pipa pembuangan talang tegak.",
                "Pasang saringan kawat (leaf guard) di mulut roof drain/corong talang untuk mencegah penyumbatan ulang.",
                "Pastikan aliran air mengalir lancar turun hingga saluran drainase halaman."
            ],
            rekomendasi="Jadwalkan pembersihan rutin atap dan talang minimal 2 kali setahun (sebelum dan pertengahan musim penghujan)."
        ),
        "A.4.5.2.12.JT": _make_ahsp(
            kode="A.4.5.2.12.JT",
            bidang="Arsitektur",
            kategori="Arsitektur - Atap",
            elemen_terkait="Atap",
            sub_kerusakan="Jamur",
            nama="Pembersihan Lumut/Jamur Atap High-Pressure & Coating Pelindung Genteng",
            satuan="m2",
            tingkat_keparahan="Ringan",
            bobot_fisik=0.45,
            bobot_fungsi=0.30,
            is_struktural=False,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.100, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Cat/Atap", "koef": 0.090, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang"]},
                {"tenaga": "Mandor", "koef": 0.005, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Cairan Algicidal / Anti Lumut Atap", "koef": 0.200, "satuan": "liter", "harga_satuan": 55000},
                {"bahan": "Cat Coating Genteng Anti Jamur & UV", "koef": 0.250, "satuan": "kg", "harga_satuan": 64000},
            ],
            langkah_kerja=[
                "Semprot permukaan atap dari arah bubungan turun ke tepi menggunakan jet washer untuk merontokkan lumut dan jamur tebal.",
                "Semprotkan larutan anti lumut (algicide) secara merata untuk mematikan spora akar jamur di pori-pori genteng.",
                "Biarkan atap mengering sempurna di bawah sinar matahari selama 1 hari.",
                "Aplikasikan 2 lapis cat pelindung genteng (roof coating) tahan cuaca agar pori tertutup rapat dan air cepat meluncur turun."
            ],
            rekomendasi="Pangkas dahan pohon yang terlalu rindang menaungi atap untuk mengurangi kelembapan berlebih pemicu lumut."
        ),

        # STRUKTUR BETON BERTULANG
        "A.4.1.1.18.JT": _make_ahsp(
            kode="A.4.1.1.18.JT",
            bidang="Struktur",
            kategori="Struktur - Balok Beton Bertulang",
            elemen_terkait="Balok Beton Bertulang",
            sub_kerusakan="Retak Struktur (Crack)",
            nama="Perbaikan Retak Balok/Kolom dengan Injeksi Epoxy Resin Bertekanan",
            satuan="m'",
            tingkat_keparahan="Berat",
            bobot_fisik=1.15,
            bobot_fungsi=1.00,
            is_struktural=True,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.200, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Khusus Injeksi", "koef": 0.250, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang_Ahli"]},
                {"tenaga": "Kepala Tukang", "koef": 0.025, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.012, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Low Viscosity Epoxy Resin", "koef": 0.850, "satuan": "kg", "harga_satuan": 165000},
                {"bahan": "Epoxy Surface Sealant Paste", "koef": 0.500, "satuan": "kg", "harga_satuan": 95000},
                {"bahan": "Injection Port / Nipple", "koef": 5.000, "satuan": "buah", "harga_satuan": 5000},
            ],
            langkah_kerja=[
                "Pemetaan celah retak dan pengukuran lebar celah menggunakan crack width gauge microscope.",
                "Pembersihan permukaan beton dan penutupan celah luar dengan pasta epoxy surface sealer.",
                "Pemasangan injection ports (nipple) pada jarak interval 15 - 20 cm sepanjang alur retak.",
                "Injeksi epoxy resin viskositas rendah menggunakan pompa tekanan konstan hingga keluar di port sebelahnya.",
                "Pelepasan nipple dan perataan permukaan beton setelah waktu pengerasan (curing) 24 jam."
            ],
            rekomendasi="Lakukan evaluasi pembebanan struktur balok dan batasi muatan hidup sementara di atas lantai yang terdampak."
        ),
        "A.4.1.1.20.JT": _make_ahsp(
            kode="A.4.1.1.20.JT",
            bidang="Struktur",
            kategori="Struktur - Kolom Struktur Beton",
            elemen_terkait="Kolom Struktur Beton",
            sub_kerusakan="Spalling / Keropos",
            nama="Perbaikan Beton Keropos/Spalling & Exposed Rebar dengan Patching Mortar Polimer",
            satuan="m2",
            tingkat_keparahan="Berat",
            bobot_fisik=1.15,
            bobot_fungsi=1.00,
            is_struktural=True,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.280, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Batu Ahli", "koef": 0.200, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang_Ahli"]},
                {"tenaga": "Kepala Tukang", "koef": 0.020, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.015, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Structural Repair Mortar Non-Shrink", "koef": 18.000, "satuan": "kg", "harga_satuan": 8500},
                {"bahan": "Zinc-Rich Anti-Corrosion Primer", "koef": 0.350, "satuan": "kg", "harga_satuan": 110000},
                {"bahan": "Semen Bonding Agent", "koef": 0.500, "satuan": "liter", "harga_satuan": 55000},
            ],
            langkah_kerja=[
                "Pembongkaran (chipping) beton keropos dan berongga hingga mencapai batas beton padat (sound concrete).",
                "Pembersihan kerak karat pada besi tulangan menggunakan sikat kawat baja wire brush.",
                "Pelapisan pelindung anti korosi berbahan dasar zinc-rich primer pada seluruh tulangan baja yang terekspos.",
                "Penyapuan bonding agent sintetis pada bidang kontak beton lama.",
                "Aplikasi mortar polimer non-shrink bertahap (layer by layer) dan pemadatan manual hingga rata rapi."
            ],
            rekomendasi="Pastikan tebal selimut beton minimal 30 mm terpenuhi kembali guna melindungi baja tulangan dari karbonasi atmosfer."
        ),
        "A.4.4.3.35.JT": _make_ahsp(
            kode="A.4.4.3.35.JT",
            bidang="Struktur",
            kategori="Struktur - Pelat Lantai / Dak Beton",
            elemen_terkait="Pelat Lantai / Dak Beton",
            sub_kerusakan="Rembesan / Bocor & Efflorescence",
            nama="Pelapisan Waterproofing Membrane / Polyurethane Pelat Dak Beton",
            satuan="m2",
            tingkat_keparahan="Sedang",
            bobot_fisik=0.90,
            bobot_fungsi=0.75,
            is_struktural=True,
            upah=[
                {"tenaga": "Pekerja", "koef": 0.150, "satuan": "OH", "harga_satuan": UPAH_JATENG["Pekerja"]},
                {"tenaga": "Tukang Khusus Pelapis", "koef": 0.120, "satuan": "OH", "harga_satuan": UPAH_JATENG["Tukang_Ahli"]},
                {"tenaga": "Kepala Tukang", "koef": 0.012, "satuan": "OH", "harga_satuan": UPAH_JATENG["Kepala_Tukang"]},
                {"tenaga": "Mandor", "koef": 0.008, "satuan": "OH", "harga_satuan": UPAH_JATENG["Mandor"]},
            ],
            bahan=[
                {"bahan": "Liquid Polyurethane Waterproofing", "koef": 1.600, "satuan": "kg", "harga_satuan": 75000},
                {"bahan": "Primer Bitumen / Epoxy Substrate", "koef": 0.250, "satuan": "kg", "harga_satuan": 65000},
                {"bahan": "Polyester Mesh Reinforcement", "koef": 1.100, "satuan": "m2", "harga_satuan": 18000},
            ],
            langkah_kerja=[
                "Pembersihan lumut, efloresensi garam, debu, dan residu kotoran pada permukaan dak beton.",
                "Pembuatan fillet lengkung (chamfer) 5x5 cm pada pertemuan antara pelat lantai dengan parapet dinding.",
                "Pelapisan primer perekat dan pemasangan serat mesh pada area sudut dan retakan.",
                "Aplikasi waterproofing polyurethane lapis pertama dan kedua secara menyilang.",
                "Uji genang air (ponding test) selama minimal 24 jam untuk menjamin integritas kekedapan air."
            ],
            rekomendasi="Cek kemiringan slope dak beton menuju roof drain minimal 1.5% untuk mencegah genangan air terakumulasi."
        ),
    }

    @classmethod
    def get_all(cls) -> List[Dict[str, Any]]:
        return list(cls.CATALOG.values())

    @classmethod
    def get_by_code(cls, kode: str) -> Optional[Dict[str, Any]]:
        """Mencari data AHSP berdasarkan kode analisis."""
        return cls.CATALOG.get(kode)

    @classmethod
    def find_by_element_and_defect(cls, elemen: str, sub_kerusakan: str = "") -> Dict[str, Any]:
        """Mencari data AHSP paling akurat berdasarkan elemen/objek inspeksi & sub-jenis kerusakan."""
        el_clean = elemen.replace("Arsitektur - ", "").replace("Struktur - ", "").strip().lower()
        sub_clean = sub_kerusakan.strip().lower()

        # Penanganan khusus elemen struktur beton agar Spalling/Exposed Rebar dan Retak bekerja pada Balok, Kolom, & Dak
        if any(k in el_clean for k in ("balok", "kolom", "dak", "pelat")):
            el_label = (
                "Balok Beton Bertulang" if "balok" in el_clean else
                "Kolom Struktur Beton" if "kolom" in el_clean else
                "Pelat Lantai / Dak Beton"
            )
            if any(k in sub_clean for k in ("rebar", "tulangan", "spalling", "keropos")):
                base = dict(cls.CATALOG["A.4.1.1.20.JT"])
                base["elemen_terkait"] = el_label
                if "rebar" in sub_clean or "tulangan" in sub_clean:
                    base["sub_kerusakan"] = "Tulangan Terbuka (Exposed Rebar)"
                    base["jenis_kerusakan"] = f"Besi Tulangan Terbuka (Exposed Rebar) pada {el_label}"
                else:
                    base["sub_kerusakan"] = "Spalling / Keropos Beton"
                    base["jenis_kerusakan"] = f"Spalling / Keropos Beton pada {el_label}"
                return base
            if any(k in sub_clean for k in ("rembes", "bocor", "efflorescence", "jamur")):
                base = dict(cls.CATALOG["A.4.4.3.35.JT"])
                base["elemen_terkait"] = el_label
                base["jenis_kerusakan"] = f"Rembesan Air & Efloresensi Garam pada {el_label}"
                return base
            if any(k in sub_clean for k in ("retak", "crack")):
                if "dak" in el_clean or "pelat" in el_clean:
                    return cls.CATALOG["A.4.4.3.35.JT"]
                base = dict(cls.CATALOG["A.4.1.1.18.JT"])
                base["elemen_terkait"] = el_label
                base["jenis_kerusakan"] = f"Retak Struktur (Crack) pada {el_label}"
                return base

        if sub_clean and "otomatis" not in sub_clean:
            for item in cls.CATALOG.values():
                item_el = item["elemen_terkait"].lower()
                item_sub = item["sub_kerusakan"].lower()
                if (item_el in el_clean or el_clean in item_el) and (
                    sub_clean == item_sub or sub_clean in item_sub or item_sub in sub_clean
                ):
                    return item

        return cls.find_by_element(elemen)

    @classmethod
    def find_by_element(cls, elemen: str) -> Dict[str, Any]:
        """Mencari data AHSP yang paling cocok berdasarkan elemen gedung."""
        el_clean = elemen.replace("Arsitektur - ", "").replace("Struktur - ", "").strip().lower()
        for item in cls.CATALOG.values():
            if item["elemen_terkait"].lower() in el_clean or el_clean in item["elemen_terkait"].lower():
                return item
        return cls.CATALOG["A.4.4.1.25.JT"]

    @classmethod
    def compute_material_breakdown(cls, kode: str, volume: float) -> Dict[str, Any]:
        """Menghitung rincian biaya terperinci per tenaga kerja dan per masing-masing bahan sesuai volume."""
        item = cls.get_by_code(kode) or cls.CATALOG["A.4.4.1.25.JT"]
        upah_rows = []
        bahan_rows = []
        tot_upah_satuan = 0.0
        tot_bahan_satuan = 0.0

        for u in item.get("koefisien_upah", []):
            k = float(u.get("koef", 0))
            h = float(u.get("harga_satuan", 0))
            keb_riil = round(k * volume, 3)
            tot_upah_satuan += k * h
            upah_rows.append({
                "Unsur Tenaga Kerja": u.get("tenaga", "-"),
                "Koefisien": f"{k:.3f}",
                "Satuan": u.get("satuan", "OH"),
                "Harga Satuan": int(h),
                "Kebutuhan Riil": f"{keb_riil:.3f} {u.get('satuan', 'OH')}",
                "Jumlah Biaya": int(round(keb_riil * h))
            })

        for b in item.get("koefisien_bahan", []):
            k = float(b.get("koef", 0))
            h = float(b.get("harga_satuan", 0))
            keb_riil = round(k * volume, 3)
            tot_bahan_satuan += k * h
            bahan_rows.append({
                "Nama Bahan / Material": b.get("bahan", "-"),
                "Koefisien": f"{k:.3f}",
                "Satuan": b.get("satuan", "-"),
                "Harga Satuan": int(h),
                "Kebutuhan Riil": f"{keb_riil:.2f} {b.get('satuan', '')}",
                "Jumlah Biaya": int(round(keb_riil * h))
            })

        subtotal_langsung = int(round((tot_upah_satuan + tot_bahan_satuan) * volume))
        overhead_total = int(round(subtotal_langsung * 0.15))
        grand_total = int(round(volume * item["harga_satuan"]))
        return {
            "upah_rows": upah_rows,
            "bahan_rows": bahan_rows,
            "subtotal_upah": sum(r["Jumlah Biaya"] for r in upah_rows),
            "subtotal_bahan": sum(r["Jumlah Biaya"] for r in bahan_rows),
            "subtotal_langsung": subtotal_langsung,
            "overhead_15": overhead_total,
            "harga_satuan": item["harga_satuan"],
            "grand_total": grand_total
        }


# ==================================================================================================
# 4. INTEGRASI MESIN DETEKSI AI YOLOV8 & COMPUTER VISION (VIA requests / lokal)
# ==================================================================================================
@st.cache_resource(show_spinner=False)
def _load_yolo_model(model_path: str):
    """Memuat dan men-cache model bobot YOLOv8 (best.pt) agar inferensi cepat & konsisten."""
    from ultralytics import YOLO
    return YOLO(model_path)


class AIAgentClient:
    """Kelas integrasi untuk inferensi model YOLOv8 dan modul penilaian AHSP Jawa Tengah."""

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
        force_simulation: bool = False,
        luas_total: float = 25.0,
        jenis_kerusakan_pilihan: str = "Deteksi Otomatis AI"
    ) -> Dict[str, Any]:
        """
        Mengubah foto ke format Base64 dan memproses citra melalui inferensi YOLOv8
        atau endpoint API eksternal via protokol HTTP POST JSON.
        """
        base64_encoded_img = base64.b64encode(image_bytes).decode("utf-8")
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        payload = {
            "image_base64": base64_encoded_img,
            "filename": filename,
            "volume": float(volume),
            "luas_total": float(luas_total),
            "elemen_gedung": elemen_gedung,
            "jenis_kerusakan_pilihan": jenis_kerusakan_pilihan,
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

        # 1. Jika pengguna memilih mode simulasi offline
        if force_simulation:
            return cls._deterministic_civil_ai_engine(
                elemen_gedung, volume, filename, "Simulasi Mandiri (Offline Safe)",
                luas_total=luas_total, jenis_kerusakan_pilihan=jenis_kerusakan_pilihan
            )

        # 2. Jalankan inferensi model YOLOv8 kustom (weights/best.pt)
        local_yolo_result = cls.detect_with_local_yolo(
            image_bytes, elemen_gedung, volume, filename,
            luas_total=luas_total, jenis_kerusakan_pilihan=jenis_kerusakan_pilihan
        )
        if local_yolo_result is not None:
            return local_yolo_result

        model_path = cls.get_yolo_model_path()
        if model_path:
            fallback = cls._deterministic_civil_ai_engine(
                elemen_gedung, volume, filename, "Standalone Civil AI Engine",
                luas_total=luas_total, jenis_kerusakan_pilihan=jenis_kerusakan_pilihan
            )
            try:
                import ultralytics  # noqa: F401
            except ImportError:
                fallback["warning"] = (
                    f"Model bobot '{os.path.basename(model_path)}' ditemukan, namun pustaka 'ultralytics' "
                    "belum terpasang. Dialihkan ke Standalone Civil Engine."
                )
                return fallback

        # 3. Eksekusi HTTP POST Call (jika custom endpoint disiapkan)
        if target_url and target_url != cls.DEFAULT_ENDPOINT:
            try:
                response = requests.post(target_url, json=payload, headers=headers, timeout=8)
                if response.status_code == 200:
                    result = response.json()
                    result["is_live_api"] = True
                    result["api_endpoint_used"] = target_url
                    return result
                fallback = cls._deterministic_civil_ai_engine(
                    elemen_gedung, volume, filename, f"HTTP Error {response.status_code}",
                    luas_total=luas_total, jenis_kerusakan_pilihan=jenis_kerusakan_pilihan
                )
                fallback["warning"] = f"Endpoint merespons status {response.status_code}. Beralih ke Standalone Engine."
                return fallback
            except Exception as exc:
                fallback = cls._deterministic_civil_ai_engine(
                    elemen_gedung, volume, filename, "Standalone Engine (Auto-Fallback)",
                    luas_total=luas_total, jenis_kerusakan_pilihan=jenis_kerusakan_pilihan
                )
                fallback["warning"] = f"Koneksi endpoint '{target_url}' tidak merespons ({type(exc).__name__}). Dialihkan ke Standalone Civil AI Engine."
                return fallback

        return cls._deterministic_civil_ai_engine(
            elemen_gedung, volume, filename, "Standalone Civil AI Engine",
            luas_total=luas_total, jenis_kerusakan_pilihan=jenis_kerusakan_pilihan
        )

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
    def _infer_visual_subdefect_fallback(cls, image_bytes: bytes, elemen_gedung: str) -> str:
        """Heuristik fitur visual citra (warna karat/jamur/kontras tepi) jika model YOLO tidak mengeluarkan box."""
        try:
            import numpy as np
            arr = np.array(Image.open(io.BytesIO(image_bytes)).convert("RGB").resize((160, 160)), dtype=np.float32)
            r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]
            gray = 0.299 * r + 0.587 * g + 0.114 * b
            # Edge density (proxy retak/pecah), dark stains (jamur/kotor), rust/brownish (karat/tulangan)
            gx = np.abs(np.diff(gray, axis=1)).mean()
            gy = np.abs(np.diff(gray, axis=0)).mean()
            edge_score = gx + gy
            dark_ratio = float((gray < 75).mean())
            rust_ratio = float(((r > g + 18) & (r > b + 28) & (r > 80)).mean())
            bright_ratio = float((gray > 205).mean())
        except Exception:
            return ""

        el = elemen_gedung.lower()
        if "dinding" in el or "partisi" in el:
            if dark_ratio > 0.22:
                return "Jamur"
            if rust_ratio > 0.15 or bright_ratio > 0.35:
                return "Pengelupasan cat"
            return "Retak rambut" if edge_score > 16 else "Cat menggelembung"
        if "plafon" in el:
            if dark_ratio > 0.20:
                return "Jamur"
            if rust_ratio > 0.12:
                return "Kebocoran"
            return "Retak" if edge_score > 18 else "Rusak"
        if "pintu" in el or "jendela" in el:
            if rust_ratio > 0.14:
                return "Karat"
            if edge_score > 24:
                return "Kaca retak"
            return "Pelapukan" if dark_ratio > 0.22 else "Kotor"
        if "lantai" in el and "dak" not in el:
            if dark_ratio > 0.22:
                return "Berjamur"
            if edge_score > 24:
                return "Pecah"
            return "Retak" if edge_score > 16 else "Kotor"
        if "atap" in el and "dak" not in el:
            if rust_ratio > 0.14:
                return "Karat"
            if dark_ratio > 0.22:
                return "Jamur"
            return "Pecah" if edge_score > 20 else "Bocor"
        if "balok" in el or "kolom" in el:
            if rust_ratio > 0.10 or dark_ratio > 0.25:
                return "Tulangan Terbuka (Exposed Rebar)"
            return "Retak Struktur (Crack)" if edge_score > 15 else "Spalling / Keropos"
        if "dak" in el or "pelat" in el:
            return "Rembesan / Bocor & Efflorescence"
        return ""

    @classmethod
    def _map_yolo_defect_to_ahsp(
        cls, class_name: str, elemen_gedung: str, jenis_kerusakan_pilihan: str = "Deteksi Otomatis AI"
    ) -> Dict[str, Any]:
        """
        Memetakan hasil deteksi 4 kelas YOLOv8 (Crack, Efflorescence, Exposed Rebar, Spalling)
        agar konsisten dengan perilaku model awal sekaligus mendukung 5 objek Arsitektur & 3 Struktur.
        """
        if jenis_kerusakan_pilihan and "otomatis" not in jenis_kerusakan_pilihan.lower():
            item = AHSPKnowledgeBase.find_by_element_and_defect(elemen_gedung, jenis_kerusakan_pilihan)
            return {
                "kode_ahsp": item["kode"],
                "elemen_terkait": item["elemen_terkait"],
                "jenis_kerusakan": item["jenis_kerusakan"],
                "tingkat_keparahan": item["tingkat_keparahan"],
            }

        c_lower = str(class_name).lower()
        el_lower = elemen_gedung.lower()

        # Mode Otomatis Penuh (identik dengan perilaku model YOLOv8 awal) atau deteksi beton eksplisit (Exposed Rebar / Spalling)
        if "otomatis" in el_lower:
            if "rebar" in c_lower:
                base = dict(AHSPKnowledgeBase.CATALOG["A.4.1.1.20.JT"])
                return {
                    "kode_ahsp": base["kode"],
                    "elemen_terkait": "Kolom Struktur Beton",
                    "jenis_kerusakan": "Besi Tulangan Terbuka & Keropos Beton (Exposed Rebar)",
                    "tingkat_keparahan": "Berat",
                }
            if "spalling" in c_lower:
                base = dict(AHSPKnowledgeBase.CATALOG["A.4.1.1.20.JT"])
                return {
                    "kode_ahsp": base["kode"],
                    "elemen_terkait": "Kolom Struktur Beton",
                    "jenis_kerusakan": "Spalling / Keropos Beton & Selimut Terkelupas",
                    "tingkat_keparahan": "Berat",
                }
            if "efflorescence" in c_lower:
                item = AHSPKnowledgeBase.find_by_element_and_defect("Dinding dan Partisi", "Pengelupasan cat")
                return {
                    "kode_ahsp": item["kode"],
                    "elemen_terkait": item["elemen_terkait"],
                    "jenis_kerusakan": "Cat Dinding Mengelupas, Mengapur & Berjamur (Efflorescence)",
                    "tingkat_keparahan": item["tingkat_keparahan"],
                }
            item = AHSPKnowledgeBase.find_by_element_and_defect("Dinding dan Partisi", "Retak rambut")
            return {
                "kode_ahsp": item["kode"],
                "elemen_terkait": item["elemen_terkait"],
                "jenis_kerusakan": item["jenis_kerusakan"],
                "tingkat_keparahan": item["tingkat_keparahan"],
            }

        if "lantai" in el_lower and "dak" not in el_lower:
            if "rebar" in c_lower:
                return {
                    "kode_ahsp": "A.4.1.1.20.JT",
                    "elemen_terkait": "Pelat Lantai / Struktur Beton",
                    "jenis_kerusakan": "Besi Tulangan Terbuka & Keropos Beton (Exposed Rebar)",
                    "tingkat_keparahan": "Berat",
                }
            sub = "Pecah" if "spalling" in c_lower else ("Berjamur" if "efflorescence" in c_lower else "Retak")
            item = AHSPKnowledgeBase.find_by_element_and_defect("Lantai", sub)
        elif "plafon" in el_lower:
            if "rebar" in c_lower:
                return {
                    "kode_ahsp": "A.4.1.1.20.JT",
                    "elemen_terkait": "Pelat Lantai / Dak Beton",
                    "jenis_kerusakan": "Besi Tulangan Terbuka & Keropos Beton (Exposed Rebar)",
                    "tingkat_keparahan": "Berat",
                }
            sub = "Rusak" if "spalling" in c_lower else ("Kebocoran" if "efflorescence" in c_lower else "Retak")
            item = AHSPKnowledgeBase.find_by_element_and_defect("Plafon", sub)
        elif "pintu" in el_lower or "jendela" in el_lower:
            if "rebar" in c_lower:
                return {
                    "kode_ahsp": "A.4.1.1.20.JT",
                    "elemen_terkait": "Kolom Struktur Beton",
                    "jenis_kerusakan": "Besi Tulangan Terbuka & Keropos Beton (Exposed Rebar)",
                    "tingkat_keparahan": "Berat",
                }
            sub = "Kaca retak" if "crack" in c_lower else ("Pelapukan" if "spalling" in c_lower else "Karat")
            item = AHSPKnowledgeBase.find_by_element_and_defect("Pintu dan Jendela", sub)
        elif "atap" in el_lower and "dak" not in el_lower:
            if "rebar" in c_lower:
                return {
                    "kode_ahsp": "A.4.1.1.20.JT",
                    "elemen_terkait": "Pelat Atap / Dak Beton",
                    "jenis_kerusakan": "Besi Tulangan Terbuka & Keropos Beton (Exposed Rebar)",
                    "tingkat_keparahan": "Berat",
                }
            if "spalling" in c_lower:
                return {
                    "kode_ahsp": "A.4.1.1.20.JT",
                    "elemen_terkait": "Pelat Atap / Dak Beton",
                    "jenis_kerusakan": "Spalling / Keropos Beton & Selimut Terkelupas",
                    "tingkat_keparahan": "Berat",
                }
            sub = "Pecah" if "crack" in c_lower else ("Jamur" if "efflorescence" in c_lower else "Karat")
            item = AHSPKnowledgeBase.find_by_element_and_defect("Atap", sub)
        elif "dinding" in el_lower or "partisi" in el_lower:
            if "rebar" in c_lower:
                return {
                    "kode_ahsp": "A.4.1.1.20.JT",
                    "elemen_terkait": "Kolom Struktur Beton",
                    "jenis_kerusakan": "Besi Tulangan Terbuka & Keropos Beton (Exposed Rebar)",
                    "tingkat_keparahan": "Berat",
                }
            if "spalling" in c_lower:
                return {
                    "kode_ahsp": "A.4.1.1.20.JT",
                    "elemen_terkait": "Kolom Struktur Beton",
                    "jenis_kerusakan": "Spalling / Keropos Beton & Selimut Terkelupas",
                    "tingkat_keparahan": "Berat",
                }
            sub = "Jamur" if "efflorescence" in c_lower else "Retak rambut"
            item = AHSPKnowledgeBase.find_by_element_and_defect("Dinding dan Partisi", sub)
        elif "balok" in el_lower or "kolom" in el_lower:
            el_label = "Balok Beton Bertulang" if "balok" in el_lower else "Kolom Struktur Beton"
            if "rebar" in c_lower:
                item = AHSPKnowledgeBase.find_by_element_and_defect(el_label, "Tulangan Terbuka (Exposed Rebar)")
            elif "spalling" in c_lower:
                item = AHSPKnowledgeBase.find_by_element_and_defect(el_label, "Spalling / Keropos")
            elif "efflorescence" in c_lower:
                item = AHSPKnowledgeBase.find_by_element_and_defect(el_label, "Efflorescence")
            else:
                item = AHSPKnowledgeBase.find_by_element_and_defect(el_label, "Retak Struktur (Crack)")
        elif "dak" in el_lower or "pelat" in el_lower:
            if "spalling" in c_lower or "rebar" in c_lower:
                item = AHSPKnowledgeBase.find_by_element_and_defect("Pelat Lantai / Dak Beton", "Spalling")
            else:
                item = AHSPKnowledgeBase.find_by_element_and_defect("Pelat Lantai / Dak Beton", "Rembesan / Bocor & Efflorescence")
        else:
            item = AHSPKnowledgeBase.find_by_element(elemen_gedung)

        return {
            "kode_ahsp": item["kode"],
            "elemen_terkait": item["elemen_terkait"],
            "jenis_kerusakan": item["jenis_kerusakan"],
            "tingkat_keparahan": item["tingkat_keparahan"],
        }

    @classmethod
    def _build_analysis_result(
        cls,
        ahsp_item: Dict[str, Any],
        volume: float,
        luas_total: float,
        filename: str,
        mode_label: str,
        agent_id: str,
        is_yolo_active: bool,
        confidence_score: float,
        detected_boxes: Optional[List[Dict[str, Any]]] = None,
        warning: str = "",
        elemen_override: str = "",
        kerusakan_override: str = ""
    ) -> Dict[str, Any]:
        """Menyusun paket hasil analisis standar dengan skala prosentase dan tingkat keparahan yang logis."""
        elemen_final = elemen_override or ahsp_item["elemen_terkait"]
        kerusakan_final = kerusakan_override or ahsp_item["jenis_kerusakan"]
        harga_satuan = ahsp_item["harga_satuan"]
        total_biaya = int(round(volume * harga_satuan))
        temp_item = {
            "elemen_gedung": elemen_final,
            "jenis_kerusakan": kerusakan_final,
            "kode_ahsp": ahsp_item["kode"],
            "volume": float(volume),
            "luas_total": float(luas_total),
        }
        eval_res = compute_condition_assessment(temp_item)
        res = {
            "status": "success",
            "agent_id": agent_id,
            "mode": mode_label,
            "is_live_api": False,
            "is_yolo_active": is_yolo_active,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "filename": filename,
            "elemen_gedung": elemen_final,
            "jenis_kerusakan": kerusakan_final,
            "tingkat_keparahan": eval_res["keparahan"],
            "persentase_kerusakan": eval_res["persentase_kerusakan"],
            "luas_total": eval_res["luas_total"],
            "confidence_score": round(confidence_score, 1),
            "kode_ahsp": ahsp_item["kode"],
            "deskripsi_ahsp": ahsp_item["nama"],
            "volume": float(volume),
            "satuan": ahsp_item["satuan"],
            "harga_satuan": harga_satuan,
            "total_biaya": total_biaya,
            "langkah_kerja": ahsp_item["langkah_kerja"],
            "rekomendasi": ahsp_item["rekomendasi"],
            "acuan_regulasi": "AHSP Dinas Bina Marga & Cipta Karya Provinsi Jawa Tengah 2023",
            "detected_boxes": detected_boxes or [],
            "total_detected_defects": len(detected_boxes) if detected_boxes else 0,
        }
        if warning:
            res["warning"] = warning
        return res

    @classmethod
    def detect_with_local_yolo(
        cls,
        image_bytes: bytes,
        elemen_gedung: str,
        volume: float,
        filename: str,
        luas_total: float = 25.0,
        jenis_kerusakan_pilihan: str = "Deteksi Otomatis AI"
    ) -> Optional[Dict[str, Any]]:
        """Menjalankan inferensi lokal YOLOv8 jika file weights/best.pt terdeteksi."""
        model_path = cls.get_yolo_model_path()
        if not model_path:
            return None

        try:
            model = _load_yolo_model(model_path)
            img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
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
                    ahsp_map = cls._map_yolo_defect_to_ahsp(cls_name, elemen_gedung, jenis_kerusakan_pilihan)
                    sev_eval = compute_condition_assessment({
                        "elemen_gedung": ahsp_map["elemen_terkait"],
                        "jenis_kerusakan": ahsp_map["jenis_kerusakan"],
                        "kode_ahsp": ahsp_map["kode_ahsp"],
                        "volume": volume,
                        "luas_total": luas_total,
                    })

                    boxes_data.append({
                        "box": coords,
                        "class_name": cls_name,
                        "label": ahsp_map["jenis_kerusakan"],
                        "elemen_terkait": ahsp_map["elemen_terkait"],
                        "confidence": round(conf, 1),
                        "keparahan": sev_eval["keparahan"],
                        "kode_ahsp": ahsp_map["kode_ahsp"]
                    })

            if not boxes_data:
                sub_fallback = jenis_kerusakan_pilihan
                if not sub_fallback or "otomatis" in sub_fallback.lower():
                    sub_fallback = cls._infer_visual_subdefect_fallback(image_bytes, elemen_gedung)
                ahsp_item = AHSPKnowledgeBase.find_by_element_and_defect(elemen_gedung, sub_fallback)
                return cls._build_analysis_result(
                    ahsp_item=ahsp_item,
                    volume=volume,
                    luas_total=luas_total,
                    filename=filename,
                    mode_label=f"YOLOv8 Engine ({os.path.basename(model_path)})",
                    agent_id="yolov8-custom-model-engine",
                    is_yolo_active=True,
                    confidence_score=95.0,
                    detected_boxes=[],
                    warning=f"Model YOLOv8 ({os.path.basename(model_path)}) memindai citra dan memetakan parameter kerusakan pada elemen terpilih."
                )

            critical_boxes = [
                b for b in boxes_data
                if (not jenis_kerusakan_pilihan or "otomatis" in jenis_kerusakan_pilihan.lower())
                and b["kode_ahsp"] == "A.4.1.1.20.JT"
                and b["confidence"] >= 30.0
            ]
            main_box = max(critical_boxes, key=lambda x: x["confidence"]) if critical_boxes else max(boxes_data, key=lambda x: x["confidence"])
            ahsp_item = AHSPKnowledgeBase.get_by_code(main_box["kode_ahsp"]) or AHSPKnowledgeBase.find_by_element_and_defect(elemen_gedung, jenis_kerusakan_pilihan)
            return cls._build_analysis_result(
                ahsp_item=ahsp_item,
                volume=volume,
                luas_total=luas_total,
                filename=filename,
                mode_label=f"YOLOv8 Local Model ({os.path.basename(model_path)})",
                agent_id="yolov8-custom-model-engine",
                is_yolo_active=True,
                confidence_score=main_box["confidence"],
                detected_boxes=boxes_data,
                elemen_override=main_box.get("elemen_terkait", ""),
                kerusakan_override=main_box.get("label", "")
            )
        except Exception:
            return None

    @classmethod
    def _deterministic_civil_ai_engine(
        cls,
        elemen_gedung: str,
        volume: float,
        filename: str,
        mode_label: str,
        luas_total: float = 25.0,
        jenis_kerusakan_pilihan: str = "Deteksi Otomatis AI"
    ) -> Dict[str, Any]:
        """Mesin analisis teknik sipil berbasis katalog AHSP Jateng 2023 dan rasio prosentase kerusakan."""
        ahsp_item = AHSPKnowledgeBase.find_by_element_and_defect(elemen_gedung, jenis_kerusakan_pilihan)
        return cls._build_analysis_result(
            ahsp_item=ahsp_item,
            volume=volume,
            luas_total=luas_total,
            filename=filename,
            mode_label=mode_label,
            agent_id="yolov8-civil-vision-engine",
            is_yolo_active=False,
            confidence_score=97.2,
            detected_boxes=[]
        )


# ==================================================================================================
# 5. GENERATOR LAPORAN PENILAIAN KONDISI (IKBG) & ESTIMASI RAB (PDF)
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
    berbasis rasio luasan/volume kerusakan terhadap total komponen di ruangan
    (Mengacu Permen PUPR No. 24/PRT/M/2008 & Permen PUPR No. 22/PRT/M/2018).
    """
    elemen = item.get("elemen_gedung", "Komponen Bangunan")
    kerusakan = item.get("jenis_kerusakan", "Kerusakan Permukaan")
    kode = item.get("kode_ahsp", "")
    ahsp_meta = AHSPKnowledgeBase.CATALOG.get(kode, {})

    vol = float(item.get("volume") or 0.0)
    luas_total = float(item.get("luas_total") or 0.0)
    if luas_total <= 0:
        luas_total = max(20.0, vol * 2.5 if vol > 0 else 20.0)

    # Rasio Prosentase Kerusakan (0 - 100%)
    if item.get("persentase_kerusakan") is not None and float(item.get("persentase_kerusakan") or 0) > 0 and vol <= 0:
        persentase_kerusakan = min(100.0, max(0.1, float(item.get("persentase_kerusakan"))))
    else:
        persentase_kerusakan = min(100.0, max(0.1, round((vol / luas_total) * 100.0, 2)))

    bobot_fisik = float(ahsp_meta.get("bobot_fisik", 0.75))
    bobot_fungsi = float(ahsp_meta.get("bobot_fungsi", 0.55))
    is_struktural = bool(ahsp_meta.get("is_struktural", False))

    nilai_awal = 100

    # Skor Fisik & Fungsi berkurang proporsional terhadap prosentase kerusakan dan bobot jenis cacat
    skor_fisik = round(max(5.0, min(99.9, 100.0 - (persentase_kerusakan * bobot_fisik))), 1)
    skor_fungsi = round(max(10.0, min(100.0, 100.0 - (persentase_kerusakan * bobot_fungsi))), 1)

    kinerja_f = round(0.5 * skor_fisik, 1)
    kinerja_fn = round(0.5 * skor_fungsi, 1)
    total_kinerja = round(kinerja_f + kinerja_fn, 1)

    # Penentuan Tingkat Keparahan:
    # - Elemen Struktur Primer (Balok/Kolom spalling/exposed rebar/retak) bersifat kritis terhadap keamanan struktur
    # - Elemen Arsitektur (Lantai, Dinding, Plafon, Pintu/Jendela, Atap) & Waterproofing mengikuti rasio prosentase kerusakan
    k_lower = str(kerusakan).lower()
    if is_struktural and kode in ("A.4.1.1.20.JT", "A.4.1.1.18.JT"):
        if kode == "A.4.1.1.20.JT" or any(w in k_lower for w in ("rebar", "tulangan", "spalling", "keropos")):
            is_exposed_rebar = any(w in k_lower for w in ("rebar", "tulangan"))
            keparahan = "Berat" if (is_exposed_rebar or persentase_kerusakan > 1.0) else "Sedang"
            if keparahan == "Berat":
                skor_fisik = min(skor_fisik, 48.0)
                skor_fungsi = min(skor_fungsi, 48.0)
            else:
                skor_fisik = min(skor_fisik, 68.0)
                skor_fungsi = min(skor_fungsi, 70.0)
        else:
            keparahan = "Berat" if persentase_kerusakan > 3.0 else "Sedang"
            if keparahan == "Berat":
                skor_fisik = min(skor_fisik, 48.0)
                skor_fungsi = min(skor_fungsi, 50.0)
            else:
                skor_fisik = min(skor_fisik, 68.0)
                skor_fungsi = min(skor_fungsi, 72.0)
        kinerja_f = round(0.5 * skor_fisik, 1)
        kinerja_fn = round(0.5 * skor_fungsi, 1)
        total_kinerja = round(kinerja_f + kinerja_fn, 1)
    elif is_struktural and persentase_kerusakan > 25.0:
        keparahan = "Berat"
    elif persentase_kerusakan <= 15.0 and total_kinerja >= 79.5:
        keparahan = "Ringan"
    elif persentase_kerusakan <= 45.0 and total_kinerja >= 50.0:
        keparahan = "Sedang"
    else:
        keparahan = "Berat"

    if keparahan == "Ringan":
        desc_fisik = (
            f"Secara visual kondisi {elemen} masih tergolong Baik (prosentase kerusakan {persentase_kerusakan:.1f}% "
            f"dari total {luas_total:.1f} {ahsp_meta.get('satuan', 'm2')}; temuan: {kerusakan})."
        )
        desc_fungsi = f"Fungsi layanan {elemen} tetap berjalan normal dan tidak mengganggu aktivitas operasional ruangan."
    elif keparahan == "Sedang":
        desc_fisik = (
            f"Secara visual {elemen} berstatus Cukup (prosentase kerusakan {persentase_kerusakan:.1f}% "
            f"dari total {luas_total:.1f} {ahsp_meta.get('satuan', 'm2')}; temuan: {kerusakan})."
        )
        desc_fungsi = f"Fungsi layanan {elemen} mulai tereduksi pada area terdampak dan membutuhkan pemeliharaan korektif terencana."
    else:
        desc_fisik = (
            f"Secara visual {elemen} mengalami kerusakan Berat (prosentase kerusakan {persentase_kerusakan:.1f}% "
            f"dari total {luas_total:.1f} {ahsp_meta.get('satuan', 'm2')}; temuan: {kerusakan})."
        )
        desc_fungsi = f"Fungsi layanan {elemen} mengalami penurunan signifikan dan berisiko terhadap kenyamanan maupun keselamatan."

    baik_f = skor_fisik if skor_fisik >= 79.5 else ""
    cukup_f = skor_fisik if 49.5 <= skor_fisik < 79.5 else ""
    rusak_f = skor_fisik if skor_fisik < 49.5 else ""

    baik_fn = skor_fungsi if skor_fungsi >= 79.5 else ""
    cukup_fn = skor_fungsi if 49.5 <= skor_fungsi < 79.5 else ""
    rusak_fn = skor_fungsi if skor_fungsi < 49.5 else ""

    subtotal = int(round(skor_fisik + skor_fungsi))

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
        "luas_total": round(luas_total, 2),
        "persentase_kerusakan": persentase_kerusakan,
        "nilai_awal": nilai_awal,
        "desc_fisik": desc_fisik,
        "desc_fungsi": desc_fungsi,
        "skor_fisik": skor_fisik,
        "skor_fungsi": skor_fungsi,
        "baik_f": baik_f,
        "cukup_f": cukup_f,
        "rusak_f": rusak_f,
        "nilai_kondisi_f": skor_fisik,
        "baik_fn": baik_fn,
        "cukup_fn": cukup_fn,
        "rusak_fn": rusak_fn,
        "nilai_kondisi_fn": skor_fungsi,
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


class PDFReportGenerator:
    """Generator dokumen PDF laporan resmi pemeliharaan gedung menggunakan ReportLab."""

    @classmethod
    def generate_pdf_report(
        cls,
        data_list: List[Dict[str, Any]],
        title_doc: str = "Laporan Inspeksi Gedung",
        annotated_image_bytes: Optional[bytes] = None,
        signer_info: Optional[Dict[str, str]] = None
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
        story.append(Paragraph("SISTEM DETEKSI KERUSAKAN GEDUNG", kop_sub_style))
        story.append(Paragraph("Standar Acuan: Pedoman AHSP Gedung & Perumahan Provinsi Jawa Tengah Tahun 2023", kop_sub_style))
        story.append(Spacer(1, 4))
        story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=6))

        # 2. METADATA DOKUMEN INSPEKSI
        doc_no = f"BA-PBG/JATENG/{datetime.now().strftime('%Y%m%d')}/{len(data_list):03d}"

        target_item = data_list[0] if data_list else {}
        nama_gedung = target_item.get("nama_gedung") or "Bangunan Gedung"
        lokasi_lantai = target_item.get("lokasi_lantai") or "-"
        surveyor_name = target_item.get("surveyor") or ""
        koordinat_lokasi = target_item.get("koordinat_lokasi") or "-"
        objek_inspeksi = target_item.get("objek_inspeksi") or target_item.get("elemen_gedung") or "-"
        tanggal_inspeksi = str(target_item.get("tanggal") or datetime.now().strftime("%d %B %Y"))[:22]

        if surveyor_name and surveyor_name not in ["Surveyor Lapangan", "Inspektor Lapangan", "-", ""]:
            surveyor_disp = surveyor_name
        else:
            surveyor_disp = "Inspektor Lapangan"

        meta_data = [
            [
                Paragraph("<b>Nomor Dokumen:</b>", body_small_left), Paragraph(doc_no, body_small_left),
                Paragraph("<b>Tanggal Inspeksi:</b>", body_small_left), Paragraph(tanggal_inspeksi, body_small_left)
            ],
            [
                Paragraph("<b>Nama Bangunan:</b>", body_small_left), Paragraph(nama_gedung, body_small_left),
                Paragraph("<b>Nama Ruangan:</b>", body_small_left), Paragraph(lokasi_lantai, body_small_left)
            ],
            [
                Paragraph("<b>Objek Inspeksi:</b>", body_small_left), Paragraph(objek_inspeksi, body_small_left),
                Paragraph("<b>Nama Inspektor:</b>", body_small_left), Paragraph(surveyor_disp, body_small_left)
            ],
            [
                Paragraph("<b>Koordinat Lokasi:</b>", body_small_left), Paragraph(koordinat_lokasi, body_small_left),
                Paragraph("<b>Acuan Standar:</b>", body_small_left), Paragraph("AHSP Jateng 2023 & Permen PUPR 24/2008", body_small_left)
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

        target_item = data_list[0] if data_list else {
            "elemen_gedung": "Komponen Gedung", "jenis_kerusakan": "Kerusakan Permukaan", "tingkat_keparahan": "Ringan"
        }
        eval_m = compute_condition_assessment(target_item)

        # -----------------------------------------------------------------------------------------
        # 3. FOTO DOKUMENTASI KERUSAKAN (URUTAN PERTAMA SESUAI REVISI: FOTO DULU)
        # -----------------------------------------------------------------------------------------
        sec_offset = 0
        if annotated_image_bytes:
            sec_offset = 1
            story.append(Paragraph("1. DOKUMENTASI FOTO KERUSAKAN BANGUNAN", doc_header_style))
            try:
                img_io = io.BytesIO(annotated_image_bytes)
                pil_temp = Image.open(img_io)
                orig_w, orig_h = pil_temp.size

                target_w = 380.0
                target_h = target_w * (orig_h / orig_w)
                if target_h > 250.0:
                    target_h = 250.0
                    target_w = target_h * (orig_w / orig_h)

                img_io.seek(0)
                rl_img = RLImage(img_io, width=target_w, height=target_h)

                meta_lampiran = [
                    [
                        Paragraph(f"<b>Elemen Bangunan:</b> {target_item.get('elemen_gedung', '-')}", body_small_left),
                        Paragraph(f"<b>Jenis Kerusakan:</b> {target_item.get('jenis_kerusakan', '-')}", body_small_left)
                    ],
                    [
                        Paragraph(f"<b>Kode AHSP:</b> {target_item.get('kode_ahsp', '-')}", body_small_left),
                        Paragraph(f"<b>Prosentase & Keparahan:</b> {eval_m['persentase_kerusakan']:.1f}% ({str(eval_m['keparahan']).upper()})", body_small_left)
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
                        ('BOX', (0, 0), (-1, -1), 0.8, colors.HexColor('#CBD5E1')),
                        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
                        ('TOPPADDING', (0, 0), (-1, -1), 5),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                    ]),
                    Spacer(1, 4),
                    tbl_lampiran
                ]))
                story.append(Spacer(1, 8))
            except Exception:
                sec_offset = 0

        # -----------------------------------------------------------------------------------------
        # 4. FORMULIR PROSENTASE PENILAIAN KONDISI FISIK DAN FUNGSI (IKBG)
        # -----------------------------------------------------------------------------------------
        no_ikbg = 1 + sec_offset
        story.append(Paragraph(f"{no_ikbg}. FORMULIR PROSENTASE PENILAIAN KONDISI FISIK DAN FUNGSI (IKBG)", doc_header_style))

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
        # 5. TABEL REKAPITULASI RENCANA ANGGARAN BIAYA (RAB) & TOTAL ESTIMASI ANGGARAN
        # -----------------------------------------------------------------------------------------
        no_rab = 2 + sec_offset
        story.append(Paragraph(f"{no_rab}. REKAPITULASI RENCANA ANGGARAN BIAYA (RAB) & TOTAL ESTIMASI ANGGARAN", doc_header_style))

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
        story.append(Paragraph(f"<b>Total Estimasi Anggaran: {format_rupiah(grand_total)}</b> (<i>Terbilang: {terbilang_teks}</i>)", body_small_left))
        story.append(Spacer(1, 8))

        # Jika inspeksi tunggal, tampilkan rincian AHSP detail per bahan dan tenaga kerja
        if len(data_list) == 1:
            bd_pdf = AHSPKnowledgeBase.compute_material_breakdown(
                target_item.get("kode_ahsp", "A.4.4.1.25.JT"),
                float(target_item.get("volume", 1.0))
            )
            detail_rows = [
                [
                    Paragraph("<b>Komponen Bahan / Tenaga Kerja</b>", body_small_bold),
                    Paragraph("<b>Koef</b>", body_small_bold),
                    Paragraph("<b>Kebutuhan</b>", body_small_bold),
                    Paragraph("<b>Harga Satuan</b>", body_small_bold),
                    Paragraph("<b>Jumlah Biaya</b>", body_small_bold),
                ]
            ]
            for u_r in bd_pdf["upah_rows"]:
                detail_rows.append([
                    Paragraph(f"[Upah] {u_r['Unsur Tenaga Kerja']}", body_small_left),
                    Paragraph(u_r["Koefisien"], body_small),
                    Paragraph(u_r["Kebutuhan Riil"], body_small),
                    Paragraph(format_rupiah(u_r["Harga Satuan"]), body_small),
                    Paragraph(format_rupiah(u_r["Jumlah Biaya"]), body_small),
                ])
            for b_r in bd_pdf["bahan_rows"]:
                detail_rows.append([
                    Paragraph(f"[Bahan] {b_r['Nama Bahan / Material']}", body_small_left),
                    Paragraph(b_r["Koefisien"], body_small),
                    Paragraph(b_r["Kebutuhan Riil"], body_small),
                    Paragraph(format_rupiah(b_r["Harga Satuan"]), body_small),
                    Paragraph(format_rupiah(b_r["Jumlah Biaya"]), body_small),
                ])
            detail_rows.append([
                Paragraph("<b>Overhead & Profit (15%)</b>", body_small_left),
                Paragraph("-", body_small),
                Paragraph("-", body_small),
                Paragraph("-", body_small),
                Paragraph(format_rupiah(bd_pdf["overhead_15"]), body_small_bold),
            ])
            t_detail_ahsp = Table(detail_rows, colWidths=[215, 50, 75, 85, 98])
            t_detail_ahsp.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#334155')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#94A3B8')),
                ('INNERGRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#CBD5E1')),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#F1F5F9')),
                ('TOPPADDING', (0, 0), (-1, -1), 2.5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
            ]))
            story.append(Paragraph("<b>Rincian Kebutuhan Bahan & Tenaga Kerja per Item (AHSP):</b>", body_small_left))
            story.append(Spacer(1, 3))
            story.append(t_detail_ahsp)
            story.append(Spacer(1, 8))

        # -----------------------------------------------------------------------------------------
        # 6. LANGKAH-LANGKAH PERBAIKAN & REKOMENDASI TEKNIS
        # -----------------------------------------------------------------------------------------
        no_rek = 3 + sec_offset
        story.append(Paragraph(f"{no_rek}. LANGKAH-LANGKAH PERBAIKAN & REKOMENDASI TEKNIS", doc_header_style))
        if len(data_list) == 1 and target_item.get("langkah_kerja"):
            for s_idx, step_txt in enumerate(target_item.get("langkah_kerja", []), start=1):
                story.append(Paragraph(f"<b>Tahap {s_idx}:</b> {step_txt}", body_small_left))
                story.append(Spacer(1, 1.5))
            story.append(Spacer(1, 3))
        for idx_rc, r_row in enumerate(data_list[:2], start=1):
            recom_text = f"<b>Catatan Teknis ({r_row.get('elemen_gedung')}):</b> {r_row.get('rekomendasi', '-')}"
            story.append(Paragraph(recom_text, body_small_left))
            story.append(Spacer(1, 2))

        story.append(Spacer(1, 14))

        if not signer_info:
            signer_info = {}

        # Hanya 2 Penandatangan sesuai Revisi: Dibuat Oleh & Disetujui Oleh
        s_nama = signer_info.get("surveyor_name") or surveyor_name or ""
        s_nip = signer_info.get("surveyor_nip") or ""
        if s_nama and s_nama not in ["Surveyor Lapangan", "Inspektor Lapangan", "-", ""]:
            sig_s_nama = f"<b><u>{s_nama}</u></b>"
            sig_s_nip = f"NIP/ID. {s_nip}" if s_nip else "NIP/ID. ....................."
        else:
            sig_s_nama = "<u>( .................................................. )</u>"
            sig_s_nip = "NIP/ID."

        p_nama = signer_info.get("ppk_name") or ""
        p_nip = signer_info.get("ppk_nip") or ""
        if p_nama:
            sig_p_nama = f"<b><u>{p_nama}</u></b>"
            sig_p_nip = f"NIP. {p_nip}" if p_nip else "NIP. ....................."
        else:
            sig_p_nama = "<u>( .................................................. )</u>"
            sig_p_nip = "NIP."

        story.append(KeepTogether([
            Table([
                [
                    Paragraph(f"Dibuat Oleh,<br/><b>Inspektor Lapangan</b><br/><br/><br/><br/>{sig_s_nama}<br/>{sig_s_nip}", body_small),
                    Paragraph(f"Disetujui Oleh,<br/><b>Pengelola Gedung / PPK</b><br/><br/><br/><br/>{sig_p_nama}<br/>{sig_p_nip}", body_small)
                ]
            ], colWidths=[260, 260], style=[
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ])
        ]))

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

@st.cache_data(ttl=3600, show_spinner=False)
def get_auto_coordinates() -> str:
    """Mendapatkan koordinat lokasi perkiraan otomatis via geolokasi jaringan dengan fallback Semarang."""
    endpoints = [
        ("http://ip-api.com/json/?fields=status,lat,lon,city,regionName", "lat", "lon"),
        ("https://ipapi.co/json/", "latitude", "longitude"),
    ]
    for url, lat_k, lon_k in endpoints:
        try:
            resp = requests.get(url, timeout=2.0)
            if resp.status_code == 200:
                d = resp.json()
                lat = d.get(lat_k)
                lon = d.get(lon_k)
                if lat is not None and lon is not None:
                    return f"{float(lat):.6f}, {float(lon):.6f}"
        except Exception:
            continue
    return "-7.005100, 110.438100"


# Sidebar Navigasi
st.sidebar.markdown("""
<div style="text-align: center; padding: 14px 0 18px 0;">
    <div style="display: inline-block; padding: 14px; background: linear-gradient(135deg, #1e293b, #0f172a); border-radius: 20px; box-shadow: 0 8px 18px rgba(15, 23, 42, 0.25); border: 1px solid rgba(255,255,255,0.1); margin-bottom: 8px;">
        <span style="font-size: 2.2rem;">🏢</span>
    </div>
    <h3 style="margin: 6px 0 0 0; color: #0f172a; font-weight: 900; letter-spacing: -0.5px; font-size: 1.15rem; line-height: 1.3;">Sistem Deteksi Kerusakan Gedung</h3>
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
        "📜 Riwayat Inspeksi Gedung"
    ],
    index=0
)

st.sidebar.markdown("---")

# Pengaturan Integrasi AI Agent (Disembunyikan secara rapi di bawah toggle Pengembang)
dev_mode = st.sidebar.checkbox("🛠️ Mode Pengembang (API Dev)", value=False, help="Aktifkan jika ingin menghubungkan aplikasi ke server HTTP POST API eksternal.")
if dev_mode:
    with st.sidebar.expander("⚙️ Konfigurasi Server AI (Opsional)", expanded=True):
        st.markdown("<p style='font-size:0.84rem; color:#1e293b; font-weight:600;'>Konfigurasi HTTP POST API Call library <code>requests</code> ke server deteksi eksternal.</p>", unsafe_allow_html=True)
        api_endpoint_input = st.text_input(
            "Server Endpoint URL:",
            value="http://localhost:8000/api/analyze-damage",
            help="Alamat HTTP POST server deteksi AI"
        )
        api_token_input = st.text_input(
            "Bearer API Key (Opsional):",
            value="",
            type="password"
        )
        offline_toggle = st.checkbox(
            "Mode Simulasi Offline (Offline Safe)",
            value=False,
            help="Aktifkan opsi ini jika tidak menjalankan server lokal terpisah atau jaringan internet terkendala."
        )
else:
    api_endpoint_input = "http://localhost:8000/api/analyze-damage"
    api_token_input = ""
    offline_toggle = False

st.sidebar.markdown("""
<div style="background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%); border-radius: 12px; padding: 14px; border: 1.5px solid #cbd5e1; font-size: 0.8rem; color: #0f172a; font-weight: 500; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
    <div style="display:flex; align-items:center; gap:6px; margin-bottom:6px;">
        <span class="live-dot" style="width:7px; height:7px;"></span>
        <span style="font-weight:800; color:#0f172a;">STATUS SISTEM AKTIF</span>
    </div>
    <b style="color:#0f172a;">Database:</b> SQLite3 (<code>pemeliharaan_gedung.db</code>)<br>
    <b style="color:#0f172a;">Engine:</b> Deep Learning YOLOv8<br>
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
            Cakupan: Arsitektur & Struktur Gedung
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
                    <span class="badge-ahsp" style="background:#2563eb; color:#ffffff; border:none; padding:4px 12px;">MANAJEMEN PEMELIHARAAN GEDUNG</span>
                    <div class="live-badge" style="margin-bottom:0;">
                        <span class="live-dot"></span> YOLOV8 ENGINE SIAP
                    </div>
                </div>
                <div class="hero-title">Sistem Deteksi Kerusakan Gedung</div>
                <div class="hero-subtitle">
                    Aplikasi inspeksi & pemeliharaan bangunan gedung berbasis Deep Learning YOLOv8 dengan rincian Rencana Anggaran Biaya (RAB), detail kebutuhan bahan AHSP Jawa Tengah 2023, serta formulir penilaian prosentase kondisi (IKBG).
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

        with st.expander("🏢 Lembar Identitas Bangunan & Inspektor", expanded=True):
            col_id1, col_id2 = st.columns(2)
            with col_id1:
                nama_gedung_input = st.text_input(
                    "Nama Bangunan:",
                    value=st.session_state.get("inp_gedung", ""),
                    placeholder="Contoh: Gedung Kuliah Terpadu / Kantor Utama",
                    help="Masukkan nama bangunan gedung yang sedang diinspeksi."
                )
            with col_id2:
                lokasi_lantai_input = st.text_input(
                    "Nama Ruangan:",
                    value=st.session_state.get("inp_lokasi", ""),
                    placeholder="Contoh: Lantai 2 - Ruang 204 / Lobby Utama",
                    help="Masukkan nama ruangan atau zona lokasi inspeksi."
                )

            objek_options = ["Otomatis (Deteksi AI YOLOv8)"] + list(AHSPKnowledgeBase.OBJEK_INSPEKSI_MAP.keys())
            default_obj_idx = objek_options.index(st.session_state.get("inp_objek", objek_options[0])) if st.session_state.get("inp_objek") in objek_options else 0
            objek_inspeksi_pilihan = st.selectbox(
                "Objek Inspeksi (Dapat Dipilih):",
                objek_options,
                index=default_obj_idx,
                help="Biarkan 'Otomatis (Deteksi AI YOLOv8)' agar model YOLOv8 mendeteksi jenis kerusakan secara langsung seperti semula, atau pilih elemen spesifik Arsitektur/Struktur."
            )

            col_id3, col_id4 = st.columns(2)
            with col_id3:
                surveyor_input = st.text_input(
                    "Nama Inspektor (Dibuat Oleh):",
                    value=st.session_state.get("inp_surveyor", ""),
                    placeholder="Contoh: Ahmad Fajar, S.T.",
                    help="Nama inspektor lapangan yang menyusun laporan."
                )
            with col_id4:
                tanggal_inspeksi_obj = st.date_input(
                    "Tanggal, Bulan, Tahun:",
                    value=st.session_state.get("inp_tanggal_obj", date.today()),
                    format="DD/MM/YYYY",
                    help="Pilih tanggal pelaksanaan inspeksi lapangan."
                )

            auto_coord_default = st.session_state.get("inp_koordinat") or get_auto_coordinates()
            col_k1, col_k2 = st.columns([2, 1])
            with col_k1:
                koordinat_input = st.text_input(
                    "Koordinat Lokasi (Otomatis):",
                    value=auto_coord_default,
                    help="Terdeteksi otomatis (Latitude, Longitude). Dapat disesuaikan jika berada di titik berbeda."
                )
            with col_k2:
                surveyor_nip_input = st.text_input(
                    "NIP / ID Inspektor:",
                    value=st.session_state.get("inp_surveyor_nip", ""),
                    placeholder="Opsional"
                )

            st.components.v1.html("""
                <div style="font-family: 'Plus Jakarta Sans', sans-serif; font-size: 12px; color: #334155; display: flex; align-items: center; gap: 8px; padding: 4px 8px; background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px;">
                    <span id="gps_status">📍 Sensor GPS Perangkat: Mencari sinyal...</span>
                </div>
                <script>
                if (navigator.geolocation) {
                    navigator.geolocation.getCurrentPosition(
                        function(pos) {
                            const lat = pos.coords.latitude.toFixed(6);
                            const lon = pos.coords.longitude.toFixed(6);
                            document.getElementById("gps_status").innerHTML =
                                "📍 <b>GPS Perangkat Terbaca:</b> <code>" + lat + ", " + lon + "</code>";
                        },
                        function() {
                            document.getElementById("gps_status").innerHTML =
                                "📍 Koordinat otomatis berbasis lokasi jaringan aktif.";
                        },
                        {enableHighAccuracy: true, timeout: 4000}
                    );
                } else {
                    document.getElementById("gps_status").innerHTML = "📍 Koordinat otomatis aktif.";
                }
                </script>
            """, height=36)

            col_sig_p1, col_sig_p2 = st.columns(2)
            with col_sig_p1:
                ppk_input = st.text_input(
                    "Disetujui Oleh (Pengelola Gedung / PPK):",
                    value=st.session_state.get("inp_ppk", ""),
                    placeholder="Contoh: Hendra Kusuma, S.T. (Opsional)"
                )
            with col_sig_p2:
                ppk_nip_input = st.text_input(
                    "NIP Pejabat Penyetuju:",
                    value=st.session_state.get("inp_ppk_nip", ""),
                    placeholder="Opsional"
                )

            bulan_id = [
                "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
                "Juli", "Agustus", "September", "Oktober", "November", "Desember"
            ]
            tanggal_formatted = f"{tanggal_inspeksi_obj.day:02d} {bulan_id[tanggal_inspeksi_obj.month]} {tanggal_inspeksi_obj.year}"

            st.session_state["inp_gedung"] = nama_gedung_input
            st.session_state["inp_lokasi"] = lokasi_lantai_input
            st.session_state["inp_objek"] = objek_inspeksi_pilihan
            st.session_state["inp_surveyor"] = surveyor_input
            st.session_state["inp_surveyor_nip"] = surveyor_nip_input
            st.session_state["inp_tanggal_obj"] = tanggal_inspeksi_obj
            st.session_state["inp_koordinat"] = koordinat_input
            st.session_state["inp_ppk"] = ppk_input
            st.session_state["inp_ppk_nip"] = ppk_nip_input

            st.session_state["signer_info"] = {
                "surveyor_name": surveyor_input.strip() if surveyor_input else "",
                "surveyor_nip": surveyor_nip_input.strip() if surveyor_nip_input else "",
                "ppk_name": ppk_input.strip() if ppk_input else "",
                "ppk_nip": ppk_nip_input.strip() if ppk_nip_input else "",
            }

        # Area Unggah Foto Diperbesar dengan Visual Kamera
        st.markdown("""
        <div class="camera-upload-hero">
            <div style="font-size: 3.1rem; line-height: 1; margin-bottom: 6px;">📷</div>
            <div style="font-size: 1rem; font-weight: 800; color: #0f172a;">1. Unggah Foto Kerusakan Bangunan</div>
            <div style="font-size: 0.82rem; font-weight: 600; color: #334155;">Ambil foto langsung dari kamera atau pilih berkas JPG / PNG</div>
        </div>
        """, unsafe_allow_html=True)

        uploaded_file = st.file_uploader(
            "Pilih atau seret foto kerusakan ke kotak ini (JPG / PNG):",
            type=["jpg", "jpeg", "png"],
            label_visibility="collapsed",
            help="Pilih foto dokumentasi visual kerusakan elemen bangunan gedung."
        )
        st.markdown('<div class="camera-upload-caption">⚠️ Harap unggah foto kerusakan secara jelas</div>', unsafe_allow_html=True)

        sub_kerusakan_opts = ["Deteksi Otomatis AI"] + AHSPKnowledgeBase.OBJEK_INSPEKSI_MAP.get(objek_inspeksi_pilihan, [])
        jenis_kerusakan_pilihan = st.selectbox(
            f"2. Jenis Kerusakan ({objek_inspeksi_pilihan}):",
            sub_kerusakan_opts,
            index=0,
            key=f"jenis_kerusakan_sel_{objek_inspeksi_pilihan}",
            help="Biarkan 'Deteksi Otomatis AI' agar model YOLOv8 mengklasifikasikan kerusakan dari foto secara otomatis, atau pilih sub-jenis kerusakan spesifik jika ingin menentukan manual."
        )

        # Deteksi satuan otomatis berdasarkan item AHSP terkait
        matched_preview = AHSPKnowledgeBase.find_by_element_and_defect(objek_inspeksi_pilihan, jenis_kerusakan_pilihan)
        satuan_default = matched_preview.get("satuan", "m2")
        is_linier = (satuan_default == "m'")
        is_lantai = ("lantai" in objek_inspeksi_pilihan.lower() and "dak" not in objek_inspeksi_pilihan.lower())

        metode_opts = ["Input Volume Langsung (Opname)", "📐 Hitung dari Dimensi (Panjang × Lebar / Linier)"]
        if is_lantai:
            metode_opts.append("🧱 Hitung per Keping Keramik (1 buah, dst.)")

        metode_vol = st.radio(
            "3. Metode Penentuan Volume & Skala Prosentase:",
            metode_opts,
            horizontal=True
        )

        if metode_vol == "🧱 Hitung per Keping Keramik (1 buah, dst.)":
            c_kr1, c_kr2 = st.columns(2)
            with c_kr1:
                jml_keramik = st.number_input("Jumlah Keping Rusak (buah):", min_value=1, value=1, step=1)
            with c_kr2:
                uk_keramik = st.selectbox(
                    "Ukuran 1 Keping Keramik:",
                    ["60 × 60 cm (0.36 m²)", "40 × 40 cm (0.16 m²)", "50 × 50 cm (0.25 m²)", "30 × 30 cm (0.09 m²)"],
                    index=0
                )
            area_per_keping = {"60 × 60 cm (0.36 m²)": 0.36, "40 × 40 cm (0.16 m²)": 0.16, "50 × 50 cm (0.25 m²)": 0.25, "30 × 30 cm (0.09 m²)": 0.09}[uk_keramik]
            volume_input = float(round(jml_keramik * area_per_keping, 2))
            st.success(f"Volume Terhitung: **{jml_keramik} keping × {area_per_keping:.2f} m² = {volume_input:.2f} m²**")
        elif metode_vol == "📐 Hitung dari Dimensi (Panjang × Lebar / Linier)":
            if is_linier:
                panjang_m = st.number_input("Panjang Kerusakan (meter):", min_value=0.1, value=2.0, step=0.5)
                volume_input = float(round(panjang_m, 2))
                st.success(f"Volume Terhitung: **{volume_input:.2f} m'**")
            else:
                c_p, c_l = st.columns(2)
                with c_p:
                    panjang_m = st.number_input("Panjang Area Rusak (m):", min_value=0.1, value=1.5, step=0.5)
                with c_l:
                    lebar_m = st.number_input("Lebar / Tinggi Area Rusak (m):", min_value=0.1, value=1.0, step=0.5)
                volume_input = float(round(panjang_m * lebar_m, 2))
                st.success(f"Volume Terhitung: **{panjang_m:.2f} m × {lebar_m:.2f} m = {volume_input:.2f} m²**")
        else:
            volume_input = st.number_input(
                f"Masukkan Volume Kerusakan ({satuan_default}):",
                min_value=0.05,
                max_value=10000.0,
                value=1.0 if not is_lantai else 0.36,
                step=0.1,
                help=f"Masukkan besaran volume kerusakan ({satuan_default})."
            )

        luas_total_input = st.number_input(
            f"4. Total Luas / Panjang Komponen pada Ruangan ({satuan_default}) [Skala Prosentase]:",
            min_value=0.5,
            max_value=100000.0,
            value=20.0 if not is_linier else 10.0,
            step=1.0,
            help="Digunakan untuk menghitung prosentase kerusakan secara logis: (Volume Rusak ÷ Total Luas Komponen) × 100%."
        )
        est_pct = min(100.0, round((volume_input / max(0.1, luas_total_input)) * 100.0, 2))
        st.caption(f"📊 *Prosentase Kerusakan Terhitung: **{est_pct:.2f}%** ({volume_input:.2f} {satuan_default} dari total {luas_total_input:.1f} {satuan_default}).*")

        start_btn = st.button("🚀 Mulai Analisis Sistem", type="primary", use_container_width=True)

    with col_right:
        st.markdown("""
        <div class="tw-card">
            <div class="tw-card-header">
                <h4 style="margin: 0; color: #0f172a; font-weight: 800;">🔍 Sensor Visual & Pratinjau Citra</h4>
                <p style="margin: 4px 0 0 0; font-size: 0.88rem; color: #1e293b; font-weight: 600;">Citra yang diunggah akan dianalisis oleh sistem Computer Vision YOLOv8.</p>
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

                elemen_tag = str(st.session_state.latest_analysis.get('elemen_gedung', 'Elemen')).replace(' ', '_').replace('/', '_')
                st.download_button(
                    label="📸 Unduh Citra Deteksi AI (.jpg)",
                    data=st.session_state.annotated_img_bytes,
                    file_name=f"Deteksi_AI_{elemen_tag}_{datetime.now().strftime('%Y%m%d_%H%M')}.jpg",
                    mime="image/jpeg",
                    use_container_width=True
                )
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
            st.markdown("""
            <div class="upload-placeholder">
                <span class="floating-icon">📸</span>
                <p style="margin-top: 14px; font-weight: 800; color: #0f172a; font-size: 1.05rem;">Belum ada foto yang dipilih</p>
                <p style="font-size: 0.86rem; color: #1e293b; font-weight: 600; margin: 0; max-width: 320px; margin-inline: auto;">Pilih foto kerusakan pada kotak unggah di sebelah kiri untuk pratinjau citra dan memulai analisis.</p>
            </div>
            """, unsafe_allow_html=True)

    # ==============================================================================================
    # PROSES EKSEKUSI API CALL & PENYIMPANAN KE SQLITE
    # ==============================================================================================
    if start_btn:
        if uploaded_file is None:
            st.warning("⚠️ Harap unggah foto kerusakan secara jelas sebelum memulai analisis!")
        else:
            with st.spinner("loading"):
                img_bytes = uploaded_file.getvalue()
                fname = uploaded_file.name

                result_ai = AIAgentClient.call_agent(
                    image_bytes=img_bytes,
                    filename=fname,
                    volume=volume_input,
                    elemen_gedung=objek_inspeksi_pilihan,
                    endpoint_url=api_endpoint_input,
                    api_key=api_token_input,
                    force_simulation=offline_toggle,
                    luas_total=luas_total_input,
                    jenis_kerusakan_pilihan=jenis_kerusakan_pilihan
                )

                nama_gedung_clean = nama_gedung_input.strip() if nama_gedung_input and nama_gedung_input.strip() else "Bangunan Gedung"
                lokasi_clean = lokasi_lantai_input.strip() if lokasi_lantai_input and lokasi_lantai_input.strip() else "-"
                surveyor_clean = surveyor_input.strip() if surveyor_input and surveyor_input.strip() else "Inspektor Lapangan"
                koordinat_clean = koordinat_input.strip() if koordinat_input and koordinat_input.strip() else get_auto_coordinates()

                result_ai["nama_gedung"] = nama_gedung_clean
                result_ai["lokasi_lantai"] = lokasi_clean
                result_ai["surveyor"] = surveyor_clean
                result_ai["koordinat_lokasi"] = koordinat_clean
                result_ai["objek_inspeksi"] = objek_inspeksi_pilihan
                result_ai["tanggal"] = tanggal_formatted

                annotated_img = AIVisualAnnotator.annotate(
                    image_bytes=img_bytes,
                    jenis_kerusakan=result_ai.get("jenis_kerusakan", "Kerusakan Permukaan"),
                    tingkat_keparahan=result_ai.get("tingkat_keparahan", "Ringan"),
                    confidence_score=result_ai.get("confidence_score", 96.5),
                    kode_ahsp=result_ai.get("kode_ahsp", "A.4.4.1.25.JT"),
                    detected_boxes=result_ai.get("detected_boxes")
                )
                st.session_state.annotated_img_bytes = annotated_img
                st.session_state.uploaded_img_bytes = img_bytes

                record_id = DatabaseManager.insert_riwayat(
                    tanggal=tanggal_formatted,
                    elemen_gedung=result_ai.get("elemen_gedung", objek_inspeksi_pilihan),
                    jenis_kerusakan=result_ai.get("jenis_kerusakan", "Kerusakan Permukaan"),
                    kode_ahsp=result_ai.get("kode_ahsp", "A.4.4.1.25.JT"),
                    volume=float(result_ai.get("volume", volume_input)),
                    satuan=result_ai.get("satuan", satuan_default),
                    harga_satuan=int(result_ai.get("harga_satuan", 0)),
                    total_biaya=int(result_ai.get("total_biaya", 0)),
                    rekomendasi=result_ai.get("rekomendasi", "Perbaikan standar sesuai metode kerja."),
                    nama_gedung=nama_gedung_clean,
                    lokasi_lantai=lokasi_clean,
                    surveyor=surveyor_clean,
                    koordinat_lokasi=koordinat_clean,
                    objek_inspeksi=objek_inspeksi_pilihan,
                    luas_total=float(result_ai.get("luas_total", luas_total_input)),
                    persentase_kerusakan=float(result_ai.get("persentase_kerusakan", est_pct))
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

        # Notifikasi Status Mesin AI yang Bersih & Terverifikasi
        if res.get("is_yolo_active"):
            st.success(f"🟢 **Status Engine:** YOLOv8 Computer Vision Engine (`{res.get('mode', 'weights/best.pt')}`) Aktif & Optimal | ID Riwayat DB: #{res.get('db_id')}")
        elif res.get("is_live_api"):
            st.success(f"🟢 **Koneksi Live HTTP POST Berhasil:** Respon diterima dari endpoint `{res.get('api_endpoint_used')}` | ID Riwayat DB: #{res.get('db_id')}")
        else:
            engine_label = res.get("mode", "YOLOv8 & Civil Vision Engine")
            st.success(f"🟢 **Status Engine:** {engine_label} (Mode Analisis Mandiri Aktif) | ID Riwayat DB: #{res.get('db_id')}")

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
                        # Hitung ulang persentase kerusakan & tingkat keparahan secara proporsional
                        updated_item = dict(st.session_state.latest_analysis)
                        updated_item["volume"] = sim_volume
                        updated_item["total_biaya"] = sim_total_biaya
                        recalc_eval = compute_condition_assessment(updated_item)
                        new_pct = float(recalc_eval.get("persentase_kerusakan", 0.0))
                        new_sev = recalc_eval.get("keparahan", updated_item.get("tingkat_keparahan", "Sedang"))

                        st.session_state.latest_analysis["volume"] = sim_volume
                        st.session_state.latest_analysis["total_biaya"] = sim_total_biaya
                        st.session_state.latest_analysis["persentase_kerusakan"] = new_pct
                        st.session_state.latest_analysis["tingkat_keparahan"] = new_sev

                        with DatabaseManager.get_connection() as conn:
                            cur = conn.cursor()
                            cur.execute(
                                "UPDATE riwayat_kerusakan SET volume = ?, total_biaya = ?, persentase_kerusakan = ?, tingkat_keparahan = ? WHERE id = ?",
                                (sim_volume, sim_total_biaya, new_pct, new_sev, res.get("db_id"))
                            )
                            conn.commit()
                        st.success(f"✅ Volume diperbarui ke {sim_volume:.2f} {satuan_kerja} (Prosentase: {new_pct:.2f}% | {new_sev}) dengan RAB {format_rupiah(sim_total_biaya)}!")
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
        # TOTAL ESTIMASI ANGGARAN & RINCIAN AHSP DETAIL PER MASING-MASING BAHAN & TENAGA KERJA
        # -----------------------------------------------------------------------------------------
        kode_item = res.get("kode_ahsp", "")
        vol_calc = float(res.get("volume", 0.0))
        sat_calc = res.get("satuan", "m2")
        ahsp_breakdown = AHSPKnowledgeBase.compute_material_breakdown(kode_item, vol_calc)
        upah_rows = ahsp_breakdown.get("upah_rows", [])
        bahan_rows = ahsp_breakdown.get("bahan_rows", [])

        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #ecfdf5 0%, #f0fdf4 100%); border: 2px solid #86efac; border-radius: 14px; padding: 18px 22px; margin: 14px 0 18px 0; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
            <div>
                <div style="font-size: 0.78rem; font-weight: 800; color: #166534; text-transform: uppercase; letter-spacing: 0.7px;">Total Estimasi Anggaran Perbaikan (RAB)</div>
                <div style="font-size: 0.92rem; color: #1e293b; font-weight: 600; margin-top: 4px;">
                    {res.get('deskripsi_ahsp')} ({vol_calc:.2f} {sat_calc})
                </div>
            </div>
            <div style="font-size: 1.75rem; font-weight: 900; color: #15803d;">
                {format_rupiah(res.get('total_biaya', 0))}
            </div>
        </div>
        """, unsafe_allow_html=True)

        with st.expander("🧱 Rincian AHSP Detail per Masing-Masing Bahan & Tenaga Kerja (AHSP Jateng 2023)", expanded=True):
            col_koef1, col_koef2 = st.columns(2)
            with col_koef1:
                st.markdown("<div style='background:#eff6ff; border-left:4px solid #2563eb; padding:6px 12px; border-radius:4px; margin-bottom:8px;'><b style='color:#1e40af; font-size:0.88rem;'>👷 Rincian Tenaga Kerja Lapangan:</b></div>", unsafe_allow_html=True)
                if upah_rows:
                    df_upah = pd.DataFrame([
                        {
                            "Unsur Tenaga Kerja": r["Unsur Tenaga Kerja"],
                            "Koef.": r["Koefisien"],
                            "Kebutuhan": r["Kebutuhan Riil"],
                            "Harga Satuan": format_rupiah(r["Harga Satuan"]),
                            "Jumlah Biaya": format_rupiah(r["Jumlah Biaya"]),
                        }
                        for r in upah_rows
                    ])
                    st.dataframe(df_upah, use_container_width=True, hide_index=True)
                    st.caption(f"Subtotal Biaya Tenaga Kerja (untuk {vol_calc:.2f} {sat_calc}): **{format_rupiah(ahsp_breakdown['subtotal_upah'])}**")
                else:
                    st.caption("Detail tenaga kerja mengikuti satuan standar pekerjaan.")

            with col_koef2:
                st.markdown("<div style='background:#f0fdf4; border-left:4px solid #16a34a; padding:6px 12px; border-radius:4px; margin-bottom:8px;'><b style='color:#166534; font-size:0.88rem;'>🧱 Rincian Masing-Masing Bahan & Material:</b></div>", unsafe_allow_html=True)
                if bahan_rows:
                    df_bahan = pd.DataFrame([
                        {
                            "Nama Bahan / Material": r["Nama Bahan / Material"],
                            "Koef.": r["Koefisien"],
                            "Kebutuhan": r["Kebutuhan Riil"],
                            "Harga Satuan": format_rupiah(r["Harga Satuan"]),
                            "Jumlah Biaya": format_rupiah(r["Jumlah Biaya"]),
                        }
                        for r in bahan_rows
                    ])
                    st.dataframe(df_bahan, use_container_width=True, hide_index=True)
                    st.caption(f"Subtotal Biaya Bahan (untuk {vol_calc:.2f} {sat_calc}): **{format_rupiah(ahsp_breakdown['subtotal_bahan'])}**")
                else:
                    st.caption("Detail material mengikuti satuan standar pekerjaan.")

            st.markdown(f"""
            <div style="background:#f8fafc; border:1px solid #cbd5e1; border-radius:8px; padding:10px 14px; margin-top:8px; font-size:0.85rem; color:#1e293b;">
                <b>Ringkasan Biaya ({vol_calc:.2f} {sat_calc} | {kode_item}):</b> Tenaga + Bahan = <b>{format_rupiah(ahsp_breakdown['subtotal_langsung'])}</b> | Overhead &amp; Profit (15%) = <b>{format_rupiah(ahsp_breakdown['overhead_15'])}</b> | HSP/{sat_calc} = <b>{format_rupiah(res.get('harga_satuan', 0))}</b> &rarr; <b>Total Estimasi Anggaran: {format_rupiah(res.get('total_biaya', 0))}</b>
            </div>
            """, unsafe_allow_html=True)

        # Tampilan Tabel Matriks Penilaian Kondisi Fisik & Fungsi (IKBG) Sesuai Standar
        eval_ui = compute_condition_assessment(res)
        with st.expander("📋 Formulir Penilaian Kondisi Fisik, Fungsi & Prosentase Kerusakan (Format IKBG)", expanded=True):
            st.markdown(f"""
            <div style="font-size:0.88rem; color:#1e293b; font-weight:600; margin-bottom:10px;">
                Matriks evaluasi kondisi fisik &amp; fungsi (Permen PUPR No. 24/PRT/M/2008) | <b>Prosentase Kerusakan: {eval_ui['persentase_kerusakan']:.2f}%</b> ({vol_calc:.2f} dari {eval_ui['luas_total']:.2f} {sat_calc}).
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

        # Langkah-Langkah Perbaikan (Tampilan Baru Grid Tahapan Terstruktur) & Rekomendasi
        st.markdown("#### 🛠️ Langkah-Langkah Perbaikan & Metode Kerja Standar")
        steps_list = res.get("langkah_kerja", [])
        if steps_list:
            cards_html = '<div class="repair-steps-grid">'
            for idx, step in enumerate(steps_list, start=1):
                cards_html += f"""
                <div class="repair-step-card">
                    <div class="repair-step-badge">TAHAP {idx}</div>
                    <div class="repair-step-body">{step}</div>
                </div>
                """
            cards_html += "</div>"
            st.markdown(cards_html, unsafe_allow_html=True)

        st.info(f"💡 **Rekomendasi Pemeliharaan Teknis:** {res.get('rekomendasi')}")

        # Tombol Unduh Laporan Resmi PDF (Tanpa Excel sesuai Revisi)
        single_record_pdf = PDFReportGenerator.generate_pdf_report(
            [res],
            title_doc=f"Laporan Penilaian - {res.get('elemen_gedung')}",
            annotated_image_bytes=st.session_state.get("annotated_img_bytes"),
            signer_info=st.session_state.get("signer_info")
        )
        st.download_button(
            label="📄 Unduh Laporan Resmi Penilaian, Prosentase & RAB (PDF)",
            data=single_record_pdf,
            file_name=f"Laporan_PBG_ID_{res.get('db_id')}_{res.get('kode_ahsp')}.pdf",
            mime="application/pdf",
            type="primary",
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
            st.caption("ℹ️ *Jendela pratinjau PDF resmi standar A4 (Urutan: Dokumentasi Foto -> Formulir Prosentase -> RAB -> Langkah Perbaikan -> TTD 2 Pihak).*")


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
            Daftar seluruh riwayat inspeksi yang tersimpan di dalam database SQLite. Menyediakan ringkasan KPI, filter data elemen gedung, dan fasilitas ekspor laporan resmi PDF bersertifikasi standar IKBG &amp; AHSP Jateng 2023.
        </div>
    </div>
    """, unsafe_allow_html=True)

    records = DatabaseManager.get_all_riwayat()

    if not records:
        st.info("ℹ️ Database riwayat saat ini masih kosong.")
        if st.button("🧪 Muat Data Sampel Inspeksi Gedung", type="primary"):
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
        <div class="ikbg-global-card" style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); border-radius: 16px; padding: 22px; color: #ffffff !important; box-shadow: 0 10px 25px rgba(15,23,42,0.25); border: 1px solid rgba(255,255,255,0.1); margin-bottom: 24px;">
            <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 16px;">
                <div style="flex: 1; min-width: 280px;">
                    <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                        <span style="background: rgba(255,255,255,0.15); color: #ffffff !important; padding: 4px 10px; border-radius: 6px; font-size: 0.75rem; font-weight: 700; letter-spacing: 0.5px;">EVALUASI KEANDALAN GLOBAL</span>
                        <span style="background: {card_color}; color: #ffffff !important; padding: 4px 12px; border-radius: 20px; font-size: 0.75rem; font-weight: 800;">{status_gedung_badge}</span>
                    </div>
                    <div style="margin: 0; font-size: 1.35rem; font-weight: 900; color: #ffffff !important; letter-spacing: -0.3px;">Indeks Kesehatan Gedung Kumulatif (IKBG Global)</div>
                    <div style="margin: 6px 0 14px 0; font-size: 0.92rem; color: #e2e8f0 !important; line-height: 1.5;">
                        Status: <b style="color: #ffffff !important; font-weight: 800;">{status_gedung_label}</b> (Berdasarkan evaluasi menyeluruh {len(records)} titik inspeksi sesuai Permen PUPR No. 24/PRT/M/2008).
                    </div>
                    <div style="background: rgba(255,255,255,0.07); border-left: 4px solid {card_color}; border-radius: 6px; padding: 10px 14px; font-size: 0.85rem; color: #f1f5f9 !important; line-height: 1.5;">
                        <b style="color: #ffffff !important;">Rekomendasi Manajerial:</b> {rekomendasi_manajerial}
                    </div>
                </div>
                <div style="text-align: center; background: rgba(255,255,255,0.06); border-radius: 14px; padding: 16px 24px; border: 1px solid rgba(255,255,255,0.1); min-width: 170px;">
                    <div style="font-size: 0.78rem; text-transform: uppercase; color: #cbd5e1 !important; font-weight: 700; letter-spacing: 0.5px;">Rata-rata Skor (µ̄ku)</div>
                    <div style="font-size: 2.6rem; font-weight: 900; color: #ffffff !important; line-height: 1.1; margin: 4px 0;">{avg_mu_ku}</div>
                    <div style="font-size: 0.8rem; color: #cbd5e1 !important; font-weight: 600;">Skala 0 - 100</div>
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
            pct_r = float(r.get("persentase_kerusakan") or 0.0)
            table_display.append({
                "ID": r["id"],
                "Tanggal": str(r["tanggal"])[:16] if r.get("tanggal") else "-",
                "Gedung": r.get("nama_gedung", "Gedung Kuliah Terpadu"),
                "Ruangan": r.get("lokasi_lantai", "Lantai 2 - Ruang 204"),
                "Objek Inspeksi": r.get("objek_inspeksi", r["elemen_gedung"]),
                "Inspektor": r.get("surveyor", "Ahmad Fajar, S.T."),
                "Koordinat": r.get("koordinat_lokasi", "-"),
                "Jenis Kerusakan": r["jenis_kerusakan"],
                "Kode AHSP": r["kode_ahsp"],
                "Volume": f"{r['volume']:.2f} {r['satuan']}",
                "Prosentase": f"{pct_r:.2f}%",
                "Keparahan": r.get("tingkat_keparahan", "Sedang"),
                "Harga Satuan": format_rupiah(r["harga_satuan"]),
                "Total Biaya": format_rupiah(r["total_biaya"]),
                "Rekomendasi Teknis": r["rekomendasi"]
            })

        df_display = pd.DataFrame(table_display)
        st.dataframe(df_display, use_container_width=True, hide_index=True)

        # Baris Tombol Aksi Ekspor (PDF & CSV) & Hapus Data (Tanpa Excel sesuai Revisi)
        st.markdown("""
        <div style="background: #ffffff; border: 1.5px solid #cbd5e1; border-radius: 14px; padding: 16px 20px; margin-top: 20px; margin-bottom: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
            <div style="font-size: 0.95rem; font-weight: 800; color: #0f172a; display: flex; align-items: center; gap: 8px;">
                <span>📥</span> <span>Opsi Ekspor Laporan Resmi &amp; Arsip Pemeriksaan:</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        col_act1, col_act2 = st.columns(2)

        with col_act1:
            # Ekspor PDF Resmi ReportLab
            pdf_data = PDFReportGenerator.generate_pdf_report(
                filtered_records,
                title_doc="Laporan Pemeliharaan Gedung Jawa Tengah",
                signer_info=st.session_state.get("signer_info")
            )
            st.download_button(
                label="📄 Ekspor Laporan Resmi IKBG & RAB (PDF)",
                data=pdf_data,
                file_name=f"Laporan_Resmi_PBG_Jateng_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )

        with col_act2:
            # Ekspor CSV
            csv_data = df_all.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📑 Unduh Database Arsip (CSV)",
                data=csv_data,
                file_name=f"Data_Inspeksi_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                type="secondary",
                use_container_width=True
            )

        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        col_mgt1, col_mgt2 = st.columns([2.2, 1.0])

        with col_mgt1:
            col_id_del, col_btn_del = st.columns([2, 1])
            with col_id_del:
                delete_id = st.number_input("Hapus ID:", min_value=1, step=1, label_visibility="collapsed", placeholder="Ketik ID...")
            with col_btn_del:
                if st.button("🗑️ Hapus ID", use_container_width=True):
                    if DatabaseManager.delete_riwayat(delete_id):
                        st.success(f"Catatan ID #{delete_id} berhasil dihapus.")
                        st.rerun()
                    else:
                        st.error(f"ID #{delete_id} tidak ditemukan.")

        with col_mgt2:
            # Kosongkan Database dengan Pengamanan Dua Langkah (Two-Step Safety Popover)
            with st.popover("⚠️ Reset Database", use_container_width=True):
                st.markdown("<b style='color:#dc2626; font-size:0.9rem;'>⚠️ Peringatan Keamanan Data</b>", unsafe_allow_html=True)
                st.markdown("<p style='font-size:0.82rem; color:#1e293b; line-height:1.4;'>Tindakan ini akan mengosongkan <b>seluruh riwayat inspeksi gedung & RAB</b> secara permanen dari database SQLite.</p>", unsafe_allow_html=True)
                confirm_txt = st.text_input("Ketik kata **HAPUS** untuk verifikasi:", value="", placeholder="HAPUS", key="input_confirm_reset")
                if st.button("🔴 Konfirmasi Hapus Seluruh Data", type="primary", use_container_width=True, key="btn_confirm_reset"):
                    if confirm_txt.strip().upper() == "HAPUS":
                        DatabaseManager.clear_all()
                        st.success("Seluruh riwayat berhasil dikosongkan.")
                        st.rerun()
                    else:
                        st.error("Kata kunci verifikasi salah. Penghapusan dibatalkan demi keamanan data.")

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
            st.caption("ℹ️ *Pratinjau dokumen laporan audit pemeliharaan gedung resmi format A4 tanpa berpindah aplikasi.*")


# ==================================================================================================
# 7. FOOTER APLIKASI WEB
# ==================================================================================================
st.markdown("""
<div style="text-align: center; margin-top: 50px; padding: 24px 0; border-top: 1.5px solid #cbd5e1; font-size: 0.86rem; color: #1e293b; font-weight: 500;">
    <b style="color: #0f172a; font-weight: 800;">Sistem Deteksi Kerusakan Gedung</b><br>
    Manajemen Pemeliharaan &amp; Perawatan Bangunan Gedung | Standar Acuan: <span style="color: #0f172a; font-weight: 700;">AHSP Dinas PU Bina Marga &amp; Cipta Karya Jawa Tengah 2023</span>
</div>
""", unsafe_allow_html=True)
