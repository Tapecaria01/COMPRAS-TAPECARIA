import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
import traceback
import plotly.express as px
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal Compras - Tapeçaria", layout="wide")

# ==========================================
# --- ESTILIZAÇÃO CSS (DESIGN PREMIUM) ---
# ==========================================
st.markdown("""
    <style>
    .stApp { background-color: #0E1117 !important; }
    div.stButton > button:first-child {
        background-color: #004A8F !important;
        color: white !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 8px 24px !important;
        font-weight: 600 !important;
        transition: 0.3s !important;
    }
    div.stButton > button:first-child:hover { background-color: #003366 !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# --- TELA DE SENHA (BLOQUEIO CENTRALIZADO) ---
# ==========================================
SENHA_ACESSO = "Tape2026"

if "liberado" not in st.session_state:
    st.session_state.liberado = False

if not st.session_state.liberado:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br><br><br>", unsafe_allow_html=True)
        sc1, sc2, sc3 = st.columns([1.5, 1, 1.5])
        with sc2:
            try: st.image("logo.png", use_container_width=True)
            except: st.markdown("<h1 style='text-align: center; color: white;'>🏢</h1>", unsafe_allow_html=True)
            
        st.markdown("<h2 style='text-align: center; color: #FFFFFF;'>Acesso Restrito</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #4DA8DA;'>Insira a senha de sistema para aceder à inteligência de compras.</p>", unsafe_allow_html=True)
        
        senha = st.text_input("Senha", type="password")
        if st.button("Entrar no Portal", use_container_width=True):
            if senha == SENHA_ACESSO:
                st.session_state.liberado = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")
    st.stop()

# ==========================================
# --- VARIÁVEIS DE SESSÃO ---
# ==========================================
if "analise_concluida" not in st.
