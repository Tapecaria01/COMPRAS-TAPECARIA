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
from openpyxl.formatting.rule import FormulaRule

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal Compras - Tapeçaria", layout="wide")

# ==========================================
# --- ESTILIZAÇÃO CSS (DESIGN PREMIUM) ---
# ==========================================
st.markdown("""
    
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
        st.markdown("
                    
