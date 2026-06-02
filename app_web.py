import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
import plotly.express as px
from io import BytesIO
from openpyxl.styles import PatternFill

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Compras Inteligente - Tapeçaria", layout="wide")

# ==========================================
# --- TELA DE SENHA (BLOQUEIO DE ACESSO) ---
# ==========================================
SENHA_ACESSO = "Tape2026"

if "liberado" not in st.session_state:
    st.session_state.liberado = False

if not st.session_state.liberado:
    st.title("🔒 Acesso Restrito - Tapeçaria")
    st.info("Por favor, insira a senha de liberação para aceder ao portal de análise.")
    senha = st.text_input("Senha", type="password")
    if st.button("Entrar"):
        if senha == SENHA_ACESSO:
            st.session_state.liberado = True
            st.rerun()
        else:
            st.error("Senha incorreta. Tente novamente.")
    st.stop()

# ==========================================
# --- FUNÇÕES DE APOIO ---
# ==========================================

def limpar_v(valor):
    if not valor: return 0.0
    s = str(valor).strip().replace('.', '').replace(',', '.')
    try:
        return float(re.sub(r'[^\d.]', '', s))
    except:
        return 0.0

def extrair_dados_pdf_web(pdf_file):
    dados = []
    nome_filial = pdf_file.name.replace(".pdf", "").upper()
    meses_encontrados = []
    fornecedor_atual = "DESCONHECIDO"
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if not texto: continue
                if len(meses_encontrados) < 4:
                    padrao_mes = r'\b(?:jan|fev|feb|mar|abr|apr|mai|may|jun|jul|ago|aug|set|sep|out|oct|nov|dez|dec)/\d{2,4}\b'
                    encontrados = re.findall(padrao_mes, texto.lower())
                    for m in encontrados:
                        m_upper = m.upper()
                        if m_upper not in meses_encontrados: 
                            meses_encontrados.append(m_upper)
                for l in texto.split('\n'):
                    if "SEGMENTO" in l.upper():
                        match = re.search(r'SEGMENTO\s*:\s*(.*)', l, re.IGNORECASE)
                        if match: 
                            fornecedor_atual = match.group(1).strip()
                    elif re.match(r'^\d{3,6}\s', l):
                        partes = l.split()
                        try:
                            dados.append({
                                'CODIGO': partes[0], 
                                'DESCRICAO': " ".join(partes[1:-11]), 
                                'EMB.': partes[-11],
                                'MES_1': limpar_v(partes[-10]), 
                                'MES_2': limpar_v(partes[-9]), 
                                'MES_3': limpar_v(partes[-8]),
                                'MES_4': limpar_v(partes[-7]), 
                                'MEDIA_SISTEMA': limpar_v(partes[-6]), 
                                'ESTOQUE': limpar_v(partes[-5]),
                                'RESERVA': limpar_v(partes[-4]), 
                                'COMPRADA': limpar_v(partes[-3]), 
                                'MESES_ESTOQUE': limpar_v(partes[-1]), 
                                'FILIAL_NOME': nome_filial, 
                                'FORNECEDOR': fornecedor_atual
                            })
                        except: 
                            continue
        return pd.DataFrame(dados), meses_encontrados
    except: 
        return pd.DataFrame(), []

# --- INTERFACE WEB (BARRA LATERAL) ---
with st.sidebar:
    try: 
        st.image("logo.png", use_container_width=True)
    except: 
        st.error("Arquivo 'logo.png' não encontrado no GitHub.")
    st.markdown("---")
    st.header("⚙️ Configurações")
    meta = st.number_input("Meta de estoque (meses)", min_value=1, value=2)
    meses_parado = st.number_input("Considerar estoque parado após (meses)", min_value=1, value=3, step=1)
    fator_pico = st.number_input("Sensibilidade de Pico (x vezes a média)", min_value=1.5, value=2.5, step=0.5)
    nome_sugerido = st.text_input("Nome do ficheiro Excel", value="Relatorio_Compras_Tapecaria")
    nome_final_xlsx = nome_sugerido if nome_sugerido.endswith(".xlsx") else f"{nome_sugerido}.xlsx"
    st.markdown("---")
    uploaded_files = st.file_uploader("Selecione os 4 PDFs das Unidades", type="pdf", accept_multiple_files=True)

# --- CORPO DO SITE ---
col1, col2 = st.columns([1, 15])
with col1:
    try: 
        st.image("simbolo.png", width=50)
    except: 
        pass
with col2:
    st.title("Compras Inteligente")

st.markdown("### Tapeçaria")

if uploaded_files:
    if st.button("🚀 Processar Análise"):
        dfs_por_filial = {}
        todos_dados = []
        meses_globais = []
        dash_qtd_comprar = 0
        dash_qtd_transferida = 0
        dash_itens_pico = 0
        
        for f in uploaded_files:
