import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
import traceback
import plotly.express as px
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal Compras - Tapeçaria", layout="wide")

# --- TELA DE SENHA (BLOQUEIO CENTRALIZADO) ---
SENHA_ACESSO = "Tape2026"
if "liberado" not in st.session_state:
    st.session_state.liberado = False

if not st.session_state.liberado:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br><br><br>", unsafe_allow_html=True)
        sc1, sc2, sc3 = st.columns([1.5, 1, 1.5])
        with sc2:
            try:
                st.image("logo.png", use_container_width=True)
            except Exception:
                st.markdown("<h1 style='text-align: center; color: white;'>🏢</h1>", unsafe_allow_html=True)
        
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

# --- VARIÁVEIS DE SESSÃO E ESTADO ---
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

if "analise_concluida" not in st.session_state:
    st.session_state.analise_concluida = False

if "df_regras" not in st.session_state:
    st.session_state.df_regras = pd.DataFrame([
        {"FORNECEDOR": "CORTTEX", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "TEX COMPANY", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "CIPATEX", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "KARSTEN", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "ETRURIA", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "TELLAIO", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "OBER", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "TEXTIL J. SERRANO", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "CKS", "MULTIPLO": 50, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "AGRO QUIMICA", "MULTIPLO": 45, "TOLERANCIA": 20, "PALAVRA_CHAVE": ""},
        {"FORNECEDOR": "ROMPLAS", "MULTIPLO": 30, "TOLERANCIA": 15, "PALAVRA_CHAVE": "URUGUA"},
        {"FORNECEDOR": "ROMA DUBLADOS", "MULTIPLO": 10, "TOLERANCIA": 5, "PALAVRA_CHAVE": ""}
    ])

def limpar_dados():
    st.session_state.uploader_key += 1
    st.session_state.analise_concluida = False
    st.rerun()

# --- FUNÇÕES DE LIMPEZA E SUPORTE DE PDF ---
def limpar_v(val):
    if not val or val in ['-', 'S/N', 'N/A']:
        return 0.0
    val_limpo = str(val).replace('.', '').replace(',', '.')
    try:
        return float(val_limpo)
    except ValueError:
        return 0.0

def extrair_dados_pdf_web(file):
    dados = []
    meses_cabecalho = []
    nome_filial = file.name.replace(".pdf", "").upper()
    fornecedor_atual = "DESCONHECIDO"

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            linhas = text.split('\n')
            for l in linhas:
                l_str = l.strip()

                if "FORNECEDOR:" in l_str.upper():
                    partes_forn = l_str.upper().split("FORNECEDOR:")
                    if len(partes_forn) > 1:
                        fornecedor_atual = partes_forn[1].split("-")[0].strip()

                if "CÓDIGO" in l_str.upper() and "DESCRIÇÃO" in l_str.upper():
                    partes_h = l_str.split()
                    cand_meses = [p for p in partes_h if len(p) == 3 and p.isalpha()]
                    if len(cand_meses) >= 4 and not meses_cabecalho:
                        meses_cabecalho = [m.upper() for m in cand_meses[:4]]

                l_limpa = re.sub(r'^\*+\s*', '', l_str)
                match_cod = re.search(r'\b\d{3,6}\b', l_limpa)

                if match_cod:
                    codigo = match_cod.group(0)
                    partes = l_limpa.split()

                    if len(partes) >= 12:
                        try:
                            item_dict = {
                                'CODIGO': codigo,
                                'DESCRICAO': " ".join([p for p in partes if not p.replace(',', '.').replace('-', '').replace('.', '').isdigit() and p != codigo])[:50],
                                'EMB.': partes[-12],
                                'MES_1': limpar_v(partes[-11]),
                                'MES_2': limpar_v(partes[-10]),
                                'MES_3': limpar_v(partes[-9]),
                                'MES_4': limpar_v(partes[-8]),
                                'MEDIA_SISTEMA': limpar_v(partes[-7]),
                                'ESTOQUE': limpar_v(partes[-6]),
                                'RESERVA': limpar_v(partes[-5]),
                                'COMPRADA': limpar_v(partes[-4]),
                                'SITUACAO': partes[-2],
                                'MESES_ESTOQUE': limpar_v(partes[-1]),
                                'FILIAL_NOME': nome_filial,
                                'FORNECEDOR': fornecedor_atual
                            }
                            dados.append(item_dict)
                        except Exception:
                            continue

    df_res = pd.DataFrame(dados)
    if not meses_cabecalho or len(meses_cabecalho) < 4:
        meses_cabecalho = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]

    return df_res, meses_cabecalho

# --- FUNÇÃO DE GERAÇÃO DO EXCEL ---
def gerar_excel_relatorio(dfs_por_filial, meses_cabecalho):
    wb = Workbook()
    wb.remove(wb.active)  # Remove aba padrão

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    fill_parado = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    fill_atipico = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_compra = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    fill_transf = PatternFill(start_color="C9DAF8", end_color="C9DAF8", fill_type="solid")
    fill_transf_morto = PatternFill(start_color="C55A11", end_color="C55A11", fill_type="solid")
    font_white = Font(name="Calibri", size=10, color="FFFFFF", bold=True)
    fill_comprada = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
    fill_ruptura = PatternFill(start_color="FFD2D2", end_color="FFD2D2", fill_type="solid")

    cols_export = [
        'CODIGO', 'DESCRICAO', 'EMB.', 
        'MES_1', 'MES_2', 'MES_3', 'MES_4', 
        'MEDIA', 'ESTOQUE', 'RESERVA', 'COMPRADA', 
        'SUGESTAO COMPRA', 'TRANS INTERNA', 'RUPTURA CRÍTICA', 
        'VENDA_ATIPICA', 'ESTOQUE PARADO', 'FORNECEDOR'
    ]

    cabecalhos_personalizados = [
        'CÓDIGO', 'DESCRIÇÃO', 'EMB.', 
        meses_cabecalho[0], meses_cabecalho[1], meses_cabecalho[2], meses_cabecalho[3], 
        'MÉDIA', 'ESTOQUE', 'RESERVA', 'COMPRADA', 
        'SUGESTÃO COMPRA', 'TRANS. INTERNA', 'RUPTURA CRÍTICA', 
        'VENDA ATÍPICA', 'ESTOQUE PARADO', 'FORNECEDOR'
    ]

    for nome_f, df_f in dfs_por_filial.items():
        ws = wb.create_sheet(title=nome_f[:30])
        ws.views.sheetView[0].showGridLines = True

        for col_idx, text_h in enumerate(cabecalhos_personalizados, 1):
            cell = ws.cell(row=1, column=col_idx, value=text_h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row_idx, (_, row) in enumerate(df_f.iterrows(), 2):
            for col_idx, col_name in enumerate(cols_export, 1):
                val = row.get(col_name, '')
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = data_font
                cell.border = border_thin

                if col_name in ['MES_1', 'MES_2', 'MES_3', 'MES_4', 'MEDIA', 'ESTOQUE', 'RESERVA', 'COMPRADA', 'SUGESTAO COMPRA']:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right")
                else:
                    cell.alignment = Alignment(horizontal="center" if col_name not in ['DESCRICAO', 'FORNECEDOR'] else "left")

                if col_name in ['ESTOQUE PARADO', 'ESTOQUE'] and '🛑 SIM' in str(row.get('ESTOQUE PARADO', '')):
                    cell.fill = fill_parado
                if col_name == 'VENDA_ATIPICA' and '⚠️ SIM' in str(row.get('VENDA_ATIPICA', '')):
                    cell.fill = fill_atipico
                if col_name == 'SUGESTAO COMPRA' and pd.to_numeric(row.get('SUGESTAO COMPRA', 0), errors='coerce') > 0:
                    cell.fill = fill_compra
                if col_name == 'TRANS INTERNA' and str(row.get('TRANS INTERNA', '')) not in ['0', 'None', '', 'nan']:
                    if 'ESTOQUE MORTO' in str(row.get('TRANS INTERNA', '')):
                        cell.fill = fill_transf_morto
                        cell.font = font_white
                    else:
                        cell.fill = fill_transf
                if col_name == 'COMPRADA' and pd.to_numeric(row.get('COMPRADA', 0), errors='coerce') > 0:
                    cell.fill = fill_comprada
                if col_name in ['RUPTURA CRÍTICA', 'ESTOQUE'] and '🚨 CRÍTICA' in str(row.get('RUPTURA CRÍTICA', '')):
                    cell.fill = fill_ruptura

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- INTERFACE WEB (BARRA LATERAL) ---
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception:
        pass

    st.markdown("---")
    st.header("📂 Nova Compra")
    uploaded_files = st.file_uploader(
        "Selecione os 4 PDFs das Unidades
