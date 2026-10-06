import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
import traceback
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal Compras - Tapeçaria", layout="wide")

# --- INICIALIZAÇÃO DE ESTADO (MÚLTIPLOS E REGRAS) ---
if "df_regras" not in st.session_state:
    st.session_state.df_regras = pd.DataFrame([
        {"FORNECEDOR": "FORNECEDOR A", "MULTIPLO": 10},
        {"FORNECEDOR": "FORNECEDOR B", "MULTIPLO": 50},
        {"FORNECEDOR": "GERAL", "MULTIPLO": 1}
    ])

# --- FUNÇÕES DE SUPORTE E LEITURA DE PDF ---
def limpar_v(val):
    """Converte valores numéricos no formato brasileiro (ex: '1.234,56' ou '0,00') para float."""
    if not val or val == '-' or val == 'S/N':
        return 0.0
    val_limpo = str(val).replace('.', '').replace(',', '.')
    try:
        return float(val_limpo)
    except ValueError:
        return 0.0

def extrair_dados_pdf_web(file):
    """
    Lê o PDF de forma ultra-resiliente usando Regex Dinâmico do final para o início.
    Garante captura total de itens, mesmo com desalinhamentos ou colunas vazias.
    """
    dados = []
    meses_cabecalho = []
    nome_filial = file.name.replace(".pdf", "").replace(".PDF", "").upper()
    fornecedor_atual = "DESCONHECIDO"

    # Regex para identificar números decimais/inteiros no formato BR (ex: 1.234,56 / 0,00 / 12) ou traço '-'
    num_pattern = r'(?:-|\d+(?:\.\d{3})*(?:,\d+)?|\d+)'

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            linhas = text.split('\n')
            for l in linhas:
                l_str = l.strip()
                if not l_str:
                    continue

                # 1. Captura de Fornecedor no Cabeçalho
                if "FORNECEDOR:" in l_str.upper():
                    partes_forn = l_str.upper().split("FORNECEDOR:")
                    if len(partes_forn) > 1:
                        fornecedor_atual = partes_forn[1].split("-")[0].strip()
                    continue

                # 2. Captura dos Meses no Cabeçalho
                if "CÓDIGO" in l_str.upper() and "DESCRIÇÃO" in l_str.upper():
                    partes_h = l_str.split()
                    cand_meses = [p for p in partes_h if len(p) == 3 and p.isalpha()]
                    if len(cand_meses) >= 4 and not meses_cabecalho:
                        meses_cabecalho = [m.upper() for m in cand_meses[:4]]
                    continue

                # Remove asteriscos do início da linha (ex: ***16759 -> 16759)
                l_limpa = re.sub(r'^\*+\s*', '', l_str)

                # Busca por um código de produto no início ou próximo ao início da linha
                match_cod = re.search(r'^\s*([A-Za-z0-9\-]{3,10})\b', l_limpa)
                if not match_cod:
                    # Segunda tentativa: procurar qualquer sequência numérica inicial
                    match_cod = re.search(r'\b\d{3,7}\b', l_limpa)
                
                if match_cod:
                    codigo = match_cod.group(1) if match_cod.groups() else match_cod.group(0)

                    # Busca todos os tokens numéricos da linha (valores da tabela)
                    tokens_numericos = re.findall(num_pattern, l_limpa)

                    # Uma linha válida deve conter ao menos os valores numéricos principais
                    if len(tokens_numericos) >= 8:
                        try:
                            # Pega os últimos valores da tabela a partir do final da linha
                            meses_est = limpar_v(tokens_numericos[-1])
                            
                            comprada = limpar_v(tokens_numericos[-2]) if len(tokens_numericos) >= 9 else 0.0
                            reserva = limpar_v(tokens_numericos[-3]) if len(tokens_numericos) >= 10 else 0.0
                            estoque = limpar_v(tokens_numericos[-4]) if len(tokens_numericos) >= 11 else 0.0
                            media = limpar_v(tokens_numericos[-5]) if len(tokens_numericos) >= 12 else 0.0

                            # Meses de vendas (4 meses)
                            m4 = limpar_v(tokens_numericos[-6]) if len(tokens_numericos) >= 13 else 0.0
                            m3 = limpar_v(tokens_numericos[-7]) if len(tokens_numericos) >= 14 else 0.0
                            m2 = limpar_v(tokens_numericos[-8]) if len(tokens_numericos) >= 15 else 0.0
                            m1 = limpar_v(tokens_numericos[-9]) if len(tokens_numericos) >= 16 else 0.0

                            # Isola a Descrição removendo o código e as partes numéricas finais
                            desc_raw = re.sub(r'^\s*' + re.escape(codigo), '', l_limpa).strip()
                            descricao = re.sub(r'(\s+' + num_pattern + r')+$', '', desc_raw).strip()

                            item_dict = {
                                'CODIGO': codigo,
                                'DESCRICAO': descricao[:60] if descricao else "ITEM SEM DESCRICAO",
                                'EMB.': 'UN',
                                'MES_1': m1,
                                'MES_2': m2,
                                'MES_3': m3,
                                'MES_4': m4,
                                'MEDIA_SISTEMA': media,
                                'ESTOQUE': estoque,
                                'RESERVA': reserva,
                                'COMPRADA': comprada,
                                'SITUACAO': 'OK',
                                'MESES_ESTOQUE': meses_est,
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

def pintar_tabela(val):
    """Aplica cores dinâmicas nas linhas do Dataframe no Streamlit."""
    status = val.get('STATUS', '')
    if 'RUPTURA' in str(status):
        return ['background-color: #ffcccc'] * len(val)
    elif 'TRANSFERIR' in str(status):
        return ['background-color: #e6f2ff'] * len(val)
    elif 'EXCESSO' in str(status):
        return ['background-color: #fff2cc'] * len(val)
    return [''] * len(val)

# --- INTERFACE WEB (BARRA LATERAL) ---
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception:
        pass

    st.markdown("---")
    st.header("📂 Nova Compra")
    uploaded_files = st.file_uploader("Selecione os PDFs das Unidades", type="pdf", accept_multiple_files=True)
    st.markdown("---")

    with st.expander("⚙️ Configurações Avançadas"):
        meta = st.number_input("Meta de estoque (meses)", min_value=1, value=2)
        meses_parado = st.number_input("Considerar estoque parado após (meses)", min_value=1, value=3, step=1)
        fator_pico = st.number_input("Sensibilidade de Pico (x vezes a média)", min_value=1.5, value=2.5, step=0.5)
        nome_sugerido = st.text_input("Nome do ficheiro Excel", value="Relatorio_Compras_Tapecaria")
        nome_final_xlsx = nome_sugerido if nome_sugerido.endswith(".xlsx") else f"{nome_sugerido}.xlsx"

    with st.expander("🏭 Fornecedores e Múltiplos"):
        st.caption("Edite ou adicione regras na última linha vazia.")
        df_regras_editado = st.data_editor(st.session_state.df_regras, num_rows="dynamic", use_container_width=True, hide_index=True)
        st.session_state.df_regras = df_regras_editado

# --- CORPO DO SITE ---
col1, col2 = st.columns([1, 15])
with col1:
    try:
        st.image("simbolo.png", width=50)
    except Exception:
        pass
with col2:
    st.title("Inteligência de Compras")

st.markdown("##### Portal Operacional - Tapeçaria")
st.markdown("
