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

# --- TELA DE SENHA ---
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

# --- VARIÁVEIS DE SESSÃO ---
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

def limpar_v(val):
    if not val or str(val).strip() in ['-', 'S/N', 'N/A', '']:
        return 0.0
    val_limpo = str(val).replace('.', '').replace(',', '.')
    try:
        return float(val_limpo)
    except ValueError:
        return 0.0

# --- EXTRAÇÃO DO PDF ---
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
                                'MARCA': fornecedor_atual,
                                'CURVA': 'A',
                                'ESTOQUE': limpar_v(partes[-6]),
                                'RESERVA': limpar_v(partes[-5]),
                                'DISPONIVEL': max(0.0, limpar_v(partes[-6]) - limpar_v(partes[-5])),
                                'EMB': partes[-12],
                                'ULT_ENTRADA': '',
                                'DIAS_SEM_ENTRADA': 0,
                                'MES_1': limpar_v(partes[-11]),
                                'MES_2': limpar_v(partes[-10]),
                                'MES_3': limpar_v(partes[-9]),
                                'MES_4': limpar_v(partes[-8]),
                                'MEDIA': limpar_v(partes[-7]),
                                'MESES_ESTOQUE': limpar_v(partes[-1]),
                                'SITUACAO': partes[-2],
                                'COMPRADA': limpar_v(partes[-4]),
                                'COMPRAR': 0.0,
                                'TRANFERIR': '0',
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

# --- GERADOR EXCEL EXATO AO MODELO ESPUMAUTO ---
def gerar_excel_relatorio(dfs_por_filial, meses_cabecalho):
    wb = Workbook()
    wb.remove(wb.active)

    # Estilos
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_white = Font(name="Calibri", size=10, color="FFFFFF", bold=True)
    
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    fill_compra = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    fill_transf = PatternFill(start_color="C9DAF8", end_color="C9DAF8", fill_type="solid")
    fill_comprada = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
    fill_ruptura = PatternFill(start_color="FFD2D2", end_color="FFD2D2", fill_type="solid")

    cols_export = [
        'CODIGO', 'DESCRICAO', 'MARCA', 'CURVA', 
        'ESTOQUE', 'RESERVA', 'DISPONIVEL', 'EMB', 
        'ULT_ENTRADA', 'DIAS_SEM_ENTRADA', 
        'MES_1', 'MES_2', 'MES_3', 'MES_4', 
        'MEDIA', 'MESES_ESTOQUE', 'SITUACAO', 
        'COMPRADA', 'COMPRAR', 'TRANFERIR'
    ]

    cabecalhos_personalizados = [
        'CODIGO', 'DESCRICAO', 'MARCA', 'CURVA', 
        'ESTOQUE', 'RESERVA', 'DISPONIVEL', 'EMB', 
        'ULT_ENTRADA', 'DIAS SEM ENTRADA', 
        meses_cabecalho[0], meses_cabecalho[1], meses_cabecalho[2], meses_cabecalho[3], 
        'MEDIA', 'MESES ESTOQUE', 'SITUACAO', 
        'COMPRADA', 'COMPRAR', 'TRANFERIR'
    ]

    for nome_f, df_f in dfs_por_filial.items():
        ws = wb.create_sheet(title=str(nome_f)[:30])
        ws.views.sheetView[0].showGridLines = True

        # Cabeçalhos
        for col_idx, text_h in enumerate(cabecalhos_personalizados, 1):
            cell = ws.cell(row=1, column=col_idx, value=text_h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Dados
        for row_idx, (_, row) in enumerate(df_f.iterrows(), 2):
            for col_idx, col_name in enumerate(cols_export, 1):
                val = row.get(col_name, '')
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = data_font
                cell.border = border_thin

                # Formatação Numérica
                if col_name in ['ESTOQUE', 'RESERVA', 'DISPONIVEL', 'MES_1', 'MES_2', 'MES_3', 'MES_4', 'MEDIA', 'MESES_ESTOQUE', 'COMPRADA', 'COMPRAR']:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center" if col_name not in ['DESCRICAO', 'MARCA'] else "left", vertical="center")

                # Destaques
                if col_name == 'COMPRAR' and pd.to_numeric(row.get('COMPRAR', 0), errors='coerce') > 0:
                    cell.fill = fill_compra
                    cell.font = font_bold
                if col_name == 'TRANFERIR' and str(row.get('TRANFERIR', '')) not in ['0', 'None', '', 'nan']:
                    cell.fill = fill_transf
                if col_name == 'COMPRADA' and pd.to_numeric(row.get('COMPRADA', 0), errors='coerce') > 0:
                    cell.fill = fill_comprada
                if col_name == 'ESTOQUE' and pd.to_numeric(row.get('ESTOQUE', 0), errors='coerce') == 0 and pd.to_numeric(row.get('MEDIA', 0), errors='coerce') > 0:
                    cell.fill = fill_ruptura
                    cell.font = font_bold

        # Larguras de Colunas
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- BARRA LATERAL ---
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception:
        pass

    st.markdown("---")
    st.header("📂 Nova Compra")
    uploaded_files = st.file_uploader("Selecione os 4 PDFs das Unidades", type="pdf", accept_multiple_files=True, key=f"pdf_uploader_{st.session_state.uploader_key}")
    
    st.button("🧹 Limpar Dados para Nova Compra", on_click=limpar_dados, use_container_width=True)
    st.markdown("---")

    with st.expander("⚙️ Configurações Avançadas"):
        meta = st.number_input("Meta de estoque (meses)", min_value=1, value=2)
        retencao_seguranca = st.number_input("Retenção de Segurança Doadora (meses)", min_value=0, value=3)
        fator_pico = st.number_input("Sensibilidade de Pico (x vezes a média)", min_value=1.5, value=2.5, step=0.5)
        nome_sugerido = st.text_input("Nome do ficheiro Excel", value="ESPUMAUTO")
        nome_final_xlsx = nome_sugerido if nome_sugerido.endswith(".xlsx") else f"{nome_sugerido}.xlsx"

    with st.expander("🏭 Fornecedores e Múltiplos"):
        st.subheader("Adicionar / Editar Regra")
        novo_forn = st.text_input("Nome Fornecedor").upper().strip()
        novo_mult = st.number_input("Múltiplo", min_value=1, value=50)
        
        if st.button("➕ Gravar Fornecedor", use_container_width=True):
            if novo_forn:
                forns_existentes = st.session_state.df_regras['FORNECEDOR'].str.upper().tolist()
                if novo_forn in forns_existentes:
                    st.warning(f"⚠️ O fornecedor '{novo_forn}' já está cadastrado!")
                else:
                    nova_linha = pd.DataFrame([{"FORNECEDOR": novo_forn, "MULTIPLO": novo_mult, "TOLERANCIA": 0, "PALAVRA_CHAVE": ""}])
                    st.session_state.df_regras = pd.concat([st.session_state.df_regras, nova_linha], ignore_index=True)
                    st.success(f"✅ Fornecedor '{novo_forn}' gravado com sucesso!")
                    st.rerun()

        st.caption("Tabela Geral de Múltiplos:")
        df_regras_editado = st.data_editor(st.session_state.df_regras, num_rows="dynamic", use_container_width=True, hide_index=True)
        st.session_state.df_regras = df_regras_editado

# --- INTERFACE PRINCIPAL ---
col1, col2 = st.columns([1, 15])
with col1:
    try:
        st.image("simbolo.png", width=50)
    except Exception:
        pass
with col2:
    st.title("Inteligência de Compras")

st.markdown("##### Portal Operacional - Tapeçaria")
st.markdown("<br>", unsafe_allow_html=True)

# --- EXECUÇÃO ---
if uploaded_files:
    with st.spinner("🔍 A executar Inteligência de Compras..."):
        dfs_por_filial = {}
        todos_dados = []
        meses_globais = []

        for f in uploaded_files:
            df, meses = extrair_dados_pdf_web(f)
            if not df.empty:
                dfs_por_filial[f.name.replace(".pdf", "").upper()] = df
                todos_dados.append(df)
                if len(meses) >= 4 and not meses_globais:
                    meses_globais = meses[:4]

        if not meses_globais:
            meses_globais = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]

        if not todos_dados:
            st.error("⚠️ O sistema não encontrou produtos compatíveis nos PDFs.")
        else:
            try:
                df_global = pd.concat(todos_dados).reset_index(drop=True)
                tracker_estoque = {}
                
                for _, row in df_global.iterrows():
                    f_nome = row['FILIAL_NOME']
                    c = row['CODIGO']
                    est = float(row['ESTOQUE'])
                    med = float(row['MEDIA'])
                    excesso = est if med == 0 else max(0.0, est - (med * retencao_seguranca))
                    tracker_estoque[(f_nome, c)] = {'EXCEDENTE': excesso, 'MEDIA': med, 'ESTOQUE_FINAL': est}

                dash_qtd_comprar = 0
                dash_qtd_transferida = 0
                regras_dict = dict(zip(st.session_state.df_regras['FORNECEDOR'].str.upper(), st.session_state.df_regras['MULTIPLO']))

                for f_nome, df_f in dfs_por_filial.items():
                    suge_compra = []
                    trans_interna = []

                    for _, row in df_f.iterrows():
                        c = row['CODIGO']
                        med = float(row['MEDIA'])
                        est = float(row['ESTOQUE'])
                        comp = float(row['COMPRADA'])
                        res = float(row['RESERVA'])
                        fornecedor = str(row['FORNECEDOR']).upper()

                        necessidade = max(0.0, (med * meta) - (est + comp - res))
                        qtd_transf = 0.0
                        origem = ""

                        if necessidade > 0:
                            for (outra_f, cod_item), d_est in tracker_estoque.items():
                                if cod_item == c and outra_f != f_nome and d_est['EXCEDENTE'] > 0:
                                    atend = min(necessidade, d_est['EXCEDENTE'])
                                    qtd_transf += atend
                                    d_est['EXCEDENTE'] -= atend
                                    necessidade -= atend
                                    origem = outra_f
                                    dash_qtd_transferida += atend
                                    if necessidade == 0:
                                        break

                        mult = regras_dict.get(fornecedor, 1)
                        qtd_compra = math.ceil(necessidade / mult) * mult if necessidade > 0 else 0
                        dash_qtd_comprar += qtd_compra

                        suge_compra.append(qtd_compra)
                        trans_interna.append(f"{qtd_transf:.0f} DE {origem}" if qtd_transf > 0 else "0")

                    df_f['COMPRAR'] = suge_compra
                    df_f['TRANFERIR'] = trans_interna

                st.session_state.dfs_por_filial = dfs_por_filial
                st.session_state.meses_globais = meses_globais
                st.session_state.dash_qtd_comprar = dash_qtd_comprar
                st.session_state.dash_qtd_transferida = dash_qtd_transferida
                st.session_state.analise_concluida = True

            except Exception as e:
                st.error(f"🚨 Erro durante os cálculos: {e}")
                st.code(traceback.format_exc())

# --- EXIBIÇÃO ---
if st.session_state.get('analise_concluida', False):
    dfs_por_filial = st.session_state.dfs_por_filial
    meses_globais = st.session_state.meses_globais

    st.success("✅ Análise concluída com sucesso!")
    
    c1, c2 = st.columns(2)
    c1.metric("🛒 Sugestão Total de Compras", f"{int(st.session_state.dash_qtd_comprar)} un.")
    c2.metric("🔄 Economia (Transferências)", f"{int(st.session_state.dash_qtd_transferida)} un.")

    st.markdown("---")
    excel_bytes = gerar_excel_relatorio(dfs_por_filial, meses_globais)
    st.download_button(
        label="📥 Baixar Relatório Excel Original (.xlsx)",
        data=excel_bytes,
        file_name=nome_final_xlsx,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.info("Aguardando PDFs na barra lateral.")
