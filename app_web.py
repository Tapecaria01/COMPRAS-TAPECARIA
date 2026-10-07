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

# ============================================================
# CONFIGURAÇÃO
# ============================================================
st.set_page_config(page_title="Portal Compras - Tapeçaria", layout="wide")

SENHA_SISTEMA = "Tape2026"

# ============================================================
# ESTADO
# ============================================================
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

if "analise_concluida" not in st.session_state:
    st.session_state.analise_concluida = False

if "df_regras" not in st.session_state:
    st.session_state.df_regras = pd.DataFrame([
        {"FORNECEDOR": "CORTTEX", "MULTIPLO": 50},
        {"FORNECEDOR": "TEX COMPANY", "MULTIPLO": 50},
        {"FORNECEDOR": "CIPATEX", "MULTIPLO": 50},
        {"FORNECEDOR": "KARSTEN", "MULTIPLO": 50},
        {"FORNECEDOR": "ETRURIA", "MULTIPLO": 50},
        {"FORNECEDOR": "TELLAIO", "MULTIPLO": 50},
        {"FORNECEDOR": "OBER", "MULTIPLO": 50},
        {"FORNECEDOR": "TEXTIL J. SERRANO", "MULTIPLO": 50},
        {"FORNECEDOR": "CKS", "MULTIPLO": 50},
        {"FORNECEDOR": "AGRO QUIMICA", "MULTIPLO": 45},
        {"FORNECEDOR": "ROMPLAS", "MULTIPLO": 30},
        {"FORNECEDOR": "ROMA DUBLADOS", "MULTIPLO": 10},
        {"FORNECEDOR": "GERAL", "MULTIPLO": 1},
    ])

for key, default in {
    "dfs_por_filial": {},
    "meses_globais": ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"],
    "dashboard": {},
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


# ============================================================
# AUTENTICAÇÃO
# ============================================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    try:
        st.image("logo.png", width=320)
    except Exception:
        pass

    st.title("Portal Compras - Tapeçaria")
    senha = st.text_input("Digite a senha de acesso", type="password")

    if st.button("Entrar", use_container_width=True):
        if senha == SENHA_SISTEMA:
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.error("Senha incorreta.")

    st.stop()


# ============================================================
# FUNÇÕES DE SUPORTE
# ============================================================
MESES_VALIDOS = {
    "JAN", "FEV", "MAR", "ABR", "MAI", "JUN",
    "JUL", "AGO", "SET", "OUT", "NOV", "DEZ",
    "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER",
}

def limpar_v(val):
    """Converte números brasileiros para float."""
    if val is None:
        return 0.0

    texto = str(val).strip().upper()

    if texto in {"", "-", "—", "S/N", "N/A", "NA", "NULL"}:
        return 0.0

    # Trata números negativos e formato brasileiro.
    texto = texto.replace(" ", "")

    try:
        if "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        else:
            # Mantém números simples como 1234 ou 12.5.
            # Se houver vários pontos, assume separador de milhar.
            if texto.count(".") > 1:
                texto = texto.replace(".", "")
        return float(texto)
    except Exception:
        return 0.0


def extrair_meses_do_texto(texto):
    """
    Procura meses no formato do relatório, por exemplo:
    JUL/2026 AUG/2026 SEP/2026 OCT/2026

    Também aceita:
    JUL-2026, JUL.2026, JUL 2026
    """
    if not texto:
        return []

    padrao = re.compile(
        r"\b("
        r"JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ|"
        r"JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|"
        r"SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER"
        r")"
        r"(?:[\/\-.]|\s+)"
        r"(\d{4})\b",
        re.IGNORECASE
    )

    encontrados = []
    vistos = set()

    for match in padrao.finditer(texto):
        mes = match.group(1).upper()
        ano = match.group(2)
        valor = f"{mes}/{ano}"

        # Evita duplicidade se o PDF repetir cabeçalho.
        if valor not in vistos:
            encontrados.append(valor)
            vistos.add(valor)

    return encontrados


def extrair_dados_pdf_web(file):
    """
    Extrai os dados do PDF.

    PRINCIPAL CORREÇÃO:
    - Os meses são capturados diretamente do cabeçalho do PDF.
    - Ex.: JUL/2026, AUG/2026, SEP/2026, OCT/2026.
    - A ordem dos 4 meses é preservada.
    - Cada PDF leva seus próprios meses para o processamento.
    """
    dados = []
    meses_cabecalho = []
    nome_filial = re.sub(r"\.pdf$", "", file.name, flags=re.IGNORECASE).upper()
    fornecedor_atual = "DESCONHECIDO"

    # Números no padrão brasileiro ou simples.
    num_pattern = r"(?:-|\d+(?:\.\d{3})*(?:,\d+)?|\d+(?:\.\d+)?)"

    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                # Primeiro procura os meses no texto completo da página.
                if len(meses_cabecalho) < 4:
                    meses_pagina = extrair_meses_do_texto(text)
                    if len(meses_pagina) >= 4:
                        meses_cabecalho = meses_pagina[:4]

                linhas = text.split("\n")

                for linha in linhas:
                    l_str = linha.strip()
                    if not l_str:
                        continue

                    upper = l_str.upper()

                    # Fornecedor
                    if "FORNECEDOR:" in upper:
                        partes = re.split(r"FORNECEDOR\s*:", l_str, flags=re.IGNORECASE)
                        if len(partes) > 1:
                            fornecedor_atual = partes[1].split(" - ")[0].strip().upper()

                    # Reforça a captura dos meses no cabeçalho.
                    if len(meses_cabecalho) < 4:
                        meses_linha = extrair_meses_do_texto(l_str)
                        if len(meses_linha) >= 4:
                            meses_cabecalho = meses_linha[:4]

                    # Ignora linhas de cabeçalho.
                    if "CÓDIGO" in upper and "DESCRIÇÃO" in upper:
                        continue

                    # Remove asteriscos antes do código.
                    l_limpa = re.sub(r"^\*+\s*", "", l_str)

                    # Código do produto.
                    match_cod = re.search(
                        r"^\s*([A-Za-z0-9][A-Za-z0-9\-_]{2,9})\b",
                        l_limpa
                    )

                    if not match_cod:
                        match_cod = re.search(r"\b\d{3,7}\b", l_limpa)

                    if not match_cod:
                        continue

                    codigo = match_cod.group(1)

                    # Todos os valores numéricos da linha.
                    tokens = re.findall(num_pattern, l_limpa)

                    # O relatório atual tem os 4 meses + média + estoque +
                    # reserva + comprada + meses de estoque.
                    if len(tokens) < 8:
                        continue

                    try:
                        # Os campos finais são lidos de trás para frente.
                        meses_est = limpar_v(tokens[-1])
                        comprada = limpar_v(tokens[-2]) if len(tokens) >= 9 else 0.0
                        reserva = limpar_v(tokens[-3]) if len(tokens) >= 10 else 0.0
                        estoque = limpar_v(tokens[-4]) if len(tokens) >= 11 else 0.0
                        media = limpar_v(tokens[-5]) if len(tokens) >= 12 else 0.0

                        # Mantém a ordem real do cabeçalho:
                        # MES_1 = primeiro mês, MES_4 = quarto mês.
                        m1 = limpar_v(tokens[-9]) if len(tokens) >= 16 else 0.0
                        m2 = limpar_v(tokens[-8]) if len(tokens) >= 15 else 0.0
                        m3 = limpar_v(tokens[-7]) if len(tokens) >= 14 else 0.0
                        m4 = limpar_v(tokens[-6]) if len(tokens) >= 13 else 0.0

                        # Descrição: remove código e números finais.
                        desc_raw = re.sub(
                            r"^\s*" + re.escape(codigo),
                            "",
                            l_limpa
                        ).strip()

                        descricao = re.sub(
                            r"(?:\s+" + num_pattern + r")+$",
                            "",
                            desc_raw
                        ).strip()

                        if not descricao:
                            descricao = "ITEM SEM DESCRICAO"

                        dados.append({
                            "CODIGO": str(codigo),
                            "DESCRICAO": descricao[:100],
                            "EMB.": "UN",
                            "MES_1": m1,
                            "MES_2": m2,
                            "MES_3": m3,
                            "MES_4": m4,
                            "MEDIA_SISTEMA": media,
                            "ESTOQUE": estoque,
                            "RESERVA": reserva,
                            "COMPRADA": comprada,
                            "SITUACAO": "OK",
                            "MESES_ESTOQUE": meses_est,
                            "FILIAL_NOME": nome_filial,
                            "FORNECEDOR": fornecedor_atual,
                        })

                    except Exception:
                        continue

    except Exception as e:
        raise RuntimeError(f"Erro lendo o PDF {file.name}: {e}") from e

    df_res = pd.DataFrame(dados)

    if len(meses_cabecalho) < 4:
        meses_cabecalho = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]

    return df_res, meses_cabecalho[:4]


def numero_seguro(valor):
    try:
        if pd.isna(valor):
            return 0.0
        return float(valor)
    except Exception:
        return 0.0


def normalizar_regras(df_regras):
    """Cria o dicionário de múltiplos de fornecedor de forma segura."""
    regras = {}

    if df_regras is None or df_regras.empty:
        return {"GERAL": 1}

    for _, linha in df_regras.iterrows():
        fornecedor = str(linha.get("FORNECEDOR", "")).strip().upper()
        multiplo = numero_seguro(linha.get("MULTIPLO", 1))

        if fornecedor:
            regras[fornecedor] = max(1, int(round(multiplo)))

    regras.setdefault("GERAL", 1)
    return regras


def nome_aba_seguro(nome, usados):
    """Nome de aba Excel <= 31 caracteres e sem caracteres proibidos."""
    base = re.sub(r'[\[\]\:\*\?\/\\]', "_", str(nome))
    base = base.strip() or "FILIAL"
    base = base[:31]

    original = base
    contador = 2

    while base in usados:
        sufixo = f"_{contador}"
        base = original[:31 - len(sufixo)] + sufixo
        contador += 1

    usados.add(base)
    return base


# ============================================================
# EXPORTAÇÃO EXCEL
# ============================================================
def gerar_excel(dfs_por_filial, meses_globais, dashboard=None):
    """
    Gera Excel sem filtros automáticos.

    Abas:
    - RESUMO
    - SUGESTÕES DE COMPRA
    - TRANSFERÊNCIAS
    - uma aba para cada filial

    Não é criado AutoFilter nem tabela do Excel.
    Portanto o arquivo NÃO abre com setas/filtros nas colunas.
    """
    dashboard = dashboard or {}

    wb = Workbook()
    ws_resumo = wb.active
    ws_resumo.title = "RESUMO"

    # Estilos.
    fill_titulo = PatternFill("solid", fgColor="1F4E78")
    fill_cabecalho = PatternFill("solid", fgColor="5B9BD5")
    fill_compra = PatternFill("solid", fgColor="E2F0D9")
    fill_transf = PatternFill("solid", fgColor="DDEBF7")
    fill_parado = PatternFill("solid", fgColor="FFF2CC")
    fill_ruptura = PatternFill("solid", fgColor="F4CCCC")

    fonte_branca = Font(color="FFFFFF", bold=True)
    fonte_negrito = Font(bold=True)
    borda_fina = Border(
        left=Side(style="thin", color="D9E1F2"),
        right=Side(style="thin", color="D9E1F2"),
        top=Side(style="thin", color="D9E1F2"),
        bottom=Side(style="thin", color="D9E1F2"),
    )

    # --------------------------------------------------------
    # RESUMO
    # --------------------------------------------------------
    ws_resumo["A1"] = "RELATÓRIO DE INTELIGÊNCIA DE COMPRAS"
    ws_resumo["A1"].fill = fill_titulo
    ws_resumo["A1"].font = Font(color="FFFFFF", bold=True, size=14)
    ws_resumo.merge_cells("A1:D1")

    resumo = [
        ("Total sugerido para compra", dashboard.get("qtd_comprar", 0)),
        ("Total sugerido para transferência", dashboard.get("qtd_transferida", 0)),
        ("Itens com venda atípica", dashboard.get("itens_pico", 0)),
        ("Itens com necessidade", dashboard.get("itens_necessidade", 0)),
        ("Itens parados > 90 dias", dashboard.get("itens_parados", 0)),
    ]

    ws_resumo["A3"] = "INDICADOR"
    ws_resumo["B3"] = "VALOR"
    for cell in ws_resumo[3]:
        cell.fill = fill_cabecalho
        cell.font = fonte_branca
        cell.border = borda_fina

    for i, (indicador, valor) in enumerate(resumo, start=4):
        ws_resumo.cell(i, 1, indicador)
        ws_resumo.cell(i, 2, valor)
        ws_resumo.cell(i, 1).border = borda_fina
        ws_resumo.cell(i, 2).border = borda_fina

    ws_resumo["A11"] = "Meses identificados nos PDFs"
    ws_resumo["A11"].font = fonte_negrito
    ws_resumo["B11"] = " | ".join(meses_globais[:4])

    ws_resumo.freeze_panes = "A3"

    # --------------------------------------------------------
    # BASE CONSOLIDADA
    # --------------------------------------------------------
    todos = []
    for _, df in dfs_por_filial.items():
        if df is not None and not df.empty:
            todos.append(df.copy())

    df_todos = pd.concat(todos, ignore_index=True) if todos else pd.DataFrame()

    colunas_base = [
        "FILIAL",
        "CODIGO",
        "DESCRICAO",
        "FORNECEDOR",
        "MEDIA",
        "ESTOQUE",
        "RESERVA",
        "COMPRADA",
        "SUGESTAO_COMPRA",
        "SUGESTAO_TRANSFERENCIA",
        "ORIGEM_TRANSFERENCIA",
        "DESTINO_TRANSFERENCIA",
        "ESTOQUE_PARADO_90D",
        "SITUACAO",
    ]

    def preparar_export(df):
        x = df.copy()

        # Garante colunas.
        for c in colunas_base:
            if c not in x.columns:
                x[c] = "" if "SITUACAO" in c or "TRANSFERENCIA" in c or "DESCRICAO" in c or c == "FORNECEDOR" else 0

        # Acrescenta os 4 meses com o nome real do PDF.
        rename_meses = {
            "MES_1": meses_globais[0] if len(meses_globais) > 0 else "MÊS 1",
            "MES_2": meses_globais[1] if len(meses_globais) > 1 else "MÊS 2",
            "MES_3": meses_globais[2] if len(meses_globais) > 2 else "MÊS 3",
            "MES_4": meses_globais[3] if len(meses_globais) > 3 else "MÊS 4",
        }

        x = x.rename(columns=rename_meses)

        ordem = [
            "FILIAL",
            "CODIGO",
            "DESCRICAO",
            "FORNECEDOR",
            rename_meses["MES_1"],
            rename_meses["MES_2"],
            rename_meses["MES_3"],
            rename_meses["MES_4"],
            "MEDIA",
            "ESTOQUE",
            "RESERVA",
            "COMPRADA",
            "SUGESTAO_COMPRA",
            "SUGESTAO_TRANSFERENCIA",
            "ORIGEM_TRANSFERENCIA",
            "DESTINO_TRANSFERENCIA",
            "ESTOQUE_PARADO_90D",
            "SITUACAO",
        ]

        for c in ordem:
            if c not in x.columns:
                x[c] = ""

        return x[ordem]

    df_export = preparar_export(df_todos)

    # --------------------------------------------------------
    # Função de escrita de planilha
    # --------------------------------------------------------
    def escrever_df(ws, df, titulo):
        ws["A1"] = titulo
        ws["A1"].fill = fill_titulo
        ws["A1"].font = Font(color="FFFFFF", bold=True, size=13)

        max_col = max(1, len(df.columns))
        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=max_col
        )

        # Cabeçalho na linha 3.
        for col_idx, coluna in enumerate(df.columns, start=1):
            cell = ws.cell(3, col_idx, coluna)
            cell.fill = fill_cabecalho
            cell.font = fonte_branca
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = borda_fina

        # Dados.
        for row_idx, (_, row) in enumerate(df.iterrows(), start=4):
            for col_idx, coluna in enumerate(df.columns, start=1):
                valor = row[coluna]

                if pd.isna(valor):
                    valor = ""

                cell = ws.cell(row_idx, col_idx, valor)
                cell.border = borda_fina
                cell.alignment = Alignment(vertical="center")

                if coluna in {
                    "MEDIA",
                    "ESTOQUE",
                    "RESERVA",
                    "COMPRADA",
                    "SUGESTAO_COMPRA",
                    "SUGESTAO_TRANSFERENCIA",
                    "ESTOQUE_PARADO_90D",
                } or coluna in meses_globais:
                    if isinstance(valor, (int, float)):
                        cell.number_format = '#,##0.00'

                if coluna == "SUGESTAO_COMPRA" and numero_seguro(valor) > 0:
                    cell.fill = fill_compra
                    cell.font = fonte_negrito

                if coluna == "SUGESTAO_TRANSFERENCIA" and numero_seguro(valor) > 0:
                    cell.fill = fill_transf
                    cell.font = fonte_negrito

                if coluna == "ESTOQUE_PARADO_90D" and numero_seguro(valor) > 0:
                    cell.fill = fill_parado

                if coluna == "SITUACAO":
                    texto = str(valor).upper()
                    if "COMPRAR" in texto or "RUPTURA" in texto:
                        cell.fill = fill_ruptura
                    elif "TRANSFER" in texto:
                        cell.fill = fill_transf
                    elif "PARADO" in texto:
                        cell.fill = fill_parado

        # Largura das colunas.
        for col_idx, coluna in enumerate(df.columns, start=1):
            letra = get_column_letter(col_idx)
            if df.empty:
                largura = len(str(coluna)) + 2
            else:
                valores = df[coluna].astype(str).head(500)
                maior = max([len(str(coluna))] + [len(v) for v in valores])
                largura = min(max(maior + 2, 10), 45)

            ws.column_dimensions[letra].width = largura

        ws.freeze_panes = "A4"
        ws.sheet_view.showGridLines = False

        # IMPORTANTE:
        # NÃO usar ws.auto_filter.ref
        # NÃO usar openpyxl.worksheet.table.Table
        # Assim o Excel abre sem setas/filtros.

    # --------------------------------------------------------
    # SUGESTÕES DE COMPRA
    # --------------------------------------------------------
    ws_compra = wb.create_sheet("SUGESTÕES DE COMPRA")
    if not df_export.empty:
        df_compra = df_export[
            pd.to_numeric(df_export["SUGESTAO_COMPRA"], errors="coerce").fillna(0) > 0
        ].copy()
    else:
        df_compra = df_export.copy()

    escrever_df(ws_compra, df_compra, "SUGESTÕES DE COMPRA")

    # --------------------------------------------------------
    # TRANSFERÊNCIAS
    # --------------------------------------------------------
    ws_transf = wb.create_sheet("TRANSFERÊNCIAS")
    if not df_export.empty:
        df_transf = df_export[
            pd.to_numeric(df_export["SUGESTAO_TRANSFERENCIA"], errors="coerce").fillna(0) > 0
        ].copy()
    else:
        df_transf = df_export.copy()

    escrever_df(ws_transf, df_transf, "SUGESTÕES DE TRANSFERÊNCIA")

    # --------------------------------------------------------
    # UMA ABA POR FILIAL
    # --------------------------------------------------------
    usados = {
        "RESUMO",
        "SUGESTÕES DE COMPRA",
        "TRANSFERÊNCIAS",
    }

    for filial, df_filial in dfs_por_filial.items():
        if df_filial is None or df_filial.empty:
            continue

        ws = wb.create_sheet(nome_aba_seguro(filial, usados))

        # Reaproveita a estrutura final.
        df_f = preparar_export(df_filial)
        escrever_df(ws, df_f, str(filial))

    # Ajustes finais.
    for ws in wb.worksheets:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# INTERFACE
# ============================================================
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception:
        pass

    st.markdown("---")
    st.header("📂 Nova Compra")

    uploaded_files = st.file_uploader(
        "Selecione os PDFs das Unidades",
        type="pdf",
        accept_multiple_files=True,
        key=f"pdf_uploader_{st.session_state.uploader_key}",
    )

    if st.button("🗑️ Limpar análise", use_container_width=True):
        st.session_state.dfs_por_filial = {}
        st.session_state.meses_globais = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]
        st.session_state.dashboard = {}
        st.session_state.analise_concluida = False
        st.session_state.uploader_key += 1
        st.rerun()

    st.markdown("---")

    with st.expander("⚙️ Configurações Avançadas"):
        meta = st.number_input(
            "Meta de estoque (meses)",
            min_value=1,
            value=2,
            step=1,
        )

        meses_parado = st.number_input(
            "Considerar estoque parado após (meses)",
            min_value=1,
            value=3,
            step=1,
            help="3 meses = 90 dias. Produtos sem venda nesse período entram na sugestão de transferência.",
        )

        fator_pico = st.number_input(
            "Sensibilidade de Pico (x vezes a média)",
            min_value=1.5,
            value=2.5,
            step=0.5,
        )

        nome_sugerido = st.text_input(
            "Nome do ficheiro Excel",
            value="Relatorio_Compras_Tapecaria",
        )

        nome_final_xlsx = (
            nome_sugerido
            if nome_sugerido.lower().endswith(".xlsx")
            else f"{nome_sugerido}.xlsx"
        )

    with st.expander("🏭 Fornecedores e Múltiplos"):
        st.caption("Edite ou adicione regras na tabela abaixo.")

        df_regras_editado = st.data_editor(
            st.session_state.df_regras,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
        )

        st.session_state.df_regras = df_regras_editado


# ============================================================
# CABEÇALHO
# ============================================================
col1, col2 = st.columns([1, 15])

with col1:
    try:
        st.image("simbolo.png", width=50)
    except Exception:
        pass

with col2:
    st.title("Inteligência de Compras")

st.markdown("##### Portal Operacional - Tapeçaria")


# ============================================================
# PROCESSAMENTO
# ============================================================
if uploaded_files:
    with st.spinner("🔍 Processando PDFs e calculando compras/transferências..."):
        try:
            dfs_por_filial = {}
            todos_dados = []
            meses_globais = []

            # ------------------------------------------------
            # 1. Ler cada PDF e guardar seus meses.
            # ------------------------------------------------
            for f in uploaded_files:
                df, meses = extrair_dados_pdf_web(f)

                if not df.empty:
                    nome_filial = re.sub(
                        r"\.pdf$",
                        "",
                        f.name,
                        flags=re.IGNORECASE
                    ).upper()

                    dfs_por_filial[nome_filial] = df
                    todos_dados.append(df)

                    # O sistema usa o primeiro conjunto de meses
                    # válido como referência global para exportação.
                    if len(meses) >= 4 and not meses_globais:
                        meses_globais = meses[:4]

            if not meses_globais:
                meses_globais = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]

            if not todos_dados:
                st.error("⚠️ O sistema não encontrou produtos compatíveis nos PDFs.")
                st.stop()

            # ------------------------------------------------
            # 2. Consolidar.
            # ------------------------------------------------
            df_global = pd.concat(todos_dados, ignore_index=True)

            for col in [
                "MES_1", "MES_2", "MES_3", "MES_4",
                "MEDIA_SISTEMA", "ESTOQUE", "RESERVA",
                "COMPRADA", "MESES_ESTOQUE"
            ]:
                if col not in df_global.columns:
                    df_global[col] = 0.0
                df_global[col] = pd.to_numeric(
                    df_global[col],
                    errors="coerce"
                ).fillna(0.0)

            df_global["ESTOQUE_DISPONIVEL"] = df_global["ESTOQUE"]

            df_global["TOTAL_VENDAS_RECENTES"] = (
                df_global["MES_1"]
                + df_global["MES_2"]
                + df_global["MES_3"]
                + df_global["MES_4"]
            )

            # ------------------------------------------------
            # 3. Regras de fornecedor.
            # ------------------------------------------------
            regras_dict = normalizar_regras(st.session_state.df_regras)
            multiplo_padrao = regras_dict.get("GERAL", 1)

            # ------------------------------------------------
            # 4. Necessidade de cada filial.
            # ------------------------------------------------
            necessidades = {}

            for _, row in df_global.iterrows():
                filial = row["FILIAL_NOME"]
                codigo = str(row["CODIGO"])
                media = numero_seguro(row["MEDIA_SISTEMA"])
                estoque = numero_seguro(row["ESTOQUE"])
                reserva = numero_seguro(row["RESERVA"])
                comprada = numero_seguro(row["COMPRADA"])

                necessidade = max(
                    0.0,
                    (media * meta) - (estoque + comprada - reserva)
                )

                necessidades[(filial, codigo)] = necessidade

            # ------------------------------------------------
            # 5. Excedentes para transferência de emergência.
            #    Mantém a regra original: estoque acima de
            #    3 meses da média pode atender outra filial.
            # ------------------------------------------------
            tracker_estoque = {}

            for _, row in df_global.iterrows():
                filial = row["FILIAL_NOME"]
                codigo = str(row["CODIGO"])
                estoque = numero_seguro(row["ESTOQUE"])
                media = numero_seguro(row["MEDIA_SISTEMA"])

                excesso = (
                    estoque
                    if media == 0
                    else max(0.0, estoque - (media * 3))
                )

                tracker_estoque[(filial, codigo)] = {
                    "EXCEDENTE": excesso,
                    "ESTOQUE_INICIAL": estoque,
                    "MEDIA": media,
                }

            # ------------------------------------------------
            # 6. Primeiro atende necessidades por transferência.
            # ------------------------------------------------
            transferencias_emergencia = {}

            for chave_destino, necessidade_inicial in list(necessidades.items()):
                if necessidade_inicial <= 0:
                    continue

                filial_destino, codigo = chave_destino
                necessidade = necessidade_inicial
                origem_textos = []

                for (filial_origem, codigo_origem), dados_est in tracker_estoque.items():
                    if necessidade <= 0:
                        break

                    if filial_origem == filial_destino:
                        continue

                    if codigo_origem != codigo:
                        continue

                    disponivel = numero_seguro(dados_est["EXCEDENTE"])
                    if disponivel <= 0:
                        continue

                    qtd = min(necessidade, disponivel)

                    dados_est["EXCEDENTE"] -= qtd
                    necessidade -= qtd

                    origem_textos.append(
                        f"{qtd:g} de {filial_origem}"
                    )

                    transferencias_emergencia[
                        (filial_destino, codigo)
                    ] = {
                        "QTD": transferencias_emergencia.get(
                            (filial_destino, codigo), {}
                        ).get("QTD", 0) + qtd,
                        "ORIGEM": (
                            transferencias_emergencia.get(
                                (filial_destino, codigo), {}
                            ).get("ORIGEM", "")
                            + (
                                (" + " if transferencias_emergencia.get(
                                    (filial_destino, codigo), {}
                                ).get("ORIGEM", "") else "")
                                + f"{qtd:g} de {filial_origem}"
                            )
                        ),
                    }

                necessidades[chave_destino] = necessidade

            # ------------------------------------------------
            # 7. Sugestão específica de transferência de
            #    produtos parados há mais de 90 dias.
            #
            #    Regra:
            #    - estoque > 0
            #    - vendas dos últimos N meses = 0
            #    - N padrão = 3 meses = 90 dias
            #    - procura outra filial com necessidade do
            #      mesmo produto.
            # ------------------------------------------------
            transferencias_parados = {}
            estoque_parado_por_chave = {}

            for _, row in df_global.iterrows():
                filial = row["FILIAL_NOME"]
                codigo = str(row["CODIGO"])
                estoque = numero_seguro(row["ESTOQUE"])

                # Usa os primeiros N meses disponíveis do relatório.
                qtd_meses_considerados = min(
                    int(meses_parado),
                    4
                )

                meses_venda = [
                    numero_seguro(row[f"MES_{i}"])
                    for i in range(1, qtd_meses_considerados + 1)
                ]

                vendas_no_periodo = sum(meses_venda)

                # Produto sem venda durante o período configurado.
                parado = estoque > 0 and vendas_no_periodo <= 0

                if parado:
                    estoque_parado_por_chave[(filial, codigo)] = estoque

                    # Procura filial de destino com necessidade.
                    melhor_destino = None
                    melhor_necessidade = 0.0

                    for (filial_destino, codigo_destino), necessidade in necessidades.items():
                        if filial_destino == filial:
                            continue
                        if codigo_destino != codigo:
                            continue

                        if necessidade > melhor_necessidade:
                            melhor_necessidade = necessidade
                            melhor_destino = filial_destino

                    if melhor_destino:
                        # Usa somente o saldo ainda disponível na filial de origem.
                        # As transferências de emergência feitas anteriormente
                        # já reduziram este excedente.
                        saldo_origem = numero_seguro(
                            tracker_estoque.get(
                                (filial, codigo),
                                {}
                            ).get("EXCEDENTE", 0)
                        )

                        qtd_transferir = min(
                            saldo_origem,
                            melhor_necessidade
                        )

                        if qtd_transferir > 0:
                            tracker_estoque[(filial, codigo)]["EXCEDENTE"] -= qtd_transferir

                            transferencias_parados[
                                (filial, codigo)
                            ] = {
                                "QTD": qtd_transferir,
                                "DESTINO": melhor_destino,
                            }

                            # Diminui a necessidade do destino.
                            necessidades[
                                (melhor_destino, codigo)
                            ] = max(
                                0.0,
                                necessidades[
                                    (melhor_destino, codigo)
                                ] - qtd_transferir
                            )

            # ------------------------------------------------
            # 8. Montar resultado final.
            # ------------------------------------------------
            resultados_por_filial = {
                filial: [] for filial in dfs_por_filial
            }

            dash_qtd_comprar = 0.0
            dash_qtd_transferida = 0.0
            dash_itens_pico = 0
            dash_itens_necessidade = 0
            dash_itens_parados = 0

            for _, row in df_global.iterrows():
                filial = row["FILIAL_NOME"]
                codigo = str(row["CODIGO"])
                media = numero_seguro(row["MEDIA_SISTEMA"])
                estoque = numero_seguro(row["ESTOQUE"])
                reserva = numero_seguro(row["RESERVA"])
                comprada = numero_seguro(row["COMPRADA"])
                fornecedor = str(row["FORNECEDOR"]).upper()

                # Pico de venda.
                vendas = [
                    numero_seguro(row["MES_1"]),
                    numero_seguro(row["MES_2"]),
                    numero_seguro(row["MES_3"]),
                    numero_seguro(row["MES_4"]),
                ]

                max_venda = max(vendas)
                eh_pico = (
                    media > 0
                    and max_venda > (media * fator_pico)
                )

                if eh_pico:
                    dash_itens_pico += 1

                necessidade_restante = numero_seguro(
                    necessidades.get((filial, codigo), 0)
                )

                necessidade_original = max(
                    0.0,
                    (media * meta) - (estoque + comprada - reserva)
                )

                if necessidade_original > 0:
                    dash_itens_necessidade += 1

                # Transferência de emergência recebida.
                transf_emergencia = transferencias_emergencia.get(
                    (filial, codigo),
                    {"QTD": 0.0, "ORIGEM": ""}
                )

                qtd_transf_recebida = numero_seguro(
                    transf_emergencia.get("QTD", 0)
                )
                origem_transf = transf_emergencia.get("ORIGEM", "")

                # Compra final depois das transferências.
                mult = regras_dict.get(fornecedor, multiplo_padrao)

                qtd_comprar = (
                    math.ceil(necessidade_restante / mult) * mult
                    if necessidade_restante > 0
                    else 0
                )

                dash_qtd_comprar += qtd_comprar

                # Sugestão de transferência de parado > 90 dias.
                transf_parado = transferencias_parados.get(
                    (filial, codigo),
                    {"QTD": 0.0, "DESTINO": ""}
                )

                qtd_transferir_parado = numero_seguro(
                    transf_parado.get("QTD", 0)
                )
                destino_transferencia = transf_parado.get(
                    "DESTINO", ""
                )

                dash_qtd_transferida += (
                    qtd_transf_recebida
                    + qtd_transferir_parado
                )

                # Estoque parado.
                estoque_parado_90d = estoque_parado_por_chave.get(
                    (filial, codigo),
                    0.0
                )

                if estoque_parado_90d > 0:
                    dash_itens_parados += 1

                # Status.
                if qtd_comprar > 0:
                    status = "🔴 RUPTURA / COMPRAR"
                elif qtd_transferir_parado > 0:
                    status = (
                        f"🟣 TRANSFERIR PARADO > 90 DIAS "
                        f"PARA {destino_transferencia}"
                    )
                elif qtd_transf_recebida > 0:
                    status = f"🔵 RECEBER TRANSFERÊNCIA"
                elif media == 0 and estoque > 0:
                    status = "🟡 EXCESSO / SEM GIRO"
                elif eh_pico:
                    status = "🟠 VENDA ATÍPICA"
                else:
                    status = "🟢 OK"

                resultados_por_filial[filial].append({
                    "FILIAL": filial,
                    "CODIGO": codigo,
                    "DESCRICAO": row["DESCRICAO"],
                    "FORNECEDOR": fornecedor,
                    "MES_1": vendas[0],
                    "MES_2": vendas[1],
                    "MES_3": vendas[2],
                    "MES_4": vendas[3],
                    "MEDIA": media,
                    "ESTOQUE": estoque,
                    "RESERVA": reserva,
                    "COMPRADA": comprada,

                    # IMPORTANTE:
                    # Esta é a coluna que fica imediatamente depois
                    # da sugestão de compra no Excel.
                    "SUGESTAO_COMPRA": qtd_comprar,

                    # Sugestão de transferência de produto parado.
                    "SUGESTAO_TRANSFERENCIA": qtd_transferir_parado,

                    "ORIGEM_TRANSFERENCIA": (
                        origem_transf
                        if qtd_transf_recebida > 0
                        else ""
                    ),

                    "DESTINO_TRANSFERENCIA": (
                        destino_transferencia
                        if qtd_transferir_parado > 0
                        else ""
                    ),

                    "ESTOQUE_PARADO_90D": estoque_parado_90d,
                    "SITUACAO": status,
                })

            # ------------------------------------------------
            # 9. DataFrames finais.
            # ------------------------------------------------
            dfs_finais = {}

            for filial, lista in resultados_por_filial.items():
                dfs_finais[filial] = pd.DataFrame(lista)

            dashboard = {
                "qtd_comprar": dash_qtd_comprar,
                "qtd_transferida": dash_qtd_transferida,
                "itens_pico": dash_itens_pico,
                "itens_necessidade": dash_itens_necessidade,
                "itens_parados": dash_itens_parados,
            }

            st.session_state.dfs_por_filial = dfs_finais
            st.session_state.meses_globais = meses_globais
            st.session_state.dashboard = dashboard
            st.session_state.analise_concluida = True

        except Exception as e:
            st.error("❌ Ocorreu um erro durante o processamento.")
            st.code(traceback.format_exc())
            st.stop()


# ============================================================
# RESULTADOS
# ============================================================
if st.session_state.analise_concluida and st.session_state.dfs_por_filial:

    dfs = st.session_state.dfs_por_filial
    meses_globais = st.session_state.meses_globais
    dashboard = st.session_state.dashboard

    st.markdown("---")

    # ========================================================
    # DOWNLOAD EXCEL
    # ========================================================
    try:
        excel_bytes = gerar_excel(
            dfs,
            meses_globais,
            dashboard
        )

        st.download_button(
            label="📥 EXPORTAR EXCEL - COMPRAS E TRANSFERÊNCIAS",
            data=excel_bytes,
            file_name=nome_final_xlsx,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.caption(
            "O Excel é gerado sem filtros automáticos nas colunas. "
            "A aba 'SUGESTÕES DE COMPRA' mostra apenas compras e a aba "
            "'TRANSFERÊNCIAS' mostra as transferências sugeridas."
        )

    except Exception:
        st.error("❌ Não foi possível gerar o Excel.")
        st.code(traceback.format_exc())

    st.markdown("---")

    # ========================================================
    # MÉTRICAS
    # ========================================================
    m1, m2, m3, m4, m5 = st.columns(5)

    m1.metric(
        "📦 Sugestão de compra",
        f"{dashboard.get('qtd_comprar', 0):,.0f}"
    )

    m2.metric(
        "🔄 Transferências",
        f"{dashboard.get('qtd_transferida', 0):,.0f}"
    )

    m3.metric(
        "📈 Vendas atípicas",
        dashboard.get("itens_pico", 0)
    )

    m4.metric(
        "⚠️ Necessidades",
        dashboard.get("itens_necessidade", 0)
    )

    m5.metric(
        "🕒 Parados > 90 dias",
        dashboard.get("itens_parados", 0)
    )

    # ========================================================
    # ABAS
    # ========================================================
    aba1, aba2, aba3, aba4 = st.tabs([
        "📊 Visão Geral",
        "🚨 Top Urgentes",
        "🕒 Estoque Parado",
        "🏭 Prévia por Filial",
    ])

    # --------------------------------------------------------
    # VISÃO GERAL
    # --------------------------------------------------------
    with aba1:
        resumo_filiais = []

        for filial, df in dfs.items():
            if df.empty:
                continue

            resumo_filiais.append({
                "FILIAL": filial,
                "COMPRA": pd.to_numeric(
                    df["SUGESTAO_COMPRA"],
                    errors="coerce"
                ).fillna(0).sum(),
                "TRANSFERÊNCIA": pd.to_numeric(
                    df["SUGESTAO_TRANSFERENCIA"],
                    errors="coerce"
                ).fillna(0).sum(),
                "PARADOS > 90D": pd.to_numeric(
                    df["ESTOQUE_PARADO_90D"],
                    errors="coerce"
                ).fillna(0).sum(),
            })

        if resumo_filiais:
            df_resumo = pd.DataFrame(resumo_filiais)
            st.dataframe(
                df_resumo,
                use_container_width=True,
                hide_index=True,
            )

            fig = px.bar(
                df_resumo,
                x="FILIAL",
                y=["COMPRA", "TRANSFERÊNCIA"],
                barmode="group",
                title="Compras x Transferências por filial",
            )
            st.plotly_chart(fig, use_container_width=True)

    # --------------------------------------------------------
    # TOP URGENTES
    # --------------------------------------------------------
    with aba2:
        todos_resultados = [
            df for df in dfs.values()
            if df is not None and not df.empty
        ]

        if todos_resultados:
            df_todos = pd.concat(
                todos_resultados,
                ignore_index=True
            )

            df_urgentes = df_todos[
                (
                    pd.to_numeric(
                        df_todos["SUGESTAO_COMPRA"],
                        errors="coerce"
                    ).fillna(0) > 0
                )
                |
                (
                    pd.to_numeric(
                        df_todos["SUGESTAO_TRANSFERENCIA"],
                        errors="coerce"
                    ).fillna(0) > 0
                )
            ].copy()

            if not df_urgentes.empty:
                st.dataframe(
                    df_urgentes[
                        [
                            "FILIAL",
                            "CODIGO",
                            "DESCRICAO",
                            "FORNECEDOR",
                            "MEDIA",
                            "ESTOQUE",
                            "SUGESTAO_COMPRA",
                            "SUGESTAO_TRANSFERENCIA",
                            "ORIGEM_TRANSFERENCIA",
                            "DESTINO_TRANSFERENCIA",
                            "SITUACAO",
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.success("Nenhuma compra ou transferência urgente encontrada.")

    # --------------------------------------------------------
    # ESTOQUE PARADO
    # --------------------------------------------------------
    with aba3:
        todos_resultados = [
            df for df in dfs.values()
            if df is not None and not df.empty
        ]

        if todos_resultados:
            df_todos = pd.concat(
                todos_resultados,
                ignore_index=True
            )

            df_parados = df_todos[
                pd.to_numeric(
                    df_todos["ESTOQUE_PARADO_90D"],
                    errors="coerce"
                ).fillna(0) > 0
            ].copy()

            if not df_parados.empty:
                st.dataframe(
                    df_parados[
                        [
                            "FILIAL",
                            "CODIGO",
                            "DESCRICAO",
                            "ESTOQUE",
                            "ESTOQUE_PARADO_90D",
                            "SUGESTAO_TRANSFERENCIA",
                            "DESTINO_TRANSFERENCIA",
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )

                grafico = (
                    df_parados
                    .groupby("FILIAL", as_index=False)[
                        "ESTOQUE_PARADO_90D"
                    ]
                    .sum()
                    .sort_values(
                        "ESTOQUE_PARADO_90D",
                        ascending=False
                    )
                )

                fig = px.bar(
                    grafico,
                    x="FILIAL",
                    y="ESTOQUE_PARADO_90D",
                    title="Estoque parado há mais de 90 dias por filial",
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True
                )
            else:
                st.success(
                    "Nenhum produto sem giro por 90 dias foi identificado."
                )

    # --------------------------------------------------------
    # PRÉVIA POR FILIAL
    # --------------------------------------------------------
    with aba4:
        for filial, df in dfs.items():
            st.subheader(f"🏭 {filial}")

            if df.empty:
                st.info("Nenhum produto encontrado nesta filial.")
                continue

            colunas_exibicao = [
                "CODIGO",
                "DESCRICAO",
                "FORNECEDOR",
                *[
                    f"MES_{i}"
                    for i in range(1, 5)
                ],
                "MEDIA",
                "ESTOQUE",
                "RESERVA",
                "COMPRADA",
                "SUGESTAO_COMPRA",
                "SUGESTAO_TRANSFERENCIA",
                "ORIGEM_TRANSFERENCIA",
                "DESTINO_TRANSFERENCIA",
                "ESTOQUE_PARADO_90D",
                "SITUACAO",
            ]

            colunas_exibicao = [
                c for c in colunas_exibicao
                if c in df.columns
            ]

            st.dataframe(
                df[colunas_exibicao],
                use_container_width=True,
                hide_index=True,
            )
