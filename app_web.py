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


# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Portal Compras - Tapeçaria",
    layout="wide"
)


# ============================================================
# TELA DE SENHA
# ============================================================

SENHA_ACESSO = "Tape2026"

if "liberado" not in st.session_state:
    st.session_state.liberado = False

if not st.session_state.liberado:

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        st.markdown(
            "<br><br><br><br>",
            unsafe_allow_html=True
        )

        sc1, sc2, sc3 = st.columns([1.5, 1, 1.5])

        with sc2:

            try:
                st.image(
                    "logo.png",
                    use_container_width=True
                )

            except Exception:

                st.markdown(
                    "<h1 style='text-align: center; color: white;'>🏢</h1>",
                    unsafe_allow_html=True
                )

        st.markdown(
            "<h2 style='text-align: center; color: #FFFFFF;'>Acesso Restrito</h2>",
            unsafe_allow_html=True
        )

        st.markdown(
            "<p style='text-align: center; color: #4DA8DA;'>"
            "Insira a senha de sistema para aceder à inteligência de compras."
            "</p>",
            unsafe_allow_html=True
        )

        senha = st.text_input(
            "Senha",
            type="password"
        )

        if st.button(
            "Entrar no Portal",
            use_container_width=True
        ):

            if senha == SENHA_ACESSO:

                st.session_state.liberado = True
                st.rerun()

            else:

                st.error(
                    "Senha incorreta. Tente novamente."
                )

    st.stop()


# ============================================================
# VARIÁVEIS DE SESSÃO
# ============================================================

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

if "analise_concluida" not in st.session_state:
    st.session_state.analise_concluida = False

if "df_regras" not in st.session_state:

    st.session_state.df_regras = pd.DataFrame([

        {
            "FORNECEDOR": "CORTTEX",
            "MULTIPLO": 50,
            "TOLERANCIA": 20,
            "PALAVRA_CHAVE": ""
        },

        {
            "FORNECEDOR": "TEX COMPANY",
            "MULTIPLO": 50,
            "TOLERANCIA": 20,
            "PALAVRA_CHAVE": ""
        },

        {
            "FORNECEDOR": "CIPATEX",
            "MULTIPLO": 50,
            "TOLERANCIA": 20,
            "PALAVRA_CHAVE": ""
        },

        {
            "FORNECEDOR": "KARSTEN",
            "MULTIPLO": 50,
            "TOLERANCIA": 20,
            "PALAVRA_CHAVE": ""
        },

        {
            "FORNECEDOR": "ETRURIA",
            "MULTIPLO": 50,
            "TOLERANCIA": 20,
            "PALAVRA_CHAVE": ""
        },

        {
            "FORNECEDOR": "TELLAIO",
            "MULTIPLO": 50,
            "TOLERANCIA": 20,
            "PALAVRA_CHAVE": ""
        },

        {
            "FORNECEDOR": "OBER",
            "MULTIPLO": 50,
            "TOLERANCIA": 20,
            "PALAVRA_CHAVE": ""
        },

        {
            "FORNECEDOR": "TEXTIL J. SERRANO",
            "MULTIPLO": 50,
            "TOLERANCIA": 20,
            "PALAVRA_CHAVE": ""
        },

        {
            "FORNECEDOR": "CKS",
            "MULTIPLO": 50,
            "TOLERANCIA": 20,
            "PALAVRA_CHAVE": ""
        },

        {
            "FORNECEDOR": "AGRO QUIMICA",
            "MULTIPLO": 45,
            "TOLERANCIA": 20,
            "PALAVRA_CHAVE": ""
        },

        {
            "FORNECEDOR": "ROMPLAS",
            "MULTIPLO": 30,
            "TOLERANCIA": 15,
            "PALAVRA_CHAVE": "URUGUA"
        },

        {
            "FORNECEDOR": "ROMA DUBLADOS",
            "MULTIPLO": 10,
            "TOLERANCIA": 5,
            "PALAVRA_CHAVE": ""
        }

    ])


# ============================================================
# LIMPAR DADOS
# ============================================================

def limpar_dados():

    st.session_state.uploader_key += 1

    st.session_state.analise_concluida = False

    for chave in [
        "dfs_por_filial",
        "df_p",
        "meses_globais",
        "dash_qtd_comprar",
        "dash_qtd_transferida",
        "dash_itens_pico",
        "dash_itens_ruptura"
    ]:

        st.session_state.pop(
            chave,
            None
        )

    st.rerun()


# ============================================================
# LIMPEZA DE VALORES
# ============================================================

def limpar_v(val):

    if val is None:
        return 0.0

    texto = str(val).strip()

    if texto in [
        "",
        "-",
        "S/N",
        "N/A",
        "nan",
        "None"
    ]:
        return 0.0

    texto = (
        texto
        .replace(".", "")
        .replace(",", ".")
    )

    try:

        return float(texto)

    except (ValueError, TypeError):

        return 0.0


# ============================================================
# EXTRAÇÃO DO PDF
# ============================================================

def extrair_dados_pdf_web(file):

    dados = []

    meses_cabecalho = []

    nome_filial = re.sub(
        r"\.pdf$",
        "",
        file.name,
        flags=re.IGNORECASE
    ).upper()

    fornecedor_atual = "DESCONHECIDO"

    with pdfplumber.open(file) as pdf:

        for page in pdf.pages:

            text = page.extract_text()

            if not text:
                continue

            linhas = text.split("\n")

            for l in linhas:

                l_str = l.strip()

                # ------------------------------------------------
                # FORNECEDOR
                # ------------------------------------------------

                if "FORNECEDOR:" in l_str.upper():

                    partes_forn = (
                        l_str
                        .upper()
                        .split("FORNECEDOR:")
                    )

                    if len(partes_forn) > 1:

                        fornecedor_atual = (
                            partes_forn[1]
                            .split("-")[0]
                            .strip()
                        )

                # ------------------------------------------------
                # CABEÇALHO DOS MESES
                # ------------------------------------------------

                if (
                    "CÓDIGO" in l_str.upper()
                    and "DESCRIÇÃO" in l_str.upper()
                ):

                    partes_h = l_str.split()

                    cand_meses = [
                        p
                        for p in partes_h
                        if len(p) == 3
                        and p.isalpha()
                    ]

                    if (
                        len(cand_meses) >= 4
                        and not meses_cabecalho
                    ):

                        meses_cabecalho = [
                            m.upper()
                            for m in cand_meses[:4]
                        ]

                # ------------------------------------------------
                # PRODUTO
                # ------------------------------------------------

                l_limpa = re.sub(
                    r"^\*+\s*",
                    "",
                    l_str
                )

                match_cod = re.search(
                    r"\b\d{3,6}\b",
                    l_limpa
                )

                if not match_cod:
                    continue

                codigo = match_cod.group(0)

                partes = l_limpa.split()

                if len(partes) < 12:
                    continue

                try:

                    item_dict = {

                        "CODIGO": codigo,

                        "DESCRICAO": " ".join([
                            p
                            for p in partes
                            if (
                                not p
                                .replace(",", ".")
                                .replace("-", "")
                                .replace(".", "")
                                .isdigit()
                                and p != codigo
                            )
                        ])[:50],

                        "EMB.": partes[-12],

                        "MES_1": limpar_v(
                            partes[-11]
                        ),

                        "MES_2": limpar_v(
                            partes[-10]
                        ),

                        "MES_3": limpar_v(
                            partes[-9]
                        ),

                        "MES_4": limpar_v(
                            partes[-8]
                        ),

                        "MEDIA_SISTEMA": limpar_v(
                            partes[-7]
                        ),

                        "ESTOQUE": limpar_v(
                            partes[-6]
                        ),

                        "RESERVA": limpar_v(
                            partes[-5]
                        ),

                        "COMPRADA": limpar_v(
                            partes[-4]
                        ),

                        "SITUACAO": partes[-2],

                        "MESES_ESTOQUE": limpar_v(
                            partes[-1]
                        ),

                        "FILIAL_NOME": nome_filial,

                        "FORNECEDOR": fornecedor_atual

                    }

                    dados.append(
                        item_dict
                    )

                except Exception:
                    continue

    df_res = pd.DataFrame(dados)

    if (
        not meses_cabecalho
        or len(meses_cabecalho) < 4
    ):

        meses_cabecalho = [
            "MÊS 1",
            "MÊS 2",
            "MÊS 3",
            "MÊS 4"
        ]

    return (
        df_res,
        meses_cabecalho
    )


# ============================================================
# FUNÇÃO DE EXPORTAÇÃO DO EXCEL
# ============================================================

def gerar_excel(
    dfs_por_filial,
    meses_globais
):

    output = BytesIO()

    wb = Workbook()

    # Remove a aba padrão
    ws_padrao = wb.active

    wb.remove(
        ws_padrao
    )

    # ========================================================
    # ESTILOS
    # ========================================================

    amarelo = PatternFill(
        fill_type="solid",
        fgColor="FFFF00"
    )

    verde = PatternFill(
        fill_type="solid",
        fgColor="D9EAD3"
    )

    azul = PatternFill(
        fill_type="solid",
        fgColor="C9DAF8"
    )

    vermelho = PatternFill(
        fill_type="solid",
        fgColor="FFD2D2"
    )

    vermelho_claro = PatternFill(
        fill_type="solid",
        fgColor="F4CCCC"
    )

    amarelo_pico = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC"
    )

    laranja = PatternFill(
        fill_type="solid",
        fgColor="C55A11"
    )

    fonte_branca = Font(
        color="FFFFFF",
        bold=True
    )

    fonte_cabecalho = Font(
        bold=True
    )

    borda_fina = Border(
        left=Side(
            style="thin",
            color="D9D9D9"
        ),

        right=Side(
            style="thin",
            color="D9D9D9"
        ),

        top=Side(
            style="thin",
            color="D9D9D9"
        ),

        bottom=Side(
            style="thin",
            color="D9D9D9"
        )
    )

    alinhamento_centro = Alignment(
        horizontal="center",
        vertical="center"
    )

    alinhamento_esquerda = Alignment(
        horizontal="left",
        vertical="center"
    )

    # ========================================================
    # GARANTE OS 4 MESES
    # ========================================================

    meses = list(
        meses_globais or []
    )

    while len(meses) < 4:

        meses.append(
            f"MÊS {len(meses) + 1}"
        )

    meses = meses[:4]

    # ========================================================
    # CRIA UMA ABA POR FILIAL
    # ========================================================

    for nome_filial, df in dfs_por_filial.items():

        # O Excel permite no máximo 31 caracteres
        nome_aba = str(
            nome_filial
        )[:31]

        nome_original = nome_aba

        contador = 1

        while nome_aba in wb.sheetnames:

            sufixo = f"_{contador}"

            nome_aba = (
                nome_original[
                    :31 - len(sufixo)
                ]
                + sufixo
            )

            contador += 1

        ws = wb.create_sheet(
            title=nome_aba
        )

        # ====================================================
        # CABEÇALHO
        # ====================================================

        colunas = [

            "CODIGO",

            "DESCRICAO",

            "EMB.",

            meses[0],
            meses[1],
            meses[2],
            meses[3],

            "MEDIA",

            "ESTOQUE",

            "RESERVA",

            "COMPRADA",

            "MESES",

            "SUGESTAO COMPRA",

            "VENDA_ATIPICA",

            "ESTOQUE PARADO",

            "RUPTURA CRÍTICA",

            # AUXILIARES OCULTAS
            "SUGESTAO ORIGINAL",

            "TRANS INTERNA",

            "TRANSF. ESTOQUE MORTO",

            "FORNECEDOR"
        ]

        for col_idx, nome_coluna in enumerate(
            colunas,
            start=1
        ):

            cell = ws.cell(
                row=1,
                column=col_idx,
                value=nome_coluna
            )

            cell.fill = amarelo

            cell.font = fonte_cabecalho

            cell.alignment = (
                alinhamento_centro
            )

            cell.border = borda_fina

        # ====================================================
        # DADOS
        # ====================================================

        for linha_idx, (_, row) in enumerate(
            df.iterrows(),
            start=2
        ):

            sugestao_compra = limpar_v(
                row.get(
                    "SUGESTAO COMPRA",
                    0
                )
            )

            transferencia = str(
                row.get(
                    "TRANS INTERNA",
                    "0"
                )
            )

            # -----------------------------------------------
            # QUANTIDADE DE TRANSFERÊNCIA
            # -----------------------------------------------

            qtd_transf = 0.0

            if transferencia not in [
                "0",
                "",
                "None",
                "nan"
            ]:

                match = re.search(
                    r"(\d+(?:[.,]\d+)?)",
                    transferencia
                )

                if match:

                    qtd_transf = limpar_v(
                        match.group(1)
                    )

            transferencia_morta = (
                "ESTOQUE MORTO"
                in transferencia.upper()
            )

            # -----------------------------------------------
            # VALORES DA LINHA
            # -----------------------------------------------

            valores = [

                row.get(
                    "CODIGO",
                    ""
                ),

                row.get(
                    "DESCRICAO",
                    ""
                ),

                row.get(
                    "EMB.",
                    ""
                ),

                row.get(
                    "MES_1",
                    0
                ),

                row.get(
                    "MES_2",
                    0
                ),

                row.get(
                    "MES_3",
                    0
                ),

                row.get(
                    "MES_4",
                    0
                ),

                row.get(
                    "MEDIA",
                    row.get(
                        "MEDIA_SISTEMA",
                        0
                    )
                ),

                row.get(
                    "ESTOQUE",
                    0
                ),

                row.get(
                    "RESERVA",
                    0
                ),

                row.get(
                    "COMPRADA",
                    0
                ),

                row.get(
                    "MESES_ESTOQUE",
                    0
                ),

                sugestao_compra,

                row.get(
                    "VENDA_ATIPICA",
                    "NÃO"
                ),

                row.get(
                    "ESTOQUE PARADO",
                    "NÃO"
                ),

                row.get(
                    "RUPTURA CRÍTICA",
                    "OK"
                ),

                # -------------------------------------------
                # AUXILIARES
                # -------------------------------------------

                sugestao_compra,

                transferencia,

                (
                    "SIM"
                    if transferencia_morta
                    else "NÃO"
                ),

                row.get(
                    "FORNECEDOR",
                    ""
                )
            ]

            # -----------------------------------------------
            # ESCREVE CÉLULAS
            # -----------------------------------------------

            for col_idx, valor in enumerate(
                valores,
                start=1
            ):

                cell = ws.cell(
                    row=linha_idx,
                    column=col_idx,
                    value=valor
                )

                cell.border = borda_fina

                if col_idx == 2:

                    cell.alignment = (
                        alinhamento_esquerda
                    )

                else:

                    cell.alignment = (
                        alinhamento_centro
                    )

            # -----------------------------------------------
            # SUGESTÃO DE COMPRA
            # -----------------------------------------------

            if sugestao_compra > 0:

                ws.cell(
                    linha_idx,
                    13
                ).fill = verde

            # -----------------------------------------------
            # VENDA ATÍPICA
            # -----------------------------------------------

            if "SIM" in str(
                row.get(
                    "VENDA_ATIPICA",
                    ""
                )
            ):

                ws.cell(
                    linha_idx,
                    14
                ).fill = amarelo_pico

            # -----------------------------------------------
            # ESTOQUE PARADO
            # -----------------------------------------------

            if "SIM" in str(
                row.get(
                    "ESTOQUE PARADO",
                    ""
                )
            ):

                ws.cell(
                    linha_idx,
                    15
                ).fill = vermelho_claro

                ws.cell(
                    linha_idx,
                    9
                ).fill = vermelho_claro

            # -----------------------------------------------
            # RUPTURA
            # -----------------------------------------------

            if "CRÍTICA" in str(
                row.get(
                    "RUPTURA CRÍTICA",
                    ""
                )
            ):

                ws.cell(
                    linha_idx,
                    16
                ).fill = vermelho

                ws.cell(
                    linha_idx,
                    9
                ).fill = vermelho

                ws.cell(
                    linha_idx,
                    16
                ).font = Font(
                    bold=True
                )

            # -----------------------------------------------
            # TRANSFERÊNCIA
            # -----------------------------------------------

            if qtd_transf > 0:

                if transferencia_morta:

                    ws.cell(
                        linha_idx,
                        18
                    ).fill = laranja

                    ws.cell(
                        linha_idx,
                        18
                    ).font = fonte_branca

                else:

                    ws.cell(
                        linha_idx,
                        18
                    ).fill = azul

            # -----------------------------------------------
            # FORMATAÇÃO NUMÉRICA
            # -----------------------------------------------

            for col_idx in range(
                4,
                14
            ):

                ws.cell(
                    linha_idx,
                    col_idx
                ).number_format = (
                    '#,##0.00'
                )

        # ====================================================
        # LARGURA DAS COLUNAS
        # ====================================================

        larguras = {

            "A": 10,

            "B": 42,

            "C": 10,

            "D": 11,
            "E": 11,
            "F": 11,
            "G": 11,

            "H": 11,

            "I": 12,

            "J": 12,

            "K": 12,

            "L": 10,

            "M": 20,

            "N": 17,

            "O": 19,

            "P": 20,

            "Q": 20,

            "R": 24,

            "S": 20,

            "T": 25
        }

        for coluna, largura in larguras.items():

            ws.column_dimensions[
                coluna
            ].width = largura

        # ====================================================
        # OCULTA COLUNAS AUXILIARES
        # ====================================================

        ws.column_dimensions[
            "Q"
        ].hidden = True

        ws.column_dimensions[
            "R"
        ].hidden = True

        ws.column_dimensions[
            "S"
        ].hidden = True

        ws.column_dimensions[
            "T"
        ].hidden = True

        # ====================================================
        # CONGELA CABEÇALHO
        # ====================================================

        ws.freeze_panes = "A2"

        # ====================================================
        # FILTRO
        # ====================================================

        if ws.max_row >= 1:

            ws.auto_filter.ref = (
                f"A1:P{ws.max_row}"
            )

        # ====================================================
        # ALTURA DO CABEÇALHO
        # ====================================================

        ws.row_dimensions[
            1
        ].height = 24

    # ========================================================
    # SALVA NA MEMÓRIA
    # ========================================================

    wb.save(
        output
    )

    output.seek(0)

    return output


# ============================================================
# BARRA LATERAL
# ============================================================

with st.sidebar:

    try:

        st.image(
            "logo.png",
            use_container_width=True
        )

    except Exception:
        pass

    st.markdown("---")

    st.header(
        "📂 Nova Compra"
    )

    uploaded_files = st.file_uploader(

        "Selecione os 4 PDFs das Unidades",

        type="pdf",

        accept_multiple_files=True,

        key=(
            f"pdf_uploader_"
            f"{st.session_state.uploader_key}"
        )
    )

    st.button(

        "🧹 Limpar Dados para Nova Compra",

        on_click=limpar_dados,

        use_container_width=True
    )

    st.markdown("---")

    # ========================================================
    # CONFIGURAÇÕES
    # ========================================================

    with st.expander(
        "⚙️ Configurações Avançadas"
    ):

        meta = st.number_input(

            "Meta de estoque (meses)",

            min_value=1,

            value=2
        )

        retencao_seguranca = st.number_input(

            "Retenção de Segurança Doadora (meses)",

            min_value=0,

            value=3
        )

        fator_pico = st.number_input(

            "Sensibilidade de Pico (x vezes a média)",

            min_value=1.5,

            value=2.5,

            step=0.5
        )

        nome_sugerido = st.text_input(

            "Nome do ficheiro Excel",

            value="Relatorio_Compras_Tapecaria"
        )

        nome_final_xlsx = (

            nome_sugerido

            if nome_sugerido.lower().endswith(
                ".xlsx"
            )

            else
            f"{nome_sugerido}.xlsx"
        )

    # ========================================================
    # FORNECEDORES
    # ========================================================

    with st.expander(
        "🏭 Fornecedores e Múltiplos"
    ):

        st.subheader(
            "Adicionar / Editar Regra"
        )

        novo_forn = st.text_input(
            "Nome Fornecedor"
        ).upper().strip()

        novo_mult = st.number_input(
            "Múltiplo",
            min_value=1,
            value=50
        )

        if st.button(
            "➕ Gravar Fornecedor",
            use_container_width=True
        ):

            if novo_forn:

                forns_existentes = (
                    st.session_state.df_regras[
                        "FORNECEDOR"
                    ]
                    .str.upper()
                    .tolist()
                )

                if novo_forn in forns_existentes:

                    st.warning(
                        f"⚠️ O fornecedor "
                        f"'{novo_forn}' já está cadastrado!"
                    )

                else:

                    nova_linha = pd.DataFrame([

                        {
                            "FORNECEDOR": novo_forn,
                            "MULTIPLO": novo_mult,
                            "TOLERANCIA": 0,
                            "PALAVRA_CHAVE": ""
                        }

                    ])

                    st.session_state.df_regras = (
                        pd.concat(
                            [
                                st.session_state.df_regras,
                                nova_linha
                            ],
                            ignore_index=True
                        )
                    )

                    st.success(
                        f"✅ Fornecedor "
                        f"'{novo_forn}' gravado com sucesso!"
                    )

                    st.rerun()

        st.caption(
            "Tabela Geral de Múltiplos:"
        )

        df_regras_editado = st.data_editor(

            st.session_state.df_regras,

            num_rows="dynamic",

            use_container_width=True,

            hide_index=True
        )

        st.session_state.df_regras = (
            df_regras_editado
        )


# ============================================================
# CABEÇALHO PRINCIPAL
# ============================================================

col1, col2 = st.columns(
    [1, 15]
)

with col1:

    try:

        st.image(
            "simbolo.png",
            width=50
        )

    except Exception:
        pass

with col2:

    st.title(
        "Inteligência de Compras"
    )

st.markdown(
    "##### Portal Operacional - Tapeçaria"
)

st.markdown(
    "<br>",
    unsafe_allow_html=True
)


# ============================================================
# PROCESSAMENTO AUTOMÁTICO
# ============================================================

if uploaded_files:

    with st.spinner(
        "🔍 A executar Inteligência de Compras..."
    ):

        dfs_por_filial = {}

        todos_dados = []

        meses_globais = []

        # ====================================================
        # LÊ OS PDFs
        # ====================================================

        for f in uploaded_files:

            df, meses = (
                extrair_dados_pdf_web(f)
            )

            if not df.empty:

                nome_filial = re.sub(

                    r"\.pdf$",

                    "",

                    f.name,

                    flags=re.IGNORECASE
                ).upper()

                dfs_por_filial[
                    nome_filial
                ] = df

                todos_dados.append(
                    df
                )

                if (
                    len(meses) >= 4
                    and not meses_globais
                ):

                    meses_globais = (
                        meses[:4]
                    )

        if not meses_globais:

            meses_globais = [

                "MÊS 1",
                "MÊS 2",
                "MÊS 3",
                "MÊS 4"

            ]

        # ====================================================
        # NENHUM DADO
        # ====================================================

        if not todos_dados:

            st.error(
                "⚠️ O sistema não encontrou "
                "produtos compatíveis nos PDFs."
            )

            st.session_state.analise_concluida = (
                False
            )

        else:

            try:

                # =================================================
                # CONSOLIDA DADOS
                # =================================================

                df_global = pd.concat(
                    todos_dados,
                    ignore_index=True
                )

                df_global[
                    "ESTOQUE_DISPONIVEL"
                ] = df_global[
                    "ESTOQUE"
                ]

                vendas_recentes = (

                    df_global["MES_1"]

                    + df_global["MES_2"]

                    + df_global["MES_3"]

                    + df_global["MES_4"]

                )

                df_global[
                    "TOTAL_VENDAS_RECENTES"
                ] = vendas_recentes

                # =================================================
                # TRACKER DE ESTOQUE
                # =================================================

                tracker_estoque = {}

                for _, row in (
                    df_global.iterrows()
                ):

                    f_nome = row[
                        "FILIAL_NOME"
                    ]

                    codigo = row[
                        "CODIGO"
                    ]

                    est = float(
                        row["ESTOQUE"]
                    )

                    med = float(
                        row["MEDIA_SISTEMA"]
                    )

                    # Estoque morto:
                    # libera 100%

                    if med == 0:

                        excesso = est

                    else:

                        excesso = max(

                            0.0,

                            est
                            - (
                                med
                                * retencao_seguranca
                            )

                        )

                    tracker_estoque[
                        (
                            f_nome,
                            codigo
                        )
                    ] = {

                        "EXCEDENTE": excesso,

                        "MEDIA": med,

                        "ESTOQUE_FINAL": est

                    }

                # =================================================
                # INDICADORES
                # =================================================

                dash_qtd_comprar = 0

                dash_qtd_transferida = 0

                dash_itens_pico = 0

                dash_itens_ruptura = 0

                # =================================================
                # REGRAS DE FORNECEDORES
                # =================================================

                regras_dict = dict(

                    zip(

                        st.session_state.df_regras[
                            "FORNECEDOR"
                        ]
                        .fillna("")
                        .astype(str)
                        .str.upper(),

                        pd.to_numeric(

                            st.session_state.df_regras[
                                "MULTIPLO"
                            ],

                            errors="coerce"

                        ).fillna(1)

                    )

                )

                # =================================================
                # PROCESSAMENTO POR FILIAL
                # =================================================

                for (
                    f_nome,
                    df_f
                ) in dfs_por_filial.items():

                    suge_compra = []

                    trans_interna = []

                    ruptura_critica = []

                    venda_atipica = []

                    est_parado = []

                    medias = []

                    for _, row in (
                        df_f.iterrows()
                    ):

                        codigo = row[
                            "CODIGO"
                        ]

                        med = float(
                            row[
                                "MEDIA_SISTEMA"
                            ]
                        )

                        est = float(
                            row[
                                "ESTOQUE"
                            ]
                        )

                        comp = float(
                            row[
                                "COMPRADA"
                            ]
                        )

                        res = float(
                            row[
                                "RESERVA"
                            ]
                        )

                        fornecedor = str(
                            row[
                                "FORNECEDOR"
                            ]
                        ).upper()

                        # =========================================
                        # RUPTURA
                        # =========================================

                        if (

                            est == 0

                            and comp == 0

                            and med > 0

                        ):

                            ruptura_critica.append(
                                "🚨 CRÍTICA"
                            )

                            dash_itens_ruptura += 1

                        else:

                            ruptura_critica.append(
                                "OK"
                            )

                        # =========================================
                        # PICO
                        # =========================================

                        max_venda = max(

                            row["MES_1"],

                            row["MES_2"],

                            row["MES_3"],

                            row["MES_4"]

                        )

                        if (

                            max_venda
                            > (
                                med
                                * fator_pico
                            )

                            and med > 0

                        ):

                            venda_atipica.append(
                                "⚠️ SIM"
                            )

                            dash_itens_pico += 1

                        else:

                            venda_atipica.append(
                                "NÃO"
                            )

                        # =========================================
                        # ESTOQUE PARADO
                        # =========================================

                        vendas_tot = (

                            row["MES_1"]

                            + row["MES_2"]

                            + row["MES_3"]

                            + row["MES_4"]

                        )

                        if (

                            vendas_tot == 0

                            and est > 0

                        ):

                            est_parado.append(
                                "🛑 SIM"
                            )

                        else:

                            est_parado.append(
                                "NÃO"
                            )

                        # =========================================
                        # NECESSIDADE
                        # =========================================

                        necessidade = max(

                            0.0,

                            (
                                med
                                * meta
                            )

                            - (

                                est
                                + comp
                                - res

                            )

                        )

                        qtd_transf = 0.0

                        origem = ""

                        veio_de_estoque_morto = False

                        # =========================================
                        # TRANSFERÊNCIA
                        # =========================================

                        if necessidade > 0:

                            for (

                                outra_f,
                                cod_item

                            ), d_est in (
                                tracker_estoque.items()
                            ):

                                if (

                                    cod_item == codigo

                                    and
                                    outra_f != f_nome

                                    and
                                    d_est[
                                        "EXCEDENTE"
                                    ] > 0

                                ):

                                    atend = min(

                                        necessidade,

                                        d_est[
                                            "EXCEDENTE"
                                        ]

                                    )

                                    qtd_transf += (
                                        atend
                                    )

                                    d_est[
                                        "EXCEDENTE"
                                    ] -= atend

                                    necessidade -= (
                                        atend
                                    )

                                    origem = (
                                        outra_f
                                    )

                                    if (
                                        d_est[
                                            "MEDIA"
                                        ] == 0
                                    ):

                                        veio_de_estoque_morto = True

                                    dash_qtd_transferida += (
                                        atend
                                    )

                                    if (
                                        necessidade
                                        <= 0
                                    ):

                                        break

                        # =========================================
                        # COMPRA
                        # =========================================

                        mult = regras_dict.get(

                            fornecedor,

                            1

                        )

                        try:

                            mult = float(
                                mult
                            )

                        except Exception:

                            mult = 1

                        if mult <= 0:
                            mult = 1

                        qtd_compra = (

                            math.ceil(

                                necessidade
                                / mult

                            )

                            * mult

                            if necessidade > 0

                            else 0

                        )

                        qtd_compra = float(
                            qtd_compra
                        )

                        dash_qtd_comprar += (
                            qtd_compra
                        )

                        suge_compra.append(
                            qtd_compra
                        )

                        # =========================================
                        # TEXTO DA TRANSFERÊNCIA
                        # =========================================

                        if qtd_transf > 0:

                            txt_transf = (
                                f"{qtd_transf:.0f} "
                                f"DE {origem}"
                            )

                        else:

                            txt_transf = "0"

                        if veio_de_estoque_morto:

                            txt_transf += (
                                " (ESTOQUE MORTO)"
                            )

                        trans_interna.append(
                            txt_transf
                        )

                        medias.append(
                            med
                        )

                    # =============================================
                    # DEVOLVE RESULTADOS AO DATAFRAME
                    # =============================================

                    df_f[
                        "SUGESTAO COMPRA"
                    ] = suge_compra

                    df_f[
                        "TRANS INTERNA"
                    ] = trans_interna

                    df_f[
                        "RUPTURA CRÍTICA"
                    ] = ruptura_critica

                    df_f[
                        "VENDA_ATIPICA"
                    ] = venda_atipica

                    df_f[
                        "ESTOQUE PARADO"
                    ] = est_parado

                    df_f[
                        "MEDIA"
                    ] = medias

                # =================================================
                # ESTOQUE PARADO GLOBAL
                # =================================================

                df_p = df_global[

                    (

                        df_global[
                            "TOTAL_VENDAS_RECENTES"
                        ] == 0

                    )

                    &

                    (

                        df_global[
                            "ESTOQUE_DISPONIVEL"
                        ] > 0

                    )

                ].copy()

                # =================================================
                # SALVA TUDO NA SESSÃO
                # =================================================

                st.session_state.dfs_por_filial = (
                    dfs_por_filial
                )

                st.session_state.dash_qtd_comprar = (
                    dash_qtd_comprar
                )

                st.session_state.dash_qtd_transferida = (
                    dash_qtd_transferida
                )

                st.session_state.dash_itens_pico = (
                    dash_itens_pico
                )

                st.session_state.dash_itens_ruptura = (
                    dash_itens_ruptura
                )

                st.session_state.df_p = (
                    df_p
                )

                # IMPORTANTE:
                # guardamos os meses para a exportação
                st.session_state.meses_globais = (
                    meses_globais
                )

                st.session_state.analise_concluida = (
                    True
                )

            except Exception as e:

                st.session_state.analise_concluida = (
                    False
                )

                st.error(
                    f"🚨 Ocorreu um erro durante os cálculos: {e}"
                )

                st.code(
                    traceback.format_exc()
                )


# ============================================================
# RESULTADOS
# ============================================================

if st.session_state.get(
    "analise_concluida",
    False
):

    dfs_por_filial = (
        st.session_state.dfs_por_filial
    )

    dash_qtd_comprar = (
        st.session_state.dash_qtd_comprar
    )

    dash_qtd_transferida = (
        st.session_state.dash_qtd_transferida
    )

    dash_itens_pico = (
        st.session_state.dash_itens_pico
    )

    dash_itens_ruptura = (
        st.session_state.dash_itens_ruptura
    )

    df_p = (
        st.session_state.df_p
    )

    meses_globais = (
        st.session_state.get(

            "meses_globais",

            [
                "MÊS 1",
                "MÊS 2",
                "MÊS 3",
                "MÊS 4"
            ]

        )
    )

    # ========================================================
    # ABAS
    # ========================================================

    tab1, tab2, tab3, tab4 = st.tabs([

        "📊 Visão Geral",

        "🚨 Top Urgentes",

        "📦 Estoque Parado",

        "🔍 Prévia por Filial"

    ])

    # ========================================================
    # VISÃO GERAL
    # ========================================================

    with tab1:

        st.subheader(
            "Indicadores de Desempenho"
        )

        c1, c2, c3, c4, c5 = (
            st.columns(5)
        )

        c1.metric(

            "🛒 Sugestão de Compra",

            f"{int(dash_qtd_comprar)} un."

        )

        c2.metric(

            "🔄 Economia (Transf.)",

            f"{int(dash_qtd_transferida)} un."

        )

        c3.metric(

            "⚠️ Picos de Vendas",

            f"{int(dash_itens_pico)} itens"

        )

        c4.metric(

            "🚨 Rupturas Críticas",

            f"{int(dash_itens_ruptura)} itens"

        )

        if not df_p.empty:

            f_p = (

                df_p
                .groupby(
                    "FILIAL_NOME"
                )[
                    "ESTOQUE_DISPONIVEL"
                ]
                .sum()
                .idxmax()

            )

        else:

            f_p = "Nenhuma"

        c5.metric(

            "📦 Maior Estoque Parado",

            f_p

        )

        st.success(
            "✅ Processamento concluído com sucesso!"
        )

        # ====================================================
        # EXPORTAÇÃO
        # ====================================================

        st.markdown(
            "---"
        )

        st.subheader(
            "📥 Exportação"
        )

        try:

            arquivo_excel = gerar_excel(

                dfs_por_filial,

                meses_globais

            )

            st.download_button(

                label=(
                    "📊 Baixar Excel - "
                    "Sugestões de Compras e Transferências"
                ),

                data=arquivo_excel.getvalue(),

                file_name=nome_final_xlsx,

                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),

                use_container_width=True

            )

            st.caption(

                "O Excel será gerado com uma aba por filial, "
                "incluindo sugestões de compra e transferências."

            )

        except Exception:

            st.error(
                "❌ Não foi possível gerar o Excel."
            )

            st.code(
                traceback.format_exc()
            )

    # ========================================================
    # TOP URGENTES
    # ========================================================

    with tab2:

        df_all = pd.concat(

            dfs_por_filial.values(),

            ignore_index=True

        )

        df_rupturas = (

            df_all[
                df_all[
                    "RUPTURA CRÍTICA"
                ]
                == "🚨 CRÍTICA"
            ]

            .sort_values(

                by="MEDIA",

                ascending=False

            )

        )

        if not df_rupturas.empty:

            st.error(

                "🚨 PRODUTOS EM RUPTURA CRÍTICA "
                "DETECTADOS (Estoque Zero + Sem Pedido em Andamento)"

            )

            st.dataframe(

                df_rupturas[
                    [
                        "CODIGO",
                        "DESCRICAO",
                        "FILIAL_NOME",
                        "MEDIA",
                        "SUGESTAO COMPRA",
                        "FORNECEDOR"
                    ]
                ],

                use_container_width=True

            )

        else:

            st.success(

                "✅ Nenhuma ruptura crítica absoluta "
                "detectada nas filiais!"

            )

        st.markdown(
            "<br><hr>",
            unsafe_allow_html=True
        )

        st.subheader(
            "🛒 Maior Volume de Compra Sugerido (Top 15)"
        )

        top_compra = (

            df_all[
                df_all[
                    "SUGESTAO COMPRA"
                ] > 0
            ]

            .sort_values(

                by="SUGESTAO COMPRA",

                ascending=False

            )

            .head(15)

        )

        st.dataframe(

            top_compra[
                [
                    "CODIGO",
                    "DESCRICAO",
                    "FILIAL_NOME",
                    "SUGESTAO COMPRA",
                    "FORNECEDOR"
                ]
            ],

            use_container_width=True

        )

        # ====================================================
        # TRANSFERÊNCIAS
        # ====================================================

        st.markdown(
            "<br>",
            unsafe_allow_html=True
        )

        st.subheader(
            "🔄 Transferências Internas Sugeridas"
        )

        df_transf = df_all[

            df_all[
                "TRANS INTERNA"
            ]
            .astype(str)
            .str.strip()
            != "0"

        ]

        if not df_transf.empty:

            st.dataframe(

                df_transf[
                    [
                        "CODIGO",
                        "DESCRICAO",
                        "FILIAL_NOME",
                        "TRANS INTERNA",
                        "FORNECEDOR"
                    ]
                ],

                use_container_width=True

            )

        else:

            st.info(
                "Nenhuma transferência interna foi sugerida."
            )

    # ========================================================
    # ESTOQUE PARADO
    # ========================================================

    with tab3:

        st.subheader(
            "Distribuição de Estoque Excedente / Sem Giro"
        )

        if not df_p.empty:

            grafico_dados = (

                df_p
                .groupby(
                    "FILIAL_NOME"
                )[
                    "ESTOQUE_DISPONIVEL"
                ]
                .sum()
                .reset_index()

            )

            fig = px.bar(

                grafico_dados,

                x="FILIAL_NOME",

                y="ESTOQUE_DISPONIVEL",

                title=(
                    "Volume de Estoque Acima do "
                    "Limite de Giro por Filial"
                ),

                color="ESTOQUE_DISPONIVEL",

                color_continuous_scale="Reds"

            )

            st.plotly_chart(

                fig,

                use_container_width=True

            )

        else:

            st.info(

                "Nenhum estoque crítico/parado foi "
                "detectado com base nos parâmetros configurados."

            )

    # ========================================================
    # PRÉVIA POR FILIAL
    # ========================================================

    with tab4:

        st.subheader(
            "Prévia Colorida dos Dados por Filial"
        )

        sel_f = st.selectbox(

            "Selecione a Filial para visualizar:",

            list(
                dfs_por_filial.keys()
            )

        )

        df_view = (
            dfs_por_filial[
                sel_f
            ].copy()
        )

        def pintar_tabela(row):

            cols = row.index

            estilos = [
                ""
                for _ in range(
                    len(cols)
                )
            ]

            def f_idx(nome):

                if nome in cols:

                    return cols.get_loc(
                        nome
                    )

                return -1

            i_parado = f_idx(
                "ESTOQUE PARADO"
            )

            i_estoque = f_idx(
                "ESTOQUE"
            )

            i_atipica = f_idx(
                "VENDA_ATIPICA"
            )

            i_compra = f_idx(
                "SUGESTAO COMPRA"
            )

            i_transf = f_idx(
                "TRANS INTERNA"
            )

            i_comprada = f_idx(
                "COMPRADA"
            )

            i_ruptura = f_idx(
                "RUPTURA CRÍTICA"
            )

            # -----------------------------------------------
            # ESTOQUE PARADO
            # -----------------------------------------------

            if (

                i_parado >= 0

                and "🛑 SIM"
                in str(
                    row.get(
                        "ESTOQUE PARADO",
                        ""
                    )
                )

            ):

                estilos[
                    i_parado
                ] = (
                    "background-color: #F4CCCC; "
                    "color: black;"
                )

                if i_estoque >= 0:

                    estilos[
                        i_estoque
                    ] = (
                        "background-color: #F4CCCC; "
                        "color: black;"
                    )

            # -----------------------------------------------
            # VENDA ATÍPICA
            # -----------------------------------------------

            if (

                i_atipica >= 0

                and "⚠️ SIM"
                in str(
                    row.get(
                        "VENDA_ATIPICA",
                        ""
                    )
                )

            ):

                estilos[
                    i_atipica
                ] = (
                    "background-color: #FFF2CC; "
                    "color: black;"
                )

            # -----------------------------------------------
            # COMPRA
            # -----------------------------------------------

            if (

                i_compra >= 0

                and pd.to_numeric(

                    row.get(
                        "SUGESTAO COMPRA",
                        0
                    ),

                    errors="coerce"

                ) > 0

            ):

                estilos[
                    i_compra
                ] = (
                    "background-color: #D9EAD3; "
                    "color: black;"
                )

            # -----------------------------------------------
            # TRANSFERÊNCIA
            # -----------------------------------------------

            if (

                i_transf >= 0

                and str(
                    row.get(
                        "TRANS INTERNA",
                        ""
                    )
                ) not in [
                    "0",
                    "None",
                    "",
                    "nan"
                ]

            ):

                if (

                    "ESTOQUE MORTO"

                    in str(
                        row.get(
                            "TRANS INTERNA",
                            ""
                        )
                    )

                ):

                    estilos[
                        i_transf
                    ] = (
                        "background-color: #C55A11; "
                        "color: white; "
                        "font-weight: bold;"
                    )

                else:

                    estilos[
                        i_transf
                    ] = (
                        "background-color: #C9DAF8; "
                        "color: black;"
                    )

            # -----------------------------------------------
            # COMPRADA
            # -----------------------------------------------

            if (

                i_comprada >= 0

                and pd.to_numeric(

                    row.get(
                        "COMPRADA",
                        0
                    ),

                    errors="coerce"

                ) > 0

            ):

                estilos[
                    i_comprada
                ] = (
                    "background-color: #FCE5CD; "
                    "color: black;"
                )

            # -----------------------------------------------
            # RUPTURA
            # -----------------------------------------------

            if (

                i_ruptura >= 0

                and "🚨 CRÍTICA"

                in str(
                    row.get(
                        "RUPTURA CRÍTICA",
                        ""
                    )
                )

            ):

                estilos[
                    i_ruptura
                ] = (
                    "background-color: #FFD2D2; "
                    "color: black; "
                    "font-weight: bold;"
                )

                if i_estoque >= 0:

                    estilos[
                        i_estoque
                    ] = (
                        "background-color: #FFD2D2; "
                        "color: black;"
                    )

            return estilos

        st.dataframe(

            df_view.style.apply(

                pintar_tabela,

                axis=1

            ),

            use_container_width=True

        )

else:

    st.info(

        "Aguardando documentos. Por favor, "
        "selecione os ficheiros PDF na barra lateral "
        "para iniciar a análise."

    )
