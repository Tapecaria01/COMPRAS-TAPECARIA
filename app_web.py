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


# --- INICIALIZAÇÃO DE ESTADO (MÚLTIPLOS E REGRAS) ---
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


# --- FUNÇÕES DE SUPORTE E LEITURA DE PDF ---
def limpar_v(val):
    """Converte valores numéricos no formato brasileiro para float."""
    if val is None or val == "" or val == "-" or val == "S/N":
        return 0.0

    val_limpo = str(val).replace(".", "").replace(",", ".")

    try:
        return float(val_limpo)
    except (ValueError, TypeError):
        return 0.0


def extrair_dados_pdf_web(file):
    """
    Lê o PDF usando Regex Dinâmico do final para o início.
    Mantém a lógica original de captura dos dados.
    """
    dados = []
    meses_cabecalho = []
    nome_filial = file.name.replace(".pdf", "").replace(".PDF", "").upper()
    fornecedor_atual = "DESCONHECIDO"

    # Números no formato BR (1.234,56 / 0,00 / 12) ou '-'
    num_pattern = r"(?:-|\d+(?:\.\d{3})*(?:,\d+)?|\d+)"

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()

            if not text:
                continue

            linhas = text.split("\n")

            for l in linhas:
                l_str = l.strip()

                if not l_str:
                    continue

                # 1. Captura de fornecedor no cabeçalho
                if "FORNECEDOR:" in l_str.upper():
                    partes_forn = l_str.upper().split("FORNECEDOR:")

                    if len(partes_forn) > 1:
                        fornecedor_atual = (
                            partes_forn[1].split("-")[0].strip()
                        )

                    continue

                # 2. Captura dos meses no cabeçalho
                if (
                    "CÓDIGO" in l_str.upper()
                    and "DESCRIÇÃO" in l_str.upper()
                ):
                    partes_h = l_str.split()

                    cand_meses = [
                        p for p in partes_h
                        if len(p) == 3 and p.isalpha()
                    ]

                    if len(cand_meses) >= 4 and not meses_cabecalho:
                        meses_cabecalho = [
                            m.upper() for m in cand_meses[:4]
                        ]

                    continue

                # Remove asteriscos do início da linha
                l_limpa = re.sub(r"^\*+\s*", "", l_str)

                # Busca código no início
                match_cod = re.search(
                    r"^\s*([A-Za-z0-9\-]{3,10})\b",
                    l_limpa
                )

                if not match_cod:
                    match_cod = re.search(r"\b\d{3,7}\b", l_limpa)

                if not match_cod:
                    continue

                codigo = (
                    match_cod.group(1)
                    if match_cod.groups()
                    else match_cod.group(0)
                )

                # Busca os tokens numéricos da linha
                tokens_numericos = re.findall(num_pattern, l_limpa)

                # Linha válida deve possuir os valores principais
                if len(tokens_numericos) < 8:
                    continue

                try:
                    meses_est = limpar_v(tokens_numericos[-1])

                    comprada = (
                        limpar_v(tokens_numericos[-2])
                        if len(tokens_numericos) >= 9
                        else 0.0
                    )

                    reserva = (
                        limpar_v(tokens_numericos[-3])
                        if len(tokens_numericos) >= 10
                        else 0.0
                    )

                    estoque = (
                        limpar_v(tokens_numericos[-4])
                        if len(tokens_numericos) >= 11
                        else 0.0
                    )

                    media = (
                        limpar_v(tokens_numericos[-5])
                        if len(tokens_numericos) >= 12
                        else 0.0
                    )

                    # Meses de vendas
                    m4 = (
                        limpar_v(tokens_numericos[-6])
                        if len(tokens_numericos) >= 13
                        else 0.0
                    )

                    m3 = (
                        limpar_v(tokens_numericos[-7])
                        if len(tokens_numericos) >= 14
                        else 0.0
                    )

                    m2 = (
                        limpar_v(tokens_numericos[-8])
                        if len(tokens_numericos) >= 15
                        else 0.0
                    )

                    m1 = (
                        limpar_v(tokens_numericos[-9])
                        if len(tokens_numericos) >= 16
                        else 0.0
                    )

                    # Isola a descrição
                    desc_raw = re.sub(
                        r"^\s*" + re.escape(codigo),
                        "",
                        l_limpa
                    ).strip()

                    descricao = re.sub(
                        r"(\s+" + num_pattern + r")+$",
                        "",
                        desc_raw
                    ).strip()

                    item_dict = {
                        "CODIGO": codigo,
                        "DESCRICAO": (
                            descricao[:60]
                            if descricao
                            else "ITEM SEM DESCRICAO"
                        ),
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
                    }

                    dados.append(item_dict)

                except Exception:
                    continue

    df_res = pd.DataFrame(dados)

    if not meses_cabecalho or len(meses_cabecalho) < 4:
        meses_cabecalho = [
            "MÊS 1",
            "MÊS 2",
            "MÊS 3",
            "MÊS 4",
        ]

    return df_res, meses_cabecalho


def pintar_tabela(val):
    """Aplica cores dinâmicas nas linhas do DataFrame."""
    status = val.get("STATUS", "")

    if "RUPTURA" in str(status):
        return ["background-color: #ffcccc"] * len(val)

    if "TRANSFERIR" in str(status):
        return ["background-color: #e6f2ff"] * len(val)

    if "EXCESSO" in str(status):
        return ["background-color: #fff2cc"] * len(val)

    return [""] * len(val)


def criar_excel(df_resultados, nome_aba="Sugestão de Compras"):
    """Gera o Excel final formatado e retorna BytesIO."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resultados.to_excel(
            writer,
            index=False,
            sheet_name=nome_aba
        )

        workbook = writer.book
        worksheet = writer.sheets[nome_aba]

        # Cabeçalho
        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )
        header_font = Font(
            bold=True,
            color="FFFFFF"
        )

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # Bordas
        thin = Side(style="thin", color="D9E1F2")
        border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(
                    vertical="center"
                )

        # Largura das colunas
        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(
                column_cells[0].column
            )

            for cell in column_cells:
                try:
                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )
                except Exception:
                    pass

            worksheet.column_dimensions[
                column_letter
            ].width = min(max_length + 2, 50)

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        # Formatação numérica
        headers = {
            cell.value: cell.column
            for cell in worksheet[1]
        }

        for nome_coluna in [
            "MEDIA",
            "ESTOQUE",
            "COMPRADA",
            "SUG_COMPRA",
            "SUG_TRANSF",
        ]:
            coluna = headers.get(nome_coluna)

            if coluna:
                for row_num in range(2, worksheet.max_row + 1):
                    worksheet.cell(
                        row=row_num,
                        column=coluna
                    ).number_format = '#,##0.00'

    output.seek(0)
    return output


# --- INTERFACE WEB (BARRA LATERAL) ---
with st.sidebar:
    try:
        st.image(
            "logo.png",
            use_container_width=True
        )
    except Exception:
        pass

    st.markdown("---")
    st.header("📂 Nova Compra")

    uploaded_files = st.file_uploader(
        "Selecione os PDFs das Unidades",
        type="pdf",
        accept_multiple_files=True
    )

    st.markdown("---")

    with st.expander("⚙️ Configurações Avançadas"):
        meta = st.number_input(
            "Meta de estoque (meses)",
            min_value=1,
            value=2
        )

        meses_parado = st.number_input(
            "Considerar estoque parado após (meses)",
            min_value=1,
            value=3,
            step=1
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
            if nome_sugerido.endswith(".xlsx")
            else f"{nome_sugerido}.xlsx"
        )

    with st.expander("🏭 Fornecedores e Múltiplos"):
        st.caption(
            "Edite ou adicione regras na tabela abaixo."
        )

        df_regras_editado = st.data_editor(
            st.session_state.df_regras,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True
        )

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


# ====================================================
# === PROCESSAMENTO AUTOMÁTICO ===
# ====================================================
if uploaded_files:
    with st.spinner(
        "🔍 Processando arquivos PDF e calculando regras de estoque..."
    ):
        try:
            dfs_por_filial = {}
            todos_dados = []
            meses_globais = []

            # ------------------------------------------
            # LEITURA DOS PDFs
            # ------------------------------------------
            for f in uploaded_files:
                df, meses = extrair_dados_pdf_web(f)

                if not df.empty:
                    nome_arquivo = (
                        f.name
                        .replace(".pdf", "")
                        .replace(".PDF", "")
                        .upper()
                    )

                    dfs_por_filial[nome_arquivo] = df
                    todos_dados.append(df)

                    if len(meses) >= 4 and not meses_globais:
                        meses_globais = meses[:4]

            if not meses_globais:
                meses_globais = [
                    "MÊS 1",
                    "MÊS 2",
                    "MÊS 3",
                    "MÊS 4",
                ]

            # ------------------------------------------
            # NENHUM PRODUTO ENCONTRADO
            # ------------------------------------------
            if not todos_dados:
                st.error(
                    "⚠️ O sistema não encontrou produtos compatíveis nos PDFs."
                )

            else:
                # --------------------------------------
                # CONSOLIDAÇÃO
                # --------------------------------------
                df_global = pd.concat(
                    todos_dados,
                    ignore_index=True
                )

                df_global["ESTOQUE_DISPONIVEL"] = (
                    df_global["ESTOQUE"]
                )

                vendas_recentes = (
                    df_global["MES_1"]
                    + df_global["MES_2"]
                    + df_global["MES_3"]
                    + df_global["MES_4"]
                )

                df_global["TOTAL_VENDAS_RECENTES"] = (
                    vendas_recentes
                )

                # --------------------------------------
                # RASTREADOR DE EXCEDENTES
                # --------------------------------------
                tracker_estoque = {}

                for _, row in df_global.iterrows():
                    f_nome = row["FILIAL_NOME"]
                    c = row["CODIGO"]
                    est = float(row["ESTOQUE"])
                    med = float(row["MEDIA_SISTEMA"])

                    # Se média for 0, estoque morto:
                    # todo estoque pode ser liberado.
                    # Caso contrário, mantém 3 meses de segurança.
                    excesso = (
                        est
                        if med == 0
                        else max(0.0, est - (med * 3))
                    )

                    tracker_estoque[
                        (f_nome, c)
                    ] = {
                        "EXCEDENTE": excesso,
                        "MEDIA": med,
                        "ESTOQUE_FINAL": est,
                    }

                # --------------------------------------
                # VARIÁVEIS DO DASHBOARD
                # --------------------------------------
                resultados = []
                dash_qtd_comprar = 0
                dash_qtd_transferida = 0
                dash_itens_pico = 0
                dash_itens_ruptura = 0

                # --------------------------------------
                # REGRAS DE MÚLTIPLOS
                # --------------------------------------
                regras_temp = (
                    st.session_state.df_regras.copy()
                )

                regras_temp["FORNECEDOR"] = (
                    regras_temp["FORNECEDOR"]
                    .fillna("")
                    .astype(str)
                    .str.upper()
                    .str.strip()
                )

                regras_temp["MULTIPLO"] = pd.to_numeric(
                    regras_temp["MULTIPLO"],
                    errors="coerce"
                ).fillna(1)

                regras_dict = dict(
                    zip(
                        regras_temp["FORNECEDOR"],
                        regras_temp["MULTIPLO"]
                    )
                )

                multiplo_padrao = regras_dict.get(
                    "GERAL",
                    1
                )

                try:
                    multiplo_padrao = float(
                        multiplo_padrao
                    )
                except (TypeError, ValueError):
                    multiplo_padrao = 1

                if multiplo_padrao <= 0:
                    multiplo_padrao = 1

                # --------------------------------------
                # CÁLCULO ITEM A ITEM
                # --------------------------------------
                for _, row in df_global.iterrows():
                    f_nome = row["FILIAL_NOME"]
                    c = row["CODIGO"]

                    med = float(
                        row["MEDIA_SISTEMA"]
                    )
                    est = float(
                        row["ESTOQUE"]
                    )
                    res = float(
                        row["RESERVA"]
                    )
                    comp = float(
                        row["COMPRADA"]
                    )

                    fornecedor = str(
                        row["FORNECEDOR"]
                    ).upper().strip()

                    # ------------------------------
                    # IDENTIFICAÇÃO DE PICO
                    # ------------------------------
                    max_venda = max(
                        row["MES_1"],
                        row["MES_2"],
                        row["MES_3"],
                        row["MES_4"]
                    )

                    eh_pico = (
                        max_venda > (med * fator_pico)
                        and med > 0
                    )

                    if eh_pico:
                        dash_itens_pico += 1

                    # ------------------------------
                    # NECESSIDADE BRUTA
                    # ------------------------------
                    necessidade = max(
                        0.0,
                        (med * meta)
                        - (est + comp - res)
                    )

                    qtd_transf = 0.0
                    origem_transf = ""

                    # ------------------------------
                    # TRANSFERÊNCIA ENTRE FILIAIS
                    # ------------------------------
                    if necessidade > 0:
                        dash_itens_ruptura += 1

                        for (
                            outra_filial,
                            cod_item
                        ), dados_est in tracker_estoque.items():

                            if (
                                cod_item == c
                                and outra_filial != f_nome
                                and dados_est["EXCEDENTE"] > 0
                            ):
                                qtd_atendida = min(
                                    necessidade,
                                    dados_est["EXCEDENTE"]
                                )

                                qtd_transf += qtd_atendida
                                dados_est["EXCEDENTE"] -= (
                                    qtd_atendida
                                )
                                necessidade -= qtd_atendida

                                if origem_transf:
                                    origem_transf += (
                                        f", {outra_filial}"
                                    )
                                else:
                                    origem_transf = (
                                        outra_filial
                                    )

                                dash_qtd_transferida += (
                                    qtd_atendida
                                )

                                if necessidade <= 0:
                                    necessidade = 0
                                    break

                    # ------------------------------
                    # MÚLTIPLO DE FORNECEDOR
                    # ------------------------------
                    mult = regras_dict.get(
                        fornecedor,
                        multiplo_padrao
                    )

                    try:
                        mult = float(mult)
                    except (TypeError, ValueError):
                        mult = multiplo_padrao

                    if mult <= 0:
                        mult = multiplo_padrao

                    qtd_comprar = (
                        math.ceil(
                            necessidade / mult
                        ) * mult
                        if necessidade > 0
                        else 0
                    )

                    dash_qtd_comprar += qtd_comprar

                    # ------------------------------
                    # STATUS VISUAL
                    # ------------------------------
                    if qtd_comprar > 0:
                        status = "🔴 RUPTURA / COMPRAR"

                    elif qtd_transf > 0:
                        status = (
                            f"🔵 TRANSFERIR DE "
                            f"{origem_transf}"
                        )

                    elif (
                        med == 0
                        and est > 0
                    ):
                        status = "🟡 EXCESSO / SEM GIRO"

                    else:
                        status = "🟢 OK"

                    # ------------------------------
                    # RESULTADO DO ITEM
                    # ------------------------------
                    item_resultado = {
                        "FILIAL": f_nome,
                        "CODIGO": c,
                        "DESCRICAO": row["DESCRICAO"],
                        "FORNECEDOR": fornecedor,
                        "MEDIA": med,
                        "ESTOQUE": est,
                        "COMPRADA": comp,
                        "SUG_COMPRA": qtd_comprar,
                        "SUG_TRANSF": qtd_transf,
                        "ORIGEM_TRANSF": origem_transf,
                        "STATUS": status,
                        "PICO": "SIM" if eh_pico else "NÃO",
                        "MES_1": row["MES_1"],
                        "MES_2": row["MES_2"],
                        "MES_3": row["MES_3"],
                        "MES_4": row["MES_4"],
                        "MESES_ESTOQUE": row["MESES_ESTOQUE"],
                    }

                    resultados.append(item_resultado)

                # --------------------------------------
                # DATAFRAME FINAL
                # --------------------------------------
                df_resultados = pd.DataFrame(
                    resultados
                )

                # --------------------------------------
                # DASHBOARD
                # --------------------------------------
                st.markdown("---")
                st.subheader("📊 Resumo da Compra")

                col_a, col_b, col_c, col_d = st.columns(4)

                with col_a:
                    st.metric(
                        "🛒 Quantidade a Comprar",
                        f"{dash_qtd_comprar:,.0f}"
                    )

                with col_b:
                    st.metric(
                        "🔄 Quantidade Transferida",
                        f"{dash_qtd_transferida:,.0f}"
                    )

                with col_c:
                    st.metric(
                        "📈 Itens com Pico",
                        dash_itens_pico
                    )

                with col_d:
                    st.metric(
                        "🔴 Itens em Ruptura",
                        dash_itens_ruptura
                    )

                # --------------------------------------
                # TABELA
                # --------------------------------------
                st.markdown("---")
                st.subheader(
                    "📋 Sugestão de Compras"
                )

                if not df_resultados.empty:
                    # Ordena primeiro os itens que precisam
                    # de compra, depois transferências e OK.
                    ordem_status = {
                        "🔴 RUPTURA / COMPRAR": 0,
                        "🟡 EXCESSO / SEM GIRO": 1,
                        "🟢 OK": 2,
                    }

                    df_resultados["_ORDEM"] = (
                        df_resultados["STATUS"]
                        .map(ordem_status)
                        .fillna(3)
                    )

                    df_resultados = (
                        df_resultados
                        .sort_values(
                            ["_ORDEM", "FILIAL", "CODIGO"]
                        )
                        .drop(columns=["_ORDEM"])
                        .reset_index(drop=True)
                    )

                    st.dataframe(
                        df_resultados,
                        use_container_width=True,
                        hide_index=True
                    )

                    # ----------------------------------
                    # EXPORTAÇÃO PARA EXCEL
                    # ----------------------------------
                    output = criar_excel(
                        df_resultados
                    )

                    st.download_button(
                        label=(
                            "📥 Baixar Excel da "
                            "Sugestão de Compras"
                        ),
                        data=output.getvalue(),
                        file_name=nome_final_xlsx,
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.spreadsheetml.sheet"
                        )
                    )

                    # ----------------------------------
                    # GRÁFICO RESUMIDO
                    # ----------------------------------
                    df_grafico = pd.DataFrame({
                        "Categoria": [
                            "Comprar",
                            "Transferir",
                            "Picos",
                        ],
                        "Quantidade": [
                            dash_qtd_comprar,
                            dash_qtd_transferida,
                            dash_itens_pico,
                        ],
                    })

                    fig = px.bar(
                        df_grafico,
                        x="Categoria",
                        y="Quantidade",
                        title="Resumo Operacional"
                    )

                    st.plotly_chart(
                        fig,
                        use_container_width=True
                    )

                else:
                    st.warning(
                        "Nenhum resultado foi gerado."
                    )

        except Exception as e:
            st.error(
                f"❌ Erro durante o processamento: {e}"
            )
            st.code(
                traceback.format_exc()
            )
