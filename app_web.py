import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
import traceback
import plotly.express as px
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Portal Compras - Tapeçaria", layout="wide")

# Chave do uploader para permitir limpar a análise e importar novos PDFs
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

SENHA = "Tape2026"
REGRAS_PADRAO = [
    ["CORTTEX", 50, 20], ["TEX COMPANY", 50, 20], ["CIPATEX", 50, 20],
    ["KARSTEN", 50, 20], ["ETRURIA", 50, 20], ["TELLAIO", 50, 20],
    ["OBER", 50, 20], ["TEXTIL J. SERRANO", 50, 20], ["CKS", 50, 20],
    ["AGRO QUIMICA", 45, 20], ["ROMPLAS", 30, 15,], ["ROMA DUBLADOS", 10, 5],
]

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if not st.session_state.autenticado:
    st.title("🔐 Portal Compras - Tapeçaria")
    try:
        st.image("logo.png", use_container_width=True)
    except Exception:
        pass
    senha = st.text_input("Digite a senha", type="password")
    if st.button("Entrar", type="primary"):
        if senha == SENHA:
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    st.stop()


def limpar_v(v):
    if v is None:
        return 0.0
    s = str(v).strip().replace(" ", "")
    if not s or s.upper() in {"-", "S/N", "N/A", "NA"}:
        return 0.0
    s = re.sub(r"[^\d,.-]", "", s)
    if not s or s in {"-", ".", ","}:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    elif re.fullmatch(r"-?\d+\.\d{3}", s):
        s = s.replace(".", "")
    try:
        return float(s)
    except Exception:
        return 0.0


def meses_do_texto(texto):
    # Aceita português e inglês, exatamente como os PDFs reais.
    pad = re.compile(r"\b(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ|AUG|SEP|OCT|DEC)[/-](20\d{2})\b", re.I)
    encontrados = []
    for m in pad.finditer(texto):
        valor = f"{m.group(1).upper()}/{m.group(2)}"
        if valor not in encontrados:
            encontrados.append(valor)
    return encontrados


def fornecedor_da_linha(texto, atual=""):
    m = re.search(r"FORNECEDOR\s*:\s*(.+)", texto, re.I)
    if m:
        return m.group(1).strip().split(" - ")[0].strip()
    m = re.search(r"SEGMENTO\s*:\s*(.+)", texto, re.I)
    if m:
        return m.group(1).strip()
    return atual


def extrair_produto(linha):
    """Extrai produto no formato real: CODIGO DESCRIÇÃO UN/1 ou MT/1 + 10 valores."""
    s = re.sub(r"\s+", " ", linha.strip())
    # Mantém asteriscos apenas como marca do relatório e remove-os do código.
    s = re.sub(r"^\*+\s*", "", s)
    m = re.match(r"^(\d{3,7})\s+(.+)$", s)
    if not m:
        return None

    codigo, resto = m.group(1), m.group(2).strip()
    # A embalagem é realmente informada no PDF como UN/1 ou MT/1.
    embs = list(re.finditer(r"\b(UN|MT)\s*/\s*\d+(?:[.,]\d+)?\b", resto, re.I))
    if not embs:
        return None
    em = embs[-1]
    descricao = resto[:em.start()].strip()
    emb = em.group(1).upper()
    tokens = resto[em.end():].strip().split()
    if len(tokens) < 10:
        return None

    tokens = tokens[:10]
    vals = [limpar_v(x) for x in tokens]
    return {
        "CODIGO": codigo,
        "DESCRICAO": descricao,
        "EMB.": emb,
        "MES_1": vals[0], "MES_2": vals[1], "MES_3": vals[2], "MES_4": vals[3],
        "MEDIA_SISTEMA": vals[4], "ESTOQUE": vals[5], "RESERVA": vals[6],
        "MESES_ESTOQUE": vals[7], "COMPRADA": vals[8], "SITUACAO": vals[9],
    }


def extrair_dados_pdf_web(file):
    dados = []
    meses_cabecalho = []
    nome_filial = file.name.rsplit(".", 1)[0].upper()
    fornecedor_atual = "DESCONHECIDO"

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not text:
                continue

            if not meses_cabecalho:
                meses_cabecalho = meses_do_texto(text)[:4]

            fornecedor_atual = fornecedor_da_linha(text, fornecedor_atual)
            mf = re.search(
                r"MEDIA\s+DE\s+VENDAS\s+FILIAL\s*-\s*(.+?)(?:\s+\d{2}/\d{2}/\d{4}|\s*$)",
                text, re.I
            )
            if mf:
                nome_filial = mf.group(1).strip()

            for linha in text.splitlines():
                produto = extrair_produto(linha)
                if produto:
                    produto["FILIAL_NOME"] = nome_filial
                    produto["FORNECEDOR"] = fornecedor_atual
                    dados.append(produto)

    if not meses_cabecalho:
        meses_cabecalho = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]
    while len(meses_cabecalho) < 4:
        meses_cabecalho.append(f"MÊS {len(meses_cabecalho)+1}")

    df = pd.DataFrame(dados)
    if not df.empty:
        df = df.drop_duplicates(["FILIAL_NOME", "CODIGO", "DESCRICAO"], keep="first").reset_index(drop=True)
    return df, meses_cabecalho[:4]


def normalizar_regras(df):
    regras = []
    for _, row in df.iterrows():
        try:
            nome = str(row.get("FORNECEDOR", "")).strip()
            if not nome:
                continue
            mult = max(1, int(float(row.get("MÚLTIPLO", row.get("MULTIPLO", 50)))))
            tol = float(row.get("TOLERÂNCIA", 20))
            regras.append((nome.upper(), mult, tol))
        except Exception:
            continue
    return regras


def regra_fornecedor(fornecedor, regras):
    fornecedor = str(fornecedor or "").upper()
    for nome, mult, tol in regras:
        if nome and nome in fornecedor:
            return mult, tol
    return 1, 0


def coluna_numerica(df, coluna, padrao=0):
    if coluna not in df.columns:
        return pd.Series(padrao, index=df.index, dtype=float)
    return pd.to_numeric(df[coluna], errors="coerce").fillna(padrao)


def preparar_export_filial(df, meses):
    """Monta uma aba no mesmo padrão do arquivo-modelo antigo."""
    x = df.copy()
    meses = list(meses[:4])
    while len(meses) < 4:
        meses.append(f"MÊS {len(meses)+1}")

    mapa = {
        "MES_1": meses[0], "MES_2": meses[1], "MES_3": meses[2], "MES_4": meses[3],
        "MEDIA_SISTEMA": "MEDIA", "MESES_ESTOQUE": "MESES",
    }
    for origem, destino in mapa.items():
        if origem in x.columns:
            x[destino] = x[origem]

    # PADRÃO DO EXCEL ANTIGO: uma aba por filial e sem FILIAL_NOME.
    # Mantemos as colunas do modelo. A coluna ORIGINAL_TRANS é substituída
    # pela nova SUGESTAO TRANSFERENCIA, na mesma posição, imediatamente após
    # SUGESTAO COMPRA.
    ordem = [
        "CODIGO", "DESCRICAO", "EMB.",
        meses[0], meses[1], meses[2], meses[3],
        "MEDIA", "ESTOQUE", "RESERVA", "COMPRADA", "MESES",
        "SUGESTAO COMPRA", "SUGESTAO TRANSFERENCIA",
        "VENDA_ATIPICA", "ESTOQUE PARADO", "RUPTURA CRÍTICA",
        "ORIGINAL_SUGESTAO", "TRANS_MORTA", "TEM_ASTERISCO"
    ]

    # Campos do modelo antigo.
    if "ORIGINAL_SUGESTAO" not in x.columns:
        x["ORIGINAL_SUGESTAO"] = x.get("SUGESTAO COMPRA", 0)
    if "TRANS_MORTA" not in x.columns:
        x["TRANS_MORTA"] = (coluna_numerica(x, "SUGESTAO TRANSFERENCIA") > 0).map({True: "SIM", False: "NÃO"})
    if "TEM_ASTERISCO" not in x.columns:
        x["TEM_ASTERISCO"] = False

    for c in ordem:
        if c not in x.columns:
            x[c] = 0 if c not in {"DESCRICAO", "EMB."} else ""

    x = x[ordem]

    texto = {"DESCRICAO", "EMB."}
    for c in x.columns:
        if c not in texto:
            # Preserve booleans/textos de controle do modelo.
            if c in {"TEM_ASTERISCO", "TRANS_MORTA"} and x[c].dtype == object:
                continue
            x[c] = pd.to_numeric(x[c], errors="coerce").fillna(0)

    return x


def _escrever_aba_modelo(ws, df):
    """Escreve o Excel visualmente no padrão do modelo antigo."""
    # Cores/estilo do modelo ESPUMAUTO.xlsx
    amarelo = PatternFill("solid", fgColor="FFE599")
    fonte = Font(bold=True)
    alinhamento_cab = Alignment(horizontal="center", vertical="center")

    for col, nome in enumerate(df.columns, 1):
        cell = ws.cell(1, col, nome)
        cell.fill = amarelo
        cell.font = fonte
        cell.alignment = alinhamento_cab

    for row in df.itertuples(index=False, name=None):
        valores = []
        for value in row:
            if pd.isna(value):
                valores.append(None)
            elif isinstance(value, float) and value.is_integer():
                valores.append(int(value))
            else:
                valores.append(value)
        ws.append(valores)

    # O modelo antigo não usa filtros e congela a primeira linha.
    ws.auto_filter.ref = None
    ws.freeze_panes = "A2"

    # Larguras iguais ao padrão do arquivo antigo.
    larguras = {
        "A": 8.0, "B": 42.0, "C": 6.0,
        "D": 10.0, "E": 11.86, "F": 11.31, "G": 11.73,
        "H": 7.0, "I": 10.76, "J": 13.0, "K": 12.98,
        "L": 7.0, "M": 20.91, "N": 17.43, "O": 19.52,
        "P": 19.66, "Q": 19.0, "R": 13.0, "S": 13.0, "T": 15.0,
    }
    for letra, largura in larguras.items():
        ws.column_dimensions[letra].width = largura

    ws.row_dimensions[1].height = 15
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 15

    # Alinhamentos do modelo.
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 2).alignment = Alignment(vertical="bottom")
        for c in range(3, ws.max_column + 1):
            ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center")


def gerar_excel(dfs_por_filial, meses):
    """Exporta exatamente no padrão antigo: uma aba para cada filial."""
    wb = Workbook()
    wb.remove(wb.active)

    usadas = set()
    for nome_filial, df in dfs_por_filial.items():
        if df is None or df.empty:
            continue
        base = re.sub(r"[\\/*?:\[\]]", "-", str(nome_filial)).strip() or "FILIAL"
        base = base[:31]
        nome_aba = base
        n = 2
        while nome_aba in usadas:
            sufixo = f"_{n}"
            nome_aba = base[:31-len(sufixo)] + sufixo
            n += 1
        usadas.add(nome_aba)

        ws = wb.create_sheet(title=nome_aba)
        export_df = preparar_export_filial(df, meses)
        _escrever_aba_modelo(ws, export_df)

    if not wb.sheetnames:
        wb.create_sheet("RELATORIO")

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# =============================
# INTERFACE ORIGINAL
# =============================
if "df_regras" not in st.session_state:
    st.session_state.df_regras = pd.DataFrame(REGRAS_PADRAO, columns=["FORNECEDOR", "MÚLTIPLO", "TOLERÂNCIA"])

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

    # Mantém o padrão do dashboard e apenas adiciona a ação solicitada:
    # limpar a análise atual para permitir importar novos fornecedores/relatórios.
    if st.button("🗑️ Limpar análise", use_container_width=True):
        for chave in [
            "dfs_por_filial", "meses_globais", "analise_concluida",
        ]:
            st.session_state.pop(chave, None)
        st.session_state.uploader_key += 1
        st.rerun()

    st.markdown("---")

    with st.expander("⚙️ Configurações Avançadas"):
        meta = st.number_input("Meta de estoque (meses)", min_value=1.0, value=2.0, step=0.5)
        meses_parado = st.number_input("Considerar estoque parado após (meses)", min_value=1.0, value=3.0, step=1.0)
        fator_pico = st.number_input("Sensibilidade de Pico (x vezes a média)", min_value=1.5, value=2.5, step=0.5)
        nome_sugerido = st.text_input("Nome do ficheiro Excel", value="Relatorio_Compras_Tapecaria")
        nome_final_xlsx = nome_sugerido if nome_sugerido.lower().endswith(".xlsx") else f"{nome_sugerido}.xlsx"

    with st.expander("🏭 Fornecedores e Múltiplos"):
        st.caption("Edite ou adicione regras na tabela abaixo.")
        st.session_state.df_regras = st.data_editor(
            st.session_state.df_regras, num_rows="dynamic", use_container_width=True, hide_index=True
        )

col1, col2 = st.columns([1, 15])
with col1:
    try:
        st.image("simbolo.png", width=50)
    except Exception:
        pass
with col2:
    st.title("Inteligência de Compras")
st.markdown("##### Portal Operacional - Tapeçaria")

# =============================
# PROCESSAMENTO AUTOMÁTICO — SEM BOTÃO NOVO
# =============================
if uploaded_files:
    try:
        dfs_por_filial = {}
        todos_dados = []
        meses_globais = []

        with st.spinner("🔍 Processando arquivos PDF e calculando regras de estoque..."):
            for arquivo in uploaded_files:
                df, meses = extrair_dados_pdf_web(arquivo)
                if not df.empty:
                    chave = arquivo.name.rsplit(".", 1)[0].upper()
                    dfs_por_filial[chave] = df
                    todos_dados.append(df)
                    if not meses_globais:
                        meses_globais = meses[:4]

            if not todos_dados:
                st.error("⚠️ O sistema não encontrou produtos compatíveis nos PDFs.")
                st.stop()

            if len(meses_globais) < 4:
                meses_globais = (meses_globais + ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"])[:4]

            df_global = pd.concat(todos_dados, ignore_index=True)
            for col in ["MES_1", "MES_2", "MES_3", "MES_4", "MEDIA_SISTEMA", "ESTOQUE", "RESERVA", "MESES_ESTOQUE", "COMPRADA", "SITUACAO"]:
                df_global[col] = pd.to_numeric(df_global[col], errors="coerce").fillna(0)
            df_global["EMB."] = df_global["EMB."].astype(str).str.upper().where(
                df_global["EMB."].astype(str).str.upper().isin(["MT", "UN"]), "UN"
            )

            regras = normalizar_regras(st.session_state.df_regras)
            tracker = {}
            for _, row in df_global.iterrows():
                chave = (str(row["FILIAL_NOME"]).strip().upper(), str(row["CODIGO"]).strip(), str(row["EMB."]).upper())
                estoque = float(row["ESTOQUE"])
                media = float(row["MEDIA_SISTEMA"])
                reserva = float(row["RESERVA"])
                excesso = estoque if media == 0 else max(0.0, estoque - reserva - media * 3)
                tracker[chave] = excesso

            # necessidade por filial/código/unidade
            necessidades = {}
            for _, row in df_global.iterrows():
                chave = (str(row["FILIAL_NOME"]).strip().upper(), str(row["CODIGO"]).strip(), str(row["EMB."]).upper())
                disponivel = float(row["ESTOQUE"]) + float(row["COMPRADA"]) - float(row["RESERVA"])
                necessidades[chave] = max(0.0, float(row["MEDIA_SISTEMA"]) * meta - disponivel)

            saida = {}
            for filial, df in df_global.groupby("FILIAL_NOME", sort=False):
                df = df.copy()
                df["MEDIA"] = df["MEDIA_SISTEMA"]
                df["SUGESTAO COMPRA"] = 0.0
                df["SUGESTAO TRANSFERENCIA"] = 0.0
                df["ORIGEM TRANSFERENCIA"] = ""
                df["DESTINO TRANSFERENCIA"] = ""
                df["RUPTURA CRÍTICA"] = 0
                df["VENDA_ATIPICA"] = 0
                df["ESTOQUE PARADO"] = 0.0
                df["PARADO_90D"] = 0

                for idx, row in df.iterrows():
                    vendas = [float(row["MES_1"]), float(row["MES_2"]), float(row["MES_3"]), float(row["MES_4"])]
                    media = float(row["MEDIA_SISTEMA"])
                    estoque = float(row["ESTOQUE"])
                    reserva = float(row["RESERVA"])
                    comprada = float(row["COMPRADA"])
                    codigo = str(row["CODIGO"]).strip()
                    emb = str(row["EMB."]).upper()
                    chave_dest = (str(filial).strip().upper(), codigo, emb)

                    df.at[idx, "RUPTURA CRÍTICA"] = int(estoque <= 0 and comprada <= 0 and media > 0)
                    df.at[idx, "VENDA_ATIPICA"] = int(media > 0 and max(vendas) > media * fator_pico)
                    df.at[idx, "ESTOQUE PARADO"] = estoque if estoque > 0 and sum(vendas) == 0 else 0
                    # +90 dias: sem venda nos 3 meses anteriores e com estoque.
                    df.at[idx, "PARADO_90D"] = int(estoque > 0 and sum(vendas[:3]) == 0)

                    necessidade = necessidades.get(chave_dest, 0.0)
                    transferencia = 0.0
                    origens = []

                    # Compra continua com a mesma lógica, mas a transferência usa também a unidade.
                    if necessidade > 0:
                        for chave_org in list(tracker.keys()):
                            origem_filial, origem_codigo, origem_emb = chave_org
                            if origem_codigo != codigo or origem_emb != emb or origem_filial == str(filial).strip().upper():
                                continue
                            disponivel_transferir = tracker[chave_org]
                            if disponivel_transferir <= 0:
                                continue
                            qtd = min(necessidade, disponivel_transferir)
                            qtd = math.floor(qtd)
                            if qtd <= 0:
                                continue
                            transferencia += qtd
                            necessidade -= qtd
                            tracker[chave_org] -= qtd
                            origens.append(f"{qtd:g} de {origem_filial}")
                            if necessidade <= 0:
                                break

                    df.at[idx, "SUGESTAO TRANSFERENCIA"] = transferencia
                    df.at[idx, "ORIGEM TRANSFERENCIA"] = " + ".join(origens)
                    df.at[idx, "DESTINO TRANSFERENCIA"] = str(filial) if transferencia > 0 else ""

                    restante = max(0.0, necessidade)
                    mult, tol = regra_fornecedor(row["FORNECEDOR"], regras)
                    if restante > 0:
                        compra = math.ceil(restante / mult) * mult
                        if compra < tol:
                            compra = mult
                        df.at[idx, "SUGESTAO COMPRA"] = compra

                saida[filial] = df

            st.session_state.dfs_por_filial = saida
            st.session_state.meses_globais = meses_globais
            st.session_state.analise_concluida = True

    except Exception:
        st.error("⚠️ Erro durante o processamento dos relatórios.")
        st.code(traceback.format_exc())

# =============================
# DASHBOARD — MESMO PADRÃO
# =============================
if st.session_state.get("analise_concluida", False):
    dfs = st.session_state.get("dfs_por_filial", {})
    meses = st.session_state.get("meses_globais", ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"])

    if not dfs:
        st.info("Importe os PDFs para iniciar a análise.")
        st.stop()

    # Compatibilidade com dados antigos em sessão.
    for filial, df in dfs.items():
        defaults = {
            "SUGESTAO COMPRA": 0.0, "SUGESTAO TRANSFERENCIA": 0.0,
            "ORIGEM TRANSFERENCIA": "", "DESTINO TRANSFERENCIA": "",
            "RUPTURA CRÍTICA": 0, "VENDA_ATIPICA": 0, "ESTOQUE PARADO": 0.0,
            "PARADO_90D": 0, "MEDIA": df["MEDIA_SISTEMA"] if "MEDIA_SISTEMA" in df.columns else 0.0,
        }
        for col, valor in defaults.items():
            if col not in df.columns:
                df[col] = valor
        dfs[filial] = df

    all_df = pd.concat(dfs.values(), ignore_index=True)
    compra = coluna_numerica(all_df, "SUGESTAO COMPRA").sum()
    transf = coluna_numerica(all_df, "SUGESTAO TRANSFERENCIA").sum()
    ruptura = int(coluna_numerica(all_df, "RUPTURA CRÍTICA").sum())
    parado = coluna_numerica(all_df, "ESTOQUE PARADO").sum()
    p90 = int(coluna_numerica(all_df, "PARADO_90D").sum())

    st.caption("Período identificado: **" + " | ".join(meses) + "**")

    a, b, c = st.columns(3)
    a.metric("Sugestão compra", f"{compra:,.0f}".replace(",", "."))
    b.metric("Transferências", f"{transf:,.0f}".replace(",", "."))
    c.metric("Rupturas", ruptura)

    # Botão de exportação, sem filtros no arquivo.
    try:
        arquivo_excel = gerar_excel(dfs, list(meses))
        st.download_button(
            "📥 EXPORTAR EXCEL",
            data=arquivo_excel,
            file_name=nome_final_xlsx,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception:
        st.error("Erro ao gerar Excel.")
        st.code(traceback.format_exc())

    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 Visão Geral", "🚨 Top Urgentes", "📦 Estoque Parado", "🏢 Prévia por Filial"
    ])

    with tab1:
        resumo = []
        for filial, df in dfs.items():
            resumo.append({
                "FILIAL": filial,
                "SUGESTÃO COMPRA": coluna_numerica(df, "SUGESTAO COMPRA").sum(),
                "TRANSFERÊNCIAS": coluna_numerica(df, "SUGESTAO TRANSFERENCIA").sum(),
            })
        st.dataframe(pd.DataFrame(resumo), use_container_width=True, hide_index=True)

    with tab2:
        urgentes = all_df[
            (coluna_numerica(all_df, "RUPTURA CRÍTICA") > 0) |
            (coluna_numerica(all_df, "SUGESTAO COMPRA") > 0)
        ].copy()
        cols = [
            "FILIAL_NOME", "CODIGO", "DESCRICAO", "EMB.", "MEDIA_SISTEMA",
            "ESTOQUE", "RESERVA", "COMPRADA", "SUGESTAO COMPRA",
            "SUGESTAO TRANSFERENCIA", "RUPTURA CRÍTICA"
        ]
        if not urgentes.empty:
            st.dataframe(
                urgentes[cols].sort_values(["RUPTURA CRÍTICA", "SUGESTAO COMPRA"], ascending=[False, False]),
                use_container_width=True, hide_index=True
            )
        else:
            st.info("Nenhum produto urgente.")

    with tab3:
        parados = all_df[coluna_numerica(all_df, "ESTOQUE PARADO") > 0]
        if not parados.empty:
            resumo_parado = (
                parados.groupby("FILIAL_NOME", as_index=False)["ESTOQUE PARADO"]
                .sum().sort_values("ESTOQUE PARADO", ascending=False)
            )
            st.plotly_chart(
                px.bar(resumo_parado, x="FILIAL_NOME", y="ESTOQUE PARADO", title="Estoque parado por filial"),
                use_container_width=True,
            )
        else:
            st.info("Nenhum estoque parado.")

        p90df = all_df[coluna_numerica(all_df, "PARADO_90D") > 0]
        st.subheader("🔄 Produtos com +90 dias sem venda")
        cols = [
            "FILIAL_NOME", "CODIGO", "DESCRICAO", "EMB.", "ESTOQUE", "MEDIA_SISTEMA",
            "SUGESTAO COMPRA", "SUGESTAO TRANSFERENCIA", "ORIGEM TRANSFERENCIA", "DESTINO TRANSFERENCIA"
        ]
        if not p90df.empty:
            st.dataframe(p90df[cols], use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum produto com +90 dias sem venda.")

    with tab4:
        filial_selecionada = st.selectbox("Selecione a filial", list(dfs.keys()))
        df = dfs[filial_selecionada]
        cols = [
            "CODIGO", "DESCRICAO", "EMB.", "MEDIA_SISTEMA", "ESTOQUE", "RESERVA", "COMPRADA",
            "SUGESTAO COMPRA", "SUGESTAO TRANSFERENCIA", "ORIGEM TRANSFERENCIA",
            "DESTINO TRANSFERENCIA", "RUPTURA CRÍTICA", "VENDA_ATIPICA", "ESTOQUE PARADO",
            "PARADO_90D", "SITUACAO", "MESES_ESTOQUE", "FORNECEDOR"
        ]
        st.dataframe(df[cols], use_container_width=True, hide_index=True)
