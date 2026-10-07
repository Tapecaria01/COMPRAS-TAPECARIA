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

st.set_page_config(page_title="Portal Compras - Tapeçaria", layout="wide")

SENHA = "Tape2026"
REGRAS_PADRAO = [
    ["CORTTEX", 50, 20], ["TEX COMPANY", 50, 20], ["CIPATEX", 50, 20],
    ["KARSTEN", 50, 20], ["ETRURIA", 50, 20], ["TELLAIO", 50, 20],
    ["OBER", 50, 20], ["TEXTIL J. SERRANO", 50, 20], ["CKS", 50, 20],
    ["AGRO QUIMICA", 45, 20], ["ROMPLAS", 30, 15], ["ROMA DUBLADOS", 10, 5],
]

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if not st.session_state.autenticado:
    st.title("🔐 Portal Compras - Tapeçaria")
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
    p = re.compile(r"\b(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ|AUG|SEP|OCT|DEC)[/-](20\d{2})\b", re.I)
    out = []
    for m in p.finditer(texto):
        x = f"{m.group(1).upper()}/{m.group(2)}"
        if x not in out:
            out.append(x)
    return out


def fornecedor_da_linha(texto, atual=""):
    m = re.search(r"FORNECEDOR\s*:\s*(.+)", texto, re.I)
    if m:
        return m.group(1).strip()
    m = re.search(r"SEGMENTO\s*:\s*(.+)", texto, re.I)
    if m:
        return m.group(1).strip()
    return atual


def extrair_produto(linha):
    """Lê o padrão real dos PDFs: CODIGO DESCRIÇÃO UN/1|MT/1 + 10 números."""
    s = re.sub(r"\s+", " ", linha.strip())
    m = re.match(r"^(?:\*+\s*)?(\d{3,6})\s+(.+)$", s)
    if not m:
        return None
    codigo, resto = m.group(1), m.group(2).strip()
    embs = list(re.finditer(r"\b(UN|MT)\s*/\s*\d+(?:[.,]\d+)?\b", resto, re.I))
    if not embs:
        return None
    em = embs[-1]
    descricao = resto[:em.start()].strip()
    emb = em.group(1).upper()
    nums = resto[em.end():].strip().split()
    # MES1 MES2 MES3 MES4 MEDIA ESTOQUE RESERVA MESES COMPRADA SITUACAO
    if len(nums) < 10:
        return None
    nums = nums[:10]
    if not all(re.fullmatch(r"-?(?:\d+(?:[.,]\d+)?|\d{1,3}(?:\.\d{3})+(?:,\d+)?)", x) for x in nums):
        return None
    vals = [limpar_v(x) for x in nums]
    return {
        "CODIGO": codigo, "DESCRICAO": descricao, "EMB.": emb,
        "MES_1": vals[0], "MES_2": vals[1], "MES_3": vals[2], "MES_4": vals[3],
        "MEDIA_SISTEMA": vals[4], "ESTOQUE": vals[5], "RESERVA": vals[6],
        "MESES_ESTOQUE": vals[7], "COMPRADA": vals[8], "SITUACAO": vals[9],
    }


def extrair_pdf(arquivo):
    registros, meses, fornecedor, filial = [], [], "", ""
    with pdfplumber.open(arquivo) as pdf:
        for page in pdf.pages:
            texto = page.extract_text() or ""
            if not texto:
                continue
            if not meses:
                meses = meses_do_texto(texto)[:4]
            fornecedor = fornecedor_da_linha(texto, fornecedor)
            mf = re.search(r"MEDIA\s+DE\s+VENDAS\s+FILIAL\s*-\s*(.+?)(?:\s+\d{2}/\d{2}/\d{4}|\s*$)", texto, re.I)
            if mf:
                filial = mf.group(1).strip()
            for linha in texto.splitlines():
                p = extrair_produto(linha)
                if p:
                    p["FILIAL_NOME"] = filial
                    p["FORNECEDOR"] = fornecedor
                    registros.append(p)
    if not meses:
        meses = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]
    df = pd.DataFrame(registros)
    if not df.empty:
        df = df.drop_duplicates(["FILIAL_NOME", "CODIGO", "DESCRICAO"], keep="first").reset_index(drop=True)
    return df, meses


def regras_df():
    if "df_regras" not in st.session_state:
        st.session_state.df_regras = pd.DataFrame(REGRAS_PADRAO, columns=["FORNECEDOR", "MÚLTIPLO", "TOLERÂNCIA"])
    return st.session_state.df_regras


def regra_fornecedor(fornecedor, regras):
    f = str(fornecedor or "").upper()
    for nome, multiplo, tol in regras:
        if str(nome).upper().strip() and str(nome).upper().strip() in f:
            return max(1, int(float(multiplo))), float(tol)
    return 50, 20


def normalizar_regras(df):
    out = []
    for _, r in df.iterrows():
        try:
            nome = str(r.get("FORNECEDOR", "")).strip()
            if nome:
                out.append([nome, max(1, int(float(r.get("MÚLTIPLO", 50)))), float(r.get("TOLERÂNCIA", 20))])
        except Exception:
            pass
    return out


def preparar_export(df, meses):
    x = df.copy()
    mapa = {"MES_1": meses[0], "MES_2": meses[1], "MES_3": meses[2], "MES_4": meses[3], "MEDIA_SISTEMA": "MEDIA"}
    for a, b in mapa.items():
        if a in x.columns:
            x[b] = x[a]
    ordem = ["FILIAL_NOME", "CODIGO", "DESCRICAO", "EMB.", mapa["MES_1"], mapa["MES_2"], mapa["MES_3"], mapa["MES_4"], "MEDIA", "ESTOQUE", "RESERVA", "COMPRADA", "SUGESTAO COMPRA", "SUGESTAO TRANSFERENCIA", "ORIGEM TRANSFERENCIA", "DESTINO TRANSFERENCIA", "RUPTURA CRÍTICA", "VENDA_ATIPICA", "ESTOQUE PARADO", "PARADO_90D", "SITUACAO", "MESES_ESTOQUE", "FORNECEDOR"]
    for c in ordem:
        if c not in x.columns:
            x[c] = ""
    x = x[ordem]
    for c in x.columns:
        if c not in {"FILIAL_NOME", "CODIGO", "DESCRICAO", "EMB.", "ORIGEM TRANSFERENCIA", "DESTINO TRANSFERENCIA", "FORNECEDOR"}:
            x[c] = pd.to_numeric(x[c], errors="coerce").fillna(0)
    return x


def escrever_planilha(ws, df):
    if df.empty:
        ws.append(["Nenhum registro"])
        return
    for c, nome in enumerate(df.columns, 1):
        ws.cell(1, c, nome)
    for row in df.itertuples(index=False, name=None):
        vals = []
        for v in row:
            vals.append("" if pd.isna(v) else (int(v) if isinstance(v, (int, float)) and float(v).is_integer() else v))
        ws.append(vals)
    fill = PatternFill("solid", fgColor="17365D")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill, cell.font = fill, font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = None  # SEM FILTROS
    heads = {c.value: c.column for c in ws[1]}
    cores = {"SUGESTAO COMPRA": "E2F0D9", "SUGESTAO TRANSFERENCIA": "DDEBF7", "RUPTURA CRÍTICA": "F4CCCC", "PARADO_90D": "FCE4D6"}
    for nome, cor in cores.items():
        if nome in heads:
            fillc = PatternFill("solid", fgColor=cor)
            for r in range(2, ws.max_row + 1): ws.cell(r, heads[nome]).fill = fillc
    for i in range(1, ws.max_column + 1):
        letter = get_column_letter(i)
        mx = max((len(str(ws.cell(r, i).value or "")) for r in range(1, ws.max_row + 1)), default=10)
        ws.column_dimensions[letter].width = min(max(mx + 2, 10), 45)


def gerar_excel(dfs, meses):
    wb = Workbook(); wb.remove(wb.active)
    # Resumo
    ws = wb.create_sheet("RESUMO")
    total_c = total_t = 0
    linhas = []
    for filial, df in dfs.items():
        c = pd.to_numeric(df["SUGESTAO COMPRA"], errors="coerce").fillna(0).sum()
        t = pd.to_numeric(df["SUGESTAO TRANSFERENCIA"], errors="coerce").fillna(0).sum()
        total_c += c; total_t += t
        linhas.append([filial, c, t])
    ws.append(["RELATÓRIO DE COMPRAS - TAPEÇARIA"]); ws.merge_cells("A1:C1")
    ws.append([]); ws.append(["INDICADOR", "VALOR"])
    ws.append(["Sugestão de compra", total_c]); ws.append(["Sugestão de transferência", total_t])
    ws.append([]); ws.append(["FILIAL", "COMPRA", "TRANSFERÊNCIA"])
    for l in linhas: ws.append(l)
    ws["A1"].font = Font(size=16, bold=True, color="FFFFFF"); ws["A1"].fill = PatternFill("solid", fgColor="17365D"); ws["A1"].alignment = Alignment(horizontal="center")
    for row in (3, 7):
        for c in ws[row]: c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 35; ws.column_dimensions["B"].width = 22; ws.column_dimensions["C"].width = 22
    ws.auto_filter.ref = None
    # Compra e transferência
    for nome, campo in [("SUGESTÕES DE COMPRA", "SUGESTAO COMPRA"), ("TRANSFERÊNCIAS", "SUGESTAO TRANSFERENCIA")]:
        ws = wb.create_sheet(nome)
        frames = []
        for filial, df in dfs.items():
            z = df[pd.to_numeric(df[campo], errors="coerce").fillna(0) > 0].copy()
            if not z.empty: frames.append(z)
        z = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        escrever_planilha(ws, preparar_export(z, meses) if not z.empty else z)
    # Filiais
    usados = {"RESUMO", "SUGESTÕES DE COMPRA", "TRANSFERÊNCIAS"}
    for filial, df in dfs.items():
        base = re.sub(r'[\[\]:*?/\\]', '_', str(filial))[:31] or "FILIAL"
        nome = base; n = 2
        while nome in usados:
            suf = f"_{n}"; nome = base[:31-len(suf)] + suf; n += 1
        usados.add(nome)
        ws = wb.create_sheet(nome)
        escrever_planilha(ws, preparar_export(df, meses))
    buf = BytesIO(); wb.save(buf); return buf.getvalue()

# Sidebar
st.sidebar.title("📋 Portal Compras")
if "uploader_key" not in st.session_state: st.session_state.uploader_key = 0
if "analise_concluida" not in st.session_state: st.session_state.analise_concluida = False

arquivos = st.sidebar.file_uploader("Importe os relatórios PDF", type=["pdf"], accept_multiple_files=True, key=f"pdfs_{st.session_state.uploader_key}")
if st.sidebar.button("🗑️ Limpar dados"):
    st.session_state.analise_concluida = False; st.session_state.dfs_por_filial = {}; st.session_state.uploader_key += 1; st.rerun()

meta = st.sidebar.number_input("Meta de estoque (meses)", min_value=0.0, value=2.0, step=0.5)
retencao = st.sidebar.number_input("Retenção de segurança (meses)", min_value=0.0, value=3.0, step=0.5)
fator_pico = st.sidebar.number_input("Fator para venda atípica", min_value=1.0, value=2.5, step=0.5)
dias_parado = st.sidebar.number_input("Dias parado para transferência", min_value=90, value=90, step=30)
nome_excel = st.sidebar.text_input("Nome do Excel", "Relatorio_Compras_Tapecaria")
if not nome_excel.lower().endswith(".xlsx"): nome_excel += ".xlsx"

st.sidebar.subheader("📦 Regras por fornecedor")
st.session_state.df_regras = st.sidebar.data_editor(regras_df(), num_rows="dynamic", use_container_width=True)

if arquivos and st.sidebar.button("🚀 Processar relatórios", type="primary"):
    try:
        todos, periodos = [], []
        with st.spinner("Lendo os PDFs..."):
            for arq in arquivos:
                df, meses = extrair_pdf(arq)
                if df.empty: st.warning(f"Nenhum produto encontrado em {arq.name}")
                else:
                    todos.append(df); periodos.append(meses)
        if not todos: st.error("Nenhum produto foi encontrado."); st.stop()
        meses = periodos[0][:4]
        if len(meses) < 4: meses = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]
        df_all = pd.concat(todos, ignore_index=True)
        for c in ["MES_1","MES_2","MES_3","MES_4","MEDIA_SISTEMA","ESTOQUE","RESERVA","MESES_ESTOQUE","COMPRADA","SITUACAO"]:
            df_all[c] = pd.to_numeric(df_all[c], errors="coerce").fillna(0)
        df_all["EMB."] = df_all["EMB."].astype(str).str.upper().where(df_all["EMB."].astype(str).str.upper().isin(["MT","UN"]), "UN")
        regras = normalizar_regras(st.session_state.df_regras)
        # Necessidade por filial/código/unidade e estoque excedente por filial/código/unidade.
        necessidades = {}
        excesso = {}
        for _, r in df_all.iterrows():
            key = (str(r.FILIAL_NOME).strip().upper(), str(r.CODIGO).strip(), str(r["EMB."]).upper())
            disponivel = r.ESTOQUE + r.COMPRADA - r.RESERVA
            necessidades[key] = max(0, r.MEDIA_SISTEMA * meta - disponivel)
            minimo = r.MEDIA_SISTEMA * retencao
            excesso[key] = max(0, r.ESTOQUE - r.RESERVA - minimo) if r.MEDIA_SISTEMA > 0 else max(0, r.ESTOQUE - r.RESERVA)
        saida = {}
        for filial, df in df_all.groupby("FILIAL_NOME", sort=False):
            df = df.copy()
            for c in ["SUGESTAO COMPRA","SUGESTAO TRANSFERENCIA","ESTOQUE PARADO"]: df[c] = 0.0
            for c in ["ORIGEM TRANSFERENCIA","DESTINO TRANSFERENCIA"]: df[c] = ""
            df["RUPTURA CRÍTICA"] = 0; df["VENDA_ATIPICA"] = 0; df["PARADO_90D"] = 0; df["MEDIA"] = df["MEDIA_SISTEMA"]
            for idx, r in df.iterrows():
                vendas = [r.MES_1, r.MES_2, r.MES_3, r.MES_4]
                media, estoque, reserva, comprada = map(float, [r.MEDIA_SISTEMA, r.ESTOQUE, r.RESERVA, r.COMPRADA])
                df.at[idx,"RUPTURA CRÍTICA"] = int(estoque <= 0 and comprada <= 0 and media > 0)
                df.at[idx,"VENDA_ATIPICA"] = int(media > 0 and max(vendas) > media * fator_pico)
                df.at[idx,"ESTOQUE PARADO"] = estoque if estoque > 0 and sum(vendas) == 0 else 0
                # Com 4 meses no relatório, ausência de venda nos 3 meses anteriores é o indicador de +90 dias.
                df.at[idx,"PARADO_90D"] = int(estoque > 0 and sum(vendas[:3]) == 0)
                key_dest = (str(filial).strip().upper(), str(r.CODIGO).strip(), str(r["EMB."]).upper())
                necessidade = necessidades.get(key_dest, 0)
                transfer = 0; origens = []
                if necessidade > 0:
                    for origem, _ in df_all.groupby("FILIAL_NOME", sort=False):
                        key_org = (str(origem).strip().upper(), str(r.CODIGO).strip(), str(r["EMB."]).upper())
                        if key_org == key_dest: continue
                        q = min(necessidade-transfer, excesso.get(key_org, 0))
                        q = math.floor(q)
                        if q > 0:
                            transfer += q; excesso[key_org] -= q; origens.append(f"{q:g} de {origem}")
                        if transfer >= necessidade: break
                df.at[idx,"SUGESTAO TRANSFERENCIA"] = transfer
                df.at[idx,"ORIGEM TRANSFERENCIA"] = " + ".join(origens)
                df.at[idx,"DESTINO TRANSFERENCIA"] = str(filial) if transfer > 0 else ""
                restante = max(0, necessidade-transfer)
                mult, tol = regra_fornecedor(r.FORNECEDOR, regras)
                if restante > 0:
                    compra = math.ceil(restante / mult) * mult
                    if compra < tol: compra = mult
                    df.at[idx,"SUGESTAO COMPRA"] = compra
            saida[filial] = df
        st.session_state.dfs_por_filial = saida
        st.session_state.meses_globais = meses
        st.session_state.analise_concluida = True
        st.success("Análise concluída: " + " | ".join(meses))
    except Exception:
        st.error("Erro durante o processamento.")
        st.code(traceback.format_exc())

if st.session_state.get("analise_concluida", False):
    dfs = st.session_state.dfs_por_filial; meses = st.session_state.meses_globais
    all_df = pd.concat(dfs.values(), ignore_index=True)
    compra = pd.to_numeric(all_df["SUGESTAO COMPRA"], errors="coerce").fillna(0).sum()
    transf = pd.to_numeric(all_df["SUGESTAO TRANSFERENCIA"], errors="coerce").fillna(0).sum()
    ruptura = int(all_df["RUPTURA CRÍTICA"].sum()); parado = pd.to_numeric(all_df["ESTOQUE PARADO"], errors="coerce").fillna(0).sum(); p90 = int(all_df["PARADO_90D"].sum())
    st.title("📊 Portal Compras - Tapeçaria")
    st.caption("Período identificado: **" + " | ".join(meses) + "**")
    a,b,c,d,e = st.columns(5)
    a.metric("Sugestão compra", f"{compra:,.0f}".replace(",",".")); b.metric("Transferências", f"{transf:,.0f}".replace(",",".")); c.metric("Rupturas", ruptura); d.metric("Estoque parado", f"{parado:,.0f}".replace(",",".")); e.metric("Parados +90d", p90)
    try:
        xlsx = gerar_excel(dfs, meses)
        st.download_button("📥 EXPORTAR EXCEL", xlsx, file_name=nome_excel, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
    except Exception:
        st.error("Erro ao gerar Excel."); st.code(traceback.format_exc())
    t1,t2,t3,t4 = st.tabs(["📋 Visão Geral","🚨 Top Urgentes","📦 Estoque Parado","🏢 Prévia por Filial"])
    with t1:
        resumo=[]
        for filial,df in dfs.items(): resumo.append({"FILIAL":filial,"SUGESTÃO COMPRA":df["SUGESTAO COMPRA"].sum(),"TRANSFERÊNCIAS":df["SUGESTAO TRANSFERENCIA"].sum(),"RUPTURAS":int(df["RUPTURA CRÍTICA"].sum()),"PARADOS +90D":int(df["PARADO_90D"].sum())})
        st.dataframe(pd.DataFrame(resumo), use_container_width=True, hide_index=True)
    with t2:
        u=all_df[(all_df["RUPTURA CRÍTICA"]>0)|(all_df["SUGESTAO COMPRA"]>0)].copy()
        cols=["FILIAL_NOME","CODIGO","DESCRICAO","EMB.","MEDIA_SISTEMA","ESTOQUE","RESERVA","COMPRADA","SUGESTAO COMPRA","SUGESTAO TRANSFERENCIA","RUPTURA CRÍTICA"]
        st.dataframe(u[cols].sort_values(["RUPTURA CRÍTICA","SUGESTAO COMPRA"], ascending=[False,False]), use_container_width=True, hide_index=True) if not u.empty else st.info("Nenhum produto urgente.")
    with t3:
        p=all_df[all_df["ESTOQUE PARADO"]>0]
        if not p.empty:
            r=p.groupby("FILIAL_NOME",as_index=False)["ESTOQUE PARADO"].sum().sort_values("ESTOQUE PARADO",ascending=False)
            st.plotly_chart(px.bar(r,x="FILIAL_NOME",y="ESTOQUE PARADO",title="Estoque parado por filial"), use_container_width=True)
        p90df=all_df[all_df["PARADO_90D"]>0]
        st.subheader("🔄 Produtos com +90 dias sem venda")
        cols=["FILIAL_NOME","CODIGO","DESCRICAO","EMB.","ESTOQUE","MEDIA_SISTEMA","SUGESTAO COMPRA","SUGESTAO TRANSFERENCIA","ORIGEM TRANSFERENCIA","DESTINO TRANSFERENCIA"]
        st.dataframe(p90df[cols],use_container_width=True,hide_index=True) if not p90df.empty else st.info("Nenhum produto com +90 dias sem venda.")
    with t4:
        f=st.selectbox("Selecione a filial", list(dfs.keys()))
        df=dfs[f]; cols=["CODIGO","DESCRICAO","EMB.","MEDIA_SISTEMA","ESTOQUE","RESERVA","COMPRADA","SUGESTAO COMPRA","SUGESTAO TRANSFERENCIA","ORIGEM TRANSFERENCIA","DESTINO TRANSFERENCIA","RUPTURA CRÍTICA","VENDA_ATIPICA","ESTOQUE PARADO","PARADO_90D","SITUACAO","MESES_ESTOQUE","FORNECEDOR"]
        st.dataframe(df[cols],use_container_width=True,hide_index=True)
