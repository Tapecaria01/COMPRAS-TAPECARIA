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
                            # Pega os últimos valores da tabela do final da linha
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

# ====================================================
# === PROCESSAMENTO AUTOMÁTICO ===
# ====================================================
if uploaded_files:
    with st.spinner("🔍 Processando arquivos PDF e calculando regras de estoque..."):
        dfs_por_filial = {}
        todos_dados = []
        meses_globais = []

        for f in uploaded_files:
            df, meses = extrair_dados_pdf_web(f)
            if not df.empty:
                dfs_por_filial[f.name.replace(".pdf", "").replace(".PDF", "").upper()] = df
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
                df_global['ESTOQUE_DISPONIVEL'] = df_global['ESTOQUE']

                vendas_recentes = df_global['MES_1'] + df_global['MES_2'] + df_global['MES_3'] + df_global['MES_4']
                df_global['TOTAL_VENDAS_RECENTES'] = vendas_recentes

                # Rastreador de excedentes para cálculo de transferências
                tracker_estoque = {}
                for _, row in df_global.iterrows():
                    f_nome = row['FILIAL_NOME']
                    c = row['CODIGO']
                    est = float(row['ESTOQUE'])
                    med = float(row['MEDIA_SISTEMA'])
                    excesso = est if med == 0 else max(0.0, est - (med * meta))
                    tracker_estoque[(f_nome, c)] = {'EXCEDENTE': excesso, 'MEDIA': med, 'ESTOQUE_FINAL': est}

                # Lógica de Sugestão de Compras e Transferências
                resultados = []
                dash_qtd_comprar = 0
                dash_qtd_transferida = 0
                dash_itens_pico = 0
                dash_itens_ruptura = 0

                # Mapeamento de múltiplos de fornecedor
                regras_dict = dict(zip(st.session_state.df_regras['FORNECEDOR'].str.upper(), st.session_state.df_regras['MULTIPLO']))
                multiplo_padrao = regras_dict.get('GERAL', 1)

                for _, row in df_global.iterrows():
                    f_nome = row['FILIAL_NOME']
                    c = row['CODIGO']
                    med = float(row['MEDIA_SISTEMA'])
                    est = float(row['ESTOQUE'])
                    res = float(row['RESERVA'])
                    comp = float(row['COMPRADA'])
                    fornecedor = str(row['FORNECEDOR']).upper()

                    # Identificação de Picos
                    max_venda = max(row['MES_1'], row['MES_2'], row['MES_3'], row['MES_4'])
                    eh_pico = max_venda > (med * fator_pico) and med > 0
                    if eh_pico:
                        dash_itens_pico += 1

                    # Necessidade bruta
                    necessidade = max(0.0, (med * meta) - (est + comp - res))

                    qtd_transf = 0.0
                    origem_transf = ""

                    # Tenta buscar transferência de outras filiais com excesso
                    if necessidade > 0:
                        dash_itens_ruptura += 1
                        for (outra_filial, cod_item), dados_est in tracker_estoque.items():
                            if cod_item == c and outra_filial != f_nome and dados_est['EXCEDENTE'] > 0:
                                qtd_atendida = min(necessidade, dados_est['EXCEDENTE'])
                                qtd_transf += qtd_atendida
                                dados_est['EXCEDENTE'] -= qtd_atendida
                                necessidade -= qtd_atendida
                                origem_transf = outra_filial
                                dash_qtd_transferida += qtd_atendida
                                if necessidade == 0:
                                    break

                    # Múltiplo de embalagem/fornecedor
                    mult = regras_dict.get(fornecedor, multiplo_padrao)
                    qtd_comprar = math.ceil(necessidade / mult) * mult if necessidade > 0 else 0
                    dash_qtd_comprar += qtd_comprar

                    # Status visual
                    if qtd_comprar > 0:
                        status = "🔴 RUPTURA / COMPRAR"
                    elif qtd_transf > 0:
                        status = f"🔵 TRANSFERIR DE {origem_transf}"
                    elif med == 0 and est > 0:
                        status = "🟡 EXCESSO / SEM GIRO"
                    else:
                        status = "🟢 OK"

                    item_resultado = {
                        'FILIAL': f_nome,
                        'CODIGO': c,
                        'DESCRICAO': row['DESCRICAO'],
                        'FORNECEDOR': fornecedor,
                        'MEDIA': med,
                        'ESTOQUE': est,
                        'COMPRADA': comp,
                        'SUG_COMPRA': qtd_comprar,
                        'SUG_TRANSF': qtd_transf,
                        'ORIGEM_TRANSF': origem_transf,
                        'STATUS': status,
                        'MES_1': row['MES_1'],
                        'MES_2': row['MES_2'],
                        'MES_3': row['MES_3'],
                        'MES_4': row['MES_4']
                    }
                    resultados.append(item_resultado)

                df_final = pd.DataFrame(resultados)

                # --- PAINEL DE MÉTRICAS (DASHBOARD KPIs) ---
                m1, m2, m3, m4 = st.columns(4)
                txt_comprar = f"{dash_qtd_comprar:,.0f}".replace(",", ".")
                txt_transf = f"{dash_qtd_transferida:,.0f}".replace(",", ".")
                
                m1.metric("Unidades a Comprar", txt_comprar)
                m2.metric("Unidades a Transferir", txt_transf)
                m3.metric("Itens em Ruptura", dash_itens_ruptura)
                m4.metric("Alertas de Pico", dash_itens_pico)

                st.markdown("---")

                # --- ABAS DE INTERFACE ---
                tab1, tab2, tab3, tab4 = st.tabs(["📊 Visão Geral", "🚨 Top Urgentes", "📦 Estoque Parado", "🔍 Prévia por Filial"])

                with tab1:
                    st.subheader("Visão Geral do Pedido de Compras")
                    
                    # Gráfico Plotly: Distribuição por Status
                    df_status_count = df_final['STATUS'].apply(lambda x: x.split('/')[0].strip()).value_counts().reset_index()
                    df_status_count.columns = ['Status', 'Quantidade']
                    
                    fig_status = px.bar(
                        df_status_count, 
                        x='Status', 
                        y='Quantidade', 
                        color='Status',
                        title="Distribuição de Itens por Categoria de Status",
                        text_auto=True
                    )
                    st.plotly_chart(fig_status, use_container_width=True)

                    df_view = df_final[['FILIAL', 'CODIGO', 'DESCRICAO', 'FORNECEDOR', 'MEDIA', 'ESTOQUE', 'COMPRADA', 'SUG_COMPRA', 'SUG_TRANSF', 'STATUS']]
                    st.dataframe(df_view.style.apply(pintar_tabela, axis=1), use_container_width=True)

                with tab2:
                    st.subheader("🚨 Itens em Ruptura Crítica (Necessidade Imediata de Compra)")
                    df_ruptura = df_final[df_final['SUG_COMPRA'] > 0].sort_values(by='SUG_COMPRA', ascending=False)
                    if not df_ruptura.empty:
                        st.dataframe(df_ruptura[['FILIAL', 'CODIGO', 'DESCRICAO', 'FORNECEDOR', 'MEDIA', 'ESTOQUE', 'SUG_COMPRA']], use_container_width=True)
                    else:
                        st.success("Nenhum item em ruptura crítica no momento!")

                with tab3:
                    st.subheader("📦 Itens com Estoque Parado / Sem Giro")
                    df_parado = df_final[(df_final['MEDIA'] == 0) & (df_final['ESTOQUE'] > 0)].sort_values(by='ESTOQUE', ascending=False)
                    if not df_parado.empty:
                        st.dataframe(df_parado[['FILIAL', 'CODIGO', 'DESCRICAO', 'FORNECEDOR', 'ESTOQUE']], use_container_width=True)
                    else:
                        st.success("Nenhum produto com estoque parado e sem vendas encontrado!")

                with tab4:
                    st.subheader("🔍 Filtrar Dados por Unidade / Filial")
                    filiais_disponiveis = df_final['FILIAL'].unique().tolist()
                    filial_sel = st.selectbox("Selecione a Filial:", filiais_disponiveis)
                    df_filial = df_final[df_final['FILIAL'] == filial_sel]
                    
