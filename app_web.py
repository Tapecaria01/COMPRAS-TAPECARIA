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
st.set_page_config(page_title="Portal de Negociações", layout="wide")

st.markdown("""
    
""", unsafe_allow_html=True)

# --- FUNÇÕES DE APOIO ---
def limpar_v(valor):
    if not valor: return 0.0
    s = str(valor).strip().replace('.', '').replace(',', '.')
    try: return float(re.sub(r'[^\d.]', '', s))
    except: return 0.0

@st.cache_data(show_spinner="A processar catálogos em PDF...")
def extrair_dados_pdf_negociacao(pdf_file):
    dados = []
    fornecedor_atual = "DESCONHECIDO"
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if not texto: continue
                
                for linha in texto.split('\n'):
                    l = linha.strip()
                    # Identificar o fornecedor no cabeçalho
                    if "SEGMENTO" in l.upper() or "FORNECEDOR" in l.upper():
                        match = re.search(r'(?:SEGMENTO|FORNECEDOR)\s*:\s*(.*)', l, re.IGNORECASE)
                        if match: fornecedor_atual = match.group(1).split('-')[0].strip()
                    else:
                        l = re.sub(r'^\*+\s*', '', l)
                        if re.match(r'^\d{3,6}\s', l):
                            partes = l.split()
                            try:
                                dados.append({
                                    'CODIGO': partes[0], 
                                    'DESCRICAO': " ".join(partes[1:-11])[:50], 
                                    'MEDIA_SISTEMA': limpar_v(partes[-6]), 
                                    'ESTOQUE': limpar_v(partes[-5]),
                                    'FORNECEDOR': fornecedor_atual
                                })
                            except: continue
        return pd.DataFrame(dados)
    except Exception: return pd.DataFrame()

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Parâmetros da Negociação")
    fornecedor_filtro = st.text_input("Filtrar Fornecedor", placeholder="Ex: YORK, CORTTEX...")
    
    st.subheader("📦 Volume Acordado")
    volume_total = st.number_input("Volume Total Negociado", min_value=0, value=30000, step=1000)
    
    st.subheader("📅 Cronograma de Entrega")
    col_l1, col_l2 = st.columns(2)
    with col_l1:
        mes_lote1 = st.text_input("Lote 1 (Mês)", value="Setembro")
        qtd_lote1 = st.number_input("Qtd Lote 1", min_value=0, value=18000, step=500)
    with col_l2:
        mes_lote2 = st.text_input("Lote 2 (Mês)", value="Outubro")
        qtd_lote2 = st.number_input("Qtd Lote 2", min_value=0, value=12000, step=500)
        
    multiplo = st.number_input("Múltiplo Padrão de Embalagem", min_value=1, value=50)

    st.markdown("---")
    uploaded_files = st.file_uploader("Carregue os PDFs dos Produtos", type="pdf", accept_multiple_files=True)

# --- CORPO PRINCIPAL ---
st.title("🤝 Portal de Negociações e Compras em Lote")
st.markdown("Ferramenta para distribuição de cota de volume fechado de fornecedores baseado na curva de vendas.")
st.markdown("---")

if uploaded_files:
    todos_dados = []
    for f in uploaded_files:
        df = extrair_dados_pdf_negociacao(f)
        if not df.empty: todos_dados.append(df)
        
    if todos_dados:
        df_global = pd.concat(todos_dados).reset_index(drop=True)
        
        # Agrupar itens únicos somando estoque e média de vendas de todas as filiais
        df_agrupado = df_global.groupby(['CODIGO', 'DESCRICAO', 'FORNECEDOR'], as_index=False).agg({
            'ESTOQUE': 'sum',
            'MEDIA_SISTEMA': 'sum'
        })
        
        # Filtro de fornecedor
        if fornecedor_filtro:
            df_agrupado = df_agrupado[df_agrupado['FORNECEDOR'].str.contains(fornecedor_filtro.upper(), na=False)]
            
        if df_agrupado.empty:
            st.warning("Nenhum produto encontrado para o fornecedor especificado.")
            st.stop()
            
        st.subheader("1️⃣ Selecione os Produtos da Negociação")
        st.caption("Marque os itens que entram no rateio. O volume de " + f"{volume_total:,.0f} será distribuído apenas entre os selecionados.")
        
        df_agrupado.insert(0, "INCLUIR", True)
        
        df_selecao = st.data_editor(
            df_agrupado.sort_values(by='MEDIA_SISTEMA', ascending=False),
            column_config={
                "INCLUIR": st.column_config.CheckboxColumn("Participa?", default=True),
                "CODIGO": "Código",
                "DESCRICAO": "Descrição",
                "ESTOQUE": st.column_config.NumberColumn("Estoque Total", format="%.0f"),
                "MEDIA_SISTEMA": st.column_config.NumberColumn("Média Global", format="%.2f")
            },
            disabled=["CODIGO", "DESCRICAO", "FORNECEDOR", "ESTOQUE", "MEDIA_SISTEMA"],
            hide_index=True,
            use_container_width=True
        )
        
        # --- LÓGICA DE RATEIO ---
        df_curva = df_selecao[df_selecao['INCLUIR'] & (df_selecao['MEDIA_SISTEMA'] > 0)].copy()
        
        if df_curva.empty:
            st.error("Selecione pelo menos um produto com histórico de vendas para realizar o rateio.")
        else:
            soma_vendas = df_curva['MEDIA_SISTEMA'].sum()
            df_curva['PARTICIPACAO_%'] = df_curva['MEDIA_SISTEMA'] / soma_vendas
            
            # Cálculo dos lotes com arredondamento no múltiplo
            df_curva['TOTAL_SUGERIDO'] = (round((df_curva['PARTICIPACAO_%'] * volume_total) / multiplo) * multiplo).astype(int)
            df_curva['LOTE_1'] = (round((df_curva['PARTICIPACAO_%'] * qtd_lote1) / multiplo) * multiplo).astype(int)
            
            # O Lote 2 recebe o saldo restante do total sugerido
            df_curva['LOTE_2'] = df_curva['TOTAL_SUGERIDO'] - df_curva['LOTE_1']
            df_curva['LOTE_2'] = df_curva['LOTE_2'].apply(lambda x: max(0, int(x)))
            
            st.markdown("---")
            st.subheader("📊 Resultado da Distribuição")
            
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("Volume Calculado (Total)", f"{df_curva['TOTAL_SUGERIDO'].sum():,.0f}".replace(',', '.'))
            kpi2.metric(f"Lote 1 ({mes_lote1})", f"{df_curva['LOTE_1'].sum():,.0f}".replace(',', '.'))
            kpi3.metric(f"Lote 2 ({mes_lote2})", f"{df_curva['LOTE_2'].sum():,.0f}".replace(',', '.'))
            kpi4.metric("SKUs Selecionados", len(df_curva))
            
            # Tabela Final
            df_final = df_curva[['CODIGO', 'DESCRICAO', 'ESTOQUE', 'MEDIA_SISTEMA', 'LOTE_1', 'LOTE_2', 'TOTAL_SUGERIDO']].copy()
            df_final.rename(columns={'MEDIA_SISTEMA': 'MÉDIA VENDAS'}, inplace=True)
            st.dataframe(df_final, use_container_width=True, hide_index=True)
            
            # --- EXPORTAÇÃO EXCEL ---
            output = BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.title = "Negociacao_Lotes"
            
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Renomear cabeçalhos para o Excel
            colunas_excel = ['CÓDIGO', 'DESCRIÇÃO', 'ESTOQUE ATUAL', 'MÉDIA VENDAS', f'LOTE 1 ({mes_lote1})', f'LOTE 2 ({mes_lote2})', 'TOTAL SUGERIDO']
            ws.append(colunas_excel)
            
            for col_idx, cell in enumerate(ws[1], 1):
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center')
                ws.column_dimensions[get_column_letter(col_idx)].width = 20
            ws.column_dimensions['B'].width = 45 # Descrição mais larga
                
            for _, row in df_final.iterrows():
                ws.append(row.tolist())
                
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.border = border
                    if isinstance(cell.value, (int, float)): cell.number_format = '#,##0'

            wb.save(output)
            output.seek(0)
            
            st.success("✅ Rateio calculado com sucesso! Ficheiro Excel pronto para envio.")
            st.download_button(
                label="📥 Descarregar Pedido Fechado (Excel)",
                data=output,
                file_name=f"Negociacao_{fornecedor_filtro if fornecedor_filtro else 'Geral'}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
else:
    st.info("Aguardando documentos. Carregue os PDFs na barra lateral e configure os lotes para iniciar o rateio.")
