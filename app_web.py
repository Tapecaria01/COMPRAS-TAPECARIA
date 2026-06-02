import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
import plotly.express as px
from io import BytesIO
from openpyxl.styles import PatternFill

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Compras Inteligente - Tapeçaria", layout="wide")

# ==========================================
# --- TELA DE SENHA (BLOQUEIO DE ACESSO) ---
# ==========================================
SENHA_ACESSO = "Tape2026"

if "liberado" not in st.session_state:
    st.session_state.liberado = False

if not st.session_state.liberado:
    st.title("🔒 Acesso Restrito - Tapeçaria")
    st.info("Por favor, insira a senha de liberação para aceder ao portal de análise.")
    senha = st.text_input("Senha", type="password")
    if st.button("Entrar"):
        if senha == SENHA_ACESSO:
            st.session_state.liberado = True
            st.rerun()
        else:
            st.error("Senha incorreta. Tente novamente.")
    st.stop()

# ==========================================
# --- FUNÇÕES DE APOIO ---
# ==========================================

def limpar_v(valor):
    if not valor: return 0.0
    s = str(valor).strip().replace('.', '').replace(',', '.')
    try:
        return float(re.sub(r'[^\d.]', '', s))
    except:
        return 0.0

def extrair_dados_pdf_web(pdf_file):
    dados = []
    nome_filial = pdf_file.name.replace(".pdf", "").upper()
    meses_encontrados = []
    fornecedor_atual = "DESCONHECIDO"
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if not texto: continue
                if len(meses_encontrados) < 4:
                    padrao_mes = r'\b(?:jan|fev|feb|mar|abr|apr|mai|may|jun|jul|ago|aug|set|sep|out|oct|nov|dez|dec)/\d{2,4}\b'
                    encontrados = re.findall(padrao_mes, texto.lower())
                    for m in encontrados:
                        m_upper = m.upper()
                        if m_upper not in meses_encontrados: meses_encontrados.append(m_upper)
                for l in texto.split('\n'):
                    if "SEGMENTO" in l.upper():
                        match = re.search(r'SEGMENTO\s*:\s*(.*)', l, re.IGNORECASE)
                        if match: fornecedor_atual = match.group(1).strip()
                    elif re.match(r'^\d{3,6}\s', l):
                        partes = l.split()
                        try:
                            dados.append({
                                'CODIGO': partes[0], 'DESCRICAO': " ".join(partes[1:-11]), 'EMB.': partes[-11],
                                'MES_1': limpar_v(partes[-10]), 'MES_2': limpar_v(partes[-9]), 'MES_3': limpar_v(partes[-8]),
                                'MES_4': limpar_v(partes[-7]), 'MEDIA_SISTEMA': limpar_v(partes[-6]), 'ESTOQUE': limpar_v(partes[-5]),
                                'RESERVA': limpar_v(partes[-4]), 'COMPRADA': limpar_v(partes[-3]), 'MESES_ESTOQUE': limpar_v(partes[-1]), 
                                'FILIAL_NOME': nome_filial, 'FORNECEDOR': fornecedor_atual
                            })
                        except: continue
        return pd.DataFrame(dados), meses_encontrados
    except: return pd.DataFrame(), []

# --- INTERFACE WEB (BARRA LATERAL) ---
with st.sidebar:
    try: st.image("logo.png", use_container_width=True)
    except: st.error("Arquivo 'logo.png' não encontrado no GitHub.")
    st.markdown("---")
    st.header("⚙️ Configurações")
    meta = st.number_input("Meta de estoque (meses)", min_value=1, value=2)
    meses_parado = st.number_input("Considerar estoque parado após (meses)", min_value=1, value=3, step=1)
    fator_pico = st.number_input("Sensibilidade de Pico (x vezes a média)", min_value=1.5, value=2.5, step=0.5)
    nome_sugerido = st.text_input("Nome do ficheiro Excel", value="Relatorio_Compras_Tapecaria")
    nome_final_xlsx = nome_sugerido if nome_sugerido.endswith(".xlsx") else f"{nome_sugerido}.xlsx"
    st.markdown("---")
    uploaded_files = st.file_uploader("Selecione os 4 PDFs das Unidades", type="pdf", accept_multiple_files=True)

# --- CORPO DO SITE ---
col1, col2 = st.columns([1, 15])
with col1:
    try: st.image("simbolo.png", width=50)
    except: pass
with col2:
    st.title("Compras Inteligente")

st.markdown("### Tapeçaria")

if uploaded_files:
    if st.button("🚀 Processar Análise"):
        dfs_por_filial = {}
        todos_dados = []
        meses_globais = []
        dash_qtd_comprar = 0
        dash_qtd_transferida = 0
        dash_itens_pico = 0
        
        for f in uploaded_files:
            df, meses = extrair_dados_pdf_web(f)
            if not df.empty:
                dfs_por_filial[f.name.replace(".pdf", "").upper()] = df
                todos_dados.append(df)
                if len(meses) >= 4 and not meses_globais: meses_globais = meses[:4]
        
        if not meses_globais: meses_globais = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]
        
        if todos_dados:
            df_global = pd.concat(todos_dados).reset_index(drop=True)
            df_global['ESTOQUE_DISPONIVEL'] = df_global['ESTOQUE']
            df_global['TOTAL_VENDAS_RECENTES'] = df_global['MES_1'] + df_global['MES_2'] + df_global['MES_3'] + df_global['MES_4']
            
            # --- CÁLCULO DE EXCEDENTE ---
            def calcular_excedente(row):
                if row['MEDIA_SISTEMA'] == 0: return row['ESTOQUE_DISPONIVEL']
                else:
                    excesso = row['ESTOQUE_DISPONIVEL'] - (row['MEDIA_SISTEMA'] * meta)
                    return max(0, excesso) 
            df_global['EXCEDENTE_DISPONIVEL'] = df_global.apply(calcular_excedente, axis=1)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for nome_destino, df_dest in dfs_por_filial.items():
                    
                    # --- REGRA DO ESTOQUE PARADO / EXCESSO ---
                    def classificar_estoque_parado(row):
                        if row['ESTOQUE'] > 0:
                            if row['MEDIA_SISTEMA'] == 0: return "🛑 SIM"
                            elif row['MESES_ESTOQUE'] > meses_parado: return "🛑 SIM"
                        return ""
                        
                    df_dest['ESTOQUE PARADO'] = df_dest.apply(classificar_estoque_parado, axis=1)
                    
                    def processar_atipico(row):
                        meses_v = [row['MES_1'], row['MES_2'], row['MES_3'], row['MES_4']]
                        pico = max(meses_v)
                        outros = meses_v.copy()
                        outros.remove(pico)
                        media_sem = sum(outros) / 3 if sum(outros) > 0 else 0
                        base_comp = min(row['MEDIA_SISTEMA'], media_sem) if row['MEDIA_SISTEMA'] > 0 and media_sem > 0 else (media_sem if media_sem > 0 else row['MEDIA_SISTEMA'])
                        if base_comp > 0 and pico >= (base_comp * fator_pico) and pico >= 30: return "⚠️ SIM", round(media_sem, 2)
                        elif base_comp == 0 and pico >= 30: return "⚠️ SIM", 0.0
                        return "Não", row['MEDIA_SISTEMA']

                    res_at = df_dest.apply(processar_atipico, axis=1)
                    df_dest['VENDA_ATIPICA'] = [x[0] for x in res_at]
                    df_dest['MEDIA_P_CALCULO'] = [x[1] for x in res_at]
                    dash_itens_pico += len(df_dest[df_dest['VENDA_ATIPICA'] == "⚠️ SIM"])
                    
                    def calcular_log(row):
                        cod = row['CODIGO']
                        necessidade = (row['MEDIA_P_CALCULO'] * meta) - (row['ESTOQUE'] + row['COMPRADA'])
                        if necessidade > 0:
                            outras = df_global[(df_global['CODIGO'] == cod) & (df_global['FILIAL_NOME'] != nome_destino) & (df_global['EXCEDENTE_DISPONIVEL'] > 0)]
                            if not outras.empty:
                                outras_ord = outras.sort_values(by=['MEDIA_SISTEMA', 'EXCEDENTE_DISPONIVEL'], ascending=[True, False])
                                trans_item = []; nec_rest = necessidade
                                for idx_g, cedente in outras_ord.iterrows():
                                    if nec_rest <= 0: break
                                    sal_ced = df_global.loc[idx_g, 'EXCEDENTE_DISPONIVEL']
                                    if sal_ced <= 0: continue
                                    qtd_a_tirar = 30 if nec_rest < 30 and sal_ced >= 30 else min(nec_rest, sal_ced)
                                    if qtd_a_tirar >= 30:
                                        df_global.loc[idx_g, 'EXCEDENTE_DISPONIVEL'] -= qtd_a_tirar
                                        df_global.loc[idx_g, 'ESTOQUE_DISPONIVEL'] -= qtd_a_tirar
                                        nec_rest -= qtd_a_tirar
                                        apelido = 'RP' if 'RIBEIR' in str(cedente['FILIAL_NOME']) else ('Lon' if 'LONDRINA' in str(cedente['FILIAL_NOME']) else ('Frc' if 'FRANCA' in str(cedente['FILIAL_NOME']) else cedente['FILIAL_NOME']))
                                        trans_item.append(f"Tirar {int(qtd_a_tirar)} de {apelido}")
                                if trans_item: return " | ".join(trans_item), round(max(0, nec_rest), 2)
                        return "0", round(max(0, necessidade), 2)

                    res_log = df_dest.apply(calcular_log, axis=1)
                    df_dest['TRANS INTERNA'] = [x[0] for x in res_log]
                    sug_base = [x[1] for x in res_log]

                    # --- MÚLTIPLOS POR FORNECEDOR ---
                    def aplicar_mult(row, sug):
                        if sug <= 0: return 0
                        forn = str(row.get('FORNECEDOR', '')).upper(); desc = str(row.get('DESCRICAO', '')).upper()
                        
                        # NOVIDADE: Adicionado "CKS" aos múltiplos de 50
                        if any(x in forn for x in ["CORTTEX", "TEX COMPANY", "CIPATEX", "KARSTEN", "ETRURIA", "TELLAIO", "TELLAIO TEXTIL", "OBER", "TEXTIL J. SERRANO", "CKS"]): 
                            mult = 50; tol = 20
                        elif "AGRO QUIMICA" in forn: 
                            mult = 45; tol = 20
                        elif "ROMPLAS" in forn and ("URUGUAI" in desc or "URUGUAY" in desc): 
                            mult = 30; tol = 15
                        elif "ROMA DUBLADOS" in forn:
                            mult = 10; tol = 5
                        else: 
                            return sug
                            
                        base = (int(sug) // mult) * mult; rest = sug % mult
                        return base + mult if rest >= tol else base

                    df_dest['SUGESTAO COMPRA'] = [aplicar_mult(row, s) for row, s in zip(df_dest.to_dict('records'), sug_base)]
                    dash_qtd_comprar += df_dest['SUGESTAO COMPRA'].sum()
                    def extrair_n(t): return sum([int(n) for n in re.findall(r'\d+', str(t))]) if str(t) != "0" else 0
                    dash_qtd_transferida += df_dest['TRANS INTERNA'].apply(extrair_n).sum()
                    
                    df_dest.rename(columns={'MES_1': meses_globais[0], 'MES_2': meses_globais[1], 'MES_3': meses_globais[2], 'MES_4': meses_globais[3], 'MEDIA_SISTEMA': 'MEDIA', 'MESES_ESTOQUE': 'MESES'}, inplace=True)
                    
                    cols_f = ['CODIGO', 'DESCRICAO', 'EMB.', meses_globais[0], meses_globais[1], meses_globais[2], meses_globais[3], 'MEDIA', 'ESTOQUE', 'RESERVA', 'COMPRADA', 'MESES', 'SUGESTAO COMPRA', 'TRANS INTERNA', 'VENDA_ATIPICA', 'ESTOQUE PARADO']
                    df_dest[cols_f].to_excel(writer, sheet_name=
