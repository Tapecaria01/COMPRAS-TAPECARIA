import streamlit as st
import pandas as pd
import pdfplumber
import re
import math
import plotly.express as px
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Portal Compras - Tapeçaria", layout="wide")

# ==========================================
# --- ESTILIZAÇÃO CSS (DESIGN PREMIUM) ---
# ==========================================
st.markdown("""
    <style>
    /* Forçar o fundo escuro nativo para evitar telas brancas */
    .stApp {
        background-color: #0E1117 !important;
    }
    /* Estilizar o botão principal para um Azul Corporativo Premium */
    div.stButton > button:first-child {
        background-color: #004A8F !important;
        color: white !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 8px 24px !important;
        font-weight: 600 !important;
        transition: 0.3s !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #003366 !important;
        color: white !important;
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# --- TELA DE SENHA (BLOQUEIO CENTRALIZADO) ---
# ==========================================
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
            except: 
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

# ==========================================
# --- VARIÁVEIS DE SESSÃO ---
# ==========================================
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

# ==========================================
# --- FUNÇÕES DE APOIO ---
# ==========================================
def limpar_v(valor):
    if not valor: 
        return 0.0
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
                if not texto: 
                    continue
                    
                if len(meses_encontrados) < 4:
                    padrao_mes = r'\b(?:jan|fev|feb|mar|abr|apr|mai|may|jun|jul|ago|aug|set|sep|out|oct|nov|dez|dec)/\d{2,4}\b'
                    encontrados = re.findall(padrao_mes, texto.lower())
                    for m in encontrados:
                        m_upper = m.upper()
                        if m_upper not in meses_encontrados: 
                            meses_encontrados.append(m_upper)
                            
                for linha in texto.split('\n'):
                    l = linha.strip()
                    
                    if "SEGMENTO" in l.upper():
                        match = re.search(r'SEGMENTO\s*:\s*(.*)', l, re.IGNORECASE)
                        if match: 
                            fornecedor_atual = match.group(1).strip()
                    else:
                        # --- EXCEÇÃO EXCLUSIVA YORK ---
                        # Se o fornecedor for YORK e a linha começar com "***", removemos os asteriscos.
                        if "YORK" in fornecedor_atual.upper() and l.startswith("***"):
                            l = re.sub(r'^\*\*\*\s*', '', l)
                            
                        # Agora o robô consegue testar e ler normalmente
                        if re.match(r'^\d{3,6}\s', l):
                            partes = l.split()
                            try:
                                item_dict = {
                                    'CODIGO': partes[0], 
                                    'DESCRICAO': " ".join(partes[1:-11]), 
                                    'EMB.': partes[-11],
                                    'MES_1': limpar_v(partes[-10]), 
                                    'MES_2': limpar_v(partes[-9]), 
                                    'MES_3': limpar_v(partes[-8]),
                                    'MES_4': limpar_v(partes[-7]), 
                                    'MEDIA_SISTEMA': limpar_v(partes[-6]), 
                                    'ESTOQUE': limpar_v(partes[-5]),
                                    'RESERVA': limpar_v(partes[-4]), 
                                    'COMPRADA': limpar_v(partes[-3]), 
                                    'MESES_ESTOQUE': limpar_v(partes[-1]), 
                                    'FILIAL_NOME': nome_filial, 
                                    'FORNECEDOR': fornecedor_atual
                                }
                                dados.append(item_dict)
                            except: 
                                continue
                            
        return pd.DataFrame(dados), meses_encontrados
    except: 
        return pd.DataFrame(), []

# --- INTERFACE WEB ---
with st.sidebar:
    try: 
        st.image("logo.png", use_container_width=True)
    except: 
        pass
        
    st.markdown("---")
    st.header("📂 Nova Compra")
    uploaded_files = st.file_uploader("Selecione os 4 PDFs das Unidades", type="pdf", accept_multiple_files=True)
    
    if not uploaded_files:
        st.session_state.analise_concluida = False
        
    st.markdown("---")
    
    with st.expander("⚙️ Configurações Avançadas"):
        meta = st.number_input("Meta de estoque (meses)", min_value=1, value=2)
        meses_parado = st.number_input("Considerar estoque parado após (meses)", min_value=1, value=3, step=1)
        fator_pico = st.number_input("Sensibilidade de Pico (x vezes a média)", min_value=1.5, value=2.5, step=0.5)
        nome_sugerido = st.text_input("Nome do ficheiro Excel", value="Relatorio_Compras_Tapecaria")
        
        if nome_sugerido.endswith(".xlsx"):
            nome_final_xlsx = nome_sugerido
        else:
            nome_final_xlsx = f"{nome_sugerido}.xlsx"
            
    with st.expander("🏭 Fornecedores e Múltiplos"):
        st.caption("Edite ou adicione regras na última linha vazia.")
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
    except: 
        pass
with col2:
    st.title("Inteligência de Compras")

st.markdown("##### Portal Operacional - Tapeçaria")
st.markdown("<br>", unsafe_allow_html=True)

if uploaded_files:
    if st.button("🚀 Processar Análise Inteligente"):
        
        with st.spinner("A processar dados, calcular múltiplos e gerar relatórios..."):
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
                    if len(meses) >= 4 and not meses_globais: 
                        meses_globais = meses[:4]
            
            if not meses_globais: 
                meses_globais = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]
            
            if todos_dados:
                df_global = pd.concat(todos_dados).reset_index(drop=True)
                df_global['ESTOQUE_DISPONIVEL'] = df_global['ESTOQUE']
                
                vendas_recentes = df_global['MES_1'] + df_global['MES_2'] + df_global['MES_3'] + df_global['MES_4']
                df_global['TOTAL_VENDAS_RECENTES'] = vendas_recentes
                
                # --- GESTOR DE MEMÓRIA DE STOCK (Evita Gastos Duplos) ---
                tracker_estoque = {}
                for _, row in df_global.iterrows():
                    f_nome = row['FILIAL_NOME']
                    c = row['CODIGO']
                    est = float(row['ESTOQUE'])
                    med = float(row['MEDIA_SISTEMA'])
                    
                    excesso = est if med == 0 else max(0.0, est - (med * meta))
                    tracker_estoque[(f_nome, c)] = {
                        'EXCEDENTE': excesso,
                        'MEDIA': med,
                        'ESTOQUE_FINAL': est
                    }

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for nome_destino, df_dest in dfs_por_filial.items():
                        
                        def classificar_estoque_parado(row):
                            if row['ESTOQUE'] > 0:
                                if row['MEDIA_SISTEMA'] == 0: 
                                    return "🛑 SIM"
                                elif row['MESES_ESTOQUE'] > meses_parado: 
                                    return "🛑 SIM"
                            return ""
                            
                        df_dest['ESTOQUE PARADO'] = df_dest.apply(classificar_estoque_parado, axis=1)
                        
                        def processar_atipico(row):
                            meses_v = [row['MES_1'], row['MES_2'], row['MES_3'], row['MES_4']]
                            pico = max(meses_v)
                            outros = meses_v.copy()
                            outros.remove(pico)
                            
                            soma_outros = sum(outros)
                            media_sem = soma_outros / 3 if soma_outros > 0 else 0
                            
                            media_sis = row['MEDIA_SISTEMA']
                            if media_sis > 0 and media_sem > 0:
                                base_comp = min(media_sis, media_sem)
                            elif media_sem > 0:
                                base_comp = media_sem
                            else:
                                base_comp = media_sis
                                
                            if base_comp > 0 and pico >= (base_comp * fator_pico) and pico >= 30: 
                                return "⚠️ SIM", round(media_sem, 2)
                            elif base_comp == 0 and pico >= 30: 
                                return "⚠️ SIM", 0.0
                                
                            return "Não", media_sis

                        res_at = df_dest.apply(processar_atipico, axis=1)
                        df_dest['VENDA_ATIPICA'] = [x[0] for x in res_at]
                        df_dest['MEDIA_P_CALCULO'] = [x[1] for x in res_at]
                        dash_itens_pico += len(df_dest[df_dest['VENDA_ATIPICA'] == "⚠️ SIM"])
                        
                        def calcular_log(row):
                            cod = row['CODIGO']
                            nec_calc = (row['MEDIA_P_CALCULO'] * meta) - (row['ESTOQUE'] + row['COMPRADA'])
                            necessidade = nec_calc
                            
                            if necessidade > 0:
                                opcoes = []
                                for f_outra in dfs_por_filial.keys():
                                    if f_outra == nome_destino: continue
                                    chave = (f_outra, cod)
                                    if chave in tracker_estoque and tracker_estoque[chave]['EXCEDENTE'] > 0:
                                        opcoes.append({
                                            'filial': f_outra,
                                            'media': tracker_estoque[chave]['MEDIA'],
                                            'excedente': tracker_estoque[chave]['EXCEDENTE']
                                        })
                                
                                if opcoes:
                                    opcoes = sorted(opcoes, key=lambda x: (x['media'], -x['excedente']))
                                    trans_item = []
                                    nec_rest = necessidade
                                    
                                    for op in opcoes:
                                        if nec_rest <= 0: break
                                        
                                        chave_ced = (op['filial'], cod)
                                        sal_ced = tracker_estoque[chave_ced]['EXCEDENTE']
                                        
                                        if sal_ced <= 0: continue
                                        
                                        if nec_rest < 30 and sal_ced >= 30:
                                            qtd_a_tirar = 30
                                        else:
                                            qtd_a_tirar = min(nec_rest, sal_ced)
                                            
                                        if qtd_a_tirar >= 30:
                                            tracker_estoque[chave_ced]['EXCEDENTE'] -= qtd_a_tirar
                                            tracker_estoque[chave_ced]['ESTOQUE_FINAL'] -= qtd_a_tirar
                                            nec_rest -= qtd_a_tirar
                                            
                                            nome_ced = op['filial']
                                            if 'RIBEIR' in nome_ced: apelido = 'RP'
                                            elif 'LONDRINA' in nome_ced: apelido = 'Lon'
                                            elif 'FRANCA' in nome_ced: apelido = 'Frc'
                                            else: apelido = nome_ced
                                                
                                            trans_item.append(f"Tirar {int(qtd_a_tirar)} de {apelido}")
                                            
                                    if trans_item: 
                                        return " | ".join(trans_item), round(max(0, nec_rest), 2)
                                        
                            return "0", round(max(0, necessidade), 2)

                        res_log = df_dest.apply(calcular_log, axis=1)
                        df_dest['TRANS INTERNA'] = [x[0] for x in res_log]
                        sug_base = [x[1] for x in res_log]

                        def aplicar_mult(row, sug):
                            if sug <= 0: 
                                return 0
                                
                            forn = str(row.get('FORNECEDOR', '')).upper()
                            desc = str(row.get('DESCRICAO', '')).upper()
                            
                            for idx, regra in df_regras_editado.iterrows():
                                f_regra = str(regra.get('FORNECEDOR', '')).
