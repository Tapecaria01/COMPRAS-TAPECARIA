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
# --- DADOS PADRÃO DE FORNECEDORES ---
# ==========================================
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
                            
                for l in texto.split('\n'):
                    if "SEGMENTO" in l.upper():
                        match = re.search(r'SEGMENTO\s*:\s*(.*)', l, re.IGNORECASE)
                        if match: 
                            fornecedor_atual = match.group(1).strip()
                    elif re.match(r'^\d{3,6}\s', l):
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

# --- INTERFACE WEB (BARRA LATERAL) ---
with st.sidebar:
    try: 
        st.image("logo.png", use_container_width=True)
    except: 
        pass
        
    st.markdown("---")
    st.header("⚙️ Configurações Gerais")
    meta = st.number_input("Meta de estoque (meses)", min_value=1, value=2)
    meses_parado = st.number_input("Considerar estoque parado após (meses)", min_value=1, value=3, step=1)
    fator_pico = st.number_input("Sensibilidade de Pico (x vezes a média)", min_value=1.5, value=2.5, step=0.5)
    nome_sugerido = st.text_input("Nome do ficheiro Excel", value="Relatorio_Compras_Tapecaria")
    
    if nome_sugerido.endswith(".xlsx"):
        nome_final_xlsx = nome_sugerido
    else:
        nome_final_xlsx = f"{nome_sugerido}.xlsx"
        
    st.markdown("---")
    
    uploaded_files = st.file_uploader("Selecione os 4 PDFs das Unidades", type="pdf", accept_multiple_files=True)
    
    with st.expander("🏭 Fornecedores e Múltiplos"):
        st.caption("Adicione ou edite regras. Use a última linha vazia para adicionar novos.")
        
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
    st.title("Compras Inteligente")

st.markdown("### Gestão de Compras e Transferências")

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
                if len(meses) >= 4 and not meses_globais: 
                    meses_globais = meses[:4]
        
        if not meses_globais: 
            meses_globais = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]
        
        if todos_dados:
            df_global = pd.concat(todos_dados).reset_index(drop=True)
            df_global['ESTOQUE_DISPONIVEL'] = df_global['ESTOQUE']
            
            vendas_recentes = df_global['MES_1'] + df_global['MES_2'] + df_global['MES_3'] + df_global['MES_4']
            df_global['TOTAL_VENDAS_RECENTES'] = vendas_recentes
            
            def calcular_excedente(row):
                if row['MEDIA_SISTEMA'] == 0: 
                    return row['ESTOQUE_DISPONIVEL']
                else:
                    excesso = row['ESTOQUE_DISPONIVEL'] - (row['MEDIA_SISTEMA'] * meta)
                    return max(0, excesso) 
            
            df_global['EXCEDENTE_DISPONIVEL'] = df_global.apply(calcular_excedente, axis=1)

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
                            filtro1 = df_global['CODIGO'] == cod
                            filtro2 = df_global['FILIAL_NOME'] != nome_destino
                            filtro3 = df_global['EXCEDENTE_DISPONIVEL'] > 0
                            outras = df_global[filtro1 & filtro2 & filtro3]
                            
                            if not outras.empty:
                                outras_ord = outras.sort_values(
                                    by=['MEDIA_SISTEMA', 'EXCEDENTE_DISPONIVEL'], 
                                    ascending=[True, False]
                                )
                                trans_item = []
                                nec_rest = necessidade
                                
                                for idx_g, cedente in outras_ord.iterrows():
                                    if nec_rest <= 0: 
                                        break
                                        
                                    sal_ced = df_global.loc[idx_g, 'EXCEDENTE_DISPONIVEL']
                                    if sal_ced <= 0: 
                                        continue
                                        
                                    if nec_rest < 30 and sal_ced >= 30:
                                        qtd_a_tirar = 30
                                    else:
                                        qtd_a_tirar = min(nec_rest, sal_ced)
                                        
                                    if qtd_a_tirar >= 30:
                                        df_global.loc[idx_g, 'EXCEDENTE_DISPONIVEL'] -= qtd_a_tirar
                                        df_global.loc[idx_g, 'ESTOQUE_DISPONIVEL'] -= qtd_a_tirar
                                        nec_rest -= qtd_a_tirar
                                        
                                        nome_ced = str(cedente['FILIAL_NOME'])
                                        if 'RIBEIR' in nome_ced:
                                            apelido = 'RP'
                                        elif 'LONDRINA' in nome_ced:
                                            apelido = 'Lon'
                                        elif 'FRANCA' in nome_ced:
                                            apelido = 'Frc'
                                        else:
                                            apelido = nome_ced
                                            
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
                            f_regra = str(regra.get('FORNECEDOR', '')).upper()
                            
                            if f_regra and f_regra != "NAN" and f_regra in forn:
                                p_chave = str(regra.get('PALAVRA_CHAVE', '')).upper().strip()
                                if p_chave and p_chave != "NAN" and p_chave != "NONE":
                                    if p_chave not in desc:
                                        continue 
                                        
                                try:
                                    mult = int(regra['MULTIPLO'])
                                    tol = int(regra['TOLERANCIA'])
                                except:
                                    continue 
                                    
                                if mult > 0:
                                    base = (int(sug) // mult) * mult
                                    rest = sug % mult
                                    return int(base + mult if rest >= tol else base)
                                    
                        # Aqui está a correção: arredonda para cima para garantir stock inteiro
                        return int(math.ceil(sug))

                    df_sug_zip = zip(df_dest.to_dict('records'), sug_base)
                    df_dest['SUGESTAO COMPRA'] = [aplicar_mult(r, s) for r, s in df_sug_zip]
                    dash_qtd_comprar += df_dest['SUGESTAO COMPRA'].sum()
                    
                    def extrair_n(t): 
                        if str(t) == "0":
                            return 0
                        numeros = re.findall(r'\d+', str(t))
                        return sum([int(n) for n in numeros])
                        
                    dash_qtd_transferida += df_dest['TRANS INTERNA'].apply(extrair_n).sum()
                    
                    renames = {
                        'MES_1': meses_globais[0], 
                        'MES_2': meses_globais[1], 
                        'MES_3': meses_globais[2], 
                        'MES_4': meses_globais[3], 
                        'MEDIA_SISTEMA': 'MEDIA', 
                        'MESES_ESTOQUE': 'MESES'
                    }
                    df_dest.rename(columns=renames, inplace=True)
                    
                    cols_f = [
                        'CODIGO', 
                        'DESCRICAO', 
                        'EMB.', 
                        meses_globais[0], 
                        meses_globais[1], 
                        meses_globais[2], 
                        meses_globais[3], 
                        'MEDIA', 
                        'ESTOQUE', 
                        'RESERVA', 
                        'COMPRADA', 
                        'MESES', 
                        'SUGESTAO COMPRA', 
                        'TRANS INTERNA', 
                        'VENDA_ATIPICA', 
                        'ESTOQUE PARADO'
                    ]
                    
                    sheet_n = nome_destino[:30]
                    df_dest[cols_f].to_excel(writer, sheet_name=sheet_n, index=False)
                    
                    ws = writer.sheets[sheet_n]
                    cv = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
                    ca = PatternFill(start_color="C9DAF8", end_color="C9DAF8", fill_type="solid")
                    cl = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
                    cy = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                    c_red = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
                    
                    idx_estoque = cols_f.index('ESTOQUE') + 1 
                    idx_comprada = cols_f.index('COMPRADA') + 1 
                    idx_compra = cols_f.index('SUGESTAO COMPRA') + 1
                    idx_transf = cols_f.index('TRANS INTERNA') + 1
                    idx_atipica = cols_f.index('VENDA_ATIPICA') + 1
                    idx_parado = cols_f.index('ESTOQUE PARADO') + 1
                    
                    for r in range(2, len(ws['A']) + 1):
                        val_comprada = limpar_v(ws.cell(r, idx_comprada).value)
                        if val_comprada > 0:
                            ws.cell(r, idx_comprada).fill = cl 
                            
                        val_compra = limpar_v(ws.cell(r, idx_compra).value)
                        if val_compra > 0:
                            ws.cell(r, idx_compra).fill = cv 
                            
                        val_transf = str(ws.cell(r, idx_transf).value)
                        if val_transf != "0" and val_transf != "None":
                            ws.cell(r, idx_transf).fill = ca 
                            
                        val_atipica = str(ws.cell(r, idx_atipica).value)
                        if "⚠️ SIM" in val_atipica:
                            ws.cell(r, idx_atipica).fill = cy 
                        
                        val_parado = ws.cell(r, idx_parado).value
                        if val_parado and "🛑 SIM" in str(val_parado): 
                            ws.cell(r, idx_parado).fill = c_red
                            ws.cell(r, idx_estoque).fill = c_red

            tab1, tab2, tab3, tab4 = st.tabs([
                "📊 Visão Geral", 
                "🚨 Top Urgentes", 
                "📦 Estoque Parado", 
                "🔍 Prévia por Filial"
            ])
            
            with tab1:
                st.subheader("Indicadores de Desempenho")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("🛒 Comprar", f"{int(dash_qtd_comprar)} un.")
                c2.metric("🔄 Economia", f"{int(dash_qtd_transferida)} un.")
                c3.metric("⚠️ Picos", f"{int(dash_itens_pico)} itens")
                
                filtro_p1 = df_global['ESTOQUE_DISPONIVEL'] > 0
                filtro_p2 = df_global['MEDIA_SISTEMA'] == 0
                filtro_p3 = df_global['MESES_ESTOQUE'] > meses_parado
                df_p = df_global[filtro_p1 & (filtro_p2 | filtro_p3)]
                
                if not df_p.empty:
                    f_p = df_p.groupby('FILIAL_NOME')['ESTOQUE'].sum().idxmax()
                else:
                    f_p = "Nenhuma"
                    
                c4.metric("📦 Maior Estoque Parado", f_p)
                
                st.success("✅ Análise concluída! O Excel está pronto para download abaixo.")
                st.download_button(
                    label="📥 Baixar Relatório Excel", 
                    data=output.getvalue(), 
                    file_name=nome_final_xlsx, 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with tab2:
                st.subheader("Produtos com maior sugestão de compra")
                df_all = pd.concat(dfs_por_filial.values())
                top_compra = df_all[df_all['SUGESTAO COMPRA'] > 0].sort_values(
                    by='SUGESTAO COMPRA', ascending=False
                ).head(15)
                cols_view = ['CODIGO', 'DESCRICAO', 'FILIAL_NOME', 'SUGESTAO COMPRA', 'FORNECEDOR']
                st.dataframe(top_compra[cols_view], use_container_width=True)

            with tab3:
                st.subheader("Análise de Estoque Morto / Excedente por Filial")
                if not df_p.empty:
                    grafico_dados = df_p.groupby('FILIAL_NOME')['ESTOQUE'].sum().reset_index()
                    fig = px.bar(
                        grafico_dados, 
                        x='FILIAL_NOME', 
                        y='ESTOQUE', 
                        title="Volume de Estoque Acima do Limite (un.)", 
                        color='ESTOQUE', 
                        color_continuous_scale='Reds'
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else: 
                    st.info("Não há estoque parado detetado com os parâmetros atuais.")

            with tab4:
                sel_f = st.selectbox("Selecione a Filial para visualizar:", list(dfs_por_filial.keys()))
                st.dataframe(dfs_por_filial[sel_f], use_container_width=True)

    else: 
        st.info("A aguardar upload. Clique em 'Processar' para gerar a inteligência.")
