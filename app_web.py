import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from openpyxl.styles import PatternFill

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Compras e Transferências", layout="wide")

# ==========================================
# --- TELA DE SENHA (BLOQUEIO DE ACESSO) ---
# ==========================================
SENHA_ACESSO = "Tape2026"

if "liberado" not in st.session_state:
    st.session_state.liberado = False

if not st.session_state.liberado:
    st.title("🔒 Acesso Restrito - Tapeçaria")
    st.info("Por favor, insira a senha de liberação para acessar o portal de análise.")
    
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
                        if m_upper not in meses_encontrados:
                            meses_encontrados.append(m_upper)

                for l in texto.split('\n'):
                    if re.match(r'^\d{3,6}\s', l):
                        partes = l.split()
                        try:
                            dados.append({
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
                                'FILIAL_NOME': nome_filial
                            })
                        except: continue
        return pd.DataFrame(dados), meses_encontrados
    except:
        return pd.DataFrame(), []

# --- INTERFACE WEB (BARRA LATERAL) ---
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except:
        st.error("Arquivo 'logo.png' não encontrado no GitHub.")
    
    st.markdown("---")
    st.header("⚙️ Configurações")
    meta = st.number_input("Meta de estoque (meses)", min_value=1, value=2)
    
    fator_pico = st.number_input("Sensibilidade de Pico (x vezes a média)", min_value=1.5, value=3.0, step=0.5)
    
    nome_sugerido = st.text_input("Nome do arquivo Excel", value="Relatorio_Compras_Tapecaria")
    nome_final_xlsx = nome_sugerido if nome_sugerido.endswith(".xlsx") else f"{nome_sugerido}.xlsx"
    
    st.markdown("---")
    uploaded_files = st.file_uploader("Selecione os 4 PDFs das Unidades", type="pdf", accept_multiple_files=True)

# --- CORPO DO SITE ---
st.title("📊 Gestão de Compras e Transferências")
st.markdown("### Tapeçaria - Inteligência de Estoque")

if uploaded_files:
    if len(uploaded_files) > 0:
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
            
            if len(meses_globais) < 4:
                meses_globais = ["MÊS 1", "MÊS 2", "MÊS 3", "MÊS 4"]
            
            if todos_dados:
                df_global = pd.concat(todos_dados).reset_index(drop=True)
                df_global['ESTOQUE_DISPONIVEL'] = df_global['ESTOQUE']
                
                df_parado = df_global[(df_global['MEDIA_SISTEMA'] == 0) | (df_global['MESES_ESTOQUE'] > 3)]
                if not df_parado.empty:
                    parado_por_filial = df_parado.groupby('FILIAL_NOME')['ESTOQUE'].sum().sort_values(ascending=False)
                    dash_filial_parada = f"{parado_por_filial.index[0]} ({int(parado_por_filial.iloc[0])} un.)"
                else:
                    dash_filial_parada = "Nenhum detectado"
                
                output = BytesIO()
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for nome_destino, df_dest in dfs_por_filial.items():
                        
                        def processar_atipico(row):
                            meses_valores = [row['MES_1'], row['MES_2'], row['MES_3'], row['MES_4']]
                            pico = max(meses_valores)
                            media_sis = row['MEDIA_SISTEMA']
                            if media_sis > 0 and pico >= (media_sis * fator_pico) and pico >= 30:
                                media_ajustada = (sum(meses_valores) - pico) / 3
                                return "⚠️ SIM", round(media_ajustada, 2)
                            return "Não", media_sis

                        res_atipico = df_dest.apply(processar_atipico, axis=1)
                        df_dest['VENDA_ATIPICA'] = [x[0] for x in res_atipico]
                        df_dest['MEDIA_P_CALCULO'] = [x[1] for x in res_atipico]
                        
                        dash_itens_pico += len(df_dest[df_dest['VENDA_ATIPICA'] == "⚠️ SIM"])
                        
                        def calcular_logistica(row):
                            cod = row['CODIGO']
                            media_usada = row['MEDIA_P_CALCULO']
                            necessidade = (media_usada * meta) - (row['ESTOQUE'] + row['COMPRADA'])
                            
                            if necessidade > 0:
                                outras = df_global[
                                    (df_global['CODIGO'] == cod) & 
                                    (df_global['FILIAL_NOME'] != nome_destino) & 
                                    (df_global['ESTOQUE_DISPONIVEL'] > 0) & 
                                    ((df_global['MESES_ESTOQUE'] > 3) | (df_global['MEDIA_SISTEMA'] == 0))
                                ]
                                
                                if not outras.empty:
                                    outras_ordenadas = outras.sort_values(by=['MEDIA_SISTEMA', 'MESES_ESTOQUE'], ascending=[True, False])
                                    
                                    transferencias_item = []
                                    necessidade_restante = necessidade
                                    
                                    for idx_global, cedente in outras_ordenadas.iterrows():
                                        if necessidade_restante <= 0:
                                            break
                                        
                                        saldo_cedente = df_global.loc[idx_global, 'ESTOQUE_DISPONIVEL']
                                        if saldo_cedente <= 0:
                                            continue
                                            
                                        qtd_a_tirar = min(necessidade_restante, saldo_cedente)
                                        
                                        if qtd_a_tirar > 0:
                                            df_global.loc[idx_global, 'ESTOQUE_DISPONIVEL'] -= qtd_a_tirar
                                            necessidade_restante -= qtd_a_tirar
                                            transferencias_item.append(f"Tirar {int(qtd_a_tirar)} de {cedente['FILIAL_NOME']}")
                                    
                                    if transferencias_item:
                                        texto_final_transf = " | ".join(transferencias_item)
                                        return texto_final_transf, round(max(0, necessidade_restante), 2)
                            
                            return "0", round(max(0, necessidade), 2)

                        res_log = df_dest.apply(calcular_logistica, axis=1)
                        df_dest['TRANS INTERNA'] = [x[0] for x in res_log]
                        df_dest['SUGESTAO COMPRA'] = [x[1] for x in res_log]
                        
                        dash_qtd_comprar += df_dest['SUGESTAO COMPRA'].sum()
                        
                        def extrair_numero_transf(texto):
                            if str(texto) == "0" or not texto: return 0
                            try:
                                return sum([int(n) for n in re.findall(r'\d+', str(texto))])
                            except: return 0
                            
                        dash_qtd_transferida += df_dest['TRANS INTERNA'].apply(extrair_numero_transf).sum()
                        
                        df_dest.rename(columns={
                            'MES_1': meses_globais[0], 'MES_2': meses_globais[1], 
                            'MES_3': meses_globais[2], 'MES_4': meses_globais[3],
                            'MEDIA_SISTEMA': 'MEDIA', 'MESES_ESTOQUE': 'MESES'
                        }, inplace=True)
                        
                        cols_finais = [
                            'CODIGO', 'DESCRICAO', 'EMB.', 
                            meses_globais[0], meses_globais[1], meses_globais[2], meses_globais[3], 
                            'MEDIA', 'ESTOQUE', 'RESERVA', 'COMPRADA', 'MESES', 
                            'SUGESTAO COMPRA', 'TRANS INTERNA', 'VENDA_ATIPICA'
                        ]
                        
                        df_dest[cols_finais].to_excel(writer, sheet_name=nome_destino[:30], index=False)
                        
                        worksheet = writer.sheets[nome_destino[:30]]
                        cor_verde = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
                        cor_azul = PatternFill(start_color="C9DAF8", end_color="C9DAF8", fill_type="solid")
                        cor_amarela = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                        
                        idx_compra = cols_finais.index('SUGESTAO COMPRA') + 1
                        idx_transf = cols_finais.index('TRANS INTERNA') + 1
                        idx_atipica = cols_finais.index('VENDA_ATIPICA') + 1
                        
                        for row_num in range(2, len(df_dest) + 2):
                            val_compra = worksheet.cell(row=row_num, column=idx_compra).value
                            val_transf = worksheet.cell(row=row_num, column=idx_transf).value
                            val_atipica = worksheet.cell(row=row_num, column=idx_atipica).value
                            
                            try:
                                if float(val_compra) > 0: worksheet.cell(row=row_num, column=idx_compra).fill = cor_verde
                            except: pass
                            
                            if str(val_transf) != "0" and val_transf is not None:
                                worksheet.cell(row=row_num, column=idx_transf).fill = cor_azul
                                
                            if val_atipica and "⚠️ SIM" in str(val_atipica):
                                worksheet.cell(row=row_num, column=idx_atipica).fill = cor_amarela
                
                # --- EXIBIÇÃO DO DASHBOARD ---
                st.markdown("---")
                st.subheader("📈 Resumo da Análise (Geral)")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric(label="🛒 Unidades a Comprar", value=f"{int(dash_qtd_comprar)} un.",
                              help="Total de mercadoria que precisa ser adquirida de fornecedores após esgotar o estoque interno.")
                with col2:
                    st.metric(label="🔄 Economia Logística", value=f"{int(dash_qtd_transferida)} un.", 
                              help="Total de mercadoria reaproveitada de estoques parados entre as filiais.")
                with col3:
                    st.metric(label="⚠️ Vendas Atípicas (Picos)", value=f"{int(dash_itens_pico)} itens",
                              help="Quantidade de produtos que tiveram picos esporádicos ignorados no cálculo para evitar compras superestimadas.")
                with col4:
                    st.metric(label="📦 Maior Estoque Parado", value=dash_filial_parada,
                              help="A filial que concentra o maior volume físico de produtos com média zero ou parados há mais de 3 meses.")
                
                st.markdown("---")
                st.success(f"✅ Arquivo pronto para download com a inteligência multi-filiais!")
                st.download_button(
                    label="📥 Baixar Relatório Avançado",
                    data=output.getvalue(),
                    file_name=nome_final_xlsx,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("Aguardando o upload dos arquivos PDF para análise.")
