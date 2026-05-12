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
                                'MEDIA_SISTEMA': limpar_v(partes[-6]), # Média vinda do PDF
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
    
    fator_pico = st.number_input("Sensibilidade de Pico (x vezes a média)", min_value=1.5, value=3.0, step=0.5, 
                                 help="Se um mês vender mais que 3x a média, o sistema ignora esse pico no cálculo.")
    
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
                df_global = pd.concat(todos_dados)
                output = BytesIO()
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for nome_destino, df_dest in dfs_por_filial.items():
                        
                        # --- LÓGICA DE DETECÇÃO E AJUSTE DE PICO ---
                        def processar_atipico(row):
                            meses_valores = [row['MES_1'], row['MES_2'], row['MES_3'], row['MES_4']]
                            pico = max(meses_valores)
                            media_sis = row['MEDIA_SISTEMA']
                            
                            # Se for atípico (Pico > Média * Sensibilidade E Pico > 30 para evitar itens minúsculos)
                            if media_sis > 0 and pico >= (media_sis * fator_pico) and pico >= 30:
                                # Calcula a média SEM o pico (soma os 3 meses menores e divide por 3)
                                media_ajustada = (sum(meses_valores) - pico) / 3
                                return "⚠️ SIM", round(media_ajustada, 2)
                            else:
                                return "Não", media_sis

                        # Aplica a inteligência
                        res_atipico = df_dest.apply(processar_atipico, axis=1)
                        df_dest['VENDA_ATIPICA'] = [x[0] for x in res_atipico]
                        df_dest['MEDIA_P_CALCULO'] = [x[1] for x in res_atipico]
                        
                        # Função de Cálculo usando a MEDIA_P_CALCULO (que já vem corrigida se houver pico)
                        def calcular_logistica(row):
                            cod = row['CODIGO']
                            media_usada = row['MEDIA_P_CALCULO']
                            necessidade = (media_usada * meta) - (row['ESTOQUE'] + row['COMPRADA'])
                            
                            if necessidade > 0:
                                outras = df_global[
                                    (df_global['CODIGO'] == cod) & 
                                    (df_global['FILIAL_NOME'] != nome_destino) & 
                                    (df_global['ESTOQUE'] > 0) & 
                                    (df_global['MESES_ESTOQUE'] > 3)
                                ]
                                if not outras.empty:
                                    cedente = outras.sort_values(by='MESES_ESTOQUE', ascending=False).iloc[0]
                                    qtd = min(necessidade, cedente['ESTOQUE'])
                                    return f"Tirar {int(qtd)} de {cedente['FILIAL_NOME']}", round(max(0, necessidade - qtd), 2)
                            
                            return "0", round(max(0, necessidade), 2)

                        res_log = df_dest.apply(calcular_logistica, axis=1)
                        df_dest['TRANSFERENCIA_INTERNA'] = [x[0] for x in res_log]
                        df_dest['SUGESTAO_COMPRA'] = [x[1] for x in res_log]
                        
                        # Organiza Colunas para o Excel
                        cols_finais = ['CODIGO', 'DESCRICAO', 'EMB.', 
                                       meses_globais[0], meses_globais[1], meses_globais[2], meses_globais[3], 
                                       'MEDIA_SISTEMA', 'MEDIA_P_CALCULO', 'ESTOQUE', 'RESERVA', 'COMPRADA', 
                                       'VENDA_ATIPICA', 'SUGESTAO_COMPRA', 'TRANSFERENCIA_INTERNA']
                        
                        # Renomeia MES_1, etc para os nomes reais antes de exportar
                        df_dest.rename(columns={'MES_1': meses_globais[0], 'MES_2': meses_globais[1], 
                                                'MES_3': meses_globais[2], 'MES_4': meses_globais[3]}, inplace=True)
                        
                        df_dest[cols_finais].to_excel(writer, sheet_name=nome_destino[:30], index=False)
                        
                        # --- ESTILIZAÇÃO ---
                        worksheet = writer.sheets[nome_destino[:30]]
                        cor_verde = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
                        cor_azul = PatternFill(start_color="C9DAF8", end_color="C9DAF8", fill_type="solid")
                        cor_amarela = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                        
                        idx_compra = cols_finais.index('SUGESTAO_COMPRA') + 1
                        idx_transf = cols_finais.index('TRANSFERENCIA_INTERNA') + 1
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
                
                st.success(f"✅ Análise de Inteligência concluída! Salvo como: {nome_final_xlsx}")
                st.download_button(
                    label="📥 Baixar Relatório com Ajuste de Pico",
                    data=output.getvalue(),
                    file_name=nome_final_xlsx,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
