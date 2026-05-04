import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from openpyxl.styles import PatternFill # <-- Nova ferramenta para colorir as células

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Compras e Transferências", layout="wide")

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
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if not texto: continue
                for l in texto.split('\n'):
                    if re.match(r'^\d{3,6}\s', l):
                        partes = l.split()
                        try:
                            dados.append({
                                'CODIGO': partes[0],
                                'DESCRICAO': " ".join(partes[1:-11]),
                                'EMB.': partes[-11],
                                'JAN/26': partes[-10],
                                'FEV/26': partes[-9],
                                'MAR/26': partes[-8],
                                'ABR/26': partes[-7],
                                'MEDIA': limpar_v(partes[-6]),
                                'ESTOQUE': limpar_v(partes[-5]),
                                'RESERVA': limpar_v(partes[-4]),
                                'COMPRADA': limpar_v(partes[-3]),
                                'MESES_ESTOQUE': limpar_v(partes[-1]),
                                'FILIAL_NOME': nome_filial
                            })
                        except: continue
        return pd.DataFrame(dados)
    except:
        return pd.DataFrame()

# --- INTERFACE WEB (BARRA LATERAL) ---
with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except:
        st.error("Arquivo 'logo.png' não encontrado no GitHub.")
    
    st.markdown("---")
    st.header("⚙️ Configurações")
    meta = st.number_input("Meta de estoque (meses)", min_value=1, value=2)
    uploaded_files = st.file_uploader("Selecione os 4 PDFs das Unidades", type="pdf", accept_multiple_files=True)

# --- CORPO DO SITE ---
st.title("📊 Gestão de Compras e Transferências")
st.markdown("### Tapeçaria")

if uploaded_files:
    if len(uploaded_files) > 0:
        if st.button("🚀 Processar Análise"):
            dfs_por_filial = {}
            todos_dados = []
            
            for f in uploaded_files:
                df = extrair_dados_pdf_web(f)
                if not df.empty:
                    dfs_por_filial[f.name.replace(".pdf", "").upper()] = df
                    todos_dados.append(df)
            
            if todos_dados:
                df_global = pd.concat(todos_dados)
                output = BytesIO()
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for nome_destino, df_dest in dfs_por_filial.items():
                        def calcular_logistica(row):
                            cod = row['CODIGO']
                            necessidade = (row['MEDIA'] * meta) - (row['ESTOQUE'] + row['COMPRADA'])
                            
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

                        res = df_dest.apply(calcular_logistica, axis=1)
                        df_dest['TRANSFERENCIA_INTERNA'] = [x[0] for x in res]
                        df_dest['SUGESTAO_COMPRA'] = [x[1] for x in res]
                        
                        cols = ['CODIGO', 'DESCRICAO', 'EMB.', 'JAN/26', 'FEV/26', 'MAR/26', 'ABR/26', 
                                'MEDIA', 'ESTOQUE', 'RESERVA', 'COMPRADA', 'MESES_ESTOQUE', 'SUGESTAO_COMPRA', 'TRANSFERENCIA_INTERNA']
                        
                        df_dest[cols].to_excel(writer, sheet_name=nome_destino[:30], index=False)
                        
                        # --- ESTILIZAÇÃO DAS CÉLULAS NO EXCEL ---
                        worksheet = writer.sheets[nome_destino[:30]]
                        
                        # Define as cores (Códigos Hexadecimais)
                        cor_verde = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
                        cor_azul = PatternFill(start_color="C9DAF8", end_color="C9DAF8", fill_type="solid")
                        
                        # Encontra a posição exata das colunas (O Excel começa a contar do 1)
                        idx_compra = cols.index('SUGESTAO_COMPRA') + 1
                        idx_transf = cols.index('TRANSFERENCIA_INTERNA') + 1
                        
                        # Passa linha por linha pintando se bater com as regras
                        for row_num in range(2, len(df_dest) + 2): # Começa na 2 porque a 1 é o cabeçalho
                            val_compra = worksheet.cell(row=row_num, column=idx_compra).value
                            val_transf = worksheet.cell(row=row_num, column=idx_transf).value
                            
                            # Regra 1: Verde se a sugestão de compra for maior que 0
                            try:
                                if float(val_compra) > 0:
                                    worksheet.cell(row=row_num, column=idx_compra).fill = cor_verde
                            except:
                                pass
                            
                            # Regra 2: Azul se a transferência for diferente de "0"
                            if str(val_transf) != "0" and val_transf is not None:
                                worksheet.cell(row=row_num, column=idx_transf).fill = cor_azul
                
                st.success("✅ Análise concluída com sucesso!")
                st.download_button(
                    label="📥 Baixar Relatório Consolidado",
                    data=output.getvalue(),
                    file_name="Relatorio_Compras_Tapecaria.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("Aguardando o upload dos arquivos PDF para análise.")
