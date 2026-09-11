import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile
import re

def limpar_valor(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace("R$", "").strip()
    if ',' in val_str and '.' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

def fmt_texto(val):
    if pd.isna(val) or val is None or str(val).strip().lower() == 'nan':
        return ""
    if isinstance(val, pd.Timestamp):
        return val.strftime('%d/%m/%Y')
    val_str = str(val).strip()
    if ' 00:00:00' in val_str:
        val_str = val_str.replace(' 00:00:00', '')
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

def fmt_moeda(val):
    return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

class PDFDemonstrativo(FPDF):
    def __init__(self, nome_projeto, logo_bytes=None):
        super().__init__(orientation='L', unit='mm', format='A4')
        self.nome_projeto = nome_projeto
        self.logo_bytes = logo_bytes
        self.set_auto_page_break(auto=True, margin=12)

    def header(self):
        if self.logo_bytes:
            self.image(self.logo_bytes, x=10, y=8, w=35)
            self.set_y(22)
        else:
            self.set_y(10)

        self.set_font('Helvetica', 'B', 11)
        self.cell(0, 5, f"Projeto: {self.nome_projeto}", border=False, ln=True, align='L')
        self.set_font('Helvetica', '', 9)
        self.cell(0, 4, "Demonstrativo Rateio Estrutura Administrativa", border=False, ln=True, align='L')
        self.set_font('Helvetica', 'I', 7)
        self.cell(0, 4, "(valores expressos em Reais)", border=False, ln=True, align='L')
        self.ln(2)

    def footer(self):
        self.set_y(-10)
        self.set_font('Helvetica', 'I', 7)
        self.cell(0, 8, f'Página {self.page_no()}', align='C')

def gerar_pdf_projeto(nome_projeto, df_planilha, col_map, col_projeto=None, logo_bytes=None):
    pdf = PDFDemonstrativo(nome_projeto, logo_bytes=logo_bytes)
    pdf.add_page()
    
    cols_w = [42, 30, 42, 26, 16, 18, 24, 20, 34, 25]
    headers = [
        "NATUREZA DA DESPESA", "GRUPO", "CREDOR", "CNPJ/CPF",
        "NOTA FISCAL", "DATA PAGTO", "VALOR Pagamento", "Data Ressarc.",
        "Ag. e Conta Corrente", "VALOR Ressarc."
    ]
    
    # Cabeçalho da Tabela
    pdf.set_font('Helvetica', 'B', 6)
    pdf.set_fill_color(220, 220, 220)
    for w, h_text in zip(cols_w, headers):
        pdf.cell(w, 6, h_text, border=1, fill=True, align='C')
    pdf.ln()
    
    pdf.set_font('Helvetica', size=5.5)
    
    tot_pagto = 0.0
    tot_ressarc = 0.0
    qtd_linhas = 0
    
    for _, row in df_planilha.iterrows():
        # Busca o valor de ressarcimento prioritariamente da coluna do projeto selecionado;
        # se zerado, recorre à coluna mapeada de ressarcimento
        v_ressarc = 0.0
        if col_projeto and col_projeto in row:
            v_ressarc = limpar_valor(row.get(col_projeto, 0))
        if v_ressarc <= 0 and col_map.get('val_ressarc'):
            v_ressarc = limpar_valor(row.get(col_map['val_ressarc'], 0))
            
        v_pagto = limpar_valor(row.get(col_map['val_pagto'], 0))
        
        # Ignora linhas que não possuem pagamento nem ressarcimento
        if v_pagto <= 0 and v_ressarc <= 0:
            continue
            
        tot_pagto += v_pagto
        tot_ressarc += v_ressarc
        qtd_linhas += 1
        
        pdf.cell(cols_w[0], 5, fmt_texto(row.get(col_map['nat'], ''))[:30], border=1)
        pdf.cell(cols_w[1], 5, fmt_texto(row.get(col_map['grupo'], ''))[:20], border=1)
        pdf.cell(cols_w[2], 5, fmt_texto(row.get(col_map['credor'], ''))[:28], border=1)
        pdf.cell(cols_w[3], 5, fmt_texto(row.get(col_map['cnpj'], '')), border=1, align='C')
        pdf.cell(cols_w[4], 5, fmt_texto(row.get(col_map['nf'], '')), border=1, align='C')
        pdf.cell(cols_w[5], 5, fmt_texto(row.get(col_map['data_p'], '')), border=1, align='C')
        pdf.cell(cols_w[6], 5, fmt_moeda(v_pagto), border=1, align='R')
        pdf.cell(cols_w[7], 5, fmt_texto(row.get(col_map['data_r'], '')), border=1, align='C')
        pdf.cell(cols_w[8], 5, fmt_texto(row.get(col_map['ag_conta'], ''))[:22], border=1)
        pdf.cell(cols_w[9], 5, fmt_moeda(v_ressarc), border=1, align='R', ln=True)
        
    if qtd_linhas > 0:
        pdf.set_font('Helvetica', 'B', 6)
        pdf.set_fill_color(230, 230, 230)
        
        w_tot_label = sum(cols_w[:6])
        pdf.cell(w_tot_label, 6, "TOTAL", border=1, fill=True, align='R')
        pdf.cell(cols_w[6], 6, fmt_moeda(tot_pagto), border=1, fill=True, align='R')
        pdf.cell(cols_w[7] + cols_w[8], 6, "", border=1, fill=True)
        pdf.cell(cols_w[9], 6, fmt_moeda(tot_ressarc), border=1, fill=True, align='R', ln=True)
        
        pdf.ln(3)
        pdf.set_font('Helvetica', 'B', 7.5)
        pdf.set_fill_color(240, 240, 240)
        
        pdf.cell(180, 5, "RESUMO DE FECHAMENTO DO PROJETO", border=1, fill=True, ln=True, align='C')
        
        pdf.set_font('Helvetica', '', 7)
        pdf.cell(120, 5, " Total do Valor de Todas as Despesas (Pagamentos):", border=1)
        pdf.cell(60, 5, f"R$ {fmt_moeda(tot_pagto)}", border=1, ln=True, align='R')
        
        pdf.cell(120, 5, " Total das Despesas Atribuídas ao Projeto (Ressarcimento):", border=1)
        pdf.set_font('Helvetica', 'B', 7)
        pdf.cell(60, 5, f"R$ {fmt_moeda(tot_ressarc)}", border=1, ln=True, align='R')
    
    return bytes(pdf.output()), qtd_linhas

# --- Interface Streamlit ---
st.title("Gerador de Demonstrativo de Despesas por Projeto 📄📋")

col_left, col_right = st.columns([2, 1])

with col_left:
    arquivo_excel = st.file_uploader("1. Envie a planilha de despesas (Excel)", type=["xlsx", "xls"])
with col_right:
    arquivo_logo = st.file_uploader("2. Logo / Timbrado (Opcional)", type=["png", "jpg", "jpeg"])

if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel)
        colunas = df.columns.tolist()
        
        st.write("---")
        st.write("⚙️ **1. Selecione os PROJETOS/CONTRATOS que deseja separar:**")
        projetos_selecionados = st.multiselect(
            "Selecione um ou mais Projetos:",
            options=colunas,
            key="select_projetos"
        )
        
        st.write("⚙️ **2. Mapeie as colunas de dados fixos das despesas:**")
        c1, c2, c3 = st.columns(3)
        
        with c1:
            col_nat = st.selectbox("Natureza da Despesa:", options=colunas)
            col_grupo = st.selectbox("Grupo:", options=colunas)
            col_credor = st.selectbox("Credor:", options=colunas)
            col_cnpj = st.selectbox("CNPJ/CPF:", options=colunas)
            
        with c2:
            col_nf = st.selectbox("Nota Fiscal:", options=colunas)
            col_data_p = st.selectbox("Data Pagamento:", options=colunas)
            col_val_p = st.selectbox("Valor do Pagamento:", options=colunas)
            
        with c3:
            col_data_r = st.selectbox("Data Ressarcimento:", options=colunas)
            col_ag_conta = st.selectbox("Ag. e Conta Corrente:", options=colunas)
            col_val_r = st.selectbox("Valor de Ressarcimento:", options=colunas)

        col_map = {
            'nat': col_nat, 'grupo': col_grupo, 'credor': col_credor,
            'cnpj': col_cnpj, 'nf': col_nf, 'data_p': col_data_p,
            'val_pagto': col_val_p, 'data_r': col_data_r,
            'ag_conta': col_ag_conta, 'val_ressarc': col_val_r
        }

        if projetos_selecionados and st.button("Gerar Demonstrativos em PDF"):
            logo_bytes = io.BytesIO(arquivo_logo.read()) if arquivo_logo else None
            zip_buffer = io.BytesIO()
            resumo_geracao = {}

            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
                for col_proj in projetos_selecionados:
                    nome_contrato = str(col_proj).strip()
                    
                    pdf_bytes, qtd_itens = gerar_pdf_projeto(
                        nome_projeto=nome_contrato, 
                        df_planilha=df, 
                        col_map=col_map, 
                        col_projeto=col_proj, 
                        logo_bytes=logo_bytes
                    )
                    
                    if qtd_itens > 0:
                        nome_limpo = re.sub(r'[\\/*?:"<>|]', '_', nome_contrato).strip()
                        zip_file.writestr(f"Demonstrativo_{nome_limpo}.pdf", pdf_bytes)
                        resumo_geracao[nome_contrato] = qtd_itens

            st.success("✅ Demonstrativos gerados com sucesso!")
            st.write("**Resumo dos PDFs gerados:**")
            for proj, qtd in resumo_geracao.items():
                st.write(f"- **{proj}**: {qtd} item(ns) de despesa")

            st.download_button(
                label="⬇️ Baixar Todos os PDFs (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Demonstrativos_Projetos.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
