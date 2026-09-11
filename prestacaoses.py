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

def fmt_moeda(val):
    return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

class PDFDemonstrativo(FPDF):
    def __init__(self, nome_projeto, logo_bytes=None):
        super().__init__(orientation='L', unit='mm', format='A4')
        self.nome_projeto = nome_projeto
        self.logo_bytes = logo_bytes
        self.set_auto_page_break(auto=True, margin=12)

    def header(self):
        # Adiciona a imagem do timbrado/logo se tiver sido enviada
        if self.logo_bytes:
            # Posiciona no canto superior esquerdo (x=10, y=8) com largura de 35mm
            self.image(self.logo_bytes, x=10, y=8, w=35)
            self.set_y(22)  # Baixa a posição vertical do texto para não sobrepor o timbrado
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

def gerar_pdf_projeto(nome_projeto, df_proj, col_map, logo_bytes=None):
    pdf = PDFDemonstrativo(nome_projeto, logo_bytes=logo_bytes)
    pdf.add_page()
    
    # 10 colunas ajustadas para a página A4 Paisagem (Largura útil = 277mm)
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
    
    # Linhas de Dados
    pdf.set_font('Helvetica', size=5.5)
    
    tot_pagto = 0.0
    tot_ressarc = 0.0
    qtd_linhas = 0
    
    for _, row in df_proj.iterrows():
        v_pagto = limpar_valor(row.get(col_map['val_pagto'], 0))
        v_ressarc = limpar_valor(row.get(col_map['val_ressarc'], 0))
        
        if v_pagto <= 0 and v_ressarc <= 0:
            continue
            
        tot_pagto += v_pagto
        tot_ressarc += v_ressarc
        qtd_linhas += 1
        
        pdf.cell(cols_w[0], 5, str(row.get(col_map['nat'], ''))[:30], border=1)
        pdf.cell(cols_w[1], 5, str(row.get(col_map['grupo'], ''))[:20], border=1)
        pdf.cell(cols_w[2], 5, str(row.get(col_map['credor'], ''))[:28], border=1)
        pdf.cell(cols_w[3], 5, str(row.get(col_map['cnpj'], '')), border=1, align='C')
        pdf.cell(cols_w[4], 5, str(row.get(col_map['nf'], '')), border=1, align='C')
        pdf.cell(cols_w[5], 5, str(row.get(col_map['data_p'], '')), border=1, align='C')
        pdf.cell(cols_w[6], 5, fmt_moeda(v_pagto), border=1, align='R')
        pdf.cell(cols_w[7], 5, str(row.get(col_map['data_r'], '')), border=1, align='C')
        pdf.cell(cols_w[8], 5, str(row.get(col_map['ag_conta'], ''))[:22], border=1)
        pdf.cell(cols_w[9], 5, fmt_moeda(v_ressarc), border=1, align='R', ln=True)
        
    # Linha Totalizadora
    if qtd_linhas > 0:
        pdf.set_font('Helvetica', 'B', 6)
        pdf.set_fill_color(240, 240, 240)
        w_tot_label = sum(cols_w[:6])
        pdf.cell(w_tot_label, 5, "TOTAL", border=1, fill=True, align='R')
        pdf.cell(cols_w[6], 5, fmt_moeda(tot_pagto), border=1, fill=True, align='R')
        pdf.cell(cols_w[7] + cols_w[8], 5, "", border=1, fill=True)
        pdf.cell(cols_w[9], 5, fmt_moeda(tot_ressarc), border=1, fill=True, align='R', ln=True)
    
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
        
        st.write("**Mapeamento das Colunas:**")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            col_proj = st.selectbox("Projeto:", options=colunas)
            col_nat = st.selectbox("Natureza da Despesa:", options=colunas)
            col_grupo = st.selectbox("Grupo:", options=colunas)
            col_credor = st.selectbox("Credor:", options=colunas)
        with c2:
            col_cnpj = st.selectbox("CNPJ/CPF:", options=colunas)
            col_nf = st.selectbox("Nota Fiscal:", options=colunas)
            col_data_p = st.selectbox("Data Pagamento:", options=colunas)
            col_val_p = st.selectbox("Valor Pagamento:", options=colunas)
        with c3:
            col_data_r = st.selectbox("Data Ressarcimento:", options=colunas)
            col_ag_conta = st.selectbox("Ag. e Conta Corrente:", options=colunas)
            col_val_r = st.selectbox("Valor Ressarcimento:", options=colunas)

        col_map = {
            'nat': col_nat, 'grupo': col_grupo, 'credor': col_credor,
            'cnpj': col_cnpj, 'nf': col_nf, 'data_p': col_data_p,
            'val_pagto': col_val_p, 'data_r': col_data_r,
            'ag_conta': col_ag_conta, 'val_ressarc': col_val_r
        }

        if st.button("Gerar Demonstrativos em PDF"):
            # Lê os bytes da logo se tiver enviado
            logo_bytes = io.BytesIO(arquivo_logo.read()) if arquivo_logo else None

            zip_buffer = io.BytesIO()
            projetos = df[col_proj].dropna().unique()
            resumo_geracao = {}

            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
                for proj in projetos:
                    df_proj = df[df[col_proj] == proj]
                    pdf_bytes, qtd_itens = gerar_pdf_projeto(str(proj), df_proj, col_map, logo_bytes=logo_bytes)
                    
                    if qtd_itens > 0:
                        nome_limpo = re.sub(r'[\\/*?:"<>|]', '_', str(proj)).strip()
                        zip_file.writestr(f"Demonstrativo_{nome_limpo}.pdf", pdf_bytes)
                        resumo_geracao[proj] = qtd_itens

            st.success("✅ Demonstrativos gerados com sucesso!")
            st.write("**Resumo dos PDFs gerados (sem despesas zeradas):**")
            for proj, qtd in resumo_geracao.items():
                st.write(f"- **{proj}**: {qtd} item(ns) de despesa")

            st.download_button(
                label="⬇️ Baixar Todos os PDFs (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Demonstrativos_Projetos.pdf.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
