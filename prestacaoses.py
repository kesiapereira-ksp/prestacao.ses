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
        self.set_auto_page_break(auto=True, margin=10)

    def header(self):
        if self.logo_bytes:
            # Imagem no topo (y=4, w=80mm)
            self.image(self.logo_bytes, x=10, y=4, w=80)
            # Texto colado logo abaixo da imagem
            self.set_y(22)
        else:
            self.set_y(8)

        self.set_font('Helvetica', 'B', 11)
        self.cell(0, 4.5, f"Projeto: {self.nome_projeto}", border=False, ln=True, align='L')
        self.set_font('Helvetica', '', 9)
        self.cell(0, 3.8, "Demonstrativo Rateio Estrutura Administrativa", border=False, ln=True, align='L')
        self.set_font('Helvetica', 'I', 7)
        self.cell(0, 3.5, "(valores expressos em Reais)", border=False, ln=True, align='L')
        self.ln(1)

    def footer(self):
        self.set_y(-10)
        self.set_font('Helvetica', 'I', 7)
        self.cell(0, 8, f'Página {self.page_no()}', align='C')

def gerar_excel_projeto(nome_projeto, linhas_dados, tot_pagto, tot_ressarc):
    buffer = io.BytesIO()
    
    dados_excel = []
    for item in linhas_dados:
        dados_excel.append({
            "NATUREZA DA DESPESA": item['nat'],
            "GRUPO": item['grupo'],
            "CREDOR": item['credor'],
            "CNPJ/CPF": item['cnpj'],
            "NOTA FISCAL": item['nf'],
            "DATA PAGTO": item['data_p'],
            "VALOR Pagamento": item['val_p_num'],
            "Data Ressarc.": item['data_r'],
            "Ag. e Conta Corrente": item['ag_conta'],
            "VALOR Ressarc.": item['val_r_num']
        })
        
    df_export = pd.DataFrame(dados_excel)
    
    linha_total = {
        "NATUREZA DA DESPESA": "", "GRUPO": "", "CREDOR": "",
        "CNPJ/CPF": "", "NOTA FISCAL": "", "DATA PAGTO": "TOTAL",
        "VALOR Pagamento": tot_pagto, "Data Ressarc.": "",
        "Ag. e Conta Corrente": "", "VALOR Ressarc.": tot_ressarc
    }
    df_export_final = pd.concat([df_export, pd.DataFrame([linha_total])], ignore_index=True)
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_export_final.to_excel(writer, sheet_name='Demonstrativo', index=False, startrow=3)
        ws = writer.sheets['Demonstrativo']
        
        ws.cell(row=1, column=1, value=f"Projeto: {nome_projeto}")
        ws.cell(row=2, column=1, value="Demonstrativo Rateio Estrutura Administrativa (valores em Reais)")
        
        last_row = len(df_export_final) + 6
        ws.cell(row=last_row, column=1, value="RESUMO DE FECHAMENTO DO PROJETO")
        ws.cell(row=last_row + 1, column=1, value="Total do Valor das Despesas (Pagamentos dos itens com ressarcimento):")
        ws.cell(row=last_row + 1, column=7, value=tot_pagto)
        ws.cell(row=last_row + 2, column=1, value="Total das Despesas Atribuídas ao Projeto (Ressarcimento):")
        ws.cell(row=last_row + 2, column=7, value=tot_ressarc)

    return buffer.getvalue()

def gerar_pdf_e_excel_projeto(nome_projeto, df_planilha, col_map, col_projeto=None, logo_bytes=None):
    pdf = PDFDemonstrativo(nome_projeto, logo_bytes=logo_bytes)
    pdf.add_page()
    
    cols_w = (38, 28, 42, 28, 16, 18, 24, 20, 36, 27)
    headers = [
        "NATUREZA DA DESPESA", "GRUPO", "CREDOR", "CNPJ/CPF",
        "NOTA FISCAL", "DATA PAGTO", "VALOR Pagamento", "Data Ressarc.",
        "Ag. e Conta Corrente", "VALOR Ressarc."
    ]
    alignments = ("LEFT", "LEFT", "LEFT", "CENTER", "CENTER", "CENTER", "RIGHT", "CENTER", "LEFT", "RIGHT")
    
    linhas_validas = []
    tot_pagto = 0.0
    tot_ressarc = 0.0
    
    for _, row in df_planilha.iterrows():
        v_ressarc = 0.0
        if col_projeto and col_projeto in row:
            v_ressarc = limpar_valor(row.get(col_projeto, 0))
        if v_ressarc <= 0 and col_map.get('val_ressarc'):
            v_ressarc = limpar_valor(row.get(col_map['val_ressarc'], 0))
            
        if v_ressarc <= 0:
            continue
            
        v_pagto = limpar_valor(row.get(col_map['val_pagto'], 0))
        tot_pagto += v_pagto
        tot_ressarc += v_ressarc
        
        linhas_validas.append({
            'nat': fmt_texto(row.get(col_map['nat'], '')),
            'grupo': fmt_texto(row.get(col_map['grupo'], '')),
            'credor': fmt_texto(row.get(col_map['credor'], '')),
            'cnpj': fmt_texto(row.get(col_map['cnpj'], '')),
            'nf': fmt_texto(row.get(col_map['nf'], '')),
            'data_p': fmt_texto(row.get(col_map['data_p'], '')),
            'val_p_str': fmt_moeda(v_pagto),
            'val_p_num': v_pagto,
            'data_r': fmt_texto(row.get(col_map['data_r'], '')),
            'ag_conta': fmt_texto(row.get(col_map['ag_conta'], '')),
            'val_r_str': fmt_moeda(v_ressarc),
            'val_r_num': v_ressarc
        })
        
    if not linhas_validas:
        return None, None, 0
        
    pdf.set_font('Helvetica', 'B', 6)
    
    with pdf.table(
        col_widths=cols_w,
        text_align=alignments,
        line_height=3.5,
        padding=1
    ) as table:
        header_row = table.row()
        for h in headers:
            header_row.cell(h)
            
        pdf.set_font('Helvetica', size=5.5)
        for item in linhas_validas:
            r = table.row()
            r.cell(item['nat'])
            r.cell(item['grupo'])
            r.cell(item['credor'])
            r.cell(item['cnpj'])
            r.cell(item['nf'])
            r.cell(item['data_p'])
            r.cell(item['val_p_str'])
            r.cell(item['data_r'])
            r.cell(item['ag_conta'])
            r.cell(item['val_r_str'])
            
        pdf.set_font('Helvetica', 'B', 6)
        r_tot = table.row()
        r_tot.cell("")
        r_tot.cell("")
        r_tot.cell("")
        r_tot.cell("")
        r_tot.cell("")
        r_tot.cell("TOTAL")
        r_tot.cell(fmt_moeda(tot_pagto))
        r_tot.cell("")
        r_tot.cell("")
        r_tot.cell(fmt_moeda(tot_ressarc))
        
    pdf.ln(3)
    pdf.set_font('Helvetica', 'B', 7.5)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(180, 5, "RESUMO DE FECHAMENTO DO PROJETO", border=1, fill=True, ln=True, align='C')
    pdf.set_font('Helvetica', '', 7)
    pdf.cell(120, 5, " Total do Valor das Despesas (Pagamentos dos itens com ressarcimento):", border=1)
    pdf.cell(60, 5, f"R$ {fmt_moeda(tot_pagto)}", border=1, ln=True, align='R')
    pdf.cell(120, 5, " Total das Despesas Atribuídas ao Projeto (Ressarcimento):", border=1)
    pdf.set_font('Helvetica', 'B', 7)
    pdf.cell(60, 5, f"R$ {fmt_moeda(tot_ressarc)}", border=1, ln=True, align='R')
    
    pdf_bytes = bytes(pdf.output())
    excel_bytes = gerar_excel_projeto(nome_projeto, linhas_validas, tot_pagto, tot_ressarc)
    
    return pdf_bytes, excel_bytes, len(linhas_validas)

# --- Interface Streamlit ---
st.title("Gerador de Demonstrativos por Projeto (PDF e Excel) 📄📊")

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
        st.write("⚙️ **1. Selecione as colunas dos PROJETOS/CONTRATOS que deseja separar:**")
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
            col_credor = st.selectbox("Credor / Fornecedor:", options=colunas)
            col_cnpj = st.selectbox("CNPJ/CPF:", options=colunas)
            
        with c2:
            col_nf = st.selectbox("Nota Fiscal:", options=colunas)
            col_data_p = st.selectbox("Data Pagamento:", options=colunas)
            col_val_p = st.selectbox("Valor do Pagamento:", options=colunas)
            
        with c3:
            col_data_r = st.selectbox("Data Ressarcimento:", options=colunas)
            col_ag_conta = st.selectbox("Ag. e Conta Corrente:", options=colunas)
            col_val_r = st.selectbox("Valor de Ressarcimento (Padrão/Geral):", options=colunas)

        col_map = {
            'nat': col_nat, 'grupo': col_grupo, 'credor': col_credor,
            'cnpj': col_cnpj, 'nf': col_nf, 'data_p': col_data_p,
            'val_pagto': col_val_p, 'data_r': col_data_r,
            'ag_conta': col_ag_conta, 'val_ressarc': col_val_r
        }

        if projetos_selecionados and st.button("Gerar Demonstrativos (PDF + Excel)"):
            logo_bytes = io.BytesIO(arquivo_logo.read()) if arquivo_logo else None
            zip_buffer = io.BytesIO()
            resumo_geracao = {}

            with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
                for col_proj in projetos_selecionados:
                    nome_contrato = str(col_proj).strip()
                    
                    pdf_bytes, excel_bytes, qtd_itens = gerar_pdf_e_excel_projeto(
                        nome_projeto=nome_contrato, 
                        df_planilha=df, 
                        col_map=col_map, 
                        col_projeto=col_proj, 
                        logo_bytes=logo_bytes
                    )
                    
                    if qtd_itens > 0 and pdf_bytes and excel_bytes:
                        nome_limpo = re.sub(r'[\\/*?:"<>|]', '_', nome_contrato).strip()
                        zip_file.writestr(f"Demonstrativo_{nome_limpo}.pdf", pdf_bytes)
                        zip_file.writestr(f"Demonstrativo_{nome_limpo}.xlsx", excel_bytes)
                        
                        resumo_geracao[nome_contrato] = qtd_itens

            if resumo_geracao:
                st.success("✅ Demonstrativos (PDF e Excel) gerados com sucesso!")
                st.write("**Resumo dos arquivos incluídos no pacote ZIP:**")
                for proj, qtd in resumo_geracao.items():
                    st.write(f"- **{proj}**: {qtd} item(ns) -> Gerados: `Demonstrativo_{proj}.pdf` e `Demonstrativo_{proj}.xlsx`")

                st.download_button(
                    label="⬇️ Baixar Pacote Completo (PDFs + Excels em ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Demonstrativos_Projetos_PDF_Excel.zip",
                    mime="application/zip"
                )
            else:
                st.warning("Nenhum item com ressarcimento maior que zero foi encontrado para os projetos selecionados.")

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
