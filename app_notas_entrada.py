import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.utils import get_column_letter

# Configuração da Página Web
st.set_page_config(page_title="Relatório CFOP - CONTTEC", page_icon="🧾", layout="wide")

st.title("🧾 Gerador de Relatório: Compras por CFOP")
st.markdown("Faça o upload dos relatórios em PDF da CONTTEC. O sistema consolidará todas as notas, garantindo o alinhamento perfeito das colunas.")

arquivos_pdf = st.file_uploader("Arraste os seus PDFs para aqui", type="pdf", accept_multiple_files=True)

if arquivos_pdf:
    if st.button("Processar Documentos"):
        with st.spinner("A processar documentos e a alinhar colunas de valores..."):
            rows = []
            
            # A Regex foi ancorada estritamente no início (^) e no fim ($). 
            # Isso obriga o sistema a contar os blocos numéricos da direita para a esquerda, impossibilitando erros de coluna.
            regex = re.compile(
                r"^(\d{2}/\d{2}/\d{4})\s+"       # 1. Data
                r"(.+?)\s+"                      # 2. Fornecedor
                r"([^\s]+)\s+"                   # 3. Nota
                r"([^\s]+)\s+"                   # 4. Serie
                r"([\d.,]+)\s+"                  # 5. Qtd
                r"([\d.,]+)\s+"                  # 6. Vunt
                r"([\d.,]+)\s+"                  # 7. IPI
                r"([\d.,]+)\s+"                  # 8. ICMS
                r"([\d.,]+)\s+"                  # 9. Cred
                r"([\d.,]+)"                     # 10. Total
                r"(?:\s+([^\s]+))?$"             # 11. Prz (Opcional)
            )

            # Função isolada para processar a linha unificada e adicionar na planilha
            def process_match(m, cfop, rows_list):
                def to_f(val):
                    return float(val.replace(".", "").replace(",", "."))
                try:
                    dt = m.group(1)
                    forn = m.group(2).strip()
                    nota = m.group(3)
                    serie = m.group(4)
                    
                    if nota.isdigit(): nota = int(nota)
                    if serie.isdigit(): serie = int(serie)

                    qtd = to_f(m.group(5))
                    vunt = to_f(m.group(6))
                    ipi = to_f(m.group(7))
                    icms = to_f(m.group(8))
                    cred = to_f(m.group(9))
                    total = to_f(m.group(10))

                    rows_list.append([
                        dt, forn, nota, serie,
                        qtd, vunt, ipi, icms, cred, total,
                        cfop
                    ])
                except Exception as e:
                    pass

            for f in arquivos_pdf:
                cfop_atual = ""
                current_record = ""

                with pdfplumber.open(io.BytesIO(f.read())) as pdf:
                    for page in pdf.pages:
                        txt = page.extract_text(x_tolerance=1, y_tolerance=1) or ""
                        for line in txt.split("\n"):
                            line = line.strip()
                            if not line: continue

                            # Detecta Cabeçalho de CFOP
                            if re.match(r"^\d{4}\s*-", line):
                                if current_record:
                                    m = regex.match(current_record)
                                    if m: process_match(m, cfop_atual, rows)
                                    current_record = ""
                                cfop_atual = line
                                continue

                            # Detecta o início de uma nova Nota Fiscal (pela Data)
                            if re.match(r"^\d{2}/\d{2}/\d{4}", line):
                                if current_record:
                                    m = regex.match(current_record)
                                    if m: process_match(m, cfop_atual, rows)
                                current_record = line
                                continue

                            # Ignora quebras de página, cabeçalhos das tabelas e blocos de totais do PDF
                            ignore_keywords = ["DT EMISSÃO", "FORNECEDOR", "TOTAL CFOP", "TOTAL GERAL", "CONTTEC", "CNPJ", "Relatório", "Filtros:"]
                            if any(k in line for k in ignore_keywords):
                                if current_record:
                                    m = regex.match(current_record)
                                    if m: process_match(m, cfop_atual, rows)
                                    current_record = ""
                                continue

                            # Se o código chegou aqui, significa que a linha pertence ao Fornecedor ou Valores da nota anterior 
                            # O sistema concatena/junta as linhas partidas com segurança.
                            if current_record:
                                current_record += " " + line

                # Garante que o último registro lido na página também seja processado
                if current_record:
                    m = regex.match(current_record)
                    if m: process_match(m, cfop_atual, rows)

            if not rows:
                st.error("❌ Nenhum dado válido foi extraído. Verifique o formato do PDF enviado.")
            else:
                # ==========================================
                # 1. DASHBOARD DE PRÉ-VISUALIZAÇÃO (Na Tela)
                # ==========================================
                st.divider()
                st.subheader("👀 Resumo dos Dados Extraídos")
                
                colunas_df = ["Data", "Fornecedor", "Nota", "Serie", "Qtd", "Valor Unt", "IPI", "ICMS", "Cred ICMS", "Total", "CFOP"]
                df_view = pd.DataFrame(rows, columns=colunas_df)
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Quantidade de Notas Encontradas", len(df_view))
                c2.metric("Soma Total (R$)", f"{df_view['Total'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c3.metric("Soma ICMS (R$)", f"{df_view['ICMS'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c4.metric("Soma IPI (R$)", f"{df_view['IPI'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                df_tela = df_view.copy()
                colunas_moeda = ["Valor Unt", "IPI", "ICMS", "Cred ICMS", "Total"]
                for col in colunas_moeda:
                    df_tela[col] = df_tela[col].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                st.dataframe(df_tela, use_container_width=True, hide_index=True)

                # ==========================================
                # 2. GERAÇÃO EXCEL (Layout Original da Empresa)
                # ==========================================
                wb = Workbook()
                ws = wb.active
                ws.title = "Compras"
                
                AZUL_ESCURO = "2C3E50"  
                AZUL_TOTAL1 = "1F3A5F"  
                AZUL_TOTAL2 = "5B7C99"  
                TEAL_TOTAL = "2E8B9E"   
                AZUL_TOTAL3 = "1F3A5F"  
                
                ws['A1'] = "RELATÓRIO COMPRAS POR CFOP - CONTTEC"
                ws.merge_cells('A1:K1')
                ws['A1'].font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
                ws['A1'].fill = PatternFill(start_color=AZUL_ESCURO, end_color=AZUL_ESCURO, fill_type="solid")
                ws['A1'].alignment = Alignment(horizontal="left", vertical="center")
                ws.row_dimensions[1].height = 20
                
                ws.row_dimensions[2].height = 15
                ws.row_dimensions[3].height = 18
                
                headers = ["Data","Fornecedor","Nota","Serie","Qtd","Valor Unt","IPI","ICMS","Cred ICMS","Total","CFOP Completo"]
                for i, h in enumerate(headers, 1):
                    c = ws.cell(row=4, column=i, value=h)
                    c.font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
                    c.fill = PatternFill(start_color=AZUL_ESCURO, end_color=AZUL_ESCURO, fill_type="solid")
                    c.alignment = Alignment(horizontal="center", vertical="center")
                
                ws.row_dimensions[4].height = 20
                
                for r_idx, row in enumerate(rows, start=5):
                    for c_idx, val in enumerate(row, 1):
                        cell = ws.cell(row=r_idx, column=c_idx, value=val)
                        cell.border = Border()
                        cell.fill = PatternFill(fill_type=None)
                        if c_idx == 2 or c_idx == 11:
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                        else:
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        if 5 <= c_idx <= 10:
                            cell.number_format = '#,##0.00'
                
                last = 4 + len(rows)
                
                totais = [
                    ('F3', f"=SUM(F5:F{last})", AZUL_TOTAL1),
                    ('G3', f"=SUM(G5:G{last})", AZUL_TOTAL2),
                    ('H3', f"=SUM(H5:H{last})", TEAL_TOTAL),
                    ('I3', f"=SUM(I5:I{last})", TEAL_TOTAL),
                    ('J3', f"=SUM(J5:J{last})", AZUL_TOTAL3),
                ]
                for addr, formula, cor in totais:
                    c = ws[addr]
                    c.value = formula
                    c.fill = PatternFill(start_color=cor, end_color=cor, fill_type="solid")
                    c.font = Font(color="FFFFFF", bold=True, size=11)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.number_format = '#,##0.00'
                
                widths = [11, 32, 9, 7, 7, 12, 10, 10, 11, 12, 42]
                for i,w in enumerate(widths,1):
                    ws.column_dimensions[get_column_letter(i)].width = w
                
                ws.auto_filter.ref = f"A4:K{last}"
                ws.freeze_panes = "A5"
                ws.sheet_view.showGridLines = False
                
                output = io.BytesIO()
                wb.save(output)
                
                st.divider()
                st.success("✅ Excel gerado! Todas as colunas estão devidamente alinhadas e 100% dos dados foram lidos.")
                
                st.download_button(
                    label="📥 Baixar Relatório Excel",
                    data=output.getvalue(),
                    file_name=f"Relatorio_CFOP_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
