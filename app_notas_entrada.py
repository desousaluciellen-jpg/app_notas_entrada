import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Configuração da Página Web
st.set_page_config(page_title="Relatório CFOP - CONTTEC", page_icon="🧾", layout="wide")

st.title("🧾 Gerador de Relatório: Compras por CFOP")
st.markdown("Faça o upload dos relatórios em PDF da CONTTEC. O sistema unificará os dados e aplicará a formatação padrão da empresa no Excel.")

# Área de Upload no Navegador
arquivos_pdf = st.file_uploader("Arraste os seus PDFs para aqui", type="pdf", accept_multiple_files=True)

if arquivos_pdf:
    if st.button("Processar Documentos"):
        with st.spinner("A analisar PDFs e a formatar o Excel..."):
            rows = []
            regex = re.compile(r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(\d+)\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+(\d+)")

            # Leitura na Memória (Adaptado para Web)
            for f in arquivos_pdf:
                cfop = ""
                with pdfplumber.open(io.BytesIO(f.read())) as pdf:
                    for page in pdf.pages:
                        txt = page.extract_text(x_tolerance=1, y_tolerance=1) or ""
                        for line in txt.split("\n"):
                            line = line.strip()
                            if re.match(r"\d{4}\s*-\s*", line):
                                cfop = line
                                continue
                            m = regex.match(line)
                            if m:
                                dt, forn, nota, serie, qtd, vunt, ipi, icms, cred, total, prz = m.groups()
                                rows.append([
                                    dt, forn, int(nota), int(serie),
                                    float(qtd.replace(".","").replace(",",".")),
                                    float(vunt.replace(".","").replace(",",".")),
                                    float(ipi.replace(".","").replace(",",".")),
                                    float(icms.replace(".","").replace(",",".")),
                                    float(cred.replace(".","").replace(",",".")),
                                    float(total.replace(".","").replace(",",".")),
                                    cfop
                                ])

            if not rows:
                st.error("Nenhum dado compatível encontrado nos PDFs enviados.")
            else:
                # ==========================================
                # 1. PRÉ-VISUALIZAÇÃO NO ECRÃ (DASHBOARD)
                # ==========================================
                st.divider()
                st.subheader("👀 Resumo dos Dados Extraídos")
                
                # Criar um DataFrame pandas apenas para visualização no site
                colunas_df = ["Data", "Fornecedor", "Nota", "Serie", "Qtd", "Valor Unt", "IPI", "ICMS", "Cred ICMS", "Total", "CFOP"]
                df_view = pd.DataFrame(rows, columns=colunas_df)
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Quantidade de Registos", len(df_view))
                c2.metric("Soma Total (R$)", f"{df_view['Total'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c3.metric("Soma ICMS (R$)", f"{df_view['ICMS'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c4.metric("Soma IPI (R$)", f"{df_view['IPI'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                # Mostra a tabela interativa formatada no ecrã
                df_tela = df_view.copy()
                colunas_moeda = ["Valor Unt", "IPI", "ICMS", "Cred ICMS", "Total"]
                for col in colunas_moeda:
                    df_tela[col] = df_tela[col].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                st.dataframe(df_tela, use_container_width=True, hide_index=True)

                # ==========================================
                # 2. GERAÇÃO DO EXCEL (A sua formatação exata)
                # ==========================================
                wb = Workbook()
                ws = wb.active
                ws.title = "Compras"
                
                # CORES EXATAS DA IMAGEM
                AZUL_ESCURO = "2C3E50"  # titulo e cabecalho
                AZUL_TOTAL1 = "1F3A5F"  # 112.581,81
                AZUL_TOTAL2 = "5B7C99"  # 2.531,28
                TEAL_TOTAL = "2E8B9E"   # 1.208,69 e 0,00
                AZUL_TOTAL3 = "1F3A5F"  # 205.956,09
                
                # LINHA 1 - TITULO
                ws['A1'] = "RELATÓRIO COMPRAS POR CFOP - CONTTEC (CONSOLIDADO)"
                ws.merge_cells('A1:K1')
                ws['A1'].font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
                ws['A1'].fill = PatternFill(start_color=AZUL_ESCURO, end_color=AZUL_ESCURO, fill_type="solid")
                ws['A1'].alignment = Alignment(horizontal="left", vertical="center")
                ws.row_dimensions[1].height = 20
                
                # LINHA 2 - vazia
                ws.row_dimensions[2].height = 15
                
                # LINHA 3 - TOTAIS (colunas F a J)
                ws.row_dimensions[3].height = 18
                
                # LINHA 4 - CABECALHO
                headers = ["Data","Fornecedor","Nota","Serie","Qtd","Valor Unt","IPI","ICMS","Cred ICMS","Total","CFOP Completo"]
                for i, h in enumerate(headers, 1):
                    c = ws.cell(row=4, column=i, value=h)
                    c.font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
                    c.fill = PatternFill(start_color=AZUL_ESCURO, end_color=AZUL_ESCURO, fill_type="solid")
                    c.alignment = Alignment(horizontal="center", vertical="center")
                
                ws.row_dimensions[4].height = 20
                
                # DADOS a partir linha 5
                for r_idx, row in enumerate(rows, start=5):
                    for c_idx, val in enumerate(row, 1):
                        cell = ws.cell(row=r_idx, column=c_idx, value=val)
                        cell.border = Border() # SEM BORDAS
                        cell.fill = PatternFill(fill_type=None) # SEM FUNDO
                        if c_idx == 2 or c_idx == 11:
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                        else:
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        if 5 <= c_idx <= 10:
                            cell.number_format = '#,##0.00'
                
                last = 4 + len(rows)
                
                # APLICAR TOTAIS COM CORES E FORMULAS DO EXCEL
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
                
                # LARGURAS
                widths = [11, 32, 9, 7, 7, 12, 10, 10, 11, 12, 42]
                for i,w in enumerate(widths,1):
                    ws.column_dimensions[get_column_letter(i)].width = w
                
                # FILTRO E PAINEL CONGELADO
                ws.auto_filter.ref = f"A4:K{last}"
                ws.freeze_panes = "A5"
                ws.sheet_view.showGridLines = False
                
                # Salvar em Memória para o Botão de Download
                output = io.BytesIO()
                wb.save(output)
                
                st.divider()
                st.success("✅ Planilha formatada com sucesso!")
                
                # Botão de Download
                st.download_button(
                    label="📥 Baixar Relatório Excel (CONTTEC)",
                    data=output.getvalue(),
                    file_name=f"Relatorio_CFOP_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
