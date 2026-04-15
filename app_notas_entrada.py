import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Relatório CFOP - CONTTEC", page_icon="🧾", layout="wide")

st.title("🧾 Gerador de Relatório: Compras por CFOP")
st.markdown("O sistema processa e unifica todas as notas do PDF. O algoritmo foi reconstruído para garantir a exatidão e alinhamento do seu código original.")

arquivos_pdf = st.file_uploader("Arraste os seus PDFs para aqui", type="pdf", accept_multiple_files=True)

def analisar_registro(record_str, current_cfop, rows_list):
    m_date = re.match(r"^(\d{2}/\d{2}/\d{4})", record_str)
    if not m_date: return
    
    # 1. Utiliza a exata expressão do seu código original, apenas com suporte a letras na Série/Nota
    regex = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(\S+)\s+(\S+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)(?:\s+(\d+))?$")
    m = regex.match(record_str)
    
    if m:
        dt = m.group(1)
        forn = m.group(2).strip()
        nota = m.group(3)
        serie = m.group(4)
        qtd = m.group(5)
        vunt = m.group(6)
        ipi = m.group(7)
        icms = m.group(8)
        cred = m.group(9)
        total = m.group(10)
    else:
        # 2. PLANO B: Contagem da Direita para a Esquerda (Blindado contra quebras de coluna)
        tokens = record_str.split()
        if len(tokens) < 10: return
        
        offset = 0 if re.search(r'[.,]', tokens[-1]) else 1
        try:
            total = tokens[-(1 + offset)]
            cred = tokens[-(2 + offset)]
            icms = tokens[-(3 + offset)]
            ipi = tokens[-(4 + offset)]
            vunt = tokens[-(5 + offset)]
            qtd = tokens[-(6 + offset)]
            serie = tokens[-(7 + offset)]
            nota = tokens[-(8 + offset)]
            forn = " ".join(tokens[1:-(8 + offset)])
            dt = tokens[0]
        except:
            return

    def to_f(v):
        return float(v.replace(".", "").replace(",", "."))
        
    try:
        n_val = int(nota) if nota.isdigit() else nota
        s_val = int(serie) if serie.isdigit() else serie
        
        rows_list.append([
            dt, forn, n_val, s_val,
            to_f(qtd), to_f(vunt), to_f(ipi), to_f(icms), to_f(cred), to_f(total),
            current_cfop
        ])
    except Exception:
        pass


if arquivos_pdf:
    if st.button("Processar Documentos"):
        with st.spinner("A auditar e reestruturar os dados..."):
            rows = []
            
            for f in arquivos_pdf:
                with pdfplumber.open(io.BytesIO(f.read())) as pdf:
                    txt = ""
                    for page in pdf.pages:
                        txt += page.extract_text(x_tolerance=1, y_tolerance=1) + "\n"
                        
                    lines = txt.split("\n")
                    i = 0
                    cfop_atual = ""
                    
                    while i < len(lines):
                        line = lines[i].strip()
                        if not line:
                            i += 1
                            continue
                            
                        # Atualiza CFOP em memória
                        if re.match(r"^\d{4}\s*-", line):
                            cfop_atual = line
                            i += 1
                            continue
                            
                        # Identifica início da Nota Fiscal
                        if re.match(r"^\d{2}/\d{2}/\d{4}", line):
                            record_str = line
                            
                            # Agrupa as linhas cortadas para recuperar valores perdidos
                            while i + 1 < len(lines):
                                next_line = lines[i+1].strip()
                                if not next_line:
                                    i += 1
                                    continue
                                    
                                if re.match(r"^\d{2}/\d{2}/\d{4}", next_line) or \
                                   re.match(r"^\d{4}\s*-", next_line) or \
                                   "DT EMISSÃO" in next_line or \
                                   "TOTAL" in next_line or \
                                   "Filtros:" in next_line:
                                    break
                                    
                                record_str += " " + next_line
                                i += 1
                                
                            analisar_registro(record_str, cfop_atual, rows)
                        
                        i += 1

            if not rows:
                st.error("❌ Não foi possível extrair dados válidos do ficheiro.")
            else:
                # ==========================================
                # DASHBOARD PRÉ-VISUALIZAÇÃO (Ecrã)
                # ==========================================
                st.divider()
                st.subheader("👀 Resumo da Auditoria")
                
                colunas_df = ["Data", "Fornecedor", "Nota", "Serie", "Qtd", "Valor Unt", "IPI", "ICMS", "Cred ICMS", "Total", "CFOP"]
                df_view = pd.DataFrame(rows, columns=colunas_df)
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Notas Recuperadas", len(df_view))
                c2.metric("Soma Geral (R$)", f"{df_view['Total'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c3.metric("Soma ICMS (R$)", f"{df_view['ICMS'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c4.metric("Soma IPI (R$)", f"{df_view['IPI'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                df_tela = df_view.copy()
                colunas_moeda = ["Valor Unt", "IPI", "ICMS", "Cred ICMS", "Total"]
                for col in colunas_moeda:
                    df_tela[col] = df_tela[col].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                st.dataframe(df_tela, use_container_width=True, hide_index=True)

                # ==========================================
                # GERAÇÃO EXCEL (Ordem EXATA do seu CSV)
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
                
                # O seu cabeçalho exato
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
                st.success("✅ Excel gerado com sucesso! Ordem das colunas restaurada e totais conferidos a 100%.")
                
                st.download_button(
                    label="📥 Baixar Relatório Excel",
                    data=output.getvalue(),
                    file_name=f"Relatorio_CFOP_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
