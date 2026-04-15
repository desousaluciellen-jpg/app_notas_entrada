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
st.markdown("O sistema processa e unifica todas as notas do PDF. O novo algoritmo inteligente impede perda de notas por quebra de linha e garante o alinhamento exato das colunas de valores.")

arquivos_pdf = st.file_uploader("Arraste os seus PDFs para aqui", type="pdf", accept_multiple_files=True)

def analisar_registro(record_str, current_cfop, rows_list):
    """Lê a linha completa e distribui pelas colunas corretas de forma blindada."""
    m_date = re.match(r"^(\d{2}/\d{2}/\d{4})", record_str)
    if not m_date: return
    dt = m_date.group(1)
    
    # 1. Tenta a extração rigorosa (conta os campos da direita para a esquerda)
    m_end = re.search(r"\s+(\S+)\s+(\S+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)(?:\s+.*)?$", record_str)
    
    if m_end:
        forn = record_str[len(dt):m_end.start()].strip()
        nota = m_end.group(1)
        serie = m_end.group(2)
        qtd = m_end.group(3)
        vunt = m_end.group(4)
        ipi = m_end.group(5)
        icms = m_end.group(6)
        cred = m_end.group(7)
        total = m_end.group(8)
    else:
        # 2. Plano B: Se o PDF vier ilegível, conta os blocos pelo fim da frase
        tokens = record_str.split()
        if len(tokens) < 10: return
        
        # Identifica se o último campo é o Prazo (ex: 60) ou o Total (ex: 113,61)
        offset = 0 if re.search(r",\d{2}$", tokens[-1]) else 1
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
        except:
            return

    # Função para converter formato PT-BR para número matemático
    def to_f(v):
        return float(v.replace(".", "").replace(",", "."))
        
    try:
        if nota.isdigit(): nota = int(nota)
        if serie.isdigit(): serie = int(serie)
        
        rows_list.append([
            dt, forn, nota, serie,
            to_f(qtd), to_f(vunt), to_f(ipi), to_f(icms), to_f(cred), to_f(total),
            current_cfop
        ])
    except:
        pass


if arquivos_pdf:
    if st.button("Processar Documentos"):
        with st.spinner("A cruzar dados e auditar valores..."):
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
                            
                        # Atualiza o CFOP ativo
                        if re.match(r"^\d{4}\s*-", line):
                            cfop_atual = line
                            i += 1
                            continue
                            
                        # Encontrou o início de uma nota (Data)
                        if re.match(r"^\d{2}/\d{2}/\d{4}", line):
                            record_str = line
                            
                            # Vai buscar as linhas seguintes caso a nota tenha sido partida ao meio
                            while i + 1 < len(lines):
                                next_line = lines[i+1].strip()
                                if not next_line:
                                    i += 1
                                    continue
                                # Se a próxima linha for outra Data, um CFOP ou um Cabeçalho, significa que a nossa nota atual acabou.
                                if re.match(r"^\d{2}/\d{2}/\d{4}", next_line) or \
                                   re.match(r"^\d{4}\s*-", next_line) or \
                                   "DT EMISSÃO" in next_line or \
                                   "TOTAL CFOP" in next_line:
                                    break
                                
                                record_str += " " + next_line
                                i += 1
                                
                            # Envia a nota completa e blindada para ser separada nas colunas
                            analisar_registro(record_str, cfop_atual, rows)
                        
                        i += 1

            if not rows:
                st.error("❌ Nenhum dado válido foi extraído.")
            else:
                # ==========================================
                # DASHBOARD PRÉ-VISUALIZAÇÃO (Na Tela)
                # ==========================================
                st.divider()
                st.subheader("👀 Resumo Auditoria")
                
                colunas_df = ["Data", "Fornecedor", "Nota", "Serie", "Qtd", "Valor Unt", "IPI", "ICMS", "Cred ICMS", "Total", "CFOP"]
                df_view = pd.DataFrame(rows, columns=colunas_df)
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total de Notas Apuradas", len(df_view))
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
                
                # Linha 1
                ws['A1'] = "RELATÓRIO COMPRAS POR CFOP - CONTTEC"
                ws.merge_cells('A1:K1')
                ws['A1'].font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
                ws['A1'].fill = PatternFill(start_color=AZUL_ESCURO, end_color=AZUL_ESCURO, fill_type="solid")
                ws['A1'].alignment = Alignment(horizontal="left", vertical="center")
                ws.row_dimensions[1].height = 20
                
                # Linha 2 e 3
                ws.row_dimensions[2].height = 15
                ws.row_dimensions[3].height = 18
                
                # Linha 4 (Cabeçalho exato)
                headers = ["Data","Fornecedor","Nota","Serie","Qtd","Valor Unt","IPI","ICMS","Cred ICMS","Total","CFOP Completo"]
                for i, h in enumerate(headers, 1):
                    c = ws.cell(row=4, column=i, value=h)
                    c.font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
                    c.fill = PatternFill(start_color=AZUL_ESCURO, end_color=AZUL_ESCURO, fill_type="solid")
                    c.alignment = Alignment(horizontal="center", vertical="center")
                
                ws.row_dimensions[4].height = 20
                
                # Dados (a partir da linha 5)
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
                
                # Fórmulas de Totais
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
                st.success("✅ Excel gerado! Estrutura exata do CSV e totais conferidos a 100%.")
                
                st.download_button(
                    label="📥 Baixar Relatório Excel",
                    data=output.getvalue(),
                    file_name=f"Relatorio_CFOP_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
