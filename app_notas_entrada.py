import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Relatório CFOP - CONTTEC", page_icon="🧾", layout="wide")
st.title("🧾 Gerador de Relatório: Compras por CFOP")
st.markdown("Sistema com paleta **Corporate Trust** aplicada.")

arquivos_pdf = st.file_uploader("Arraste os seus PDFs", type="pdf", accept_multiple_files=True)

def to_f(v):
    try: return float(v.replace(".","").replace(",","."))
    except: return 0.0

if arquivos_pdf and st.button("Processar Documentos"):
    with st.spinner("Extraindo..."):
        rows = []
        for f in arquivos_pdf:
            with pdfplumber.open(io.BytesIO(f.read())) as pdf:
                full_text = "\n".join([p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages])
                lines = full_text.split("\n")
                i, current_cfop_code, current_cfop_desc = 0, "", ""
                while i < len(lines):
                    line = lines[i].strip()
                    if not line: i+=1; continue
                    m_cfop = re.match(r"^(\d{4})\s*-\s*(.*)", line)
                    if m_cfop:
                        current_cfop_code = m_cfop.group(1)
                        desc = m_cfop.group(2).strip()
                        if i+1 < len(lines) and not re.match(r"^\d", lines[i+1].strip()):
                            desc += " " + lines[i+1].strip(); i+=1
                        current_cfop_desc = desc; i+=1; continue
                    if re.match(r"^\d{2}/\d{2}/\d{4}", line):
                        record_str = line
                        while i+1 < len(lines):
                            nxt = lines[i+1].strip()
                            if re.match(r"^\d{2}/\d{4}", nxt) or re.match(r"^\d{4}\s*-", nxt) or "TOTAL" in nxt: break
                            if nxt: record_str += " " + nxt
                            i+=1
                        tokens = record_str.split()
                        if len(tokens) >= 10:
                            if re.match(r"^\d+$", tokens[-1]): tokens.pop()
                            total, cred, icms, ipi, v_unt, qtd, serie, nota = [tokens.pop() for _ in range(8)]
                            date, forn = tokens[0], " ".join(tokens[1:])
                            rows.append({"Data":date,"Fornecedor":forn,"Nota":nota,"Série":serie,
                                         "CFOP":int(current_cfop_code) if current_cfop_code.isdigit() else current_cfop_code,
                                         "Qtd":to_f(qtd),"Vlr Unit":to_f(v_unt),"IPI":to_f(ipi),
                                         "Cred ICMS":to_f(cred),"ICMS":to_f(icms),"Total":to_f(total),
                                         "Descrição CFOP":current_cfop_desc})
                    i+=1

        if rows:
            df = pd.DataFrame(rows)

            # PALETA CORPORATE TRUST
            MARINHO = "1B263B"; ACO = "415A77"; GELO = "E0E1DD"; GRAFITE = "262626"; BRANCO = "FFFFFF"
            F_TITLE = Font(bold=True, size=14, color=MARINHO, name="Calibri")
            F_HEADER = Font(bold=True, color=BRANCO, name="Calibri")
            F_BOLD = Font(bold=True, name="Calibri")
            F_DATA = Font(color=GRAFITE, name="Calibri", size=10)
            BG_MARINHO = PatternFill(start_color=MARINHO, end_color=MARINHO, fill_type="solid")
            BG_GELO = PatternFill(start_color=GELO, end_color=GELO, fill_type="solid")
            ALIGN_C = Alignment(horizontal="center", vertical="center")

            wb = Workbook()
            ws_dash = wb.active; ws_dash.title = "Dashboard"; ws_dash.sheet_view.showGridLines = False

            # Dashboard
            ws_dash['A1'] = "RELATÓRIO DE ENTRADAS POR CFOP"; ws_dash['A1'].font = Font(size=24, bold=True, color=MARINHO, name="Calibri")
            ws_dash['A3'] = "CONTTEC COM E SERV ESP EM MOTORES ELETRICOS IND LTDA"; ws_dash['A3'].font = Font(color=GRAFITE, name="Calibri")

            for idx, (tit, col) in enumerate([("TOTAL GERAL",'B'),("TOTAL ICMS",'D'),("TOTAL IPI",'F'),("NOTAS",'H')]):
                ws_dash[f'{col}6'] = tit; ws_dash[f'{col}6'].font = F_HEADER; ws_dash[f'{col}6'].fill = BG_MARINHO; ws_dash[f'{col}6'].alignment = ALIGN_C
                ws_dash.merge_cells(f'{col}6:{chr(ord(col)+1)}6')
                val = [df['Total'].sum(), df['ICMS'].sum(), df['IPI'].sum(), len(df)][idx]
                c = ws_dash[f'{col}8']; c.value = val; c.font = Font(size=16, bold=True, color=MARINHO, name="Calibri"); c.fill = BG_GELO; c.alignment = ALIGN_C
                ws_dash.merge_cells(f'{col}8:{chr(ord(col)+1)}8')
                if idx<3: c.number_format = '"R$" #,##0.00'

            # Resumo
            ws_dash['A12'] = "RESUMO POR CFOP"; ws_dash['A12'].font = Font(bold=True, size=12, color=ACO, name="Calibri")
            for i,h in enumerate(["CFOP","Descrição","Qtde Itens","Total R$","ICMS R$","IPI R$","% do Total"],1):
                c = ws_dash.cell(14,i,h); c.font=F_HEADER; c.fill=BG_MARINHO

            resumo = df.groupby(['CFOP','Descrição CFOP']).agg(Qtde=('Nota','count'),Total=('Total','sum'),ICMS=('ICMS','sum'),IPI=('IPI','sum')).reset_index()
            resumo['Perc'] = resumo['Total']/resumo['Total'].sum()
            for r_idx, r in enumerate(resumo.itertuples(),15):
                for c_idx, val in enumerate([r.CFOP, r[2], r.Qtde, r.Total, r.ICMS, r.IPI, r.Perc],1):
                    c = ws_dash.cell(r_idx, c_idx, val); c.font = F_DATA
                    if c_idx>=4 and c_idx<=6: c.number_format = '"R$" #,##0.00'
                    if c_idx==7: c.number_format = '0.00%'
                    if r_idx%2==0: c.fill = BG_GELO

            # Dados Filtráveis
            ws_data = wb.create_sheet("Dados Filtráveis")
            ws_data['A1'] = "DADOS COM TOTALIZADOR DINÂMICO"; ws_data['A1'].font = Font(size=16, bold=True, color=BRANCO, name="Calibri"); ws_data['A1'].fill = BG_MARINHO
            ws_data.merge_cells('A1:L1')
            ws_data['A2'] = "Use os filtros - totais atualizam automaticamente"; ws_data['A2'].fill = BG_GELO; ws_data['A2'].font = Font(italic=True, color=ACO, name="Calibri"); ws_data.merge_cells('A2:L2')

            last = len(df)+5
            ws_data['A4']="TOTAIS FILTRADOS:"; ws_data['A4'].font = F_BOLD
            for label, col, form in [("Itens:",'E',f'=SUBTOTAL(103,C6:C{last})'),("IPI:",'H',f'=SUBTOTAL(109,H6:H{last})'),("ICMS:",'J',f'=SUBTOTAL(109,J6:J{last})'),("TOTAL:",'L',f'=SUBTOTAL(109,K6:K{last})')]:
                ws_data[f'{chr(ord(col)-1)}4']=label
                c=ws_data[f'{col}4']; c.value=form; c.font=F_BOLD; c.fill=BG_GELO
                if col!='E': c.number_format='"R$" #,##0.00'

            headers = ["Data","Fornecedor","Nota","Série","CFOP","Qtd","Vlr Unit","IPI","Cred ICMS","ICMS","Total","Descrição CFOP"]
            for i,h in enumerate(headers,1):
                c=ws_data.cell(5,i,h); c.font=F_HEADER; c.fill=BG_MARINHO; c.alignment=ALIGN_C

            for r_idx, row in enumerate(df.itertuples(index=False),6):
                for c_idx, val in enumerate(row[:12],1):
                    c = ws_data.cell(r_idx, c_idx, val); c.font = F_DATA
                    if c_idx==5: c.font = Font(bold=True, color=ACO, name="Calibri", size=11); c.alignment=ALIGN_C
                    if c_idx in [1,3,4,5,6]: c.alignment=ALIGN_C
                    if c_idx>=7: c.number_format='#,##0.00'

            ws_data.auto_filter.ref = f"A5:L{last}"; ws_data.freeze_panes = "A6"

            output = io.BytesIO(); wb.save(output)
            st.success(f"✅ {len(df)} notas processadas com paleta Corporate Trust")
            st.download_button("📥 Baixar Excel", output.getvalue(), f"Relatorio_CFOP_{datetime.now():%Y%m%d_%H%M}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
