
import pdfplumber, pandas as pd, re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

PASTA = Path(__file__).parent
PDFS = list(PASTA.glob("*.pdf")) + list(PASTA.glob("*.PDF"))

for PDF_PATH in PDFS:
    rows = []
    cfop = ""
    regex = re.compile(r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(\d+)\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+(\d+)")

    with pdfplumber.open(PDF_PATH) as pdf:
        for page in pdf.pages:
            txt = page.extract_text(x_tolerance=1, y_tolerance=1) or ""
            for line in txt.split("\n"):
                line=line.strip()
                if re.match(r"\d{4}\s*-\s*", line):
                    cfop = line
                    continue
                m = regex.match(line)
                if m:
                    dt, forn, nota, serie, qtd, vunt, ipi, icms, cred, total, prz = m.groups()
                    rows.append([dt, forn, int(nota), int(serie),
                        float(qtd.replace(".","").replace(",",".")),
                        float(vunt.replace(".","").replace(",",".")),
                        float(ipi.replace(".","").replace(",",".")),
                        float(icms.replace(".","").replace(",",".")),
                        float(cred.replace(".","").replace(",",".")),
                        float(total.replace(".","").replace(",",".")),
                        cfop])

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
    ws['A1'] = "RELATÓRIO COMPRAS POR CFOP - CONTTEC"
    ws.merge_cells('A1:K1')
    ws['A1'].font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    ws['A1'].fill = PatternFill(start_color=AZUL_ESCURO, end_color=AZUL_ESCURO, fill_type="solid")
    ws['A1'].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 20
    
    # LINHA 2 - vazia
    ws.row_dimensions[2].height = 15
    
    # LINHA 3 - TOTAIS (colunas F a J)
    ws.row_dimensions[3].height = 18
    # sem bordas, sem fundo nas celulas vazias
    
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
            # SEM BORDAS, SEM FUNDO
            cell.border = Border()
            cell.fill = PatternFill(fill_type=None)
            if c_idx == 2 or c_idx == 11:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if 5 <= c_idx <= 10:
                cell.number_format = '#,##0.00'
    
    last = 4 + len(rows)
    
    # APLICAR TOTAIS COM CORES
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
    from openpyxl.utils import get_column_letter
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    # FILTRO
    ws.auto_filter.ref = f"A4:K{last}"
    ws.freeze_panes = "A5"
    
    # REMOVER LINHAS DE GRADE DO EXCEL
    ws.sheet_view.showGridLines = False
    
    wb.save(PDF_PATH.with_suffix('.xlsx'))
    print(f"Gerado: {PDF_PATH.stem}.xlsx")

print("Concluido")
