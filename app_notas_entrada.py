import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# Configuração da Página
st.set_page_config(page_title="Portal de Notas - Entrada", page_icon="📝", layout="wide")

st.title("📝 Consolidador de Notas de Entrada (CFOP)")
st.markdown("Carregue os relatórios PDF da CONTTEC para consolidar as entradas, calcular impostos e gerar o Excel profissional.")

# Upload de Ficheiros
ficheiros_pdf = st.file_uploader("Selecione os PDFs de Notas de Entrada", type="pdf", accept_multiple_files=True)

if ficheiros_pdf:
    if st.button("Processar e Consolidar"):
        with st.spinner("A extrair dados dos documentos..."):
            rows = []
            cfop_atual = ""
            # Expressão regular para capturar a linha da nota fiscal
            regex = re.compile(r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(\d+)\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+(\d+)")

            for f in ficheiros_pdf:
                with pdfplumber.open(io.BytesIO(f.read())) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text(x_tolerance=1, y_tolerance=1) or ""
                        for line in text.split("\n"):
                            line = line.strip()
                            # Identifica mudança de CFOP
                            if re.match(r"\d{4}\s*-\s*", line):
                                cfop_atual = line
                                continue
                            
                            m = regex.match(line)
                            if m:
                                dt, forn, nota, serie, qtd, vunt, ipi, icms, cred, total, prz = m.groups()
                                rows.append({
                                    "CFOP": cfop_atual,
                                    "DATA EMISSÃO": dt,
                                    "FORNECEDOR": forn,
                                    "NOTA": int(nota),
                                    "SÉRIE": int(serie),
                                    "QTD": float(qtd.replace(".","").replace(",",".")),
                                    "VLR UNT": float(vunt.replace(".","").replace(",",".")),
                                    "IPI": float(ipi.replace(".","").replace(",",".")),
                                    "ICMS": float(icms.replace(".","").replace(",",".")),
                                    "CRED ICMS": float(cred.replace(".","").replace(",",".")),
                                    "TOTAL": float(total.replace(".","").replace(",",".")),
                                    "PRZ MED": int(prz),
                                    "ORIGEM": f.name
                                })

            if not rows:
                st.error("Não foram encontrados dados compatíveis nos PDFs enviados.")
            else:
                df = pd.DataFrame(rows)

                # --- PRÉ-VISUALIZAÇÃO NO ECRÃ ---
                st.divider()
                st.subheader("📊 Resumo Consolidado")
                
                c1, c2, c3, c4 = st.columns(4)
                vlr_total = df["TOTAL"].sum()
                ipi_total = df["IPI"].sum()
                icms_total = df["ICMS"].sum()
                
                c1.metric("Total de Notas", len(df))
                c2.metric("Valor Total", f"R$ {vlr_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c3.metric("Total IPI", f"R$ {ipi_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                c4.metric("Total ICMS", f"R$ {icms_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

                # Tabela Interativa
                st.dataframe(df, use_container_width=True, hide_index=True)

                # --- GERAÇÃO DO EXCEL FORMATADO ---
                output = io.BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Notas de Entrada"
                
                # Cabeçalho da Tabela
                headers = list(df.columns)
                ws.append(headers)
                
                for r in df.values.tolist():
                    ws.append(r)

                # Formatação Profissional (Cores baseadas no seu padrão)
                last_row = len(df) + 1
                last_col = get_column_letter(len(headers))
                
                tab = Table(displayName="TabelaNotas", ref=f"A1:{last_col}{last_row}")
                tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight15", showRowStripes=True)
                ws.add_table(tab)
                
                # Formato Moeda
                for col_num in range(6, 12): # Colunas de valores
                    letra = get_column_letter(col_num)
                    for cell in ws[letra]:
                        if cell.row > 1:
                            cell.number_format = '#,##0.00'

                wb.save(output)
                
                st.download_button(
                    label="📥 Descarregar Folha de Cálculo Excel",
                    data=output.getvalue(),
                    file_name=f"Consolidado_Entradas_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )