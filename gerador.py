"""
Gerador de Modelo de Frete
Informe o número de dias e os prazos das carteiras.
Todos os demais parâmetros são lidos direto do template Excel.
"""

import io
import datetime
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Gerador — Modelo de Frete", page_icon="🚛")

st.title("🚛 Gerador — Modelo de Frete")
st.divider()

col1, col2 = st.columns(2)

with col1:
    n_dias = st.number_input("Número de dias", min_value=50, max_value=1800, value=395, step=10)

with col2:
    d1 = st.number_input("Prazo carteira 1 (dias)", min_value=1, max_value=365, value=18, step=1)
    d2 = st.number_input("Prazo carteira 2 (dias)", min_value=1, max_value=365, value=20, step=1)
    d3 = st.number_input("Prazo carteira 3 (dias)", min_value=1, max_value=365, value=22, step=1)
    if not (d1 <= d2 <= d3):
        st.warning("Os prazos devem estar em ordem crescente (d1 ≤ d2 ≤ d3).")

st.divider()


# ════════════════════════════════════════════════════════════════════════════
# GERADOR EXCEL
# ════════════════════════════════════════════════════════════════════════════

def gerar_excel(template_bytes, n_dias, d1, d2, d3):
    N    = n_dias
    ORIG = N - max(d1, d2, d3) + 1   # último recebimento = ORIG + (d3-1) = N
    WIND = ORIG + 5
    LR   = N + 4
    START = datetime.datetime(2026, 6, 1)

    wb   = load_workbook(io.BytesIO(template_bytes))
    ws_f = wb["Fundo"]
    ws_c = wb["Custos"]
    ws_d = wb["Dashboard"]

    def sc(ws, row, col, val):
        cell = ws.cell(row, col)
        if not isinstance(cell, MergedCell):
            cell.value = val

    def clear_rows(ws, from_row):
        for r in range(from_row, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                sc(ws, r, col, None)

    # Labels dos prazos no Dashboard
    ws_d["H9"].value  = f"% {d1} dias"
    ws_d["H10"].value = f"% {d2} dias"
    ws_d["H11"].value = f"% {d3} dias"

    # Refs dinâmicas à última linha
    ws_d["E5"].value  = f"=Fundo!AD5+Fundo!AD{LR}"
    ws_d["E6"].value  = f"=Fundo!AJ5+Fundo!AJ{LR}"
    ws_d["E7"].value  = f"=Fundo!AN5+Fundo!AN{LR}"
    ws_d["C23"].value = f"=Fundo!U{LR}/Dashboard!D14"

    # Linha 4 e IRRs
    ws_f["U4"].value  = "=Z5+AF5+AL5"
    ws_f["X2"].value  = f"=IRR(X5:X{LR},0.001)"
    ws_f["AD2"].value = f"=IRR(AD5:AD{LR},0.001)"
    ws_f["AJ2"].value = f"=IRR(AJ5:AJ{LR},0.001)"
    ws_f["AN2"].value = f"=IRR(AN5:AN{LR},0.001)"

    clear_rows(ws_f, 5)
    clear_rows(ws_c, 5)

    yellow = PatternFill(patternType="solid", fgColor="FFFFFFCC")

    for n in range(1, N + 1):
        r     = n + 4
        pr    = r - 1
        first = (n == 1)
        last  = (n == N)
        wind  = (r == WIND)
        q_eim = (r >= WIND)

        sc(ws_f,r,2, "=Dashboard!D14" if first else f"=B{pr}*(((1+Dashboard!$C$20)^(1/30)))")
        sc(ws_f,r,3, n)
        sc(ws_f,r,4, START if first else f"=D{pr}+1")

        if first:
            sc(ws_f,r,5,"=Q5*Dashboard!$I$9"); sc(ws_f,r,9,"=Q5*Dashboard!$I$10"); sc(ws_f,r,13,"=Q5*Dashboard!$I$11")
        elif wind:
            sc(ws_f,r,5,0); sc(ws_f,r,9,0); sc(ws_f,r,13,0)
        else:
            sc(ws_f,r,5,f"=E{pr}"); sc(ws_f,r,9,f"=I{pr}"); sc(ws_f,r,13,f"=M{pr}")

        sc(ws_f,r,6,  f"=E{r}*Dashboard!$I$3")
        sc(ws_f,r,10, f"=I{r}*Dashboard!$I$3")
        sc(ws_f,r,14, f"=M{r}*Dashboard!$I$3")
        sc(ws_f,r,7,  0 if r < d1+4 else f"=E{r-(d1-1)}*Dashboard!$I$4")
        sc(ws_f,r,11, 0 if r < d2+4 else f"=I{r-(d2-1)}*Dashboard!$I$4")
        sc(ws_f,r,15, 0 if r < d3+4 else f"=M{r-(d3-1)}*Dashboard!$I$4")

        sc(ws_f,r,17, "=Dashboard!I8" if first else (f"=E{r}+I{r}+M{r}" if q_eim else f"=Q{pr}"))
        sc(ws_f,r,18, f"=F{r}+J{r}+N{r}")
        sc(ws_f,r,19, f"=G{r}+K{r}+O{r}")
        sc(ws_f,r,20, 0 if r < d1+4 else f"=S{r}-R{r-(d1-1)}")
        sc(ws_f,r,22, 0 if first else f"=U{pr}*(((1+Dashboard!$C$17)^(1/30))-1)")

        if first:   sc(ws_f,r,21, f"=U4-R{r}+S{r}+V{r}-Custos!D{r}")
        elif n==2:  sc(ws_f,r,21, f"=U{pr}-R{r}+S{r}+V{r}-Custos!D{r}-AB{pr}-AH{pr}-AM{pr}")
        else:       sc(ws_f,r,21, f"=U{pr}-R{r}+S{r}+V{r}-Custos!D{r}+Z{pr}+AF{pr}+AL{pr}-AB{pr}-AH{pr}-AM{pr}")

        sc(ws_f,r,23, f"=U{r}+R{r}")
        sc(ws_f,r,24, f"=-R{r}+S{r}-Custos!D{r}")

        sc(ws_f,r,26, "=Z4" if first else 0)
        sc(ws_f,r,27, f"=AC{pr}*(((1+Dashboard!$C$20)^(1/30))-1)")
        sc(ws_f,r,28, f"=AC{r}" if last else 0)
        sc(ws_f,r,29, f"=Z{r}" if first else f"=AC{pr}+AA{r}-AB{pr}+Z{r}")
        sc(ws_f,r,30, f"=-Z{r}+AB{r}")

        sc(ws_f,r,32, "=AF4" if first else 0)
        sc(ws_f,r,33, f"=AI{pr}*(((1+Dashboard!$C$21)^(1/30))-1)")
        sc(ws_f,r,34, f"=AI{r}" if last else 0)
        sc(ws_f,r,35, f"=AF{r}" if first else f"=AI{pr}+AG{r}-AH{pr}+AF{r}")
        sc(ws_f,r,36, f"=-AF{r}+AH{r}")

        sc(ws_f,r,38, "=AL4" if first else 0)
        sc(ws_f,r,39, f"=U{r}-AB{r}-AH{r}" if last else 0)
        sc(ws_f,r,40, f"=-AL{r}+AM{r}")

        if not last:
            for col in [26, 28, 32, 34, 38, 39]:
                cell = ws_f.cell(r, col)
                if not isinstance(cell, MergedCell):
                    cell.fill = yellow

        sc(ws_c,r,2, n)
        sc(ws_c,r,3, START if first else f"=C{pr}+1")
        sc(ws_c,r,5, "=Dashboard!$J$16/30" if first else f"=MAX(Dashboard!$J$16,(Dashboard!$I$16*Fundo!W{pr})/360)/30")
        sc(ws_c,r,6, "=Dashboard!$J$15/30" if first else f"=MAX(Dashboard!$J$15,(Dashboard!$I$15*Fundo!W{pr})/360)/30")
        sc(ws_c,r,7, 0 if first else f"=Fundo!R{pr}*Dashboard!$I$14")
        sc(ws_c,r,8, 0)
        sc(ws_c,r,9, 10000 if first else None)
        sc(ws_c,r,4, f"=E{r}+F{r}+G{r}+H{r}+I{r}")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# DOWNLOAD
# ════════════════════════════════════════════════════════════════════════════

try:
    with open("Modelagem_Frete_v5.xlsx", "rb") as f:
        template_bytes = f.read()
except FileNotFoundError:
    st.error("❌ Arquivo `Modelagem_Frete_v5.xlsx` não encontrado no repositório.")
    st.stop()

with st.spinner("Gerando Excel..."):
    excel_bytes = gerar_excel(template_bytes, n_dias, d1, d2, d3)

st.download_button(
    label=f"⬇️ Baixar Modelagem_Frete_{n_dias}d.xlsx",
    data=excel_bytes,
    file_name=f"Modelagem_Frete_{n_dias}d.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
