"""
Gerador de Modelo de Frete
Preencha os parâmetros e baixe o Excel gerado.
"""

import io
import datetime
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell import MergedCell

st.set_page_config(page_title="Gerador — Modelo de Frete", page_icon="🚛")

st.title("🚛 Gerador — Modelo de Frete")
st.markdown("Preencha os parâmetros e clique em **Gerar Excel**.")
st.divider()


# ════════════════════════════════════════════════════════════════════════════
# PARÂMETROS
# ════════════════════════════════════════════════════════════════════════════

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("Operação")
    n_dias = st.number_input("Número de dias", min_value=50, max_value=1800, value=395, step=10)
    pc     = st.number_input("Preço de compra (R$)", value=1678, step=10)
    pv     = st.number_input("Preço de venda (R$)",  value=1750, step=10)
    orig   = st.number_input("Originações por dia",  value=30,   step=1)
    st.markdown(f"Lucro por op: **R$ {pv - pc:,.0f}**")

with col2:
    st.subheader("Prazos e Estrutura")
    p18 = st.number_input("% 18 dias", min_value=0.0, max_value=1.0, value=0.30, step=0.05, format="%.2f")
    p20 = st.number_input("% 20 dias", min_value=0.0, max_value=1.0, value=0.30, step=0.05, format="%.2f")
    p22 = st.number_input("% 22 dias", min_value=0.0, max_value=1.0, value=0.40, step=0.05, format="%.2f")
    if abs(p18 + p20 + p22 - 1.0) > 0.01:
        st.warning(f"Soma dos prazos = {p18+p20+p22:.0%} (deve ser 100%)")
    st.markdown("---")
    pl     = st.number_input("PL Total (R$)", value=1_000_000, step=50_000)
    pct_sr = st.number_input("% Sênior",   min_value=0.0, max_value=1.0, value=0.60, step=0.05, format="%.2f")
    pct_mz = st.number_input("% Mezanino", min_value=0.0, max_value=1.0, value=0.10, step=0.05, format="%.2f")
    pct_jr = round(1 - pct_sr - pct_mz, 4)
    st.markdown(f"% Júnior (residual): **{pct_jr:.0%}**")

with col3:
    st.subheader("Taxas e Custos")
    cdi   = st.number_input("CDI a.a. (%)",       value=14.5, step=0.5, format="%.1f") / 100
    sr_sp = st.number_input("Sênior CDI+ (%)",    value=3.0,  step=0.5, format="%.1f") / 100
    mz_sp = st.number_input("Mezanino CDI+ (%)",  value=5.0,  step=0.5, format="%.1f") / 100
    st.markdown("---")
    cons   = st.number_input("Consultoria (% desembolso)", value=1.0, step=0.1, format="%.1f") / 100
    gest   = st.number_input("Gestão a.a. (%)",   value=0.5,  step=0.1, format="%.1f") / 100
    piso_g = st.number_input("Piso Gestão (R$/mês)", value=15000, step=1000)
    adm    = st.number_input("Adm a.a. (%)",      value=0.2,  step=0.1, format="%.1f") / 100
    piso_a = st.number_input("Piso Adm (R$/mês)",    value=15000, step=1000)

st.divider()


# ════════════════════════════════════════════════════════════════════════════
# GERADOR EXCEL
# ════════════════════════════════════════════════════════════════════════════

def gerar_excel(template_bytes, p):
    N     = p["n_dias"]
    ORIG  = N - 21
    WIND  = ORIG + 5
    LR    = N + 4
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

    # Hipóteses
    ws_d["I3"].value  = p["pc"];     ws_d["I4"].value  = p["pv"]
    ws_d["I8"].value  = p["orig"];   ws_d["I9"].value  = p["p18"]
    ws_d["I10"].value = p["p20"];    ws_d["I11"].value = p["p22"]
    ws_d["D14"].value = p["pl"]
    ws_d["C10"].value = p["pct_sr"]; ws_d["C11"].value = p["pct_mz"]
    ws_d["C12"].value = round(1 - p["pct_sr"] - p["pct_mz"], 4)
    ws_d["D17"].value = p["cdi"];    ws_d["D18"].value = p["sr_sp"]
    ws_d["D19"].value = p["mz_sp"]
    ws_d["I14"].value = p["cons"];   ws_d["I15"].value = p["gest"]
    ws_d["J15"].value = p["piso_g"]; ws_d["I16"].value = p["adm"]
    ws_d["J16"].value = p["piso_a"]

    # Dashboard dinâmico
    ws_d["C4"].value  = f"=((Fundo!U{LR}/Dashboard!D14)^(365/{N}))-1"
    ws_d["D4"].value  = "=((1+C4)^12)-1"
    ws_d["E5"].value  = f"=Fundo!AD5+Fundo!AD{LR}"
    ws_d["E6"].value  = f"=Fundo!AJ5+Fundo!AJ{LR}"
    ws_d["E7"].value  = f"=Fundo!AN5+Fundo!AN{LR}"
    ws_d["C23"].value = f"=Fundo!U{LR}/Dashboard!D14"

    # IRRs
    ws_f["X2"].value  = f"=IRR(X5:X{LR},0.001)"
    ws_f["AD2"].value = f"=IRR(AD5:AD{LR},0.001)"
    ws_f["AJ2"].value = f"=IRR(AJ5:AJ{LR},0.001)"
    ws_f["AN2"].value = f"=IRR(AN5:AN{LR})"

    clear_rows(ws_f, 5)
    clear_rows(ws_c, 5)

    for n in range(1, N + 1):
        r  = n + 4;  pr = r - 1
        first = (n == 1);  last = (n == N)
        wind  = (r == WIND);  q_eim = (r >= WIND)

        sc(ws_f,r,2, "=Dashboard!D14" if first else f"=B{pr}*(((1+Dashboard!$C$20)^(1/30)))")
        sc(ws_f,r,3, n)
        sc(ws_f,r,4, START if first else f"=D{pr}+1")

        if first:
            sc(ws_f,r,5,"=Q5*Dashboard!$I$9"); sc(ws_f,r,9,"=Q5*Dashboard!$I$10"); sc(ws_f,r,13,"=Q5*Dashboard!$I$11")
        elif wind:
            sc(ws_f,r,5,0); sc(ws_f,r,9,0); sc(ws_f,r,13,0)
        else:
            sc(ws_f,r,5,f"=E{pr}"); sc(ws_f,r,9,f"=I{pr}"); sc(ws_f,r,13,f"=M{pr}")

        sc(ws_f,r,6, f"=E{r}*Dashboard!$I$3"); sc(ws_f,r,10,f"=I{r}*Dashboard!$I$3"); sc(ws_f,r,14,f"=M{r}*Dashboard!$I$3")
        sc(ws_f,r,7,  0 if r<22 else f"=E{r-17}*Dashboard!$I$4")
        sc(ws_f,r,11, 0 if r<24 else f"=I{r-19}*Dashboard!$I$4")
        sc(ws_f,r,15, 0 if r<26 else f"=M{r-21}*Dashboard!$I$4")

        sc(ws_f,r,17,"=Dashboard!I8" if first else (f"=E{r}+I{r}+M{r}" if q_eim else f"=Q{pr}"))
        sc(ws_f,r,18,f"=F{r}+J{r}+N{r}"); sc(ws_f,r,19,f"=G{r}+K{r}+O{r}")
        sc(ws_f,r,20,0 if r<22 else f"=S{r}-R{r-17}")
        sc(ws_f,r,22,0 if first else f"=U{pr}*(((1+Dashboard!$C$17)^(1/30))-1)")

        if first:  sc(ws_f,r,21,f"=U4-R{r}+S{r}+V{r}-Custos!D{r}")
        elif n==2: sc(ws_f,r,21,f"=U{pr}-R{r}+S{r}+V{r}-Custos!D{r}-AB{pr}-AH{pr}-AM{pr}")
        else:      sc(ws_f,r,21,f"=U{pr}-R{r}+S{r}+V{r}-Custos!D{r}+Z{pr}+AF{pr}+AL{pr}-AB{pr}-AH{pr}-AM{pr}")

        sc(ws_f,r,23,f"=U{r}+R{r}"); sc(ws_f,r,24,f"=-R{r}+S{r}-Custos!D{r}")

        sc(ws_f,r,26,"=Z4" if first else 0)
        sc(ws_f,r,27,f"=AC{pr}*(((1+Dashboard!$C$20)^(1/30))-1)")
        sc(ws_f,r,28,f"=AC{r}" if last else 0)
        sc(ws_f,r,29,f"=Z{r}" if first else f"=AC{pr}+AA{r}-AB{pr}")
        sc(ws_f,r,30,f"=-Z{r}+AB{r}")

        sc(ws_f,r,32,"=AF4" if first else 0)
        sc(ws_f,r,33,f"=AI{pr}*(((1+Dashboard!$C$21)^(1/30))-1)")
        sc(ws_f,r,34,f"=AI{r}" if last else 0)
        sc(ws_f,r,35,f"=AF{r}" if first else f"=AI{pr}+AG{r}-AH{pr}")
        sc(ws_f,r,36,f"=-AF{r}+AH{r}")

        sc(ws_f,r,38,"=AL4" if first else 0)
        sc(ws_f,r,39,f"=U{r}-AB{r}-AH{r}" if last else 0)
        sc(ws_f,r,40,f"=-AL{r}+AM{r}")

        sc(ws_c,r,2,n)
        sc(ws_c,r,3,START if first else f"=C{pr}+1")
        sc(ws_c,r,5,"=Dashboard!$J$16/30" if first else f"=MAX(Dashboard!$J$16,(Dashboard!$I$16*Fundo!W{pr})/360)/30")
        sc(ws_c,r,6,"=Dashboard!$J$15/30" if first else f"=MAX(Dashboard!$J$15,(Dashboard!$I$15*Fundo!W{pr})/360)/30")
        sc(ws_c,r,7,0 if first else f"=Fundo!R{pr}*Dashboard!$I$14")
        sc(ws_c,r,8,0)
        sc(ws_c,r,9,10000 if first else None)
        sc(ws_c,r,4,f"=E{r}+F{r}+G{r}+H{r}+I{r}")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# BOTÃO GERAR
# ════════════════════════════════════════════════════════════════════════════

params = dict(
    n_dias=n_dias, pc=pc, pv=pv, orig=orig,
    p18=p18, p20=p20, p22=p22,
    pl=pl, pct_sr=pct_sr, pct_mz=pct_mz,
    cdi=cdi, sr_sp=sr_sp, mz_sp=mz_sp,
    cons=cons, gest=gest, piso_g=piso_g,
    adm=adm, piso_a=piso_a,
)

try:
    with open("Modelagem_Frete_v5.xlsx", "rb") as f:
        template_bytes = f.read()
except FileNotFoundError:
    st.error("❌ Arquivo `Modelagem_Frete_v5.xlsx` não encontrado no repositório.")
    st.stop()

with st.spinner("Gerando Excel..."):
    excel_bytes = gerar_excel(template_bytes, params)

st.download_button(
    label=f"⬇️ Baixar Modelagem_Frete_{n_dias}d.xlsx",
    data=excel_bytes,
    file_name=f"Modelagem_Frete_{n_dias}d.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
