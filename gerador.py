"""
Dashboard — Modelagem de Frete
Suba este arquivo + Modelagem_Frete_v4.xlsx + requirements.txt no GitHub
e faça deploy no share.streamlit.io
"""

import io
import datetime
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from openpyxl import load_workbook
from openpyxl.cell import MergedCell

# ── Configuração da página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Modelo de Frete",
    page_icon="🚛",
    layout="wide",
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

    html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

    h1, h2, h3 { font-family: 'IBM Plex Mono', monospace; }

    .metric-box {
        background: #0f1117;
        border: 1px solid #2a2d3a;
        border-left: 3px solid #00d4aa;
        border-radius: 4px;
        padding: 16px 20px;
        margin-bottom: 8px;
    }
    .metric-label { color: #8b8fa8; font-size: 11px; text-transform: uppercase; letter-spacing: 1px; }
    .metric-value { color: #ffffff; font-size: 24px; font-family: 'IBM Plex Mono', monospace; font-weight: 600; }
    .metric-sub   { color: #8b8fa8; font-size: 12px; margin-top: 2px; }

    .stSidebar { background-color: #0a0c12; }
    [data-testid="stSidebar"] label { color: #c0c4d6 !important; font-size: 13px; }
    [data-testid="stSidebar"] .stNumberInput label { color: #c0c4d6 !important; }

    .section-header {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 11px;
        font-weight: 600;
        color: #00d4aa;
        letter-spacing: 2px;
        text-transform: uppercase;
        border-bottom: 1px solid #2a2d3a;
        padding-bottom: 6px;
        margin: 20px 0 12px 0;
    }
    div[data-testid="stDownloadButton"] button {
        background: #00d4aa;
        color: #0a0c12;
        font-family: 'IBM Plex Mono', monospace;
        font-weight: 600;
        border: none;
        border-radius: 4px;
        padding: 10px 24px;
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# GERADOR DO MODELO (integrado ao app)
# ════════════════════════════════════════════════════════════════════════════

def gerar_modelo(template_bytes, params):
    N  = params["n_dias"]
    ORIG = N - 21
    WIND = ORIG + 5
    LR   = N + 4
    START_DATE = datetime.datetime(2026, 6, 1)

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

    # ── Hipóteses do Dashboard ────────────────────────────────────────────────
    ws_d["I3"].value  = params["media_compra"]
    ws_d["I4"].value  = params["media_venda"]
    ws_d["I8"].value  = params["orig_dia"]
    ws_d["I9"].value  = params["pct_18"]
    ws_d["I10"].value = params["pct_20"]
    ws_d["I11"].value = params["pct_22"]
    ws_d["D14"].value = params["pl_total"]
    ws_d["C10"].value = params["pct_senior"]
    ws_d["C11"].value = params["pct_meza"]
    ws_d["C12"].value = params["pct_jr"]
    ws_d["D17"].value = params["cdi"]
    ws_d["D18"].value = params["sr_spread"]
    ws_d["D19"].value = params["meza_spread"]
    ws_d["I14"].value = params["taxa_consultoria"]
    ws_d["I15"].value = params["taxa_gestao"]
    ws_d["J15"].value = params["piso_gestao"]
    ws_d["I16"].value = params["taxa_adm"]
    ws_d["J16"].value = params["piso_adm"]
    ws_d["I17"].value = params["taxa_perf"]

    # ── Dashboard: refs dinâmicas ─────────────────────────────────────────────
    ws_d["E5"].value  = f"=Fundo!AD5+Fundo!AD{LR}"
    ws_d["E6"].value  = f"=Fundo!AJ5+Fundo!AJ{LR}"
    ws_d["E7"].value  = f"=Fundo!AN5+Fundo!AN{LR}"
    ws_d["C23"].value = f"=Fundo!U{LR}/Dashboard!D14"

    # ── IRRs ─────────────────────────────────────────────────────────────────
    ws_f["X2"].value  = f"=IRR(X5:X{LR},0.001)"
    ws_f["AD2"].value = f"=IRR(AD5:AD{LR})"
    ws_f["AJ2"].value = f"=IRR(AJ5:AJ{LR},0.001)"
    ws_f["AN2"].value = f"=IRR(AN5:AN{LR})"

    clear_rows(ws_f, 5)
    clear_rows(ws_c, 5)

    for n in range(1, N + 1):
        r  = n + 4
        pr = r - 1
        first = (n == 1)
        last  = (n == N)
        wind  = (r == WIND)
        q_eim = (r >= WIND)

        sc(ws_f, r, 2,  "=Dashboard!D14" if first else f"=B{pr}*(((1+Dashboard!$C$20)^(1/30)))")
        sc(ws_f, r, 3,  n)
        sc(ws_f, r, 4,  START_DATE if first else f"=D{pr}+1")

        if first:
            sc(ws_f, r, 5,  "=Q5*Dashboard!$I$9")
            sc(ws_f, r, 9,  "=Q5*Dashboard!$I$10")
            sc(ws_f, r, 13, "=Q5*Dashboard!$I$11")
        elif wind:
            sc(ws_f, r, 5, 0); sc(ws_f, r, 9, 0); sc(ws_f, r, 13, 0)
        else:
            sc(ws_f, r, 5,  f"=E{pr}")
            sc(ws_f, r, 9,  f"=I{pr}")
            sc(ws_f, r, 13, f"=M{pr}")

        sc(ws_f, r, 6,  f"=E{r}*Dashboard!$I$3")
        sc(ws_f, r, 10, f"=I{r}*Dashboard!$I$3")
        sc(ws_f, r, 14, f"=M{r}*Dashboard!$I$3")

        sc(ws_f, r, 7,  0 if r < 22 else f"=E{r-17}*Dashboard!$I$4")
        sc(ws_f, r, 11, 0 if r < 24 else f"=I{r-19}*Dashboard!$I$4")
        sc(ws_f, r, 15, 0 if r < 26 else f"=M{r-21}*Dashboard!$I$4")

        sc(ws_f, r, 17, "=Dashboard!I8" if first else (f"=E{r}+I{r}+M{r}" if q_eim else f"=Q{pr}"))
        sc(ws_f, r, 18, f"=F{r}+J{r}+N{r}")
        sc(ws_f, r, 19, f"=G{r}+K{r}+O{r}")
        sc(ws_f, r, 20, 0 if r < 22 else f"=S{r}-R{r-17}")
        sc(ws_f, r, 22, 0 if first else f"=U{pr}*(((1+Dashboard!$C$17)^(1/30))-1)")

        if first:
            sc(ws_f, r, 21, f"=U4-R{r}+S{r}+V{r}-Custos!D{r}")
        elif n == 2:
            sc(ws_f, r, 21, f"=U{pr}-R{r}+S{r}+V{r}-Custos!D{r}-AB{pr}-AH{pr}-AM{pr}")
        else:
            sc(ws_f, r, 21, f"=U{pr}-R{r}+S{r}+V{r}-Custos!D{r}+Z{pr}+AF{pr}+AL{pr}-AB{pr}-AH{pr}-AM{pr}")

        sc(ws_f, r, 23, f"=U{r}+R{r}")
        sc(ws_f, r, 24, f"=-R{r}+S{r}-Custos!D{r}")

        sc(ws_f, r, 26, "=Z4" if first else 0)
        sc(ws_f, r, 27, f"=AC{pr}*(((1+Dashboard!$C$20)^(1/30))-1)")
        sc(ws_f, r, 28, f"=AC{r}" if last else 0)
        sc(ws_f, r, 29, f"=Z{r}" if first else f"=AC{pr}+AA{r}-AB{pr}")
        sc(ws_f, r, 30, f"=-Z{r}+AB{r}")

        sc(ws_f, r, 32, "=AF4" if first else 0)
        sc(ws_f, r, 33, f"=AI{pr}*(((1+Dashboard!$C$21)^(1/30))-1)")
        sc(ws_f, r, 34, f"=AI{r}" if last else 0)
        sc(ws_f, r, 35, f"=AF{r}" if first else f"=AI{pr}+AG{r}-AH{pr}")
        sc(ws_f, r, 36, f"=-AF{r}+AH{r}")

        sc(ws_f, r, 38, "=AL4" if first else 0)
        sc(ws_f, r, 39, f"=U{r}-AB{r}-AH{r}" if last else 0)
        sc(ws_f, r, 40, f"=-AL{r}+AM{r}")

        # Custos
        sc(ws_c, r, 2, n)
        sc(ws_c, r, 3, START_DATE if first else f"=C{r-1}+1")
        sc(ws_c, r, 5, "=Dashboard!$J$16/30" if first else f"=MAX(Dashboard!$J$16,(Dashboard!$I$16*Fundo!W{pr})/360)/30")
        sc(ws_c, r, 6, "=Dashboard!$J$15/30" if first else f"=MAX(Dashboard!$J$15,(Dashboard!$I$15*Fundo!W{pr})/360)/30")
        sc(ws_c, r, 7, 0 if first else f"=Fundo!R{pr}*Dashboard!$I$14/30")
        sc(ws_c, r, 8, 0 if first else f"=MAX(0,(Fundo!V{pr}+MAX(0,Fundo!T{pr})-Fundo!W{pr}*((1+Dashboard!$I$18)^(1/360)-1))*Dashboard!$I$17)")
        sc(ws_c, r, 9, 10000 if first else None)
        sc(ws_c, r, 4, f"=E{r}+F{r}+G{r}+H{r}+I{r}")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# SIMULAÇÃO NUMÉRICA (para os gráficos — sem abrir Excel)
# ════════════════════════════════════════════════════════════════════════════

def simular(params):
    N    = params["n_dias"]
    ORIG = N - 21
    WIND_DAY = ORIG + 1

    pc   = params["media_compra"]
    pv   = params["media_venda"]
    q    = params["orig_dia"]
    p18  = params["pct_18"]
    p20  = params["pct_20"]
    p22  = params["pct_22"]
    pl0  = params["pl_total"]
    cdi_aa = params["cdi"]
    sr_aa  = ((1 + cdi_aa) * (1 + params["sr_spread"])) - 1
    mz_aa  = ((1 + cdi_aa) * (1 + params["meza_spread"])) - 1

    cdi_d  = (1 + cdi_aa) ** (1/360) - 1
    sr_d   = (1 + sr_aa)  ** (1/30)  - 1
    mz_d   = (1 + mz_aa)  ** (1/30)  - 1

    sr_pl0  = pl0 * params["pct_senior"]
    mz_pl0  = pl0 * params["pct_meza"]
    jr_pl0  = pl0 * params["pct_jr"]

    piso_adm  = params["piso_adm"]
    piso_gest = params["piso_gestao"]
    taxa_adm  = params["taxa_adm"]
    taxa_gest = params["taxa_gestao"]
    taxa_cons = params["taxa_consultoria"]
    taxa_perf = params["taxa_perf"]
    perf_over = (1 + cdi_aa) * (1 + params["sr_spread"]) - 1
    perf_d    = (1 + perf_over) ** (1/360) - 1

    dias, caixas, pl_fundo, pl_sr, pl_mz, pl_jr = [], [], [], [], [], []
    desemb_hist = []

    caixa = pl0
    sr    = sr_pl0
    mz    = mz_pl0
    jr    = jr_pl0

    for n in range(1, N + 1):
        d = n

        # originações
        if d <= ORIG:
            orig = q
        else:
            orig = 0

        desembolso = orig * pc
        desemb_hist.append(desembolso)

        # recebimentos por prazo
        rec18 = desemb_hist[n - 18] / pc * pv * p18 if n >= 18 else 0
        rec20 = desemb_hist[n - 20] / pc * pv * p20 if n >= 20 else 0
        rec22 = desemb_hist[n - 22] / pc * pv * p22 if n >= 22 else 0
        recebimento = rec18 + rec20 + rec22

        # rendimento caixa
        rec_lm = caixa * cdi_d

        # custos
        custo_adm  = max(piso_adm,  (taxa_adm  * (caixa + desembolso)) / 360) / 30
        custo_gest = max(piso_gest, (taxa_gest * (caixa + desembolso)) / 360) / 30
        custo_cons = desembolso * taxa_cons / 30 if n > 1 else 0
        receita_d  = rec_lm + max(0, recebimento - desembolso)
        hurdle_d   = (caixa + desembolso) * perf_d
        custo_perf = max(0, (receita_d - hurdle_d) * taxa_perf) if n > 1 else 0
        outros     = 10000 if n == 1 else 0
        custo_total = custo_adm + custo_gest + custo_cons + custo_perf + outros

        # juros cotas
        jr_sr = sr * sr_d
        jr_mz = mz * mz_d

        # atualiza caixa e cotas
        caixa += recebimento - desembolso + rec_lm - custo_total - jr_sr - jr_mz
        sr    += jr_sr
        mz    += jr_mz
        jr     = max(0, caixa - sr - mz)

        dias.append(d)
        caixas.append(caixa)
        pl_fundo.append(caixa + desembolso)
        pl_sr.append(sr)
        pl_mz.append(mz)
        pl_jr.append(jr)

    return pd.DataFrame({
        "dia": dias,
        "caixa": caixas,
        "pl_fundo": pl_fundo,
        "pl_senior": pl_sr,
        "pl_meza": pl_mz,
        "pl_jr": pl_jr,
    })


# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR — HIPÓTESES
# ════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🚛 Modelo de Frete")
    st.markdown("---")

    st.markdown('<div class="section-header">Período</div>', unsafe_allow_html=True)
    n_dias = st.number_input("Número de dias", min_value=50, max_value=1800, value=395, step=10)

    st.markdown('<div class="section-header">Operação</div>', unsafe_allow_html=True)
    media_compra = st.number_input("Média de compra (R$)", value=1678, step=10)
    media_venda  = st.number_input("Média de venda (R$)",  value=1750, step=10)
    orig_dia     = st.number_input("Originações por dia",  value=30,   step=1)

    st.markdown('<div class="section-header">Prazos de Recebimento</div>', unsafe_allow_html=True)
    pct_18 = st.slider("% 18 dias", 0.0, 1.0, 0.30, 0.05, format="%.0%%")
    pct_20 = st.slider("% 20 dias", 0.0, 1.0, 0.30, 0.05, format="%.0%%")
    pct_22 = st.slider("% 22 dias", 0.0, 1.0, 0.40, 0.05, format="%.0%%")
    if abs(pct_18 + pct_20 + pct_22 - 1.0) > 0.01:
        st.warning(f"⚠️ Soma = {pct_18+pct_20+pct_22:.0%} (deve ser 100%)")

    st.markdown('<div class="section-header">Estrutura do Fundo</div>', unsafe_allow_html=True)
    pl_total  = st.number_input("PL Total (R$)", value=5_000_000, step=100_000)
    pct_sr    = st.slider("% Sênior",   0.0, 1.0, 0.60, 0.05, format="%.0%%")
    pct_mz    = st.slider("% Mezanino", 0.0, 1.0, 0.10, 0.05, format="%.0%%")
    pct_jr    = round(1 - pct_sr - pct_mz, 2)
    st.caption(f"% Júnior (residual): {pct_jr:.0%}")

    st.markdown('<div class="section-header">Taxas</div>', unsafe_allow_html=True)
    cdi       = st.number_input("CDI a.a. (%)",       value=14.5, step=0.5) / 100
    sr_spread = st.number_input("Sênior CDI+ (%)",    value=3.0,  step=0.5) / 100
    mz_spread = st.number_input("Mezanino CDI+ (%)",  value=5.0,  step=0.5) / 100

    st.markdown('<div class="section-header">Custos</div>', unsafe_allow_html=True)
    taxa_cons  = st.number_input("Consultoria (%)", value=1.0, step=0.1) / 100
    taxa_gest  = st.number_input("Gestão a.a. (%)", value=0.5, step=0.1) / 100
    piso_gest  = st.number_input("Piso Gestão (R$/mês)", value=15000, step=1000)
    taxa_adm   = st.number_input("Adm a.a. (%)",    value=0.2, step=0.1) / 100
    piso_adm   = st.number_input("Piso Adm (R$/mês)",   value=15000, step=1000)
    taxa_perf  = st.number_input("Performance (%)", value=20.0, step=5.0) / 100

    st.markdown("---")
    template_file = st.file_uploader("Template .xlsx (opcional)", type=["xlsx"])


# ════════════════════════════════════════════════════════════════════════════
# PARÂMETROS CONSOLIDADOS
# ════════════════════════════════════════════════════════════════════════════

params = dict(
    n_dias=n_dias,
    media_compra=media_compra, media_venda=media_venda, orig_dia=orig_dia,
    pct_18=pct_18, pct_20=pct_20, pct_22=pct_22,
    pl_total=pl_total, pct_senior=pct_sr, pct_meza=pct_mz, pct_jr=pct_jr,
    cdi=cdi, sr_spread=sr_spread, meza_spread=mz_spread,
    taxa_consultoria=taxa_cons, taxa_gestao=taxa_gest, piso_gestao=piso_gest,
    taxa_adm=taxa_adm, piso_adm=piso_adm, taxa_perf=taxa_perf,
)


# ════════════════════════════════════════════════════════════════════════════
# SIMULAÇÃO
# ════════════════════════════════════════════════════════════════════════════

df = simular(params)

lucro_op   = media_venda - media_compra
taxa_ant   = media_venda / media_compra - 1
orig_total = min(n_dias, n_dias - 21) * orig_dia
pl_final   = df["pl_fundo"].iloc[-1]
moic       = pl_final / pl_total
retorno_jr = df["pl_jr"].iloc[-1] / (pl_total * pct_jr) - 1 if pct_jr > 0 else 0


# ════════════════════════════════════════════════════════════════════════════
# LAYOUT PRINCIPAL
# ════════════════════════════════════════════════════════════════════════════

st.markdown("# Modelagem de Frete")
st.markdown(f"**{n_dias} dias** · Originação até dia {n_dias-21} · {orig_dia} operações/dia")
st.markdown("---")

# ── KPIs ─────────────────────────────────────────────────────────────────────
c1, c2, c3, c4, c5 = st.columns(5)

def kpi(col, label, value, sub=""):
    col.markdown(f"""
    <div class="metric-box">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

kpi(c1, "Lucro por Op.",  f"R$ {lucro_op:,.0f}",   f"Antecipação: {taxa_ant:.1%}")
kpi(c2, "PL Final",       f"R$ {pl_final/1e6:.2f}M", f"Inicial: R$ {pl_total/1e6:.1f}M")
kpi(c3, "MOIC",           f"{moic:.3f}x",           f"{'▲' if moic>1 else '▼'} vs 1.0x")
kpi(c4, "Retorno Jr",     f"{retorno_jr:.1%}",      f"Hurdle Sr: {((1+cdi)*(1+sr_spread)-1):.1%} a.a.")
kpi(c5, "Orig. Total",    f"{orig_total:,.0f}",     f"ops × R$ {media_compra:,}")

st.markdown("")

# ── Gráfico 1: PL Fundo e cotas ──────────────────────────────────────────────
st.markdown('<div class="section-header">Evolução do PL</div>', unsafe_allow_html=True)

fig1 = go.Figure()
fig1.add_trace(go.Scatter(x=df["dia"], y=df["pl_fundo"]/1e6,  name="PL Fundo",
    line=dict(color="#00d4aa", width=2.5)))
fig1.add_trace(go.Scatter(x=df["dia"], y=df["pl_senior"]/1e6, name="Sênior",
    line=dict(color="#4a9eff", width=1.5, dash="dot")))
fig1.add_trace(go.Scatter(x=df["dia"], y=df["pl_meza"]/1e6,   name="Mezanino",
    line=dict(color="#f5a623", width=1.5, dash="dot")))
fig1.add_trace(go.Scatter(x=df["dia"], y=df["pl_jr"]/1e6,     name="Júnior",
    line=dict(color="#e05c5c", width=1.5, dash="dot")))

fig1.update_layout(
    plot_bgcolor="#0a0c12", paper_bgcolor="#0a0c12",
    font=dict(color="#c0c4d6", family="IBM Plex Mono"),
    legend=dict(bgcolor="#0f1117", bordercolor="#2a2d3a", borderwidth=1),
    xaxis=dict(title="Dia", gridcolor="#1a1d2a", showgrid=True),
    yaxis=dict(title="R$ Milhões", gridcolor="#1a1d2a", showgrid=True),
    margin=dict(l=0, r=0, t=10, b=0),
    height=340,
)
st.plotly_chart(fig1, use_container_width=True)

# ── Gráfico 2: Caixa e composição ────────────────────────────────────────────
col_a, col_b = st.columns(2)

with col_a:
    st.markdown('<div class="section-header">Caixa ao longo do tempo</div>', unsafe_allow_html=True)
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=df["dia"], y=df["caixa"]/1e6,
        fill="tozeroy", fillcolor="rgba(0,212,170,0.08)",
        line=dict(color="#00d4aa", width=2), name="Caixa"))
    fig2.update_layout(
        plot_bgcolor="#0a0c12", paper_bgcolor="#0a0c12",
        font=dict(color="#c0c4d6", family="IBM Plex Mono"),
        xaxis=dict(title="Dia", gridcolor="#1a1d2a"),
        yaxis=dict(title="R$ Milhões", gridcolor="#1a1d2a"),
        showlegend=False, margin=dict(l=0, r=0, t=10, b=0), height=280,
    )
    st.plotly_chart(fig2, use_container_width=True)

with col_b:
    st.markdown('<div class="section-header">Composição final do PL</div>', unsafe_allow_html=True)
    vals   = [df["pl_senior"].iloc[-1], df["pl_meza"].iloc[-1], df["pl_jr"].iloc[-1]]
    labels = ["Sênior", "Mezanino", "Júnior"]
    colors = ["#4a9eff", "#f5a623", "#e05c5c"]
    fig3 = go.Figure(go.Pie(
        labels=labels, values=vals,
        hole=0.55,
        marker=dict(colors=colors, line=dict(color="#0a0c12", width=2)),
        textfont=dict(family="IBM Plex Mono", size=12),
    ))
    fig3.update_layout(
        plot_bgcolor="#0a0c12", paper_bgcolor="#0a0c12",
        font=dict(color="#c0c4d6", family="IBM Plex Mono"),
        legend=dict(bgcolor="#0f1117", bordercolor="#2a2d3a"),
        margin=dict(l=0, r=0, t=10, b=0), height=280,
    )
    st.plotly_chart(fig3, use_container_width=True)

# ── Download Excel ────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-header">Exportar modelo</div>', unsafe_allow_html=True)

col_dl1, col_dl2, col_dl3 = st.columns([2, 1, 1])

with col_dl1:
    template_bytes = None
    if template_file:
        template_bytes = template_file.read()
    else:
        try:
            with open("Modelagem_Frete_v4.xlsx", "rb") as f:
                template_bytes = f.read()
        except FileNotFoundError:
            st.info("📁 Para gerar o Excel, suba o `Modelagem_Frete_v4.xlsx` no repositório GitHub junto com este app.py — ou use o upload abaixo.")

    if template_bytes:
        with st.spinner("Gerando Excel..."):
            excel_bytes = gerar_modelo(template_bytes, params)
        st.download_button(
            label=f"⬇ Baixar Modelagem_{n_dias}d.xlsx",
            data=excel_bytes,
            file_name=f"Modelagem_Frete_{n_dias}d.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ── Tabela resumo ─────────────────────────────────────────────────────────────
with st.expander("Ver dados da simulação"):
    df_show = df.copy()
    for col in ["caixa","pl_fundo","pl_senior","pl_meza","pl_jr"]:
        df_show[col] = df_show[col].map(lambda x: f"R$ {x:,.0f}")
    df_show.columns = ["Dia","Caixa","PL Fundo","PL Sênior","PL Mezanino","PL Júnior"]
    st.dataframe(df_show, use_container_width=True, height=300)
