import sys
import datetime
from openpyxl import load_workbook
from openpyxl.cell import MergedCell

# ── Argumentos ───────────────────────────────────────────────────────────────
if len(sys.argv) < 2:
    print("Uso: python3 gerar_modelo_v4.py <N_DIAS> [saida.xlsx]")
    sys.exit(1)

N_DIAS     = int(sys.argv[1])
TEMPLATE   = "Modelagem_Frete_v4.xlsx"
OUTPUT     = sys.argv[2] if len(sys.argv) > 2 else f"Modelagem_Frete_{N_DIAS}d.xlsx"

# ── Parâmetros derivados ─────────────────────────────────────────────────────
ORIG_DAYS  = N_DIAS - 21     # último dia com originação
WIND_ROW   = ORIG_DAYS + 5   # primeira linha onde E/I/M = 0
LAST_ROW   = N_DIAS + 4      # última linha de dados (Fundo e Custos)

print(f"N_DIAS={N_DIAS} | ORIG_DAYS={ORIG_DAYS} | WIND_ROW={WIND_ROW} | LAST_ROW={LAST_ROW}")

# ── Carregar template ────────────────────────────────────────────────────────
wb = load_workbook(TEMPLATE)
ws_f = wb["Fundo"]
ws_c = wb["Custos"]
ws_d = wb["Dashboard"]


# ════════════════════════════════════════════════════════════════════════════
# HELPER
# ════════════════════════════════════════════════════════════════════════════

def set_cell(ws, row, col, value):
    """Escreve somente em células normais (ignora MergedCell)."""
    cell = ws.cell(row, col)
    if not isinstance(cell, MergedCell):
        cell.value = value

def clear_rows(ws, from_row, to_row=5000):
    """Apaga o conteúdo de todas as linhas de from_row até to_row."""
    for r in range(from_row, to_row + 1):
        if r > ws.max_row:
            break
        for col in range(1, ws.max_column + 1):
            set_cell(ws, r, col, None)


# ════════════════════════════════════════════════════════════════════════════
# 1. DASHBOARD — atualizar referências dinâmicas à última linha
# ════════════════════════════════════════════════════════════════════════════

LR = LAST_ROW   # alias curto

ws_d["E5"].value  = f"=Fundo!AD5+Fundo!AD{LR}"
ws_d["E6"].value  = f"=Fundo!AJ5+Fundo!AJ{LR}"
ws_d["E7"].value  = f"=Fundo!AN5+Fundo!AN{LR}"
ws_d["C23"].value = f"=Fundo!U{LR}/Dashboard!D14"


# ════════════════════════════════════════════════════════════════════════════
# 2. FUNDO — apagar dados antigos e reescrever
# ════════════════════════════════════════════════════════════════════════════

clear_rows(ws_f, from_row=5)

# ── Row 2: IRR / XIRR ───────────────────────────────────────────────────────
# (não mexemos nas rows 2-4; elas ficam como no template)
# Apenas atualizamos os ranges das IRRs
ws_f["X2"].value  = f"=IRR(X5:X{LR},0.001)"
ws_f["AD2"].value = f"=IRR(AD5:AD{LR})"
ws_f["AJ2"].value = f"=IRR(AJ5:AJ{LR},0.001)"
ws_f["AN2"].value = f"=IRR(AN5:AN{LR})"


# ── Linhas de dados (rows 5 … LAST_ROW) ─────────────────────────────────────
START_DATE = datetime.datetime(2026, 6, 1)

for n in range(1, N_DIAS + 1):
    r  = n + 4       # linha atual
    pr = r - 1       # linha anterior

    first = (n == 1)
    last  = (n == N_DIAS)
    wind  = (r == WIND_ROW)          # primeira linha de vento (E/I/M = 0)
    orig  = (r < WIND_ROW)           # ainda em período de originação
    q_eim = (r >= WIND_ROW)          # Q usa fórmula E+I+M

    # ── B: Hurdle benchmark ──────────────────────────────────────────────────
    set_cell(ws_f, r, 2,
             "=Dashboard!D14" if first
             else f"=B{pr}*(((1+Dashboard!$C$20)^(1/30)))")

    # ── C: Contador de período ───────────────────────────────────────────────
    set_cell(ws_f, r, 3, n)

    # ── D: Data ──────────────────────────────────────────────────────────────
    set_cell(ws_f, r, 4,
             START_DATE if first
             else f"=D{pr}+1")

    # ── E/I/M: Originações por prazo ─────────────────────────────────────────
    if first:
        set_cell(ws_f, r, 5,  "=Q5*Dashboard!$I$9")
        set_cell(ws_f, r, 9,  "=Q5*Dashboard!$I$10")
        set_cell(ws_f, r, 13, "=Q5*Dashboard!$I$11")
    elif wind:
        set_cell(ws_f, r, 5,  0)
        set_cell(ws_f, r, 9,  0)
        set_cell(ws_f, r, 13, 0)
    else:
        set_cell(ws_f, r, 5,  f"=E{pr}")
        set_cell(ws_f, r, 9,  f"=I{pr}")
        set_cell(ws_f, r, 13, f"=M{pr}")

    # ── F/J/N: Desembolsos ───────────────────────────────────────────────────
    set_cell(ws_f, r, 6,  f"=E{r}*Dashboard!$I$3")
    set_cell(ws_f, r, 10, f"=I{r}*Dashboard!$I$3")
    set_cell(ws_f, r, 14, f"=M{r}*Dashboard!$I$3")

    # ── G/K/O: Levantamentos (recebimentos por prazo) ────────────────────────
    # Offset: prazo - 1 linhas atrás (ex: 18 dias → r-17)
    set_cell(ws_f, r, 7,  0 if r < 22 else f"=E{r-17}*Dashboard!$I$4")
    set_cell(ws_f, r, 11, 0 if r < 24 else f"=I{r-19}*Dashboard!$I$4")
    set_cell(ws_f, r, 15, 0 if r < 26 else f"=M{r-21}*Dashboard!$I$4")

    # ── Q: Total de originações ──────────────────────────────────────────────
    if first:
        set_cell(ws_f, r, 17, "=Dashboard!I8")
    elif q_eim:
        set_cell(ws_f, r, 17, f"=E{r}+I{r}+M{r}")
    else:
        set_cell(ws_f, r, 17, f"=Q{pr}")

    # ── R/S: Totais desembolso / levantamento ────────────────────────────────
    set_cell(ws_f, r, 18, f"=F{r}+J{r}+N{r}")
    set_cell(ws_f, r, 19, f"=G{r}+K{r}+O{r}")

    # ── T: P&L Ativos ────────────────────────────────────────────────────────
    # T = S{r} - R{r-17}  (compara levantamentos com desembolsos de 18d atrás)
    set_cell(ws_f, r, 20, 0 if r < 22 else f"=S{r}-R{r-17}")

    # ── V: Rendimento sobre caixa (CDI diário) ───────────────────────────────
    set_cell(ws_f, r, 22,
             0 if first
             else f"=U{pr}*(((1+Dashboard!$C$17)^(1/30))-1)")

    # ── U: Caixa ─────────────────────────────────────────────────────────────
    if first:
        set_cell(ws_f, r, 21, f"=U4-R{r}+S{r}+V{r}-Custos!D{r}")
    elif n == 2:
        set_cell(ws_f, r, 21,
                 f"=U{pr}-R{r}+S{r}+V{r}-Custos!D{r}"
                 f"-AB{pr}-AH{pr}-AM{pr}")
    else:
        set_cell(ws_f, r, 21,
                 f"=U{pr}-R{r}+S{r}+V{r}-Custos!D{r}"
                 f"+Z{pr}+AF{pr}+AL{pr}"
                 f"-AB{pr}-AH{pr}-AM{pr}")

    # ── W: PL Fundo ──────────────────────────────────────────────────────────
    set_cell(ws_f, r, 23, f"=U{r}+R{r}")

    # ── X: Fluxo ─────────────────────────────────────────────────────────────
    set_cell(ws_f, r, 24, f"=-R{r}+S{r}-Custos!D{r}")

    # ── SÊNIOR ───────────────────────────────────────────────────────────────
    # Z: Integralização (apenas dia 1)
    set_cell(ws_f, r, 26, "=Z4" if first else 0)

    # AA: Juros Sr
    set_cell(ws_f, r, 27,
             f"=AC{pr}*(((1+Dashboard!$C$20)^(1/30))-1)")

    # AB: Amortização Sr (0 salvo no último dia = liquidação total)
    set_cell(ws_f, r, 28, f"=AC{r}" if last else 0)

    # AC: PL Sr
    set_cell(ws_f, r, 29,
             f"=Z{r}" if first
             else f"=AC{pr}+AA{r}-AB{pr}")

    # AD: Fluxo Sr
    set_cell(ws_f, r, 30, f"=-Z{r}+AB{r}")

    # ── MEZANINO ─────────────────────────────────────────────────────────────
    # AF: Integralização (apenas dia 1)
    set_cell(ws_f, r, 32, "=AF4" if first else 0)

    # AG: Juros Meza
    set_cell(ws_f, r, 33,
             f"=AI{pr}*(((1+Dashboard!$C$21)^(1/30))-1)")

    # AH: Amortização Meza (0 salvo no último dia)
    set_cell(ws_f, r, 34, f"=AI{r}" if last else 0)

    # AI: PL Meza
    set_cell(ws_f, r, 35,
             f"=AF{r}" if first
             else f"=AI{pr}+AG{r}-AH{pr}")

    # AJ: Fluxo Meza
    set_cell(ws_f, r, 36, f"=-AF{r}+AH{r}")

    # ── JÚNIOR ───────────────────────────────────────────────────────────────
    # AL: Integralização (apenas dia 1)
    set_cell(ws_f, r, 38, "=AL4" if first else 0)

    # AM: Amortização Jr (0 salvo no último dia = residual após Sr e Meza)
    set_cell(ws_f, r, 39, f"=U{r}-AB{r}-AH{r}" if last else 0)

    # AN: Fluxo Jr
    set_cell(ws_f, r, 40, f"=-AL{r}+AM{r}")


# ════════════════════════════════════════════════════════════════════════════
# 3. CUSTOS — apagar dados antigos e reescrever
# ════════════════════════════════════════════════════════════════════════════

clear_rows(ws_c, from_row=5)

for n in range(1, N_DIAS + 1):
    r  = n + 4    # linha Custos
    fr = r        # linha Fundo correspondente (mesmo índice)
    pr = r - 1    # linha Fundo do dia anterior (lag de 1 dia)

    first = (n == 1)

    # ── B: Contador ──────────────────────────────────────────────────────────
    set_cell(ws_c, r, 2, n)

    # ── C: Data ──────────────────────────────────────────────────────────────
    set_cell(ws_c, r, 3,
             START_DATE if first
             else f"=C{r-1}+1")

    # ── E: Custo Adm ─────────────────────────────────────────────────────────
    set_cell(ws_c, r, 5,
             "=Dashboard!$J$16/30" if first
             else f"=MAX(Dashboard!$J$16,(Dashboard!$I$16*Fundo!W{pr})/360)/30")

    # ── F: Custo Gestão ──────────────────────────────────────────────────────
    set_cell(ws_c, r, 6,
             "=Dashboard!$J$15/30" if first
             else f"=MAX(Dashboard!$J$15,(Dashboard!$I$15*Fundo!W{pr})/360)/30")

    # ── G: Consultoria ───────────────────────────────────────────────────────
    set_cell(ws_c, r, 7,
             0 if first
             else f"=Fundo!R{pr}*Dashboard!$I$14/30")

    # ── H: Performance ───────────────────────────────────────────────────────
    set_cell(ws_c, r, 8,
             0 if first
             else (f"=MAX(0,(Fundo!V{pr}+MAX(0,Fundo!T{pr})"
                   f"-Fundo!W{pr}*((1+Dashboard!$I$18)^(1/360)-1))"
                   f"*Dashboard!$I$17)"))

    # ── I: Outros ────────────────────────────────────────────────────────────
    set_cell(ws_c, r, 9, 10000 if first else None)

    # ── D: TOTAL (soma de E:I) ────────────────────────────────────────────────
    set_cell(ws_c, r, 4, f"=E{r}+F{r}+G{r}+H{r}+I{r}")


# ════════════════════════════════════════════════════════════════════════════
# 4. SALVAR
# ════════════════════════════════════════════════════════════════════════════

wb.save(OUTPUT)
print(f"Salvo: {OUTPUT}")
print(f"  Fundo:  rows 5-{LAST_ROW} ({N_DIAS} dias)")
print(f"  Custos: rows 5-{LAST_ROW} ({N_DIAS} dias)")
print(f"  Originações até dia {ORIG_DAYS}, wind-down a partir do dia {ORIG_DAYS+1}")
