"""Replicate Lumentum 12-month P/E cases from the workbook assumptions."""
from __future__ import annotations

Q4_COMP = 649.4
Q4_SYS = 356.9
SHARES = 102.0
VAR_OPEX = 0.06
COMP_TRAIN = 0.40
SYS_TRAIN = 0.35
PRICE = 893.61

# quarterly: growth, gm, fixed opex, price/mix, train haircut
BEAR_Q = dict(
    cg=[0.20, 0.04, 0.00, 0.00],
    sg=[0.25, 0.02, 0.00, 0.02],
    px=[0.00, -0.02, -0.02, -0.01],
    hair=[0.00, 0.10, 0.20, 0.20],
    gm=[0.52, 0.49, 0.46, 0.45],
    fix=[84, 84, 84, 84],
    tax=0.15,
)
BASE_Q = dict(
    cg=[0.232, 0.125, 0.122, 0.089],
    sg=[0.261, 0.156, 0.173, 0.164],
    px=[0, 0, 0, 0],
    hair=[0, 0, 0, 0],
    gm=[0.525, 0.54, 0.55, 0.555],
    fix=[81, 82, 84, 86],
    tax=0.15,
)
BULL_Q = dict(
    cg=[0.2598, 0.195, 0.19, 0.18],
    sg=[0.28, 0.27, 0.26, 0.22],
    px=[0, 0.005, 0.005, 0.005],
    hair=[0, 0, 0, 0],
    gm=[0.53, 0.555, 0.575, 0.58],
    fix=[82, 85, 90, 96],
    tax=0.15,
)

ANNUAL = {
    "Bear": dict(cg=0.05, sg=0.08, gm=0.46, opex_g=0.05, pe=25, tax=0.17),
    "Base": dict(cg=0.45, sg=0.50, gm=0.56, opex_g=0.12, pe=30, tax=0.17),
    "Bull": dict(cg=0.60, sg=0.65, gm=0.59, opex_g=0.18, pe=35, tax=0.17),
}


def fy27(q):
    c = Q4_COMP
    s = Q4_SYS
    rows = []
    for i in range(4):
        c *= (1 + q["cg"][i]) * (1 + q["px"][i])
        s *= (1 + q["sg"][i]) * (1 + q["px"][i])
        c_rev = c * (1 - q["hair"][i] * COMP_TRAIN)
        s_rev = s * (1 - q["hair"][i] * SYS_TRAIN)
        rev = c_rev + s_rev
        gp = rev * q["gm"][i]
        ebit = gp - q["fix"][i] - VAR_OPEX * rev
        ni = ebit * (1 - q["tax"])
        rows.append(dict(c=c_rev, s=s_rev, rev=rev, ebit=ebit, ni=ni, eps=ni / SHARES))
        c, s = c_rev / (1 - q["hair"][i] * COMP_TRAIN), s_rev / (1 - q["hair"][i] * SYS_TRAIN)
        # keep pre-haircut delivered volume path; haircut is incremental on exposed share
        # Rebuild pre-haircut stocks from growth chain instead:
    # redo cleanly
    c = Q4_COMP
    s = Q4_SYS
    rows = []
    c_pre, s_pre = Q4_COMP, Q4_SYS
    for i in range(4):
        c_pre = c_pre * (1 + q["cg"][i]) * (1 + q["px"][i])
        s_pre = s_pre * (1 + q["sg"][i]) * (1 + q["px"][i])
        c_rev = c_pre * (1 - q["hair"][i] * COMP_TRAIN)
        s_rev = s_pre * (1 - q["hair"][i] * SYS_TRAIN)
        rev = c_rev + s_rev
        gp = rev * q["gm"][i]
        ebit = gp - q["fix"][i] - VAR_OPEX * rev
        ni = ebit * (1 - q["tax"])
        rows.append(dict(c=c_rev, s=s_rev, c_pre=c_pre, s_pre=s_pre, rev=rev, ebit=ebit, ni=ni, eps=ni / SHARES, gm=q["gm"][i], fix=q["fix"][i]))
    fy = {k: sum(r[k] for r in rows) for k in ["c", "s", "rev", "ebit", "ni"]}
    fy["eps"] = fy["ni"] / SHARES
    fy["fix"] = sum(r["fix"] for r in rows)
    return rows, fy


def fy28(fy, annual):
    c = fy["c"] * (1 + annual["cg"])
    s = fy["s"] * (1 + annual["sg"])
    rev = c + s
    gp = rev * annual["gm"]
    fix = fy["fix"] * (1 + annual["opex_g"])
    ebit = gp - fix - VAR_OPEX * rev
    ni = ebit * (1 - annual["tax"])
    eps = ni / SHARES
    pt = eps * annual["pe"]
    return dict(c=c, s=s, rev=rev, ebit=ebit, ni=ni, eps=eps, pe=annual["pe"], pt=pt, ret=pt / PRICE - 1)


for name, q, key in [("Bear", BEAR_Q, "Bear"), ("Base", BASE_Q, "Base"), ("Bull", BULL_Q, "Bull")]:
    rows, fy = fy27(q)
    y28 = fy28(fy, ANNUAL[key])
    print(name)
    print("  FY27 rev", round(fy["rev"], 1), "EPS", round(fy["eps"], 2), "Q1 EPS", round(rows[0]["eps"], 2), "Q1 rev", round(rows[0]["rev"], 1))
    print("  FY28 rev", round(y28["rev"], 1), "EPS", round(y28["eps"], 2), "PT", round(y28["pt"], 2), "ret", round(y28["ret"] * 100, 1))
