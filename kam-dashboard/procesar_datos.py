"""
procesar_datos.py — KAM Dashboard Multi-KAM
Genera docs/data.json con datos de TODOS los KAMs.
Forecast = mismo mes año anterior × FORECAST_GROWTH.
"""

import sys
import json
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERROR: pip install pandas openpyxl")
    sys.exit(1)

# ── Configuración ──────────────────────────────────────────────────────────────
SHEET_NAME      = "CUADRATURAS"
REF_DATE        = pd.Timestamp("2026-07-01")
OUTPUT_FILE     = Path("docs/data.json")
FORECAST_GROWTH = 1.35

# KAMs válidos (excluir BACK, #N/D, etc.)
VALID_KAMS = ["BC", "BG", "LJ", "CF", "AA", "DA", "SC", "EC", "SG"]

KAM_NAMES = {
    "BC": "BC", "BG": "BG", "LJ": "LJ", "CF": "CF",
    "AA": "AA", "DA": "DA", "SC": "SC", "EC": "EC", "SG": "SG"
}

ALL_POSSIBLE_MONTHS = [
    "2025-01","2025-02","2025-03","2025-04","2025-05","2025-06",
    "2025-07","2025-08","2025-09","2025-10","2025-11","2025-12",
    "2026-01","2026-02","2026-03","2026-04","2026-05","2026-06",
    "2026-07","2026-08","2026-09","2026-10","2026-11","2026-12",
]

MONTH_LABELS = {
    "2025-01":"Ene 25","2025-02":"Feb 25","2025-03":"Mar 25","2025-04":"Abr 25",
    "2025-05":"May 25","2025-06":"Jun 25","2025-07":"Jul 25","2025-08":"Ago 25",
    "2025-09":"Sep 25","2025-10":"Oct 25","2025-11":"Nov 25","2025-12":"Dic 25",
    "2026-01":"Ene 26","2026-02":"Feb 26","2026-03":"Mar 26","2026-04":"Abr 26",
    "2026-05":"May 26","2026-06":"Jun 26",
    "2026-07":"Jul 26*","2026-08":"Ago 26*","2026-09":"Sep 26*",
    "2026-10":"Oct 26*","2026-11":"Nov 26*","2026-12":"Dic 26*",
}

FORECAST_BASE = {
    "2026-07":"2025-07","2026-08":"2025-08","2026-09":"2025-09",
    "2026-10":"2025-10","2026-11":"2025-11","2026-12":"2025-12",
}

def pct(a, b):
    if not b or b == 0: return None
    return round((a - b) / b * 100, 1)

def fmt_month_short(m):
    if not m: return ""
    names = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    return names[int(m.split("-")[1])]

# ── Categorización de productos ─────────────────────────────────────────────────
import re as _re
def categorize_product(producto):
    """Agrupa nombres de producto inconsistentes en categorías limpias."""
    if pd.isna(producto):
        return "Sin clasificar"
    p = str(producto).strip().lower()

    # Gift Card con número de orden (PO.xxxxx) y variantes de escritura
    if _re.match(r"^gif?t?t?\s*card\s*-?\s*po\.", p) or p.startswith("gift card -") or p.startswith("gift card po") or p.startswith("gitt card"):
        return "Gift Card"
    if "gift card" in p or "giftcard" in p or "gc rewards" in p or "gc market" in p:
        return "Gift Card"
    if "puntos" in p:
        return "Puntos"
    if "software" in p:
        return "Software"
    if "agencia" in p:
        return "Agencia"
    if "comisión" in p or "comision" in p:
        return "Comisión"
    if "efectivo" in p:
        return "Efectivo"
    if "fee" in p or "saas" in p:
        return "Fee SaaS"
    if "plataforma" in p:
        return "Plataforma"
    if "rebaja" in p:
        return "Ajuste"
    return "Otros"

# ── Carga ──────────────────────────────────────────────────────────────────────
def load_files(paths):
    frames = []
    for p in paths:
        p = Path(p)
        if not p.exists():
            print(f"  ✗ No encontré: {p}"); continue
        print(f"  ✓ Cargando {p.name}...")
        df = pd.read_excel(p, sheet_name=SHEET_NAME)
        df.columns = [c.strip() for c in df.columns]

        # Detectar columna KAM
        if "Kam" in df.columns:
            df = df.rename(columns={"Kam": "KAM"})
        elif "Codigo Vendedor" in df.columns:
            df = df.rename(columns={"Codigo Vendedor": "KAM"})
        else:
            print(f"  ⚠ Sin columna KAM en {p.name}"); continue

        # Detectar columna precio
        for col in ["Precio","Prrecio","precio","PRECIO"]:
            if col in df.columns:
                df = df.rename(columns={col:"Precio"}); break

        df = df[df["Precio"] != 0].copy()
        df["Fecha Emision"] = pd.to_datetime(df["Fecha Emision"])
        df["Mes"] = df["Fecha Emision"].dt.to_period("M").astype(str)
        df = df[df["KAM"].isin(VALID_KAMS)]

        if "Producto" in df.columns:
            df["Categoria"] = df["Producto"].apply(categorize_product)
        else:
            df["Categoria"] = "Sin clasificar"

        frames.append(df[["KAM","Cliente","Precio","Mes","Categoria"]])
        print(f"    → {len(df)} registros ({df['KAM'].nunique()} KAMs)")

    if not frames:
        print("ERROR: No se cargó ningún archivo."); sys.exit(1)
    return pd.concat(frames, ignore_index=True)

# ── Dataset por KAM ────────────────────────────────────────────────────────────
def build_kam_data(df_kam, kam_code):
    monthly = df_kam.groupby(["Cliente","Mes"])["Precio"].sum().reset_index()

    real_months = sorted(monthly["Mes"].unique().tolist())
    hist_months = [m for m in ALL_POSSIBLE_MONTHS if m in real_months]
    forecast_months = sorted([m for m in FORECAST_BASE if m not in real_months])
    all_months = hist_months + forecast_months

    months_2025 = [m for m in hist_months if m.startswith("2025")]
    months_2026 = [m for m in hist_months if m.startswith("2026")]
    yoy_pairs = [(f"2025-{m.split('-')[1]}", m) for m in months_2026
                 if f"2025-{m.split('-')[1]}" in months_2025]

    clients = {}
    for client, grp in monthly.groupby("Cliente"):
        vd = dict(zip(grp["Mes"], grp["Precio"]))
        hist_vals = [int(vd.get(m, 0)) for m in hist_months]
        forecast_vals = [int(vd.get(FORECAST_BASE[m], 0) * FORECAST_GROWTH) for m in forecast_months]
        all_vals = hist_vals + forecast_vals

        total = sum(hist_vals)
        active = [hist_months[i] for i, v in enumerate(hist_vals) if v > 0]
        freq = len(active)
        last_m = active[-1] if active else None
        first_m = active[0] if active else None
        days = int((REF_DATE - pd.Timestamp(last_m + "-28")).days) if last_m else 999

        rev_2025 = sum(vd.get(m, 0) for m in months_2025)
        rev_2026 = sum(vd.get(m, 0) for m in months_2026)
        ytd25 = sum(vd.get(p[0], 0) for p in yoy_pairs)
        ytd26 = sum(vd.get(p[1], 0) for p in yoy_pairs)

        vals_2026 = [int(vd.get(m, 0)) for m in months_2026]
        avg_2026 = int(sum(vals_2026) / len(vals_2026)) if vals_2026 else 0

        mensual = [{"lbl": fmt_month_short(p[0]), "v25": int(vd.get(p[0],0)),
                    "v26": int(vd.get(p[1],0)), "pct": pct(vd.get(p[1],0), vd.get(p[0],0))}
                   for p in yoy_pairs]
        q1_25 = sum(vd.get(m,0) for m in months_2025 if m.endswith(("-01","-02","-03")))
        q1_26 = sum(vd.get(m,0) for m in months_2026 if m.endswith(("-01","-02","-03")))

        last3 = hist_months[-3:]
        has_recent = any(vd.get(m,0) > 0 for m in last3)
        mid = hist_months[max(0,len(hist_months)-7):-3]
        has_mid = any(vd.get(m,0) > 0 for m in mid)
        if has_recent:
            status = "nuevo" if not rev_2025 else "activo"
        elif rev_2026 > 0 or has_mid:
            status = "en riesgo"
        else:
            status = "churn"

        clients[client] = {
            "total": total, "freq": freq, "primera": first_m, "ultima": last_m,
            "dias": days, "status": status,
            "rev_2025": int(rev_2025), "rev_2026": int(rev_2026),
            "ytd25": int(ytd25), "ytd26": int(ytd26), "ytd_pct": pct(ytd26, ytd25),
            "avg_2026": avg_2026,
            "vals": all_vals, "hist_count": len(hist_months),
            "forecast_vals": [int(vd.get(FORECAST_BASE[m],0)*FORECAST_GROWTH) for m in forecast_months],
            "forecast_total": int(sum(int(vd.get(FORECAST_BASE[m],0)*FORECAST_GROWTH) for m in forecast_months)),
            "forecast_has_data": any(vd.get(FORECAST_BASE[m],0)>0 for m in forecast_months),
            "mensual": mensual,
            "trimestral": [{"lbl":"Q1","v25":int(q1_25),"v26":int(q1_26),"pct":pct(q1_26,q1_25)}],
            "q1_pct": pct(q1_26, q1_25),
            "may_pct": next((p["pct"] for p in mensual if p["lbl"]=="May"), None),
        }

    # Cartera del KAM
    macro = {}
    for m in hist_months:
        macro[m] = int(monthly[monthly["Mes"]==m]["Precio"].sum())
    for m in forecast_months:
        macro[m] = int(monthly[monthly["Mes"]==FORECAST_BASE[m]]["Precio"].sum() * FORECAST_GROWTH)

    macro_hist = [macro.get(m,0) for m in hist_months]
    macro_fc   = [macro.get(m,0) for m in forecast_months]
    ytd_c25 = sum(macro.get(p[0],0) for p in yoy_pairs)
    ytd_c26 = sum(macro.get(p[1],0) for p in yoy_pairs)
    q1_c25  = sum(macro.get(m,0) for m in months_2025 if m.endswith(("-01","-02","-03")))
    q1_c26  = sum(macro.get(m,0) for m in months_2026 if m.endswith(("-01","-02","-03")))

    cartera = {
        "total": sum(macro_hist),
        "vals": macro_hist + macro_fc,
        "hist_count": len(hist_months),
        "forecast_vals": macro_fc,
        "forecast_total": sum(macro_fc),
        "mensual": [{"lbl": fmt_month_short(p[0]), "v25": macro.get(p[0],0),
                     "v26": macro.get(p[1],0), "pct": pct(macro.get(p[1],0), macro.get(p[0],0))}
                    for p in yoy_pairs],
        "trimestral": [{"lbl":"Q1","v25":int(q1_c25),"v26":int(q1_c26),"pct":pct(q1_c26,q1_c25)}],
        "ytd25": int(ytd_c25), "ytd26": int(ytd_c26), "ytd_pct": pct(ytd_c26,ytd_c25),
        "q1_pct": pct(q1_c26,q1_c25),
        "may_pct": pct(macro.get("2026-05",0), macro.get("2025-05",0)),
    }

    # Clientes activos por mes
    monthly_net = df_kam.groupby(["Cliente","Mes"])["Precio"].sum().reset_index()
    monthly_net = monthly_net[monthly_net["Precio"] > 0]
    clientes_por_mes = {}
    for m in hist_months:
        cm = monthly_net[monthly_net["Mes"]==m]
        clientes_por_mes[m] = {
            "count": int(cm["Cliente"].nunique()),
            "revenue": int(cm["Precio"].sum()),
            "label": MONTH_LABELS.get(m, m)
        }

    # Tickets por mes (facturas positivas)
    tickets_raw = df_kam[df_kam["Precio"] > 0].groupby("Mes").agg(
        tickets=("Precio","count"),
        revenue=("Precio","sum")
    ).reset_index()
    tickets_por_mes = {}
    for m in hist_months:
        row = tickets_raw[tickets_raw["Mes"]==m]
        if len(row)>0:
            t=int(row["tickets"].values[0]); r=int(row["revenue"].values[0])
            tickets_por_mes[m]={"tickets":t,"revenue":r,"ticket_prom":int(r/t) if t>0 else 0,"label":MONTH_LABELS.get(m,m)}
        else:
            tickets_por_mes[m]={"tickets":0,"revenue":0,"ticket_prom":0,"label":MONTH_LABELS.get(m,m)}

    return {
        "months": all_months,
        "month_labels": [MONTH_LABELS.get(m, m) for m in all_months],
        "hist_count": len(hist_months),
        "forecast_months": forecast_months,
        "cartera": cartera,
        "clients": clients,
        "clientes_por_mes": clientes_por_mes,
        "tickets_por_mes": tickets_por_mes,
    }

# ── Main ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    files = sys.argv[1:] if len(sys.argv) > 1 else []
    if not files:
        files = list(Path(".").glob("*.xlsx")) + list(Path("data").glob("*.xlsx"))
        if not files:
            print("USO: python procesar_datos.py archivo.xlsx"); sys.exit(1)

    print(f"\n{'='*52}")
    print(f"  KAM Dashboard — Multi-KAM")
    print(f"{'='*52}\n")

    df = load_files(files)
    kams_presentes = [k for k in VALID_KAMS if k in df["KAM"].values]
    print(f"\n  KAMs encontrados: {kams_presentes}\n")

    output = {
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "kams": kams_presentes,
        "kam_names": {k: KAM_NAMES.get(k,k) for k in kams_presentes},
    }

    for kam in kams_presentes:
        df_kam = df[df["KAM"] == kam].copy()
        print(f"  Procesando {kam}: {df_kam['Cliente'].nunique()} clientes...")
        output[kam] = build_kam_data(df_kam, kam)

    # ── Productos por categoría: revenue mensual por KAM ────────────────────────
    all_months_sorted = sorted(df["Mes"].unique().tolist())
    productos_data = {}
    for kam in kams_presentes:
        df_kam = df[df["KAM"] == kam]
        if len(df_kam) == 0:
            continue
        cat_monthly = df_kam.groupby(["Categoria","Mes"])["Precio"].sum().reset_index()
        cats = sorted(df_kam["Categoria"].unique().tolist())
        cat_data = {}
        for cat in cats:
            vals = []
            for m in all_months_sorted:
                row = cat_monthly[(cat_monthly["Categoria"]==cat) & (cat_monthly["Mes"]==m)]
                vals.append(int(row["Precio"].values[0]) if len(row) else 0)
            total = sum(vals)
            if total > 0:
                cat_data[cat] = {"vals": vals, "total": total}
        productos_data[kam] = cat_data

    output["productos_data"] = productos_data
    output["productos_months"] = all_months_sorted

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",",":"))

    kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"\n  ✓ Generado: {OUTPUT_FILE}  ({kb:.1f} KB)")
    for kam in kams_presentes:
        total = output[kam]["cartera"]["total"]
        nclientes = len(output[kam]["clients"])
        print(f"    {kam}: {nclientes} clientes | ${total:,.0f}")
    print()
    print("  Categorías de producto detectadas:")
    for cat in sorted(df["Categoria"].unique()):
        rev = df[(df["Categoria"]==cat) & (df["Precio"]>0)]["Precio"].sum()
        print(f"    {cat}: ${rev:,.0f}")
    print()
