"""
procesar_datos.py — KAM Dashboard Multi-KAM
Genera docs/data.json con datos de TODOS los KAMs.
Soporta hojas: CUADRATURAS o OTROS_CUADRATURADOCUMENTOSINGRE
Forecast = mismo mes año anterior × FORECAST_GROWTH.
"""

import sys
import json
import re as _re
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERROR: pip install pandas openpyxl")
    sys.exit(1)

# ── Configuración ──────────────────────────────────────────────────────────────
SHEET_NAMES     = ["CUADRATURAS", "OTROS_CUADRATURADOCUMENTOSINGRE"]
REF_DATE        = pd.Timestamp("2026-08-01")
OUTPUT_FILE     = Path("docs/data.json")
FORECAST_GROWTH = 1.35
VALID_KAMS      = ["BC", "BG", "LJ", "CF", "AA", "DA", "SC", "EC", "SG", "BACK"]
KAM_NAMES       = {k: k for k in VALID_KAMS}

ALL_POSSIBLE_MONTHS = [
    "2024-01","2024-02","2024-03","2024-04","2024-05","2024-06",
    "2024-07","2024-08","2024-09","2024-10","2024-11","2024-12",
    "2025-01","2025-02","2025-03","2025-04","2025-05","2025-06",
    "2025-07","2025-08","2025-09","2025-10","2025-11","2025-12",
    "2026-01","2026-02","2026-03","2026-04","2026-05","2026-06",
    "2026-07","2026-08","2026-09","2026-10","2026-11","2026-12",
]

MONTH_LABELS = {
    "2024-01":"Ene 24","2024-02":"Feb 24","2024-03":"Mar 24","2024-04":"Abr 24",
    "2024-05":"May 24","2024-06":"Jun 24","2024-07":"Jul 24","2024-08":"Ago 24",
    "2024-09":"Sep 24","2024-10":"Oct 24","2024-11":"Nov 24","2024-12":"Dic 24",
    "2025-01":"Ene 25","2025-02":"Feb 25","2025-03":"Mar 25","2025-04":"Abr 25",
    "2025-05":"May 25","2025-06":"Jun 25","2025-07":"Jul 25","2025-08":"Ago 25",
    "2025-09":"Sep 25","2025-10":"Oct 25","2025-11":"Nov 25","2025-12":"Dic 25",
    "2026-01":"Ene 26","2026-02":"Feb 26","2026-03":"Mar 26","2026-04":"Abr 26",
    "2026-05":"May 26","2026-06":"Jun 26","2026-07":"Jul 26",
    "2026-08":"Ago 26*","2026-09":"Sep 26*","2026-10":"Oct 26*",
    "2026-11":"Nov 26*","2026-12":"Dic 26*",
}

FORECAST_BASE = {
    "2026-08":"2025-08","2026-09":"2025-09","2026-10":"2025-10",
    "2026-11":"2025-11","2026-12":"2025-12",
}

# ── Categorización de productos ────────────────────────────────────────────────
def categorize_product(producto):
    if pd.isna(producto): return "Sin clasificar"
    p = str(producto).strip().lower()
    if _re.match(r"^gif?t?t?\s*card\s*-?\s*po\.", p) or p.startswith("gift card -") or p.startswith("gift card po") or p.startswith("gitt card"):
        return "Gift Card"
    if "gift card" in p or "giftcard" in p or "gc rewards" in p or "gc market" in p:
        return "Gift Card"
    if "puntos" in p: return "Puntos"
    if "software" in p: return "Software"
    if "agencia" in p: return "Agencia"
    if "comisión" in p or "comision" in p: return "Comisión"
    if "efectivo" in p: return "Efectivo"
    if "fee" in p or "saas" in p: return "Fee SaaS"
    if "plataforma" in p: return "Plataforma"
    if "rebaja" in p: return "Ajuste"
    return "Otros"

def fmt_month_short(m):
    if not m: return ""
    names = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    return names[int(m.split("-")[1])]

def pct(a, b):
    if not b or b == 0: return None
    return round((a - b) / b * 100, 1)

# ── Carga ──────────────────────────────────────────────────────────────────────
def build_kam_map(path):
    """Lee la hoja Clientes y construye un mapa Cliente → VENDEDOR ACTUAL."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        if "Clientes" not in wb.sheetnames:
            return {}
        df_cl = pd.read_excel(path, sheet_name="Clientes")
        df_cl.columns = [str(c).strip() for c in df_cl.columns]
        if "PANEL" not in df_cl.columns or "VENDEDOR ACTUAL" not in df_cl.columns:
            return {}
        df_cl = df_cl[df_cl["VENDEDOR ACTUAL"].notna()].copy()
        df_cl["PANEL"] = df_cl["PANEL"].astype(str).str.strip()
        mapa = dict(zip(df_cl["PANEL"], df_cl["VENDEDOR ACTUAL"].astype(str).str.strip()))
        print(f"    → Mapa KAM: {len(mapa)} clientes desde hoja Clientes")
        return mapa
    except Exception as e:
        print(f"  ⚠ No se pudo leer hoja Clientes: {e}")
        return {}

def load_files(paths):
    frames = []
    kam_map = {}  # Cliente → KAM actual (desde hoja Clientes)

    for p in paths:
        p = Path(p)
        if not p.exists():
            print(f"  ✗ No encontré: {p}"); continue
        print(f"  ✓ Cargando {p.name}...")

        # Intentar cargar mapa KAM desde hoja Clientes
        mapa = build_kam_map(p)
        if mapa:
            kam_map.update(mapa)

        # Detectar hoja de facturas
        import openpyxl
        wb = openpyxl.load_workbook(p, read_only=True)
        sheet = None
        for s in SHEET_NAMES:
            if s in wb.sheetnames:
                sheet = s; break
        if not sheet:
            print(f"  ⚠ No encontré hojas conocidas en {p.name}. Disponibles: {wb.sheetnames}"); continue

        df = pd.read_excel(p, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]

        # Normalizar columna KAM
        for col in ["Kam", "KAM", "Codigo Vendedor"]:
            if col in df.columns:
                df = df.rename(columns={col: "Kam"}); break
        if "Kam" not in df.columns:
            print(f"  ⚠ Sin columna KAM en {p.name}"); continue

        # Normalizar precio
        for col in ["Precio","Prrecio","precio","PRECIO"]:
            if col in df.columns:
                df = df.rename(columns={col:"Precio"}); break

        df = df[df["Precio"].notna() & (df["Precio"] != 0)].copy()
        df["Fecha Emision"] = pd.to_datetime(df["Fecha Emision"])
        df["Mes"] = df["Fecha Emision"].dt.to_period("M").astype(str)

        # Reasignar KAM usando hoja Clientes (fuente de verdad)
        if kam_map:
            df["Cliente_clean"] = df["Cliente"].astype(str).str.strip()
            df["Kam"] = df["Cliente_clean"].map(kam_map).fillna(df["Kam"])
            asignados = df["Cliente_clean"].isin(kam_map).sum()
            print(f"    → {asignados}/{len(df)} registros reasignados con VENDEDOR ACTUAL")

        df = df[df["Kam"].isin(VALID_KAMS)]

        if "Producto" in df.columns:
            df["Categoria"] = df["Producto"].apply(categorize_product)
        else:
            df["Categoria"] = "Sin clasificar"

        frames.append(df[["Kam","Cliente","Precio","Mes","Categoria"]])
        print(f"    → {len(df)} registros válidos · hoja: {sheet}")

    if not frames:
        print("ERROR: No se cargó ningún archivo."); sys.exit(1)

    combined = pd.concat(frames, ignore_index=True)
    print("\n  Distribución por KAM (con asignación correcta):")
    for kam, n in combined[combined["Precio"]>0].groupby("Kam")["Precio"].sum().sort_values(ascending=False).items():
        print(f"    {kam}: ${n:,.0f}")
    return combined

# ── Dataset por KAM ────────────────────────────────────────────────────────────
def build_kam_data(df_kam):
    monthly = df_kam.groupby(["Cliente","Mes"])["Precio"].sum().reset_index()
    real_months = sorted(monthly["Mes"].unique().tolist())
    hist_months = [m for m in ALL_POSSIBLE_MONTHS if m in real_months and m <= "2026-12"]
    forecast_months = sorted([m for m in FORECAST_BASE if m not in real_months])
    all_months = hist_months + forecast_months

    # Años disponibles
    years = sorted(set(m[:4] for m in hist_months))
    months_by_year = {y: [m for m in hist_months if m.startswith(y)] for y in years}

    # Pares YoY comparables (año actual vs año anterior)
    current_year = max(years)
    prev_year = str(int(current_year) - 1)
    yoy_pairs = []
    for m_cur in months_by_year.get(current_year, []):
        m_prev = prev_year + "-" + m_cur[5:]
        if m_prev in hist_months:
            yoy_pairs.append((m_prev, m_cur))

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

        rev_by_year = {y: sum(vd.get(m, 0) for m in months_by_year.get(y, [])) for y in years}
        ytd_prev = sum(vd.get(p[0], 0) for p in yoy_pairs)
        ytd_cur  = sum(vd.get(p[1], 0) for p in yoy_pairs)

        vals_cur_year = [int(vd.get(m, 0)) for m in months_by_year.get(current_year, [])]
        avg_cur = int(sum(vals_cur_year) / len(vals_cur_year)) if vals_cur_year else 0

        mensual = [{"lbl": fmt_month_short(p[0]), "v25": int(vd.get(p[0],0)),
                    "v26": int(vd.get(p[1],0)), "pct": pct(vd.get(p[1],0), vd.get(p[0],0))}
                   for p in yoy_pairs]

        # Trimestral YoY
        q1_prev = sum(vd.get(m,0) for m in months_by_year.get(prev_year,[]) if m.endswith(("-01","-02","-03")))
        q1_cur  = sum(vd.get(m,0) for m in months_by_year.get(current_year,[]) if m.endswith(("-01","-02","-03")))
        trimestral = [{"lbl":"Q1","v25":int(q1_prev),"v26":int(q1_cur),"pct":pct(q1_cur,q1_prev)}]

        last3 = hist_months[-3:]
        has_recent = any(vd.get(m,0) > 0 for m in last3)
        mid = hist_months[max(0,len(hist_months)-7):-3]
        has_mid = any(vd.get(m,0) > 0 for m in mid)
        has_prev_year = rev_by_year.get(prev_year, 0) > 0

        if has_recent:
            status = "nuevo" if not has_prev_year else "activo"
        elif rev_by_year.get(current_year, 0) > 0 or has_mid:
            status = "en riesgo"
        else:
            status = "churn"

        clients[client] = {
            "total": total, "freq": freq, "primera": first_m, "ultima": last_m,
            "dias": days, "status": status,
            "rev_by_year": {y: int(v) for y,v in rev_by_year.items()},
            "rev_2025": int(rev_by_year.get("2025", 0)),
            "rev_2026": int(rev_by_year.get("2026", 0)),
            "ytd25": int(ytd_prev), "ytd26": int(ytd_cur),
            "ytd_pct": pct(ytd_cur, ytd_prev),
            "avg_2026": avg_cur,
            "vals": all_vals, "hist_count": len(hist_months),
            "forecast_vals": forecast_vals,
            "forecast_total": int(sum(forecast_vals)),
            "forecast_has_data": any(vd.get(FORECAST_BASE[m],0) > 0 for m in forecast_months),
            "mensual": mensual, "trimestral": trimestral,
            "q1_pct": pct(q1_cur, q1_prev),
        }

    # Cartera macro
    macro = {}
    for m in hist_months:
        macro[m] = int(monthly[monthly["Mes"]==m]["Precio"].sum())
    for m in forecast_months:
        macro[m] = int(monthly[monthly["Mes"]==FORECAST_BASE[m]]["Precio"].sum() * FORECAST_GROWTH)

    macro_hist = [macro.get(m,0) for m in hist_months]
    macro_fc   = [macro.get(m,0) for m in forecast_months]

    rev_macro_by_year = {y: sum(macro.get(m,0) for m in months_by_year.get(y,[])) for y in years}
    ytd_c_prev = sum(macro.get(p[0],0) for p in yoy_pairs)
    ytd_c_cur  = sum(macro.get(p[1],0) for p in yoy_pairs)
    q1_c_prev  = sum(macro.get(m,0) for m in months_by_year.get(prev_year,[]) if m.endswith(("-01","-02","-03")))
    q1_c_cur   = sum(macro.get(m,0) for m in months_by_year.get(current_year,[]) if m.endswith(("-01","-02","-03")))

    cartera = {
        "total": sum(macro_hist),
        "vals": macro_hist + macro_fc,
        "hist_count": len(hist_months),
        "forecast_vals": macro_fc,
        "forecast_total": sum(macro_fc),
        "rev_by_year": {y: int(v) for y,v in rev_macro_by_year.items()},
        "mensual": [{"lbl": fmt_month_short(p[0]), "v25": macro.get(p[0],0),
                     "v26": macro.get(p[1],0), "pct": pct(macro.get(p[1],0), macro.get(p[0],0))}
                    for p in yoy_pairs],
        "trimestral": [{"lbl":"Q1","v25":int(q1_c_prev),"v26":int(q1_c_cur),"pct":pct(q1_c_cur,q1_c_prev)}],
        "ytd25": int(ytd_c_prev), "ytd26": int(ytd_c_cur),
        "ytd_pct": pct(ytd_c_cur, ytd_c_prev),
        "q1_pct": pct(q1_c_cur, q1_c_prev),
        "years": years,
    }

    # Tickets por mes
    tickets_df = df_kam[df_kam["Precio"] > 0]
    tix_monthly = tickets_df.groupby("Mes").agg(tickets=("Precio","count"), revenue=("Precio","sum")).reset_index()
    tickets_por_mes = {}
    for m in hist_months:
        row = tix_monthly[tix_monthly["Mes"]==m]
        if len(row):
            t = int(row["tickets"].values[0])
            r = int(row["revenue"].values[0])
            tickets_por_mes[m] = {"tickets": t, "revenue": r, "ticket_prom": int(r/t) if t else 0, "label": MONTH_LABELS.get(m,m)}
        else:
            tickets_por_mes[m] = {"tickets": 0, "revenue": 0, "ticket_prom": 0, "label": MONTH_LABELS.get(m,m)}

    # Clientes activos por mes
    monthly_net = df_kam.groupby(["Cliente","Mes"])["Precio"].sum().reset_index()
    monthly_net = monthly_net[monthly_net["Precio"] > 0]
    clientes_por_mes = {}
    for m in hist_months:
        cm = monthly_net[monthly_net["Mes"]==m]
        clientes_por_mes[m] = {"count": int(cm["Cliente"].nunique()), "revenue": int(cm["Precio"].sum()), "label": MONTH_LABELS.get(m,m)}

    return {
        "months": all_months,
        "month_labels": [MONTH_LABELS.get(m, m) for m in all_months],
        "hist_count": len(hist_months),
        "forecast_months": forecast_months,
        "years": years,
        "cartera": cartera,
        "clients": clients,
        "tickets_por_mes": tickets_por_mes,
        "clientes_por_mes": clientes_por_mes,
    }

# ── Main ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    files = sys.argv[1:] if len(sys.argv) > 1 else []
    if not files:
        files = list(Path(".").glob("*.xlsx")) + list(Path("data").glob("*.xlsx"))
        if not files:
            print("USO: python procesar_datos.py archivo.xlsx"); sys.exit(1)

    print(f"\n{'='*54}")
    print(f"  KAM Dashboard — Multi-KAM (2024–2026)")
    print(f"{'='*54}\n")

    df = load_files(files)
    kams_presentes = [k for k in VALID_KAMS if k in df["Kam"].values]
    print(f"\n  KAMs: {kams_presentes}")

    output = {
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "kams": kams_presentes,
        "kam_names": {k: KAM_NAMES.get(k,k) for k in kams_presentes},
    }

    for kam in kams_presentes:
        df_kam = df[df["Kam"] == kam].copy()
        print(f"  {kam}: {df_kam['Cliente'].nunique()} clientes · {df_kam['Mes'].min()} → {df_kam['Mes'].max()}")
        output[kam] = build_kam_data(df_kam)

    # Productos por categoría
    all_months_sorted = sorted(df["Mes"].unique().tolist())
    productos_data = {}
    for kam in kams_presentes:
        df_kam = df[df["Kam"] == kam]
        if not len(df_kam): continue
        cat_monthly = df_kam.groupby(["Categoria","Mes"])["Precio"].sum().reset_index()
        cat_counts = df_kam[df_kam["Precio"]>0].groupby("Categoria")["Precio"].agg(["sum","count"]).reset_index()
        cat_data = {}
        for cat in sorted(df_kam["Categoria"].unique()):
            vals = []
            for m in all_months_sorted:
                row = cat_monthly[(cat_monthly["Categoria"]==cat) & (cat_monthly["Mes"]==m)]
                vals.append(int(row["Precio"].values[0]) if len(row) else 0)
            total = sum(vals)
            if total > 0:
                cr = cat_counts[cat_counts["Categoria"]==cat]
                n  = int(cr["count"].values[0]) if len(cr) else 0
                rev_pos = int(cr["sum"].values[0]) if len(cr) else total
                cat_data[cat] = {"vals": vals, "total": total, "tickets": n,
                                 "ticket_prom": int(rev_pos/n) if n else 0}
        productos_data[kam] = cat_data

    output["productos_data"]   = productos_data
    output["productos_months"] = all_months_sorted

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",",":"))

    kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"\n  ✓ {OUTPUT_FILE} · {kb:.0f} KB")
    print(f"  ✓ Años cubiertos: {output[kams_presentes[0]]['years']}")
    for kam in kams_presentes:
        t = output[kam]["cartera"]["total"]
        n = len(output[kam]["clients"])
        print(f"    {kam}: {n} clientes · ${t:,.0f}")
    print()
