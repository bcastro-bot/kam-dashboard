"""
procesar_datos.py — KAM Dashboard BC
Genera docs/data.json a partir del Excel de ventas.
Forecast = mismo mes año anterior × FORECAST_GROWTH, solo para meses futuros.
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
KAM_CODE        = "BC"
SHEET_NAME      = "CUADRATURAS"
REF_DATE        = pd.Timestamp("2026-07-01")
OUTPUT_FILE     = Path("docs/data.json")
FORECAST_GROWTH = 1.35   # +35% sobre mismo mes año anterior

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

# Meses 2026 con base en 2025 para forecast
FORECAST_BASE = {
    "2026-07":"2025-07","2026-08":"2025-08","2026-09":"2025-09",
    "2026-10":"2025-10","2026-11":"2025-11","2026-12":"2025-12",
}

# ── Helpers ────────────────────────────────────────────────────────────────────
def pct(a, b):
    if not b or b == 0: return None
    return round((a - b) / b * 100, 1)

def fmt_month_short(m):
    if not m: return ""
    names = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    return names[int(m.split("-")[1])]

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
        if "Kam" in df.columns:
            df = df[df["Kam"] == KAM_CODE].copy()
        elif "Codigo Vendedor" in df.columns:
            df = df[df["Codigo Vendedor"] == KAM_CODE].copy()
        else:
            print(f"  ⚠ Sin columna KAM en {p.name}"); continue
        for col in ["Precio","Prrecio","precio","PRECIO"]:
            if col in df.columns:
                df = df.rename(columns={col:"Precio"}); break
        df = df[df["Precio"] != 0].copy()  # incluir NC (negativos)
        df["Fecha Emision"] = pd.to_datetime(df["Fecha Emision"])
        df["Mes"] = df["Fecha Emision"].dt.to_period("M").astype(str)
        frames.append(df[["Cliente","Precio","Mes"]])
        print(f"    → {len(df)} registros BC")
    if not frames:
        print("ERROR: No se cargó ningún archivo."); sys.exit(1)
    combined = pd.concat(frames, ignore_index=True)
    # No hacer drop_duplicates porque las NC tienen precios negativos únicos
    return combined

# ── Dataset ────────────────────────────────────────────────────────────────────
def build_dataset(df):
    monthly = df.groupby(["Cliente","Mes"])["Precio"].sum().reset_index()

    # Detectar meses reales presentes en los datos
    real_months_in_data = sorted(monthly["Mes"].unique().tolist())
    # Solo usar meses que estén en ALL_POSSIBLE_MONTHS y sean datos reales (no forecast)
    hist_months = [m for m in ALL_POSSIBLE_MONTHS if m in real_months_in_data and m <= "2026-12"]
    # Meses de forecast: los de 2026 que NO están en los datos reales
    forecast_months = [m for m in FORECAST_BASE.keys() if m not in real_months_in_data]
    all_months = hist_months + sorted(forecast_months)

    print(f"  Meses reales: {hist_months[0]} → {hist_months[-1]} ({len(hist_months)} meses)")
    print(f"  Meses forecast: {forecast_months}")

    # Meses 2025 y 2026 reales para comparaciones
    months_2025 = [m for m in hist_months if m.startswith("2025")]
    months_2026_real = [m for m in hist_months if m.startswith("2026")]

    # Pares YoY comparables
    yoy_pairs = []
    for m26 in months_2026_real:
        m25 = "2025-" + m26.split("-")[1]
        if m25 in months_2025:
            yoy_pairs.append((m25, m26))

    clients = {}
    for client, grp in monthly.groupby("Cliente"):
        vd = dict(zip(grp["Mes"], grp["Precio"]))

        # Histórico real
        hist_vals = [int(vd.get(m, 0)) for m in hist_months]

        # Forecast
        forecast_vals = [int(vd.get(FORECAST_BASE[m], 0) * FORECAST_GROWTH) for m in sorted(forecast_months)]
        forecast_has_data = any(v > 0 for v in forecast_vals)
        all_vals = hist_vals + forecast_vals

        total = sum(hist_vals)
        active = [hist_months[i] for i, v in enumerate(hist_vals) if v > 0]
        freq = len(active)
        last_m = active[-1] if active else None
        first_m = active[0] if active else None
        days = int((REF_DATE - pd.Timestamp(last_m + "-28")).days) if last_m else 999

        rev_2025 = sum(vd.get(m, 0) for m in months_2025)
        rev_2026 = sum(vd.get(m, 0) for m in months_2026_real)

        # YTD: primeros N meses comparables
        ytd25 = sum(vd.get(p[0], 0) for p in yoy_pairs)
        ytd26 = sum(vd.get(p[1], 0) for p in yoy_pairs)
        ytd_pct = pct(ytd26, ytd25)

        # Promedio mensual 2026 real (sobre todos los meses transcurridos, incluyendo $0)
        vals_2026_real = [int(vd.get(m, 0)) for m in months_2026_real]
        avg_2026 = int(sum(vals_2026_real) / len(vals_2026_real)) if vals_2026_real else 0

        # Mensual YoY
        mensual = [{"lbl": fmt_month_short(p[0]), "v25": int(vd.get(p[0],0)),
                    "v26": int(vd.get(p[1],0)), "pct": pct(vd.get(p[1],0), vd.get(p[0],0))}
                   for p in yoy_pairs]

        # Trimestral
        q1_25 = sum(vd.get(m,0) for m in months_2025 if m.endswith(("-01","-02","-03")))
        q1_26 = sum(vd.get(m,0) for m in months_2026_real if m.endswith(("-01","-02","-03")))
        trimestral = [{"lbl":"Q1","v25":int(q1_25),"v26":int(q1_26),"pct":pct(q1_26,q1_25)}]

        # Status
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
            "ytd25": int(ytd25), "ytd26": int(ytd26), "ytd_pct": ytd_pct,
            "avg_2026": avg_2026,
            "vals": all_vals, "hist_count": len(hist_months),
            "forecast_vals": forecast_vals, "forecast_total": int(sum(forecast_vals)),
            "forecast_has_data": forecast_has_data,
            "mensual": mensual, "trimestral": trimestral,
            "q1_pct": trimestral[0]["pct"],
            "may_pct": next((p["pct"] for p in mensual if p["lbl"]=="May"), None),
        }

    # Cartera
    macro = {}
    for m in hist_months:
        macro[m] = int(monthly[monthly["Mes"]==m]["Precio"].sum())
    for m in sorted(forecast_months):
        macro[m] = int(monthly[monthly["Mes"]==FORECAST_BASE[m]]["Precio"].sum() * FORECAST_GROWTH)

    macro_hist = [macro.get(m,0) for m in hist_months]
    macro_fc   = [macro.get(m,0) for m in sorted(forecast_months)]

    ytd_c25 = sum(macro.get(p[0],0) for p in yoy_pairs)
    ytd_c26 = sum(macro.get(p[1],0) for p in yoy_pairs)
    q1_c25  = sum(macro.get(m,0) for m in months_2025 if m.endswith(("-01","-02","-03")))
    q1_c26  = sum(macro.get(m,0) for m in months_2026_real if m.endswith(("-01","-02","-03")))

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

    return {
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "kam": KAM_CODE,
        "months": all_months,
        "month_labels": [MONTH_LABELS.get(m, m) for m in all_months],
        "hist_count": len(hist_months),
        "forecast_months": sorted(forecast_months),
        "cartera": cartera,
        "clients": clients,
    }

# ── Main ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    files = sys.argv[1:] if len(sys.argv) > 1 else []
    if not files:
        files = list(Path(".").glob("*.xlsx")) + list(Path("data").glob("*.xlsx"))
        if not files:
            print("USO: python procesar_datos.py archivo.xlsx"); sys.exit(1)

    print(f"\n{'='*52}")
    print(f"  KAM Dashboard — Procesando datos")
    print(f"{'='*52}\n")

    df = load_files(files)
    print(f"\n  Registros totales: {len(df)}")
    print(f"  Clientes únicos:   {df['Cliente'].nunique()}")

    data = build_dataset(df)
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",",":"))

    kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"\n  ✓ Generado: {OUTPUT_FILE}  ({kb:.1f} KB)")
    print(f"  ✓ Clientes: {len(data['clients'])}")
    print(f"  ✓ Forecast H2 2026: ${data['cartera']['forecast_total']:,.0f} CLP\n")
    counts = {}
    for c in data["clients"].values():
        counts[c["status"]] = counts.get(c["status"],0) + 1
    for s,n in sorted(counts.items()):
        print(f"    {s:12s}: {n}")
    print()
