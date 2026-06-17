"""
procesar_datos.py
=================
Convierte tu Excel de ventas en data.json para el dashboard.
Incluye forecast Jun–Dic 2026 basado en mismo mes 2025.

USO:
    python procesar_datos.py LETRITAS_2025_2026.xlsx [LETRITAS_MAYO_2026.xlsx ...]
"""

import sys
import json
import math
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERROR: Falta pandas. Instala con:  pip install pandas openpyxl")
    sys.exit(1)

# ── Configuración ─────────────────────────────────────────────────────────────
KAM_CODE         = "BC"
SHEET_NAME       = "CUADRATURAS"
REF_DATE         = pd.Timestamp("2026-06-01")
OUTPUT_FILE      = Path("docs/data.json")
FORECAST_GROWTH  = 1.35   # +35% sobre el mismo mes del año anterior

# Meses reales (histórico)
HIST_MONTHS = [
    "2025-01","2025-02","2025-03","2025-04","2025-05","2025-06",
    "2025-07","2025-08","2025-09","2025-10","2025-11","2025-12",
    "2026-01","2026-02","2026-03","2026-04","2026-05"
]

# Meses de forecast (Jun–Dic 2026) → basado en mismo mes 2025
FORECAST_MONTHS = [
    "2026-06","2026-07","2026-08","2026-09","2026-10","2026-11","2026-12"
]
FORECAST_BASE = {
    "2026-06": "2025-06",
    "2026-07": "2025-07",
    "2026-08": "2025-08",
    "2026-09": "2025-09",
    "2026-10": "2025-10",
    "2026-11": "2025-11",
    "2026-12": "2025-12",
}

ALL_MONTHS  = HIST_MONTHS + FORECAST_MONTHS
ALL_LABELS  = [
    "Ene 25","Feb 25","Mar 25","Abr 25","May 25","Jun 25",
    "Jul 25","Ago 25","Sep 25","Oct 25","Nov 25","Dic 25",
    "Ene 26","Feb 26","Mar 26","Abr 26","May 26",
    "Jun 26*","Jul 26*","Ago 26*","Sep 26*","Oct 26*","Nov 26*","Dic 26*"
]

# ── Helpers ───────────────────────────────────────────────────────────────────
def pct(a, b):
    if not b or b == 0: return None
    return round((a - b) / b * 100, 1)

def fmt_month(m):
    if not m: return ""
    y, mo = m.split("-")
    names = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    return f"{names[int(mo)]} {y[2:]}"

# ── Carga ─────────────────────────────────────────────────────────────────────
def load_files(paths):
    frames = []
    for p in paths:
        p = Path(p)
        if not p.exists():
            print(f"  ✗ No encontré: {p}")
            continue
        print(f"  ✓ Cargando {p.name}...")
        df = pd.read_excel(p, sheet_name=SHEET_NAME)
        df.columns = [c.strip() for c in df.columns]

        if "Kam" in df.columns:
            df = df[df["Kam"] == KAM_CODE].copy()
        elif "Codigo Vendedor" in df.columns:
            df = df[df["Codigo Vendedor"] == KAM_CODE].copy()
        else:
            print(f"  ⚠ No encontré columna KAM en {p.name}")
            continue

        for col in ["Precio", "Prrecio", "precio", "PRECIO"]:
            if col in df.columns:
                df = df.rename(columns={col: "Precio"})
                break

        df = df[df["Precio"] > 0].copy()
        df["Fecha Emision"] = pd.to_datetime(df["Fecha Emision"])
        df["Mes"] = df["Fecha Emision"].dt.to_period("M").astype(str)
        frames.append(df[["Cliente", "Precio", "Mes"]])
        print(f"    → {len(df)} registros BC")

    if not frames:
        print("ERROR: No se cargó ningún archivo.")
        sys.exit(1)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.drop_duplicates(subset=["Cliente", "Mes", "Precio"], keep="last")
    return combined

# ── Dataset ───────────────────────────────────────────────────────────────────
def build_dataset(df):
    monthly = df.groupby(["Cliente", "Mes"])["Precio"].sum().reset_index()

    clients = {}
    for client, grp in monthly.groupby("Cliente"):
        vals_dict = dict(zip(grp["Mes"], grp["Precio"]))

        # Histórico real (17 meses)
        hist_vals = [int(vals_dict.get(m, 0)) for m in HIST_MONTHS]

        # Forecast (7 meses): mismo mes año anterior
        forecast_vals = []
        forecast_has_data = False
        for m26 in FORECAST_MONTHS:
            m25 = FORECAST_BASE[m26]
            v = int(vals_dict.get(m25, 0) * FORECAST_GROWTH)
            forecast_vals.append(v)
            if v > 0:
                forecast_has_data = True

        # Todos los valores (17 hist + 7 forecast)
        all_vals = hist_vals + forecast_vals

        total      = sum(hist_vals)
        active_m   = [HIST_MONTHS[i] for i, v in enumerate(hist_vals) if v > 0]
        freq       = len(active_m)
        last_month = active_m[-1] if active_m else None
        first_month= active_m[0]  if active_m else None
        days_since = int((REF_DATE - pd.Timestamp(last_month + "-28")).days) if last_month else 999

        # Revenue splits
        rev_2025 = sum(hist_vals[0:12])
        rev_2026 = sum(hist_vals[12:])

        # YTD Ene–May comparado
        ytd25 = sum(hist_vals[0:5])
        ytd26 = sum(hist_vals[12:17])
        ytd_pct_val = pct(ytd26, ytd25)

        # Promedio mensual año actual (2026): solo meses con compra
        meses_activos_2026 = [v for v in hist_vals[12:] if v > 0]
        avg_2026 = int(sum(meses_activos_2026) / len(meses_activos_2026)) if meses_activos_2026 else 0

        # Forecast totales
        forecast_total    = sum(forecast_vals)
        forecast_h2_2025  = sum(int(vals_dict.get(FORECAST_BASE[m], 0)) for m in FORECAST_MONTHS)

        # Comparaciones mensuales YoY (solo meses con datos en ambos años)
        mensual = []
        for m25, m26 in [("2025-01","2026-01"),("2025-02","2026-02"),("2025-03","2026-03"),
                          ("2025-04","2026-04"),("2025-05","2026-05")]:
            v25 = int(vals_dict.get(m25, 0))
            v26 = int(vals_dict.get(m26, 0))
            mensual.append({"lbl": fmt_month(m25).split()[0], "v25": v25, "v26": v26, "pct": pct(v26, v25)})

        # Trimestral
        q1_25 = sum(vals_dict.get(m, 0) for m in ["2025-01","2025-02","2025-03"])
        q1_26 = sum(vals_dict.get(m, 0) for m in ["2026-01","2026-02","2026-03"])
        trimestral = [{"lbl":"Q1","v25":int(q1_25),"v26":int(q1_26),"pct":pct(q1_26,q1_25)}]

        # Status
        last3 = HIST_MONTHS[-3:]
        has_recent = any(vals_dict.get(m, 0) > 0 for m in last3)
        has_2026   = rev_2026 > 0
        has_2025   = rev_2025 > 0
        mid_months = HIST_MONTHS[max(0, len(HIST_MONTHS)-7):-3]
        has_mid    = any(vals_dict.get(m, 0) > 0 for m in mid_months)

        if has_recent:
            status = "nuevo" if not has_2025 else "activo"
        elif has_2026 or has_mid:
            status = "en riesgo"
        else:
            status = "churn"

        clients[client] = {
            "total":           total,
            "freq":            freq,
            "primera":         first_month,
            "ultima":          last_month,
            "dias":            days_since,
            "status":          status,
            "rev_2025":        int(rev_2025),
            "rev_2026":        int(rev_2026),
            "ytd25":           int(ytd25),
            "ytd26":           int(ytd26),
            "ytd_pct":         ytd_pct_val,
            "avg_2026":        avg_2026,
            "vals":            all_vals,          # 24 valores (17 hist + 7 forecast)
            "hist_count":      len(HIST_MONTHS),  # = 17, para saber dónde empieza el forecast
            "forecast_vals":   forecast_vals,
            "forecast_total":  int(forecast_total),
            "forecast_has_data": forecast_has_data,
            "mensual":         mensual,
            "trimestral":      trimestral,
            "q1_pct":          trimestral[0]["pct"],
            "may_pct":         mensual[-1]["pct"] if mensual else None,
        }

    # Cartera agregada
    macro_dict = {}
    for m in ALL_MONTHS:
        macro_dict[m] = int(monthly[monthly["Mes"] == m]["Precio"].sum()) if m in monthly["Mes"].values else 0

    # Forecast cartera = suma de H2 2025 × multiplicador de crecimiento
    for m26 in FORECAST_MONTHS:
        m25 = FORECAST_BASE[m26]
        macro_dict[m26] = int(monthly[monthly["Mes"] == m25]["Precio"].sum() * FORECAST_GROWTH)

    macro_hist_vals     = [macro_dict.get(m, 0) for m in HIST_MONTHS]
    macro_forecast_vals = [macro_dict.get(m, 0) for m in FORECAST_MONTHS]
    macro_all_vals      = macro_hist_vals + macro_forecast_vals

    ytd_c25 = sum(macro_dict.get(m, 0) for m in ["2025-01","2025-02","2025-03","2025-04","2025-05"])
    ytd_c26 = sum(macro_dict.get(m, 0) for m in ["2026-01","2026-02","2026-03","2026-04","2026-05"])
    q1_c25  = sum(macro_dict.get(m, 0) for m in ["2025-01","2025-02","2025-03"])
    q1_c26  = sum(macro_dict.get(m, 0) for m in ["2026-01","2026-02","2026-03"])

    cartera = {
        "total":            sum(macro_hist_vals),
        "vals":             macro_all_vals,
        "hist_count":       len(HIST_MONTHS),
        "forecast_vals":    macro_forecast_vals,
        "forecast_total":   sum(macro_forecast_vals),
        "mensual": [
            {"lbl": fmt_month(m25).split()[0],
             "v25": macro_dict.get(m25, 0),
             "v26": macro_dict.get(m26, 0),
             "pct": pct(macro_dict.get(m26, 0), macro_dict.get(m25, 0))}
            for m25, m26 in [("2025-01","2026-01"),("2025-02","2026-02"),("2025-03","2026-03"),
                             ("2025-04","2026-04"),("2025-05","2026-05")]
        ],
        "trimestral": [{"lbl":"Q1","v25":int(q1_c25),"v26":int(q1_c26),"pct":pct(q1_c26,q1_c25)}],
        "ytd25":    int(ytd_c25),
        "ytd26":    int(ytd_c26),
        "ytd_pct":  pct(ytd_c26, ytd_c25),
        "q1_pct":   pct(q1_c26, q1_c25),
        "may_pct":  pct(macro_dict.get("2026-05",0), macro_dict.get("2025-05",0)),
    }

    return {
        "generated_at":  pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "kam":           KAM_CODE,
        "months":        ALL_MONTHS,
        "month_labels":  ALL_LABELS,
        "hist_count":    len(HIST_MONTHS),
        "forecast_months": FORECAST_MONTHS,
        "cartera":       cartera,
        "clients":       clients,
    }

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    files = sys.argv[1:] if len(sys.argv) > 1 else []
    if not files:
        files = list(Path(".").glob("*.xlsx")) + list(Path("data").glob("*.xlsx"))
        if not files:
            print("USO: python procesar_datos.py archivo1.xlsx [archivo2.xlsx ...]")
            sys.exit(1)

    print(f"\n{'='*52}")
    print(f"  KAM Dashboard — Procesando datos + forecast")
    print(f"{'='*52}\n")

    df = load_files(files)
    print(f"\n  Total registros: {len(df)}")
    print(f"  Clientes únicos: {df['Cliente'].nunique()}")
    print(f"  Rango histórico: {df['Mes'].min()} → {df['Mes'].max()}")
    print(f"  Forecast:        2026-06 → 2026-12 (= mismo mes 2025)\n")

    data = build_dataset(df)

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

    size_kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"  ✓ Generado: {OUTPUT_FILE}  ({size_kb:.1f} KB)")
    print(f"  ✓ Clientes: {len(data['clients'])}")
    print(f"  ✓ Forecast cartera H2 2026: ${data['cartera']['forecast_total']:,.0f} CLP\n")

    status_counts = {}
    for c in data["clients"].values():
        s = c["status"]
        status_counts[s] = status_counts.get(s, 0) + 1
    for s, n in sorted(status_counts.items()):
        print(f"    {s:12s}: {n}")
    print()
