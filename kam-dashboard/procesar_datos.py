"""
procesar_datos.py
=================
Convierte tu Excel de ventas en data.json para el dashboard.

USO:
    python procesar_datos.py LETRITAS_2025_2026.xlsx

El script detecta automáticamente la hoja CUADRATURAS y filtra por KAM = BC.
Genera docs/data.json que el dashboard lee.
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
KAM_CODE      = "BC"          # Tus iniciales en el archivo
SHEET_NAME    = "CUADRATURAS"
REF_DATE      = pd.Timestamp("2026-06-01")   # Fecha de referencia para calcular días
OUTPUT_FILE   = Path("docs/data.json")


# ── Helpers ───────────────────────────────────────────────────────────────────
def pct(a, b):
    if not b or b == 0:
        return None
    return round((a - b) / b * 100, 1)

def fmt_month(m):
    """'2025-05' → 'May 25'"""
    if not m:
        return ""
    y, mo = m.split("-")
    names = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
    return f"{names[int(mo)]} {y[2:]}"


# ── Carga y limpieza ──────────────────────────────────────────────────────────
def load_files(paths):
    """Carga uno o varios Excel y los combina. Detecta columna de precio automáticamente."""
    frames = []
    for p in paths:
        p = Path(p)
        if not p.exists():
            print(f"  ✗ No encontré: {p}")
            continue
        print(f"  ✓ Cargando {p.name}...")
        df = pd.read_excel(p, sheet_name=SHEET_NAME)
        df.columns = [c.strip() for c in df.columns]

        # Columna KAM (puede llamarse 'Kam' o 'Codigo Vendedor')
        if "Kam" in df.columns:
            df = df[df["Kam"] == KAM_CODE].copy()
        elif "Codigo Vendedor" in df.columns:
            df = df[df["Codigo Vendedor"] == KAM_CODE].copy()
        else:
            print(f"  ⚠ No encontré columna KAM en {p.name}, saltando.")
            continue

        # Columna precio (puede llamarse 'Precio' o 'Prrecio')
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
    # Si hay dos archivos con mayo solapado, el más reciente (último) gana
    combined = combined.drop_duplicates(subset=["Cliente", "Mes", "Precio"], keep="last")
    return combined


# ── Construcción del dataset ──────────────────────────────────────────────────
def build_dataset(df):
    monthly = df.groupby(["Cliente", "Mes"])["Precio"].sum().reset_index()
    all_months = sorted(monthly["Mes"].unique().tolist())

    # Períodos fijos comparables (ajusta si cambias el rango de fechas)
    MONTHS_2025 = [m for m in all_months if m.startswith("2025")]
    MONTHS_2026 = [m for m in all_months if m.startswith("2026")]

    # Índices para comparaciones YoY (solo meses que existen en ambos años)
    common_months = []
    for m26 in MONTHS_2026:
        m25 = "2025-" + m26.split("-")[1]
        if m25 in all_months:
            common_months.append((m25, m26))

    def sum_months(vals_dict, month_list):
        return sum(vals_dict.get(m, 0) for m in month_list)

    clients = {}
    for client, grp in monthly.groupby("Cliente"):
        vals_dict = dict(zip(grp["Mes"], grp["Precio"]))
        vals = [int(vals_dict.get(m, 0)) for m in all_months]

        total = sum(vals)
        active_months = [all_months[i] for i, v in enumerate(vals) if v > 0]
        freq = len(active_months)
        last_month  = active_months[-1] if active_months else None
        first_month = active_months[0]  if active_months else None

        days_since = int((REF_DATE - pd.Timestamp(last_month + "-28")).days) if last_month else 999

        # Revenue por período
        rev_2025 = sum_months(vals_dict, MONTHS_2025)
        rev_2026 = sum_months(vals_dict, MONTHS_2026)

        # YTD: primeros N meses comunes entre 2025 y 2026
        ytd_months_25 = [p[0] for p in common_months]
        ytd_months_26 = [p[1] for p in common_months]
        ytd25 = sum_months(vals_dict, ytd_months_25)
        ytd26 = sum_months(vals_dict, ytd_months_26)
        ytd_pct_val = pct(ytd26, ytd25)

        # Mensual YoY
        mensual = []
        for m25, m26 in common_months:
            v25 = int(vals_dict.get(m25, 0))
            v26 = int(vals_dict.get(m26, 0))
            mensual.append({
                "lbl": fmt_month(m25).split()[0],  # "Ene", "Feb"...
                "v25": v25, "v26": v26,
                "pct": pct(v26, v25)
            })

        # Trimestral YoY (Q1 y Q2 parcial)
        q1_25 = [m for m in MONTHS_2025 if m.endswith(("-01","-02","-03"))]
        q1_26 = [m for m in MONTHS_2026 if m.endswith(("-01","-02","-03"))]
        q2_25 = [m for m in MONTHS_2025 if m.endswith(("-04","-05"))]
        q2_26 = [m for m in MONTHS_2026 if m.endswith(("-04","-05"))]
        trimestral = []
        if q1_25 and q1_26:
            a, b = sum_months(vals_dict, q1_25), sum_months(vals_dict, q1_26)
            trimestral.append({"lbl": "Q1", "v25": int(a), "v26": int(b), "pct": pct(b, a)})
        if q2_25 and q2_26:
            a, b = sum_months(vals_dict, q2_25), sum_months(vals_dict, q2_26)
            trimestral.append({"lbl": "Abr–May", "v25": int(a), "v26": int(b), "pct": pct(b, a)})

        # Status
        last3_months = all_months[-3:] if len(all_months) >= 3 else all_months
        has_recent   = any(vals_dict.get(m, 0) > 0 for m in last3_months)
        has_2026     = rev_2026 > 0
        has_2025     = rev_2025 > 0
        mid_months   = all_months[max(0,len(all_months)-7):-3] if len(all_months) > 6 else []
        has_mid      = any(vals_dict.get(m, 0) > 0 for m in mid_months)

        if has_recent:
            status = "nuevo" if not has_2025 else "activo"
        elif has_2026 or has_mid:
            status = "en riesgo"
        else:
            status = "churn"

        clients[client] = {
            "total":    total,
            "freq":     freq,
            "primera":  first_month,
            "ultima":   last_month,
            "dias":     days_since,
            "status":   status,
            "rev_2025": int(rev_2025),
            "rev_2026": int(rev_2026),
            "ytd25":    int(ytd25),
            "ytd26":    int(ytd26),
            "ytd_pct":  ytd_pct_val,
            "vals":     vals,
            "mensual":  mensual,
            "trimestral": trimestral,
            "q1_pct":   trimestral[0]["pct"] if trimestral else None,
            "may_pct":  mensual[-1]["pct"] if mensual else None,
        }

    # Cartera agregada
    macro_vals = [
        int(monthly[monthly["Mes"] == m]["Precio"].sum()) for m in all_months
    ]
    macro_dict = dict(zip(all_months, macro_vals))

    cartera_mensual = [
        {"lbl": fmt_month(m25).split()[0],
         "v25": int(macro_dict.get(m25, 0)),
         "v26": int(macro_dict.get(m26, 0)),
         "pct": pct(macro_dict.get(m26, 0), macro_dict.get(m25, 0))}
        for m25, m26 in common_months
    ]

    q1_cartera_25 = sum(macro_dict.get(m,0) for m in all_months if m.startswith("2025") and m.endswith(("-01","-02","-03")))
    q1_cartera_26 = sum(macro_dict.get(m,0) for m in all_months if m.startswith("2026") and m.endswith(("-01","-02","-03")))
    ytd_c25 = sum(macro_dict.get(p[0],0) for p in common_months)
    ytd_c26 = sum(macro_dict.get(p[1],0) for p in common_months)

    cartera = {
        "total":       sum(macro_vals),
        "vals":        macro_vals,
        "mensual":     cartera_mensual,
        "trimestral":  [
            {"lbl":"Q1","v25":int(q1_cartera_25),"v26":int(q1_cartera_26),"pct":pct(q1_cartera_26,q1_cartera_25)}
        ],
        "ytd25":       int(ytd_c25),
        "ytd26":       int(ytd_c26),
        "ytd_pct":     pct(ytd_c26, ytd_c25),
        "q1_pct":      pct(q1_cartera_26, q1_cartera_25),
        "may_pct":     cartera_mensual[-1]["pct"] if cartera_mensual else None,
    }

    return {
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "kam":          KAM_CODE,
        "months":       all_months,
        "month_labels": [fmt_month(m) for m in all_months],
        "cartera":      cartera,
        "clients":      clients,
    }


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    files = sys.argv[1:] if len(sys.argv) > 1 else []
    if not files:
        # Auto-detectar Excel en la carpeta actual
        files = list(Path(".").glob("*.xlsx")) + list(Path("data").glob("*.xlsx"))
        if not files:
            print("USO: python procesar_datos.py archivo1.xlsx [archivo2.xlsx ...]")
            sys.exit(1)
        print(f"Auto-detectados: {[str(f) for f in files]}")

    print(f"\n{'='*50}")
    print(f"  KAM Dashboard — Procesando datos")
    print(f"{'='*50}\n")

    df = load_files(files)
    print(f"\n  Total registros combinados: {len(df)}")
    print(f"  Clientes únicos: {df['Cliente'].nunique()}")
    print(f"  Rango de meses: {df['Mes'].min()} → {df['Mes'].max()}\n")

    data = build_dataset(df)

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

    size_kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"  ✓ Generado: {OUTPUT_FILE}  ({size_kb:.1f} KB)")
    print(f"  ✓ Clientes procesados: {len(data['clients'])}")
    print(f"  ✓ Meses: {data['months'][0]} → {data['months'][-1]}")

    status_counts = {}
    for c in data["clients"].values():
        s = c["status"]
        status_counts[s] = status_counts.get(s, 0) + 1
    print(f"\n  Estado clientes:")
    for s, n in sorted(status_counts.items()):
        print(f"    {s:12s}: {n}")
    print(f"\n  ✓ Listo. Sube los cambios a GitHub para actualizar el dashboard.\n")
