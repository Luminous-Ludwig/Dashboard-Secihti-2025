"""
Preprocesa el Excel de becas CONAHCYT 2025 y genera data.json columnar
para que el dashboard lo consuma en el navegador.
"""
import json
import math
from pathlib import Path

import pandas as pd

SRC = Path("C:/Users/Luminous/Desktop/Workspace/becas 2025.xlsx")
OUT = Path(__file__).parent / "data.json"


def clean_str(v):
    if pd.isna(v):
        return "NO ESPECIFICADO"
    return str(v).strip().upper()


def main():
    df = pd.read_excel(SRC, sheet_name=0)

    rename = {
        "INSTITUCIÓN": "institucion",
        "PAÍS": "pais",
        "ÁREA DEL CONOCIMIENTO": "area",
        "NIVEL DE ESTUDIOS": "nivel",
        "MODALIDAD": "modalidad",
        "ENTIDAD": "entidad",
        "CONVOCATORIA": "convocatoria",
        "INICIO DE BECA": "inicio",
        "IMPORTE TOTAL PAGADO ENERO-DICIEMBRE": "importe",
    }
    # Soporta variaciones de encoding en headers
    cols_map = {}
    for c in df.columns:
        key = c.strip()
        for k, v in rename.items():
            if k in key or key in k:
                cols_map[c] = v
                break
    df = df.rename(columns=cols_map)

    needed = ["institucion", "pais", "area", "nivel", "modalidad",
              "entidad", "convocatoria", "inicio", "importe"]
    for col in needed:
        if col not in df.columns:
            raise SystemExit(f"Falta columna {col}. Encontradas: {list(df.columns)}")

    df["inicio"] = pd.to_datetime(df["inicio"], errors="coerce")
    df["anio_inicio"] = df["inicio"].dt.year.fillna(0).astype(int)

    for col in ["institucion", "pais", "area", "nivel", "modalidad", "entidad", "convocatoria"]:
        df[col] = df[col].map(clean_str)

    df["importe"] = pd.to_numeric(df["importe"], errors="coerce").fillna(0.0)

    lookups = {}
    encoded = {}
    for col in ["modalidad", "nivel", "institucion", "pais", "entidad", "area", "convocatoria"]:
        values = sorted(df[col].unique().tolist())
        idx = {v: i for i, v in enumerate(values)}
        lookups[col] = values
        encoded[col] = df[col].map(idx).astype(int).tolist()

    encoded["anio_inicio"] = df["anio_inicio"].tolist()
    encoded["importe"] = [round(float(x), 2) for x in df["importe"].tolist()]

    payload = {
        "meta": {
            "total_rows": int(len(df)),
            "total_importe": round(float(df["importe"].sum()), 2),
            "anios": sorted([int(y) for y in df["anio_inicio"].unique() if y > 0]),
        },
        "lookups": lookups,
        "rows": encoded,
    }

    with OUT.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))

    size_mb = OUT.stat().st_size / 1024 / 1024
    print(f"OK: {OUT.name} generado")
    print(f"  filas: {len(df):,}")
    print(f"  tamaño: {size_mb:.2f} MB")
    print(f"  países: {len(lookups['pais'])}")
    print(f"  instituciones: {len(lookups['institucion'])}")
    print(f"  niveles: {len(lookups['nivel'])}")
    print(f"  áreas: {len(lookups['area'])}")
    print(f"  modalidades: {len(lookups['modalidad'])}")


if __name__ == "__main__":
    main()
