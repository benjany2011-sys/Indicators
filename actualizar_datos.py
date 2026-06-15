"""
Actualización diaria: inflación (FRED, sin llave) + Henry Hub (EIA, con llave).
Pensado para correr en GitHub Actions.

La llave de la EIA NO va aquí: se lee de una variable de entorno que en
GitHub se configura como "Secret" (EIA_API_KEY).
"""

import os
from datetime import datetime

import requests
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # backend sin ventana (corre en servidor)
import matplotlib.pyplot as plt

# Las llaves llegan desde los Secrets de GitHub (no se escriben en el código)
EIA_API_KEY = os.environ.get("EIA_API_KEY", "")
FRED_API_KEY = os.environ.get("FRED_API_KEY", "")

FECHA_INICIO = "2021-01-01"
CARPETA_SALIDA = "resultados"


# ---------------------------------------------------------------------------
# EIA: Henry Hub diario (Natural Gas Spot Price, Daily, $/MMBtu)
# ---------------------------------------------------------------------------
def obtener_henry_hub(fecha_inicio=FECHA_INICIO):
    if not EIA_API_KEY:
        raise RuntimeError("Falta el Secret EIA_API_KEY en GitHub")

    url = "https://api.eia.gov/v2/seriesid/NG.RNGWHHD.D"
    params = {
        "api_key": EIA_API_KEY,
        "start": fecha_inicio,
        "sort[0][column]": "period",
        "sort[0][direction]": "asc",
    }
    r = requests.get(url, params=params, timeout=30)
    r.raise_for_status()
    registros = r.json()["response"]["data"]

    df = pd.DataFrame(registros)
    df["fecha"] = pd.to_datetime(df["period"])
    df["precio"] = pd.to_numeric(df["value"], errors="coerce")
    df = df[["fecha", "precio"]].dropna().sort_values("fecha").reset_index(drop=True)
    return df


# ---------------------------------------------------------------------------
# FRED: CPI -> inflación interanual (YoY %).  Vía API oficial, con llave.
# ---------------------------------------------------------------------------
def obtener_inflacion(fecha_inicio=FECHA_INICIO):
    if not FRED_API_KEY:
        raise RuntimeError("Falta el Secret FRED_API_KEY en GitHub")

    url = "https://api.stlouisfed.org/fred/series/observations"
    params = {
        "series_id": "CPIAUCSL",  # CPI urbano, ajustado estacionalmente, mensual
        "api_key": FRED_API_KEY,
        "file_type": "json",
    }
    r = requests.get(url, params=params, timeout=30)
    r.raise_for_status()
    obs = r.json()["observations"]

    df = pd.DataFrame(obs)
    df["fecha"] = pd.to_datetime(df["date"])
    df["cpi"] = pd.to_numeric(df["value"], errors="coerce")
    df = df[["fecha", "cpi"]].dropna().sort_values("fecha").reset_index(drop=True)

    # Inflación interanual: cambio % contra el mismo mes del año anterior
    df["inflacion_yoy"] = df["cpi"].pct_change(12) * 100
    df = df[df["fecha"] >= pd.to_datetime(fecha_inicio)].reset_index(drop=True)
    return df


# ---------------------------------------------------------------------------
# Gráfica (se guarda como imagen)
# ---------------------------------------------------------------------------
def graficar(hh, infl, ruta_png):
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8), sharex=True)

    ax1.plot(hh["fecha"], hh["precio"], color="tab:blue", linewidth=0.9)
    ax1.set_title("Henry Hub – precio diario spot ($/MMBtu)")
    ax1.set_ylabel("$/MMBtu")
    ax1.grid(alpha=0.3)

    ax2.plot(infl["fecha"], infl["inflacion_yoy"], color="tab:red", linewidth=1.4)
    ax2.axhline(2, color="gray", linestyle="--", linewidth=0.8, label="Meta 2%")
    ax2.set_title("Inflación interanual EE.UU. (CPI YoY %)")
    ax2.set_ylabel("% YoY")
    ax2.legend()
    ax2.grid(alpha=0.3)

    plt.tight_layout()
    fig.savefig(ruta_png, dpi=120)
    plt.close(fig)


if __name__ == "__main__":
    os.makedirs(CARPETA_SALIDA, exist_ok=True)
    hoy = f"{datetime.now():%Y-%m-%d}"

    hh = obtener_henry_hub()
    infl = obtener_inflacion()

    # Datos en Excel (dos hojas), siempre como "más reciente"
    ruta_xlsx = os.path.join(CARPETA_SALIDA, "datos_mas_reciente.xlsx")
    with pd.ExcelWriter(ruta_xlsx, engine="openpyxl") as writer:
        hh.to_excel(writer, sheet_name="HenryHub_diario", index=False)
        infl.to_excel(writer, sheet_name="Inflacion_CPI", index=False)

    # Gráfica
    graficar(hh, infl, os.path.join(CARPETA_SALIDA, "grafico_mas_reciente.png"))

    # Resumen visible en el log de GitHub
    precio = hh["precio"].iloc[-1]
    ult = infl.dropna(subset=["inflacion_yoy"]).iloc[-1]
    print(f"OK {hoy}")
    print(f"Henry Hub: ${precio:.2f}/MMBtu ({hh['fecha'].iloc[-1].date()})")
    print(f"Inflacion: {ult['inflacion_yoy']:.2f}% ({ult['fecha'].date()})")
