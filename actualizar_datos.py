"""
actualizar_datos.py
===================================================================
Descarga y consolida en un solo Excel (bien formateado) las series:

  - Henry Hub (gas natural, spot diario)        -> EIA
  - Inflación de EE. UU. (CPI y variación YoY)  -> FRED
  - S&P 500 (índice diario)                     -> FRED
  - WTI (petróleo, spot diario)                 -> FRED
  - Brent (petróleo, spot diario)               -> FRED
  - Tipo de cambio de 10 divisas vs. USD        -> Frankfurter (BCE)
  - Acereras (acciones, índices base 100)        -> Twelve Data

Todas las series diarias arrancan en 2021-01-01.

Las divisas vienen de Frankfurter (tipos de referencia del Banco Central
Europeo). Con base=USD salen ya como "unidades por 1 USD" (18 pesos, 157
yenes, etc.), sin necesidad de invertir, y se publican el mismo día hábil
(~16:00 hora de Europa central), más fresco que FRED.

Cada vez que corre crea (junto al script) una carpeta `resultados` con:
  - mercados_AAAA-MM-DD.xlsx  y  mercados_reciente.xlsx
  - graficos_AAAA-MM-DD.png   y  graficos_reciente.png
  - registro.log

-------------------------------------------------------------------
Requisitos (una sola vez):
    pip install requests pandas matplotlib openpyxl

Llaves (EIA, FRED y Twelve Data):
  - En tu PC: archivo `.env` en la misma carpeta, con
        EIA_API_KEY=tu_llave_de_eia
        FRED_API_KEY=tu_llave_de_fred
        TWELVEDATA_API_KEY=tu_llave_de_twelvedata
    (necesita además `pip install python-dotenv`)
  - En GitHub Actions: como Secrets del repo; el workflow las pasa
    como variables de entorno. No hace falta .env ni python-dotenv.
  Frankfurter no necesita llave. Si falta TWELVEDATA_API_KEY, el panel
  corre igual pero la sección de acereras se omite.
===================================================================
"""

import os
import sys
import json
import time
import datetime as dt
from pathlib import Path

import requests
import pandas as pd
import matplotlib
matplotlib.use("Agg")                 # sin ventana, solo guarda PNG
import matplotlib.pyplot as plt

# python-dotenv es opcional: en tu PC lee el .env; en GitHub Actions
# las llaves vienen de los Secrets y este import simplemente se omite.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ModuleNotFoundError:
    pass

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

# ------------------------------------------------------------------
# 0) Configuración
# ------------------------------------------------------------------
EIA_API_KEY = os.getenv("EIA_API_KEY")
FRED_API_KEY = os.getenv("FRED_API_KEY")
# Acereras (acciones) vienen ahora de Twelve Data. El workflow ya inyecta este
# Secret como variable de entorno; en tu PC va en el .env como las otras.
TWELVEDATA_API_KEY = os.getenv("TWELVEDATA_API_KEY")

FECHA_INICIO = "2021-01-01"           # arranque común de todas las series

CARPETA = Path(__file__).resolve().parent / "resultados"
CARPETA.mkdir(exist_ok=True)
HOY = dt.date.today().isoformat()
LOG = CARPETA / "registro.log"

# Henry Hub en $/MMBtu (viene de EIA)
COL_HH = "Henry Hub ($/MMBtu)"
FMT_HH = "#,##0.00"

# Series diarias de FRED: id -> (nombre de columna, formato numérico)
FRED_DIARIAS = {
    "SP500":        ("S&P 500",       "#,##0.00"),
    "DCOILWTICO":   ("WTI ($/bbl)",   "#,##0.00"),
    "DCOILBRENTEU": ("Brent ($/bbl)", "#,##0.00"),
}

# Divisas vía Frankfurter (BCE). base=USD -> "unidades por 1 USD".
# código ISO -> (nombre de columna, formato numérico)
MONEDAS = {
    "EUR": ("EUR por USD", "#,##0.0000"),
    "JPY": ("JPY por USD", "#,##0.00"),
    "CNY": ("CNY por USD", "#,##0.0000"),
    "GBP": ("GBP por USD", "#,##0.0000"),
    "MXN": ("MXN por USD", "#,##0.0000"),
    "CAD": ("CAD por USD", "#,##0.0000"),
    "CHF": ("CHF por USD", "#,##0.0000"),
    "AUD": ("AUD por USD", "#,##0.0000"),
    "INR": ("INR por USD", "#,##0.00"),
    "KRW": ("KRW por USD", "#,##0.00"),
}

# Saneamiento: rangos realistas por serie. Cualquier dato fuera del rango se
# considera erróneo (picos falsos de la fuente) y se marca como vacío; la línea
# de la gráfica simplemente lo salta. Nota: el límite alto de Henry Hub (25)
# conserva el pico real del invierno 2021 (tormenta Uri) pero descarta valores
# imposibles como ~30. Ajusta estos números si algún día resultan muy estrechos.
RANGOS_VALIDOS = {
    COL_HH:           (0.5, 25.0),    # Henry Hub $/MMBtu
    "WTI ($/bbl)":    (5.0, 250.0),   # WTI $/bbl
    "Brent ($/bbl)":  (5.0, 250.0),   # Brent $/bbl
}


def sanear(df, col):
    """Marca como vacío los valores de 'col' fuera de su rango realista."""
    if col not in RANGOS_VALIDOS or col not in df.columns:
        return df
    lo, hi = RANGOS_VALIDOS[col]
    mask = df[col].notna() & ((df[col] < lo) | (df[col] > hi))
    n = int(mask.sum())
    if n:
        log(f"  Saneando {col}: {n} valor(es) fuera de [{lo}, {hi}] -> vacío")
        df.loc[mask, col] = float("nan")
    return df


# Acereras que cotizan en bolsa, vía Twelve Data (símbolo -> nombre).
# En el plan GRATIS de Twelve Data solo entran de forma fiable las acciones
# listadas en EE. UU. (NYSE/Nasdaq). Por eso para las internacionales se usa su
# ADR/listado en EE. UU. Las que solo cotizan fuera (Tata y JSW en India) y las
# OTC inciertas se dejan abajo comentadas: si subes a un plan de pago, las
# descomentas y listo. Si un símbolo no responde, el script lo salta y lo
# reporta en "fallaron" (el índice equal-weight tolera entradas/salidas).
ACERERAS_MUNDIAL = {
    "MT":     "ArcelorMittal",     # NYSE
    "PKX":    "POSCO",             # NYSE (ADR)
    "TX":     "Ternium",           # NYSE
    "GGB":    "Gerdau",            # NYSE
    "NPSCY":  "Nippon Steel",      # OTC (ADR) — puede requerir plan de pago
    "TKAMY":  "thyssenkrupp",      # OTC (ADR) — puede requerir plan de pago
    "SSAAY":  "SSAB",              # OTC (ADR) — puede requerir plan de pago
    # "TATASTEEL:NSE": "Tata Steel",  # India: requiere plan de pago en Twelve Data
    # "JSWSTEEL:NSE":  "JSW Steel",   # India: requiere plan de pago en Twelve Data
}
ACERERAS_EEUU = {
    "NUE":   "Nucor",
    "STLD":  "Steel Dynamics",
    "CLF":   "Cleveland-Cliffs",
    "CMC":   "Commercial Metals",
    "RS":    "Reliance",
    "WOR":   "Worthington",
    "ATI":   "ATI",
    "CRS":   "Carpenter Technology",
    "ZEUS":  "Olympic Steel",
    # "X":   "U.S. Steel",  # deslistada el 18-jun-2025 (compra de Nippon); ya no cotiza
}

# Para cada divisa se agrega además la columna inversa "USD por <ISO>"
# (cuántos dólares vale 1 unidad de esa moneda = 1 / "X por USD").
# Se usan 6 decimales porque las inversas van de ~1.16 (EUR) a ~0.0008 (KRW).
FMT_INV = "#,##0.000000"


def _fx_columns():
    """(nombre, formato) de las columnas de divisas: directa + inversa, en orden."""
    cols = []
    for iso, (label, fmt) in MONEDAS.items():
        cols.append((label, fmt))
        cols.append((f"USD por {iso}", FMT_INV))
    return cols


def log(msg):
    linea = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(linea)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(linea + "\n")


# ------------------------------------------------------------------
# 1) Descarga de datos
# ------------------------------------------------------------------
def obtener_henry_hub():
    """Henry Hub spot diario desde EIA (serie NG.RNGWHHD.D)."""
    url = "https://api.eia.gov/v2/seriesid/NG.RNGWHHD.D"
    r = requests.get(url, params={"api_key": EIA_API_KEY}, timeout=60)
    r.raise_for_status()
    datos = r.json()["response"]["data"]
    df = pd.DataFrame(datos)[["period", "value"]]
    df.columns = ["fecha", COL_HH]
    df["fecha"] = pd.to_datetime(df["fecha"])
    df[COL_HH] = pd.to_numeric(df[COL_HH], errors="coerce")
    df = df[df["fecha"] >= FECHA_INICIO].sort_values("fecha")
    return df.reset_index(drop=True)


def obtener_fred(series_id, observation_start=FECHA_INICIO):
    """Una serie diaria/mensual de FRED -> DataFrame[fecha, valor]."""
    url = "https://api.stlouisfed.org/fred/series/observations"
    params = {
        "series_id": series_id,
        "api_key": FRED_API_KEY,
        "file_type": "json",
        "observation_start": observation_start,
    }
    r = requests.get(url, params=params, timeout=60)
    r.raise_for_status()
    obs = r.json()["observations"]
    df = pd.DataFrame(obs)[["date", "value"]]
    df.columns = ["fecha", "valor"]
    df["fecha"] = pd.to_datetime(df["fecha"])
    df["valor"] = pd.to_numeric(df["valor"].replace(".", None), errors="coerce")
    return df


def obtener_fx(desde=FECHA_INICIO, reintentos=3):
    """
    Tipo de cambio diario desde Frankfurter (tipos de referencia del BCE).
    base=USD devuelve "unidades por 1 USD" directamente, sin invertir.
    Pide SOLO las divisas de MONEDAS (symbols) para que la respuesta sea
    chica y no se corte; reintenta si la conexión se cae.
    """
    url = "https://api.frankfurter.dev/v2/rates"
    params = {
        "base": "USD",
        "from": desde,
        "quotes": ",".join(MONEDAS.keys()),   # solo tus 10 divisas (param v2)
    }
    ultimo_error = None
    for intento in range(1, reintentos + 1):
        try:
            r = requests.get(url, params=params, timeout=120)
            r.raise_for_status()
            data = r.json()
            break
        except (requests.exceptions.ChunkedEncodingError,
                requests.exceptions.ConnectionError,
                requests.exceptions.Timeout) as e:
            ultimo_error = e
            log(f"  Frankfurter falló (intento {intento}/{reintentos}): {e}")
            if intento == reintentos:
                raise
    else:
        raise ultimo_error

    # Frankfurter v2 devuelve una lista de registros {date, base, quote, rate}.
    # Dejamos el parser tolerante por si el rango llega como dict anidado.
    if isinstance(data, list):
        df = pd.DataFrame(data)
    elif isinstance(data, dict) and "rates" in data:
        rates = data["rates"]
        registros = []
        if rates and all(isinstance(v, dict) for v in rates.values()):
            for fecha, d in rates.items():
                for cur, val in d.items():
                    registros.append({"date": fecha, "quote": cur, "rate": val})
        else:
            for cur, val in rates.items():
                registros.append({"date": data.get("date"), "quote": cur, "rate": val})
        df = pd.DataFrame(registros)
    else:
        raise ValueError("Respuesta de Frankfurter con formato inesperado")

    df = df[df["quote"].isin(MONEDAS)].copy()
    df["fecha"] = pd.to_datetime(df["date"])
    df["rate"] = pd.to_numeric(df["rate"], errors="coerce")
    wide = df.pivot_table(index="fecha", columns="quote", values="rate").reset_index()
    wide = wide.rename(columns={c: MONEDAS[c][0] for c in MONEDAS if c in wide.columns})
    wide.columns.name = None

    # agrega la columna inversa (USD por X) junto a cada divisa y ordena
    orden = ["fecha"]
    for iso, (label, _fmt) in MONEDAS.items():
        if label in wide.columns:
            inv = f"USD por {iso}"
            wide[inv] = 1.0 / wide[label]
            orden += [label, inv]
    return wide[orden]


def obtener_inflacion():
    """CPI mensual (CPIAUCSL) + variación interanual."""
    # tomamos 13 meses antes del inicio para poder calcular el YoY de enero-2021
    df = obtener_fred("CPIAUCSL", observation_start="2019-12-01")
    df = df.rename(columns={"valor": "CPI"}).sort_values("fecha")
    df["Inflacion YoY"] = df["CPI"].pct_change(12)   # fracción (0.034 = 3.4%)
    df = df[df["fecha"] >= FECHA_INICIO].reset_index(drop=True)
    return df


def obtener_macro():
    """
    Indicadores macro de EE. UU. (FRED), en una sola tabla mensual:
      - Inflación YoY (CPI)            fracción (0.03 = 3%)
      - Desempleo                      % (UNRATE), mensual
      - PIB real (crec. anualizado)    % (A191RL1Q225SBEA), trimestral
    El PIB es trimestral, así que solo tiene dato en meses de fin de trimestre;
    el resto de filas queda en blanco, lo cual es normal.
    """
    infl = obtener_inflacion()[["fecha", "CPI", "Inflacion YoY"]]

    log("Descargando desempleo (FRED:UNRATE)...")
    des = obtener_fred("UNRATE").rename(columns={"valor": "Desempleo"})

    log("Descargando PIB (FRED:A191RL1Q225SBEA)...")
    pib = obtener_fred("A191RL1Q225SBEA").rename(columns={"valor": "PIB crec."})

    macro = infl.merge(des, on="fecha", how="outer").merge(pib, on="fecha", how="outer")
    macro = macro[macro["fecha"] >= FECHA_INICIO].sort_values("fecha").reset_index(drop=True)
    return macro


# ------------------------------------------------------------------
# 1b) Acereras: precios de acciones (Twelve Data) -> índices y correlación
# ------------------------------------------------------------------
def obtener_twelvedata(symbol, reintentos=4):
    """
    Cierre diario de una acción desde Twelve Data (JSON, con llave).
    Devuelve una Serie indexada por fecha, o None. Registra POR QUÉ falla.

    Notas sobre el plan GRATIS de Twelve Data:
      - Límite de frecuencia: 8 peticiones por minuto (el throttle real va en
        'jalar', que espera entre símbolo y símbolo).
      - Límite diario: 800 créditos; cada llamada aquí cuesta 1 crédito.
      - Si se topa con el límite (code 429 / "credits"), espera ~65 s y reintenta.
      - Si el símbolo no existe o el plan no lo cubre, no insiste: lo salta.
    Twelve Data devuelve HTTP 200 incluso en errores lógicos; el detalle viene
    en el campo 'status'/'code' del JSON, así que eso es lo que revisamos.
    """
    if not TWELVEDATA_API_KEY:
        log(f"  TwelveData {symbol}: falta TWELVEDATA_API_KEY (Secret del repo / .env)")
        return None

    url = "https://api.twelvedata.com/time_series"
    params = {
        "symbol":     symbol,
        "interval":   "1day",
        "start_date": FECHA_INICIO,
        "outputsize": 5000,          # cubre de 2021 a hoy con margen de sobra
        "order":      "ASC",
        "apikey":     TWELVEDATA_API_KEY,
        "format":     "JSON",
    }
    for intento in range(1, reintentos + 1):
        try:
            r = requests.get(url, params=params, timeout=60)
            try:
                data = r.json()
            except ValueError:
                log(f"  TwelveData {symbol}: respuesta no-JSON (status {r.status_code}) "
                    f"-> '{(r.text or '')[:60]}'")
                time.sleep(3.0 * intento)
                continue

            # Error lógico: {"status":"error","code":...,"message":...}
            if isinstance(data, dict) and data.get("status") == "error":
                code = data.get("code")
                msg = (data.get("message") or "")[:90]
                if code == 429 or "credit" in msg.lower() or "limit" in msg.lower():
                    espera = 65
                    log(f"  TwelveData {symbol}: LÍMITE (code {code}) -> espero {espera}s -> '{msg}'")
                    time.sleep(espera)
                    continue
                # Símbolo inexistente o no cubierto por el plan: no insistir.
                log(f"  TwelveData {symbol}: ERROR (code {code}) -> '{msg}'")
                return None

            valores = data.get("values") if isinstance(data, dict) else None
            if not valores:
                log(f"  TwelveData {symbol}: sin 'values' -> '{str(data)[:90]}'")
                return None

            df = pd.DataFrame(valores)
            if "datetime" not in df.columns or "close" not in df.columns:
                log(f"  TwelveData {symbol}: columnas raras {list(df.columns)[:5]}")
                return None

            s = pd.Series(pd.to_numeric(df["close"], errors="coerce").values,
                          index=pd.to_datetime(df["datetime"], errors="coerce"))
            s.index.name = "fecha"
            s = s.dropna().sort_index()
            s = s[s.index >= FECHA_INICIO]
            if not s.empty:
                return s
            log(f"  TwelveData {symbol}: 0 filas tras filtrar desde {FECHA_INICIO}")
            return None
        except Exception as e:
            log(f"  TwelveData {symbol} intento {intento}: {e}")
            time.sleep(3.0 * intento)
    return None


def _indice_grupo(precios):
    """
    Índice equal-weight (base 100) a partir de los rendimientos diarios promedio
    de las acciones del grupo. Así entran/salen empresas sin saltos artificiales.
    Recorta movimientos diarios a ±50% para evitar artefactos por splits.
    """
    rets = {}
    for tk, s in precios.items():
        s = s.dropna().sort_index()
        if len(s) < 2:
            continue
        rets[tk] = s.pct_change().clip(-0.5, 0.5)
    if not rets:
        return None
    R = pd.DataFrame(rets).sort_index()
    prom = R.mean(axis=1, skipna=True).fillna(0.0)   # rendimiento equal-weight
    return 100.0 * (1 + prom).cumprod()


def construir_acereras():
    """
    Descarga las acereras de cada grupo, arma dos índices base 100 y calcula
    su correlación (global y móvil a 90 días). Devuelve (df, info).
    """
    if not TWELVEDATA_API_KEY:
        log("  Aviso: falta TWELVEDATA_API_KEY; se omite la sección de acereras.")
        return None, {"corr_global": None, "ok_mundial": [], "ok_eeuu": [],
                      "fallaron": ["(sin TWELVEDATA_API_KEY)"]}

    def jalar(dic, etiqueta):
        precios, ok, fallaron = {}, [], []
        for tk, nom in dic.items():
            log(f"Descargando {nom} ({tk}) [{etiqueta}]...")
            s = obtener_twelvedata(tk)
            if s is None or s.empty:
                fallaron.append(f"{nom} ({tk})")
            else:
                precios[tk] = s
                ok.append(nom)
            time.sleep(8.0)   # plan gratis: máx. 8 peticiones/min -> 1 cada 8 s
        return precios, ok, fallaron

    p_m, ok_m, fail_m = jalar(ACERERAS_MUNDIAL, "mundial")
    p_u, ok_u, fail_u = jalar(ACERERAS_EEUU, "EE.UU.")

    idx_m = _indice_grupo(p_m)
    idx_u = _indice_grupo(p_u)
    if idx_m is None or idx_u is None:
        log("  Aviso: faltan datos de acereras en algún grupo; se omite la sección.")
        return None, {"corr_global": None, "ok_mundial": ok_m, "ok_eeuu": ok_u,
                      "fallaron": fail_m + fail_u}

    df = pd.concat({"Índice mundial": idx_m, "Índice EE. UU.": idx_u}, axis=1)
    df = df.sort_index()
    df = df[df.index >= FECHA_INICIO].dropna(how="all")
    # rebasar ambos a 100 en su primera fecha común para comparar en la gráfica
    df = df.dropna()
    for c in df.columns:
        df[c] = df[c] / df[c].iloc[0] * 100.0

    # Correlación sobre rendimientos SEMANALES (cierre de viernes). La semana
    # da tiempo a que todas las bolsas (Asia, Europa, EE. UU.) reflejen lo mismo,
    # así que elimina el desfase de husos horarios que ensucia la versión diaria.
    sem = df.resample("W-FRI").last()
    rmw = sem["Índice mundial"].pct_change()
    ruw = sem["Índice EE. UU."].pct_change()
    corr_global = float(rmw.corr(ruw))
    corr_movil = rmw.rolling(13).corr(ruw)            # ~1 trimestre (13 semanas)

    corr_fecha = [d.strftime("%Y-%m-%d") for d in sem.index]
    corr_series = [None if pd.isna(v) else round(float(v), 3) for v in corr_movil]

    df = df.reset_index().rename(columns={"index": "fecha"})
    if "fecha" not in df.columns:        # por si el índice no se llamaba 'fecha'
        df = df.rename(columns={df.columns[0]: "fecha"})

    info = {"corr_global": corr_global, "corr_fecha": corr_fecha,
            "corr_series": corr_series, "ok_mundial": ok_m, "ok_eeuu": ok_u,
            "fallaron": fail_m + fail_u}
    log(f"Correlación semanal acereras mundial vs. EE. UU. (desde {FECHA_INICIO}): {corr_global:.2f}")
    return df, info


def construir_diario():
    """Une todas las series diarias en una sola tabla por fecha."""
    log("Descargando Henry Hub (EIA)...")
    diario = obtener_henry_hub()

    for sid, (nombre, _fmt) in FRED_DIARIAS.items():
        log(f"Descargando {nombre} (FRED:{sid})...")
        s = obtener_fred(sid).rename(columns={"valor": nombre})
        diario = diario.merge(s, on="fecha", how="outer")

    log("Descargando divisas (Frankfurter/BCE)...")
    fx = obtener_fx()
    diario = diario.merge(fx, on="fecha", how="outer")

    diario = diario[diario["fecha"] >= FECHA_INICIO].sort_values("fecha")
    diario = diario.reset_index(drop=True)

    # quitar picos falsos de la fuente en precios de energía
    for col in RANGOS_VALIDOS:
        diario = sanear(diario, col)

    return diario


# ------------------------------------------------------------------
# 2) Escritura del Excel con formato
# ------------------------------------------------------------------
# Paleta (misma del panel web: carbón, naranja, rojo, dorado, franja cálida)
AZUL   = "EA580C"   # fila de encabezados (naranja)
AZUL2  = "26262B"   # barra de título (carbón)
ORO    = "FFC400"   # texto de títulos / acentos (dorado)
GRIS   = "FDF1E7"   # franjas alternas (cálido claro)
BLANCO = "FFFFFF"
borde_fino = Border(*(Side(style="thin", color="EADFD3"),) * 4)


def _formatos_diario():
    """Mapa nombre_de_columna -> formato numérico para la hoja Diario."""
    fmt = {COL_HH: FMT_HH}
    fmt.update({n: f for (n, f) in FRED_DIARIAS.values()})
    fmt.update({n: f for (n, f) in _fx_columns()})
    return fmt


def _series_resumen():
    """Lista ordenada (nombre, formato) para la hoja Resumen."""
    s = [(COL_HH, FMT_HH)]
    s += [(n, f) for (n, f) in FRED_DIARIAS.values()]
    s += _fx_columns()
    return s


def _estilizar_hoja(ws, df, formatos, titulo, col_fecha="fecha"):
    """Vuelca un DataFrame con encabezado, franjas, filtros y formato."""
    cols = list(df.columns)
    n_col = len(cols)

    # ---- fila de título (fila 1) ----
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_col)
    c = ws.cell(row=1, column=1, value=titulo)
    c.font = Font(name="Calibri", size=14, bold=True, color=ORO)
    c.fill = PatternFill("solid", fgColor=AZUL2)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26

    # ---- subtítulo con fecha de generación (fila 2) ----
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_col)
    c = ws.cell(row=2, column=1,
                value=f"Generado: {dt.datetime.now():%Y-%m-%d %H:%M}  ·  "
                      f"Fuentes: EIA, FRED y Frankfurter (BCE)  ·  Inicio: {FECHA_INICIO}")
    c.font = Font(name="Calibri", size=9, italic=True, color="808080")
    c.alignment = Alignment(horizontal="left", indent=1)

    fila_enc = 3
    # ---- encabezados ----
    for j, nombre in enumerate(cols, start=1):
        cell = ws.cell(row=fila_enc, column=j, value=nombre)
        cell.font = Font(name="Calibri", size=10, bold=True, color=BLANCO)
        cell.fill = PatternFill("solid", fgColor=AZUL)
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = borde_fino
    ws.row_dimensions[fila_enc].height = 30

    # ---- datos ----
    for i, (_, fila) in enumerate(df.iterrows()):
        r = fila_enc + 1 + i
        franja = PatternFill("solid", fgColor=GRIS) if i % 2 else None
        for j, nombre in enumerate(cols, start=1):
            val = fila[nombre]
            if pd.isna(val):
                val = None
            cell = ws.cell(row=r, column=j, value=val)
            cell.border = borde_fino
            if franja:
                cell.fill = franja
            if nombre == col_fecha:
                cell.number_format = "yyyy-mm-dd"
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.number_format = formatos.get(nombre, "#,##0.00")
                cell.alignment = Alignment(horizontal="right")

    # ---- anchos de columna ----
    for j, nombre in enumerate(cols, start=1):
        letra = get_column_letter(j)
        if nombre == col_fecha:
            ws.column_dimensions[letra].width = 12
        else:
            ws.column_dimensions[letra].width = max(13, min(len(nombre) + 2, 20))

    # ---- congelar encabezado + primera columna, y autofiltro ----
    ws.freeze_panes = ws.cell(row=fila_enc + 1, column=2)
    ultima = get_column_letter(n_col)
    ws.auto_filter.ref = f"A{fila_enc}:{ultima}{fila_enc + len(df)}"


def _hoja_resumen(ws, diario, infl):
    """Hoja de portada con los últimos valores de cada serie."""
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:C1")
    c = ws.cell(row=1, column=1, value="Resumen de mercados")
    c.font = Font(name="Calibri", size=16, bold=True, color=ORO)
    c.fill = PatternFill("solid", fgColor=AZUL2)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:C2")
    c = ws.cell(row=2, column=1,
                value=f"Último dato disponible · Generado {dt.datetime.now():%Y-%m-%d %H:%M}")
    c.font = Font(size=9, italic=True, color="808080")
    c.alignment = Alignment(indent=1)

    enc = ["Serie", "Último valor", "Fecha del dato"]
    for j, t in enumerate(enc, start=1):
        cell = ws.cell(row=4, column=j, value=t)
        cell.font = Font(bold=True, color=BLANCO)
        cell.fill = PatternFill("solid", fgColor=AZUL)
        cell.alignment = Alignment(horizontal="center")
        cell.border = borde_fino

    r = 5
    for nombre, fmt in _series_resumen():
        if nombre not in diario.columns:
            continue
        sub = diario[["fecha", nombre]].dropna()
        if sub.empty:
            continue
        ult = sub.iloc[-1]
        franja = PatternFill("solid", fgColor=GRIS) if (r % 2 == 0) else None
        v1 = ws.cell(row=r, column=1, value=nombre)
        v2 = ws.cell(row=r, column=2, value=float(ult[nombre]))
        v3 = ws.cell(row=r, column=3, value=ult["fecha"].to_pydatetime())
        v2.number_format = fmt
        v3.number_format = "yyyy-mm-dd"
        for cell in (v1, v2, v3):
            cell.border = borde_fino
            if franja:
                cell.fill = franja
        v1.alignment = Alignment(horizontal="left")
        v2.alignment = Alignment(horizontal="right")
        v3.alignment = Alignment(horizontal="center")
        r += 1

    # macro mensual/trimestral
    macro = infl
    filas_macro = [
        ("Inflación YoY (CPI)", "Inflacion YoY", "0.0%"),
        ("Desempleo (%)",       "Desempleo",     '0.0"%"'),
        ("PIB real (crec. anual.)", "PIB crec.",  '0.0"%"'),
    ]
    for etiqueta, col, fmt in filas_macro:
        if col not in macro.columns:
            continue
        sub = macro[["fecha", col]].dropna()
        if sub.empty:
            continue
        ult = sub.iloc[-1]
        franja = PatternFill("solid", fgColor=GRIS) if (r % 2 == 0) else None
        v1 = ws.cell(row=r, column=1, value=etiqueta)
        v2 = ws.cell(row=r, column=2, value=float(ult[col]))
        v3 = ws.cell(row=r, column=3, value=ult["fecha"].to_pydatetime())
        v2.number_format = fmt
        v3.number_format = "yyyy-mm-dd"
        for cell in (v1, v2, v3):
            cell.border = borde_fino
            if franja:
                cell.fill = franja
        v1.alignment = Alignment(horizontal="left")
        v2.alignment = Alignment(horizontal="right")
        v3.alignment = Alignment(horizontal="center")
        r += 1

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16


def _grafico_excel(ws, fila_enc, n_filas, col_idx, titulo, ancla):
    """Inserta un mini gráfico de línea nativo de Excel en la hoja Diario."""
    ch = LineChart()
    ch.title = titulo
    ch.height = 7
    ch.width = 16
    ch.style = 2
    datos = Reference(ws, min_col=col_idx, max_col=col_idx,
                      min_row=fila_enc, max_row=fila_enc + n_filas)
    cats = Reference(ws, min_col=1, min_row=fila_enc + 1, max_row=fila_enc + n_filas)
    ch.add_data(datos, titles_from_data=True)
    ch.set_categories(cats)
    ch.legend = None
    ws.add_chart(ch, ancla)


def _hoja_graficos_fx(wb, ws_diario, diario):
    """Hoja con una gráfica de línea nativa por divisa (unidades por 1 USD)."""
    ws = wb.create_sheet("Gráficos FX")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:R1")
    c = ws.cell(row=1, column=1, value="Tipos de cambio  ·  unidades por 1 USD")
    c.font = Font(name="Calibri", size=14, bold=True, color=ORO)
    c.fill = PatternFill("solid", fgColor=AZUL2)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26

    fila_enc = 3                         # fila de encabezados en la hoja Diario
    n = len(diario)
    cols = list(diario.columns)
    anclas_col = ["A", "J"]              # dos columnas de gráficas
    paso_fila = 14                       # filas entre una gráfica y la de abajo
    idx = 0
    for iso, (label, _fmt) in MONEDAS.items():
        if label not in cols:
            continue
        col_idx = cols.index(label) + 1
        ch = LineChart()
        ch.title = label
        ch.height = 6.5
        ch.width = 12
        ch.style = 2
        ch.legend = None
        datos = Reference(ws_diario, min_col=col_idx, max_col=col_idx,
                          min_row=fila_enc, max_row=fila_enc + n)
        cats = Reference(ws_diario, min_col=1,
                         min_row=fila_enc + 1, max_row=fila_enc + n)
        ch.add_data(datos, titles_from_data=True)
        ch.set_categories(cats)
        fila = 3 + (idx // 2) * paso_fila
        ws.add_chart(ch, f"{anclas_col[idx % 2]}{fila}")
        idx += 1


def escribir_excel(diario, infl, ruta, acereras=None):
    wb = Workbook()

    # Hoja 1: Resumen
    ws_res = wb.active
    ws_res.title = "Resumen"
    _hoja_resumen(ws_res, diario, infl)

    # Hoja 2: Diario
    ws_dia = wb.create_sheet("Diario")
    _estilizar_hoja(ws_dia, diario, _formatos_diario(),
                    "Series diarias  ·  Henry Hub, S&P 500, WTI, Brent y divisas")

    # gráficos nativos para Henry Hub y S&P 500
    fila_enc = 3
    n = len(diario)
    cols = list(diario.columns)
    ancla_col = get_column_letter(len(cols) + 2)
    if COL_HH in cols:
        _grafico_excel(ws_dia, fila_enc, n, cols.index(COL_HH) + 1,
                       COL_HH, f"{ancla_col}3")
    if "S&P 500" in cols:
        _grafico_excel(ws_dia, fila_enc, n, cols.index("S&P 500") + 1,
                       "S&P 500", f"{ancla_col}18")

    # Hoja 3: Macro EE. UU.
    ws_mac = wb.create_sheet("Macro EE. UU.")
    _estilizar_hoja(ws_mac, infl,
                    {"CPI": "#,##0.000", "Inflacion YoY": "0.0%",
                     "Desempleo": '0.0"%"', "PIB crec.": '0.0"%"'},
                    "Macro EE. UU.  ·  inflación, desempleo y PIB")

    # Hoja 4: Acereras (índices; la correlación semanal va en el panel web)
    if acereras is not None and not acereras.empty:
        ws_ac = wb.create_sheet("Acereras")
        _estilizar_hoja(ws_ac, acereras,
                        {"Índice mundial": "#,##0.0", "Índice EE. UU.": "#,##0.0"},
                        "Acereras  ·  índice mundial vs. EE. UU. (base 100 = 2021)")

    # Hoja final: una gráfica por divisa
    _hoja_graficos_fx(wb, ws_dia, diario)

    wb.save(ruta)


# ------------------------------------------------------------------
# 3) Gráficos PNG (panel resumen)
# ------------------------------------------------------------------
def crear_graficos(diario, macro, ruta_png):
    fig, ax = plt.subplots(3, 2, figsize=(13, 12))
    fig.suptitle("Mercados y macro — resumen", fontsize=14, fontweight="bold")

    d = diario.set_index("fecha")
    if COL_HH in d:
        ax[0, 0].plot(d.index, d[COL_HH], color="#FF7A18")
        ax[0, 0].set_title(COL_HH)
    for col, c in [("WTI ($/bbl)", "#FF7A18"), ("Brent ($/bbl)", "#E62315")]:
        if col in d:
            ax[0, 1].plot(d.index, d[col], label=col, color=c)
    ax[0, 1].set_title("Petróleo: WTI vs. Brent")
    ax[0, 1].legend(fontsize=8)
    if "S&P 500" in d:
        ax[1, 0].plot(d.index, d["S&P 500"], color="#888c92")
        ax[1, 0].set_title("S&P 500")

    m = macro.set_index("fecha")
    if "Inflacion YoY" in m:
        ai = m["Inflacion YoY"].dropna() * 100
        ax[1, 1].plot(ai.index, ai.values, color="#BF8F00")
        ax[1, 1].set_title("Inflación YoY (%)")
    if "Desempleo" in m:
        de = m["Desempleo"].dropna()
        ax[2, 0].plot(de.index, de.values, color="#2E5496")
        ax[2, 0].set_title("Desempleo (%)")
    if "PIB crec." in m:
        pb = m["PIB crec."].dropna()
        ax[2, 1].plot(pb.index, pb.values, color="#548235", marker="o", markersize=3)
        ax[2, 1].set_title("PIB real — crec. anualizado (%)")

    for a in ax.flat:
        a.grid(True, alpha=0.3)
        a.tick_params(labelsize=8)
    fig.text(0.99, 0.01,
             f"Fuentes: Henry Hub (EIA); S&P 500, WTI, Brent, inflación, "
             f"desempleo y PIB (FRED).  Generado {dt.date.today():%Y-%m-%d}.",
             ha="right", va="bottom", fontsize=7.5, color="#808080")
    fig.tight_layout(rect=[0, 0.02, 1, 0.97])
    fig.savefig(ruta_png, dpi=130)
    plt.close(fig)


def crear_graficos_fx(diario, ruta_png):
    """Panel con una gráfica por divisa (unidades por 1 USD)."""
    monedas = [label for (label, _f) in MONEDAS.values() if label in diario.columns]
    if not monedas:
        return
    ncols = 2
    nfilas = (len(monedas) + ncols - 1) // ncols
    fig, ax = plt.subplots(nfilas, ncols, figsize=(13, 3 * nfilas))
    fig.suptitle("Tipos de cambio — unidades por 1 USD", fontsize=14, fontweight="bold")
    d = diario.set_index("fecha")
    ejes = ax.flat
    for i, label in enumerate(monedas):
        a = ejes[i]
        a.plot(d.index, d[label], color="#2E5496")
        a.set_title(label, fontsize=10)
        a.grid(True, alpha=0.3)
        a.tick_params(labelsize=8)
    for j in range(len(monedas), len(ejes)):     # apaga ejes sobrantes
        ejes[j].axis("off")
    fig.text(0.99, 0.01,
             f"Fuente: Frankfurter (tipos de referencia del Banco Central Europeo).  "
             f"Generado {dt.date.today():%Y-%m-%d}.",
             ha="right", va="bottom", fontsize=7.5, color="#808080")
    fig.tight_layout(rect=[0, 0.02, 1, 0.97])
    fig.savefig(ruta_png, dpi=130)
    plt.close(fig)


# ------------------------------------------------------------------
# 3b) Exportar datos a JSON (para el panel web / GitHub Pages)
# ------------------------------------------------------------------
def escribir_json(diario, macro, ruta, acereras=None, info_acereras=None):
    """Vuelca los datos en un JSON compacto que consume el index.html."""
    def limpia(serie, dec=6):
        out = []
        for v in serie:
            out.append(None if pd.isna(v) else round(float(v), dec))
        return out

    series = {col: limpia(diario[col]) for col in diario.columns if col != "fecha"}

    resumen = []
    for nombre, fmt in _series_resumen():
        if nombre not in diario.columns:
            continue
        sub = diario[["fecha", nombre]].dropna()
        if sub.empty:
            continue
        ult = sub.iloc[-1]
        resumen.append({"serie": nombre, "valor": round(float(ult[nombre]), 6),
                        "fecha": ult["fecha"].strftime("%Y-%m-%d"), "fmt": fmt})

    # indicadores macro (mensuales/trimestrales). 'escala' lleva el valor a %.
    macro_def = [
        ("Inflación YoY (CPI)",        "Inflacion YoY", 100),
        ("Desempleo",                  "Desempleo",       1),
        ("PIB real (crec. anualizado)", "PIB crec.",      1),
    ]
    macro_out = []
    for nombre, col, escala in macro_def:
        if col not in macro.columns:
            continue
        sub = macro[["fecha", col]].dropna()
        if sub.empty:
            continue
        macro_out.append({
            "nombre": nombre, "escala": escala,
            "fecha": sub["fecha"].dt.strftime("%Y-%m-%d").tolist(),
            "valores": limpia(sub[col], 5),
        })

    ahora_utc = dt.datetime.now(dt.timezone.utc)
    obj = {
        "generado": ahora_utc.strftime("%Y-%m-%d %H:%M UTC"),
        "generado_iso": ahora_utc.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "fuentes": "Henry Hub: EIA · S&P 500, WTI, Brent, inflación, desempleo, PIB: "
                   "FRED · Divisas: Frankfurter (BCE) · Acciones: Twelve Data",
        "fechas": diario["fecha"].dt.strftime("%Y-%m-%d").tolist(),
        "series": series,
        "monedas": [label for (_iso, (label, _f)) in MONEDAS.items()],
        "macro": macro_out,
        "resumen": resumen,
    }

    if acereras is not None and not acereras.empty:
        ia = info_acereras or {}
        cg = ia.get("corr_global")
        obj["acereras"] = {
            "fecha": acereras["fecha"].dt.strftime("%Y-%m-%d").tolist(),
            "mundial": limpia(acereras["Índice mundial"], 2),
            "eeuu": limpia(acereras["Índice EE. UU."], 2),
            "corr_fecha": ia.get("corr_fecha", []),
            "corr": ia.get("corr_series", []),
            "corr_global": (None if cg is None else round(cg, 3)),
            "ok_mundial": ia.get("ok_mundial", []),
            "ok_eeuu": ia.get("ok_eeuu", []),
            "fallaron": ia.get("fallaron", []),
        }

    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False)


# ------------------------------------------------------------------
# 4) Programa principal
# ------------------------------------------------------------------
def main():
    if not EIA_API_KEY or not FRED_API_KEY:
        log("ERROR: faltan llaves EIA_API_KEY y/o FRED_API_KEY "
            "(.env en tu PC, o Secrets en GitHub Actions).")
        sys.exit(1)

    log("===== Inicio =====")
    try:
        diario = construir_diario()
        log("Descargando indicadores macro (FRED)...")
        macro = obtener_macro()
        log("Descargando acereras (Twelve Data)...")
        try:
            acereras, info_ac = construir_acereras()
        except Exception as e:
            log(f"  Aviso: la sección de acereras falló y se omite: {e}")
            acereras, info_ac = None, None

        xlsx_fecha = CARPETA / f"mercados_{HOY}.xlsx"
        xlsx_recie = CARPETA / "mercados_reciente.xlsx"
        png_fecha = CARPETA / f"graficos_{HOY}.png"
        png_recie = CARPETA / "graficos_reciente.png"
        png_fx_fecha = CARPETA / f"graficos_divisas_{HOY}.png"
        png_fx_recie = CARPETA / "graficos_divisas_reciente.png"

        escribir_excel(diario, macro, xlsx_fecha, acereras=acereras)
        escribir_excel(diario, macro, xlsx_recie, acereras=acereras)
        crear_graficos(diario, macro, png_fecha)
        crear_graficos(diario, macro, png_recie)
        crear_graficos_fx(diario, png_fx_fecha)
        crear_graficos_fx(diario, png_fx_recie)

        # datos para el panel web (GitHub Pages)
        escribir_json(diario, macro, CARPETA / "datos.json",
                      acereras=acereras, info_acereras=info_ac)

        ult_hh = diario.dropna(subset=[COL_HH]).iloc[-1]
        log(f"Henry Hub más reciente: ${ult_hh[COL_HH]:.2f}/MMBtu ({ult_hh['fecha'].date()})")
        if "MXN por USD" in diario.columns:
            ult_mxn = diario.dropna(subset=["MXN por USD"]).iloc[-1]
            log(f"Peso más reciente: {ult_mxn['MXN por USD']:.4f} MXN/USD ({ult_mxn['fecha'].date()})")
        inf = macro.dropna(subset=['Inflacion YoY']).iloc[-1]
        log(f"Inflación YoY más reciente: {inf['Inflacion YoY']*100:.1f}% ({inf['fecha'].date()})")
        log(f"Excel guardado en: {xlsx_recie}")
        log("===== Fin OK =====")
    except Exception as e:
        log(f"ERROR: {e}")
        raise


if __name__ == "__main__":
    main()
