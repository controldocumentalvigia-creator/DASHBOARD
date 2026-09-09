
from __future__ import annotations

from io import BytesIO
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import base64
import re
import calendar

import numpy as np
import pandas as pd
from pandas.io.formats.style import Styler
import openpyxl
import plotly.graph_objects as go
import requests
import streamlit as st


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
st.set_page_config(
    page_title="Centro de Control Gerencial VSE",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROYAL_BLUE = "#2563EB"
DARK_BLUE = "#60A5FA"
SECONDARY_BLUE = "#60A5FA"
LIGHT_BLUE = "#16263A"
BG = "#07111F"
CARD = "#101C2C"
PANEL = "#16263A"
TEXT = "#F8FAFC"
TEXT_2 = "#CBD5E1"
TEXT_3 = "#94A3B8"
GREEN = "#22C55E"
RED = "#EF4444"
AMBER = "#F59E0B"
ORANGE = "#F97316"
GRID = "#29415E"

st.markdown(
    f"""
    <style>
        :root {{
            --vse-bg: {BG};
            --vse-card: {CARD};
            --vse-panel: {PANEL};
            --vse-border: {GRID};
            --vse-text: {TEXT};
            --vse-muted: {TEXT_2};
            --vse-blue: {ROYAL_BLUE};
        }}

        .stApp {{
            background: {BG} !important;
            color: {TEXT} !important;
        }}
        html, body, [class*="css"] {{
            color: {TEXT};
        }}
        .block-container {{
            padding-top: 0.65rem;
            padding-bottom: 2.2rem;
            max-width: 1800px;
        }}
        h1, h2, h3, h4, h5, h6 {{
            color: {TEXT} !important;
            letter-spacing: -0.01em;
        }}
        p, label, small {{
            color: {TEXT_2};
        }}

        [data-testid="stSidebar"] {{
            background: #0B1626 !important;
            border-right: 1px solid {GRID} !important;
        }}
        [data-testid="stSidebar"] * {{
            color: #E8EEF7 !important;
        }}

        .hero {{
            background: linear-gradient(105deg, #0A2D59 0%, #114A95 48%, #2563EB 100%);
            padding: 19px 24px;
            border-radius: 14px;
            color: white;
            margin-bottom: 12px;
            box-shadow: 0 8px 26px rgba(0,0,0,.24);
            border: 1px solid rgba(96,165,250,.18);
        }}
        .hero-title {{
            font-size: 28px;
            font-weight: 850;
            margin: 0;
            color: #FFFFFF;
        }}
        .hero-sub {{
            font-size: 14px;
            font-weight: 600;
            margin-top: 5px;
            color: #DCE8F8;
        }}

        .base-status {{
            background: #0D1A2B;
            border: 1px solid {GRID};
            border-left: 4px solid {GREEN};
            border-radius: 10px;
            padding: 9px 13px;
            margin: 0 0 12px 0;
            color: {TEXT_2};
            font-size: 12px;
        }}
        .base-status strong {{ color: {TEXT}; }}

        .section-title {{
            background: #0F2D52;
            color: #FFFFFF;
            padding: 8px 12px;
            border-radius: 9px 9px 0 0;
            border: 1px solid {GRID};
            font-weight: 800;
            font-size: 13px;
            margin-top: 7px;
            letter-spacing: .25px;
        }}
        .subsection-tag {{
            color: {TEXT_2};
            font-size: 11px;
            font-weight: 800;
            letter-spacing: .7px;
            text-transform: uppercase;
            margin: 5px 0 7px 0;
        }}

        .kpi {{
            background: {CARD};
            border: 1px solid {GRID};
            border-left: 5px solid {ROYAL_BLUE};
            border-radius: 12px;
            padding: 13px 14px;
            min-height: 116px;
            box-shadow: 0 4px 14px rgba(0,0,0,.18);
        }}
        .kpi-label {{
            color: {TEXT_2};
            font-size: 11px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: .35px;
        }}
        .kpi-value {{
            color: {TEXT};
            font-size: 23px;
            font-weight: 850;
            margin-top: 6px;
        }}
        .kpi-delta {{
            margin-top: 7px;
            font-size: 11px;
            font-weight: 700;
        }}
        .kpi-note {{
            color: {TEXT_3};
            font-size: 10.5px;
            margin-top: 4px;
            line-height: 1.25;
        }}
        .small-note {{
            font-size: 11px;
            color: {TEXT_3};
        }}

        .fin-kpi {{
            background: {CARD};
            border: 1px solid {GRID};
            border-left: 5px solid {ROYAL_BLUE};
            border-radius: 12px;
            padding: 13px 14px;
            min-height: 116px;
            box-shadow: 0 4px 14px rgba(0,0,0,.18);
        }}
        .fin-kpi .label {{
            color: {TEXT_2};
            font-size: 11px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: .35px;
        }}
        .fin-kpi .value {{
            color: {TEXT};
            font-size: 23px;
            font-weight: 850;
            margin-top: 6px;
        }}
        .fin-kpi .note {{
            color: {TEXT_3};
            font-size: 10.5px;
            margin-top: 5px;
            line-height: 1.25;
        }}
        .fin-kpi.good {{ border-left-color: {GREEN}; }}
        .fin-kpi.warn {{ border-left-color: {AMBER}; }}
        .fin-kpi.bad {{ border-left-color: {RED}; }}

        .fin-panel {{
            background: {CARD};
            border: 1px solid {GRID};
            border-radius: 12px;
            padding: 13px 15px;
            margin: 7px 0 12px 0;
        }}
        .fin-panel strong {{ color: {TEXT}; }}
        .fin-note {{ color: {TEXT_2}; font-size: 11px; }}

        div[data-testid="stDataFrame"] {{
            background: {CARD} !important;
            border: 1px solid {GRID} !important;
            border-radius: 10px;
            overflow: hidden;
        }}

        div[data-baseweb="select"] > div,
        div[data-testid="stDateInput"] input,
        div[data-testid="stTextInput"] input,
        div[data-testid="stNumberInput"] input {{
            background: {CARD} !important;
            color: {TEXT} !important;
            border-color: {GRID} !important;
        }}
        div[data-baseweb="popover"] > div,
        ul[role="listbox"] {{
            background: {CARD} !important;
            color: {TEXT} !important;
        }}

        div[data-testid="stFileUploader"] section {{
            background: {CARD} !important;
            border-color: {GRID} !important;
        }}

        div[data-testid="stDownloadButton"] button,
        div[data-testid="stButton"] button {{
            background: #12335A;
            color: #FFFFFF !important;
            border: 1px solid #315A86;
            border-radius: 9px;
        }}
        div[data-testid="stDownloadButton"] button:hover,
        div[data-testid="stButton"] button:hover {{
            border-color: {SECONDARY_BLUE};
            background: #174776;
            color: #FFFFFF !important;
        }}

        div[data-testid="stSelectbox"] label,
        div[data-testid="stMultiSelect"] label,
        div[data-testid="stDateInput"] label,
        div[data-testid="stRadio"] label,
        div[data-testid="stFileUploader"] label {{
            font-weight: 700 !important;
            color: {TEXT_2} !important;
        }}

        [data-testid="stExpander"] {{
            background: #0D1A2B;
            border: 1px solid {GRID};
            border-radius: 10px;
        }}

        hr {{ border-color: {GRID}; }}
    </style>
    """,
    unsafe_allow_html=True,
)



# =========================================================
# SISTEMA VISUAL GLOBAL DE TABLAS · DARK NAVY EJECUTIVO
# =========================================================
st.markdown(
    """
    <style>
        .exec-table-wrap {
            width: 100%;
            overflow: auto;
            background: #101C2C;
            border: 1px solid #29415E;
            border-radius: 11px;
            margin: 4px 0 14px 0;
            box-shadow: 0 4px 14px rgba(0,0,0,.15);
        }

        .exec-table-wrap table {
            border-collapse: separate !important;
            border-spacing: 0 !important;
            width: 100% !important;
            min-width: 720px;
            margin: 0 !important;
            font-size: 11.5px !important;
            color: #F8FAFC !important;
        }

        .exec-table-wrap thead th {
            position: sticky;
            top: 0;
            z-index: 5;
            background: #12365E !important;
            color: #FFFFFF !important;
            font-weight: 800 !important;
            text-align: center !important;
            padding: 9px 10px !important;
            border-right: 1px solid #29415E !important;
            border-bottom: 1px solid #3B5C82 !important;
            white-space: nowrap !important;
        }

        .exec-table-wrap thead th:first-child {
            text-align: left !important;
        }

        .exec-table-wrap tbody td {
            background: #101C2C !important;
            color: #F8FAFC !important;
            padding: 8px 10px !important;
            border-right: 1px solid #20364F !important;
            border-bottom: 1px solid #20364F !important;
            text-align: right !important;
            white-space: nowrap !important;
            vertical-align: middle !important;
        }

        .exec-table-wrap tbody tr:nth-child(even) td {
            background: #13243A !important;
        }

        .exec-table-wrap tbody tr:hover td {
            background: #19334F !important;
        }

        .exec-table-wrap tbody td:first-child {
            text-align: left !important;
            font-weight: 650 !important;
            white-space: normal !important;
            min-width: 190px;
        }

        /* Acentos semánticos: discretos, no tipo semáforo */
        .exec-table-wrap td.cell-accent {
            color: #93C5FD !important;
            font-weight: 700 !important;
        }
        .exec-table-wrap td.cell-good {
            color: #86EFAC !important;
            font-weight: 700 !important;
        }
        .exec-table-wrap td.cell-warn {
            color: #FBBF24 !important;
            font-weight: 700 !important;
        }
        .exec-table-wrap td.cell-bad {
            color: #FCA5A5 !important;
            font-weight: 700 !important;
        }

        .exec-table-wrap::-webkit-scrollbar {
            height: 9px;
            width: 9px;
        }
        .exec-table-wrap::-webkit-scrollbar-track {
            background: #0B1626;
        }
        .exec-table-wrap::-webkit-scrollbar-thumb {
            background: #315A86;
            border-radius: 8px;
        }

        .exec-table-note {
            color: #94A3B8;
            font-size: 10px;
            margin-top: -7px;
            margin-bottom: 8px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


def _table_cell_classes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica acentos cromáticos discretos según el significado de la columna.
    El fondo permanece Dark Navy; sólo cambia el color tipográfico de celdas clave.
    """
    classes = pd.DataFrame("", index=df.index, columns=df.columns)

    for col in df.columns:
        label = str(col).strip().upper()

        # Pendientes / exposición: ámbar.
        if "PENDIENTE" in label:
            classes[col] = "cell-warn"

        # Porcentajes, participación, rentabilidad y variaciones: azul claro.
        elif (
            "%" in label
            or "RENTABILIDAD" in label
            or "PARTICIPACIÓN" in label
            or "PARTICIPACION" in label
            or "VARIACIÓN" in label
            or "VARIACION" in label
            or "COBERTURA" in label
        ):
            classes[col] = "cell-accent"

        # Dinero efectivamente realizado / margen: verde suave.
        elif (
            "RECAUDADO" in label
            or "PAGADO" in label
            or label == "MARGEN"
            or "CAJA OPERATIVA" in label
        ):
            classes[col] = "cell-good"

    # Valores negativos en columnas económicas relevantes: rojo suave.
    for col in df.columns:
        label = str(col).strip().upper()
        if any(k in label for k in ["MARGEN", "CAJA OPERATIVA", "VARIACIÓN", "VARIACION"]):
            numeric = pd.to_numeric(df[col], errors="coerce")
            neg_mask = numeric.lt(0)
            if neg_mask.any():
                classes.loc[neg_mask, col] = "cell-bad"

    # Estados operativos: acento semántico moderado.
    for col in df.columns:
        if str(col).strip().upper() in {"ESTADO", "ESTADO OP", "ESTADO OPERATIVO"}:
            values = df[col].astype("string").str.upper().fillna("")
            classes.loc[values.str.contains("CUMPLIDO", na=False), col] = "cell-good"
            classes.loc[values.str.contains("PROGRAM", na=False), col] = "cell-warn"
            classes.loc[values.str.contains("TRANSITO", na=False), col] = "cell-warn"

    return classes


def executive_dataframe(
    data,
    use_container_width=True,
    hide_index=True,
    height=None,
    **kwargs,
):
    """
    Sustituto visual de st.dataframe para todas las tablas del dashboard.
    Conserva los formatos de pandas Styler y añade:
    encabezado azul, filas dark navy, zebra suave, hover, scroll y contraste alto.
    """
    if isinstance(data, Styler):
        styler = data
        df_table = styler.data
    elif isinstance(data, pd.DataFrame):
        df_table = data
        styler = data.style
    else:
        # Compatibilidad defensiva.
        try:
            df_table = pd.DataFrame(data)
            styler = df_table.style
        except Exception:
            return st.markdown(str(data))

    # Ocultar índice para mantener estética ejecutiva.
    if hide_index:
        try:
            styler = styler.hide(axis="index")
        except Exception:
            pass

    # Clases semánticas generales.
    try:
        styler = styler.set_td_classes(_table_cell_classes(df_table))
    except Exception:
        pass

    # El header/filas se fuerzan por CSS global con !important.
    try:
        styler = styler.set_table_attributes('class="exec-table"')
    except Exception:
        pass

    html_table = styler.to_html()

    max_h = int(height) if isinstance(height, (int, float)) else 460
    max_h = max(160, min(max_h, 620))

    st.markdown(
        f'<div class="exec-table-wrap" style="max-height:{max_h}px;">{html_table}</div>',
        unsafe_allow_html=True,
    )


# Intercepta TODAS las tablas existentes sin reescribir 20 bloques por separado.
st.dataframe = executive_dataframe


# =========================================================
# UTILIDADES
# =========================================================
MONTHS_ES = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"
}

REQUIRED = [
    "CARGA", "FECHA CUM", "CLIENTE", "Quien Creó", "ESTADO OP",
    "TRAY. PROP", "LINEA NEG", "T.NEGOCIO",
    "V.CLIENTE", "V.CONDUCT", "T. VEHICULO",
    "PLACA", "CONDUCTO.1", "ESTADO FA", "FAC PROVE"
]


def norm_state(s: pd.Series) -> pd.Series:
    return (
        s.astype("string")
        .str.strip()
        .str.upper()
        .str.replace("Á", "A", regex=False)
        .str.replace("É", "E", regex=False)
        .str.replace("Í", "I", regex=False)
        .str.replace("Ó", "O", regex=False)
        .str.replace("Ú", "U", regex=False)
    )


@st.cache_data(show_spinner=False)
def load_excel(raw: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(raw), sheet_name="TODOS", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError("Faltan columnas requeridas: " + ", ".join(missing))

    # Fecha maestra: CARGA.
    df["CARGA"] = pd.to_datetime(df["CARGA"], errors="coerce")
    df["FECHA_CARGA"] = df["CARGA"].dt.normalize()

    for c in ["FECHA R", "F PRECUMP", "FECHA CUM"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    for c in ["V.CLIENTE", "V.CONDUCT", "TRAY. PROP"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["ESTADO OP N"] = norm_state(df["ESTADO OP"])
    df["Tipo Flota"] = np.select(
        [df["TRAY. PROP"].eq(1), df["TRAY. PROP"].eq(0)],
        ["FLOTA PROPIA", "TERCEROS"],
        default="SIN CLASIFICAR",
    )

    # Tipología: campo T. VEHICULO.
    df["TIPOLOGIA"] = (
        df["T. VEHICULO"]
        .astype("string")
        .fillna("SIN TIPOLOGÍA")
        .str.strip()
    )

    # Extrae capacidad numérica cuando el nombre contiene "XX PASAJEROS".
    cap = (
        df["TIPOLOGIA"]
        .str.extract(r"(\d+)\s*PASAJER", flags=re.IGNORECASE, expand=False)
    )
    df["CAPACIDAD_PAX"] = pd.to_numeric(cap, errors="coerce")

    # Conductor oficial: columna W en el archivo = CONDUCTO.1
    df["CONDUCTOR_NOMBRE"] = (
        df["CONDUCTO.1"].astype("string").fillna("SIN CONDUCTOR").str.strip()
    )

    df["PLACA"] = (
        df["PLACA"]
        .astype("string")
        .fillna("SIN PLACA")
        .str.strip()
        .str.upper()
    )
    df["CLIENTE"] = (
        df["CLIENTE"]
        .astype("string")
        .fillna("SIN CLIENTE")
        .str.strip()
        .str.upper()
    )
    df["Quien Creó"] = df["Quien Creó"].astype("string").fillna("SIN COORDINADOR").str.strip()

    df["AÑO"] = df["FECHA_CARGA"].dt.year
    df["MES_NUM"] = df["FECHA_CARGA"].dt.month
    df["MES"] = df["MES_NUM"].map(MONTHS_ES)
    df["BIMESTRE_NUM"] = ((df["MES_NUM"] - 1) // 2 + 1).astype("Int64")
    df["TRIMESTRE_NUM"] = ((df["MES_NUM"] - 1) // 3 + 1).astype("Int64")
    df["SEMESTRE_NUM"] = np.where(df["MES_NUM"] <= 6, 1, 2)

    return df


def valid_services(df: pd.DataFrame) -> pd.DataFrame:
    return df.loc[df["ESTADO OP N"].ne("ANULADO")].copy()


def fmt_int(x) -> str:
    if pd.isna(x):
        return "0"
    return f"{int(round(x)):,}".replace(",", ".")


def fmt_money(x) -> str:
    """COP completo y fácil de leer: $7.020.000.000."""
    if pd.isna(x):
        return "$0"
    return "$" + f"{int(round(float(x))):,}".replace(",", ".")


def fmt_pct(x) -> str:
    if pd.isna(x):
        return "—"
    return f"{x*100:.2f}%".replace(".", ",")


def delta_html(delta, good_when_up=True, suffix="%"):
    if delta is None or pd.isna(delta):
        return ""
    favorable = delta >= 0 if good_when_up else delta <= 0
    color = GREEN if favorable else RED
    arrow = "▲" if delta > 0 else ("▼" if delta < 0 else "▬")

    # Tanto porcentajes como puntos porcentuales parten de una razón decimal.
    if suffix in ("%", " p.p."):
        val = abs(delta) * 100
    else:
        val = abs(delta)

    label = f"{val:.2f}{suffix}".replace(".", ",")
    return (
        f"<span style='color:{color}'>{arrow} {label}</span> "
        f"<span style='color:{TEXT_2}'>último periodo seleccionado vs anterior</span>"
    )


def kpi_card(label, value, delta=None, good_when_up=True, note="", suffix="%", show_delta=True):
    d = delta_html(delta, good_when_up, suffix) if show_delta else ""
    delta_block = f'<div class="kpi-delta">{d}</div>' if (show_delta and d) else ""
    note_block = f'<div class="kpi-note">{note}</div>' if note else ""

    html = (
        '<div class="kpi">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f'{delta_block}'
        f'{note_block}'
        '</div>'
    )
    st.markdown(html, unsafe_allow_html=True)


def safe_div(a, b):
    return a / b if b not in (0, None) and not pd.isna(b) else np.nan


def aggregate_period(df: pd.DataFrame, grain: str) -> pd.DataFrame:
    x = df.copy()
    if grain == "Mensual":
        x["PERIODO_ORDEN"] = x["AÑO"] * 100 + x["MES_NUM"]
        x["PERIODO"] = x.apply(lambda r: f"{MONTHS_ES.get(r['MES_NUM'], '')} {int(r['AÑO'])}", axis=1)
    elif grain == "Bimestral":
        x["PERIODO_ORDEN"] = x["AÑO"] * 10 + x["BIMESTRE_NUM"].astype(int)
        x["PERIODO"] = x.apply(lambda r: f"B{int(r['BIMESTRE_NUM'])} {int(r['AÑO'])}", axis=1)
    elif grain == "Trimestral":
        x["PERIODO_ORDEN"] = x["AÑO"] * 10 + x["TRIMESTRE_NUM"].astype(int)
        x["PERIODO"] = x.apply(lambda r: f"T{int(r['TRIMESTRE_NUM'])} {int(r['AÑO'])}", axis=1)
    else:
        x["PERIODO_ORDEN"] = x["AÑO"] * 10 + x["SEMESTRE_NUM"].astype(int)
        x["PERIODO"] = x.apply(lambda r: f"S{int(r['SEMESTRE_NUM'])} {int(r['AÑO'])}", axis=1)

    g = (
        x.groupby(["PERIODO_ORDEN", "PERIODO"], dropna=False)
        .agg(
            Servicios=("CLIENTE", "size"),
            Facturacion=("V.CLIENTE", "sum"),
            Costos=("V.CONDUCT", "sum"),
        )
        .reset_index()
        .sort_values("PERIODO_ORDEN")
    )
    g["Margen"] = g["Facturacion"] - g["Costos"]
    g["Rentabilidad"] = np.where(g["Facturacion"].ne(0), g["Margen"] / g["Facturacion"], np.nan)
    g["Var_Servicios"] = g["Servicios"].pct_change()
    g["Var_Facturacion"] = g["Facturacion"].pct_change()
    g["Var_Margen"] = g["Margen"].pct_change()
    g["Var_Rent_PP"] = g["Rentabilidad"].diff()

    # Protección estadística:
    # si el último periodo está incompleto, NO se calcula una variación contra
    # un periodo anterior completo, porque sería una comparación no homogénea.
    max_date = x["FECHA_CARGA"].max()
    periodo_incompleto = False

    if pd.notna(max_date):
        year = int(max_date.year)
        month = int(max_date.month)

        if grain == "Mensual":
            last_day = calendar.monthrange(year, month)[1]
            periodo_incompleto = int(max_date.day) < last_day

        elif grain == "Bimestral":
            end_month = ((month - 1) // 2 + 1) * 2
            last_day = calendar.monthrange(year, end_month)[1]
            periodo_incompleto = not (
                month == end_month and int(max_date.day) == last_day
            )

        elif grain == "Trimestral":
            end_month = ((month - 1) // 3 + 1) * 3
            last_day = calendar.monthrange(year, end_month)[1]
            periodo_incompleto = not (
                month == end_month and int(max_date.day) == last_day
            )

        else:
            end_month = 6 if month <= 6 else 12
            last_day = calendar.monthrange(year, end_month)[1]
            periodo_incompleto = not (
                month == end_month and int(max_date.day) == last_day
            )

    g["Periodo_Incompleto"] = False
    if len(g) and periodo_incompleto:
        idx = g.index[-1]
        g.loc[idx, "Periodo_Incompleto"] = True
        g.loc[idx, ["Var_Servicios", "Var_Facturacion", "Var_Margen", "Var_Rent_PP"]] = np.nan

    return g


def trend_chart(g: pd.DataFrame, value_col: str, var_col: str, title: str, money=False):
    vals = g[value_col]
    labels = [fmt_money(v) if money else fmt_int(v) for v in vals]

    var_real_pct = g[var_col].astype(float) * 100
    visual_cap = 200.0
    var_visual_pct = var_real_pct.clip(-visual_cap, visual_cap)

    text = []
    custom = []
    for v in var_real_pct:
        if pd.isna(v):
            text.append("")
            custom.append(["Sin comparación homogénea"])
        else:
            arrow = "▲" if v > 0 else ("▼" if v < 0 else "▬")
            marker = " *" if abs(v) > visual_cap else ""
            text.append(f"{arrow} {abs(v):.2f}%{marker}")
            custom.append([f"{v:.2f}%"])

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=g["PERIODO"],
            y=vals,
            name=title,
            marker_color=ROYAL_BLUE,
            text=labels,
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{x}<br>%{text}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=g["PERIODO"],
            y=var_visual_pct,
            name="Variación %",
            mode="lines+markers+text",
            line=dict(color=SECONDARY_BLUE, width=2),
            marker=dict(size=8),
            text=text,
            customdata=np.array(custom, dtype=object),
            textposition="top center",
            yaxis="y2",
            hovertemplate="%{x}<br>Variación real: %{customdata[0]}<extra></extra>",
        )
    )

    finite = var_visual_pct.dropna()
    if len(finite):
        limit = max(30.0, min(visual_cap, float(np.nanmax(np.abs(finite))) * 1.2))
    else:
        limit = 30.0

    fig.update_layout(
        height=360,
        margin=dict(l=20, r=30, t=40, b=20),
        plot_bgcolor="#101C2C",
        paper_bgcolor="#07111F",
        font=dict(color=TEXT),
        legend=dict(orientation="h", y=1.12, x=0),
        yaxis=dict(gridcolor=GRID, zeroline=False, title="COP" if money else "Servicios"),
        yaxis2=dict(
            overlaying="y",
            side="right",
            title="Variación %",
            ticksuffix="%",
            showgrid=False,
            range=[-limit, limit],
        ),
        hovermode="x unified",
    )
    return fig


def top_table(df, group, value="Servicios", n=7):
    v = valid_services(df)
    g = (
        v.groupby(group, dropna=False)
        .agg(
            Servicios=("CLIENTE", "size"),
            Facturacion=("V.CLIENTE", "sum"),
            Costos=("V.CONDUCT", "sum"),
        )
        .reset_index()
    )
    g["Margen"] = g["Facturacion"] - g["Costos"]
    g["Rentabilidad %"] = np.where(g["Facturacion"].ne(0), g["Margen"] / g["Facturacion"], np.nan)
    if value == "Facturacion":
        g = g.sort_values("Facturacion", ascending=False)
    elif value == "Margen":
        g = g.sort_values("Margen", ascending=False)
    else:
        g = g.sort_values("Servicios", ascending=False)
    return g.head(n)


# =========================================================
# NAVEGACIÓN PRINCIPAL
# =========================================================
with st.sidebar:
    st.markdown("### 🧭 Navegación")
    page_mode = st.radio(
        "Módulo",
        ["CENTRO DE CONTROL", "POSICIÓN FINANCIERA Y PROYECCIONES"],
        index=0,
        label_visibility="collapsed",
        key="main_navigation",
    )

if page_mode == "CENTRO DE CONTROL":
    st.markdown(
        """
        <div class="hero">
            <div class="hero-title">📊 CENTRO DE CONTROL GERENCIAL VSE</div>
            <div class="hero-sub">Operación · Servicios · Producción · Rentabilidad · Cierre</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        """
        <div class="hero">
            <div class="hero-title">💹 POSICIÓN FINANCIERA Y PROYECCIONES</div>
            <div class="hero-sub">Producción · Recaudo · Costos · Obligaciones · Margen · Proyección próximo mes</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# =========================================================
# BASE VIGENTE · CARGA AUTOMÁTICA + ACTUALIZACIÓN COMPARTIDA
# =========================================================
MASTER_FILE = Path("TRAYECTOS TODOS__2026.xlsx")
BOGOTA_TZ = ZoneInfo("America/Bogota")


def _secret(name: str, default: str = "") -> str:
    """Lee Streamlit Secrets sin impedir el arranque si aún no están configurados."""
    try:
        value = st.secrets.get(name, default)
        return str(value).strip() if value is not None else default
    except Exception:
        return default


GITHUB_TOKEN = _secret("GITHUB_TOKEN")
GITHUB_REPO = _secret("GITHUB_REPO")  # Ej.: empresa/SEGUIMIENTO_OP_VSE
GITHUB_BRANCH = _secret("GITHUB_BRANCH", "main")
GITHUB_MASTER_PATH = _secret("GITHUB_MASTER_PATH", MASTER_FILE.name)

GITHUB_SHARED_READY = bool(GITHUB_TOKEN and GITHUB_REPO and GITHUB_MASTER_PATH)


def github_headers() -> dict:
    return {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def fetch_master_from_github():
    """Descarga el Excel maestro como binario RAW autenticado."""
    if not GITHUB_SHARED_READY:
        return None

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_MASTER_PATH}"

    # Metadata para verificar que existe y obtener SHA.
    meta_headers = github_headers()
    meta_headers["Accept"] = "application/vnd.github+json"
    meta_response = requests.get(
        url, headers=meta_headers, params={"ref": GITHUB_BRANCH}, timeout=25
    )
    if meta_response.status_code != 200:
        return None

    meta = meta_response.json()

    # Para Excel de varios MB no usamos el campo content/base64 del JSON.
    # Pedimos directamente los bytes RAW reales del archivo.
    raw_headers = github_headers()
    raw_headers["Accept"] = "application/vnd.github.raw+json"
    raw_response = requests.get(
        url, headers=raw_headers, params={"ref": GITHUB_BRANCH}, timeout=60
    )
    if raw_response.status_code != 200:
        return None

    raw_file = raw_response.content
    if not raw_file:
        return None

    if raw_file.startswith(b"version https://git-lfs.github.com/spec"):
        raise ValueError(
            "El archivo de GitHub es un puntero Git LFS, no el Excel real."
        )

    # Los .xlsx son contenedores ZIP y comienzan normalmente por PK.
    if not raw_file.startswith(b"PK"):
        raise ValueError(
            "GitHub no devolvió un archivo .xlsx válido."
        )

    updated_at = None
    commits_url = f"https://api.github.com/repos/{GITHUB_REPO}/commits"
    commits = requests.get(
        commits_url,
        headers=github_headers(),
        params={"path": GITHUB_MASTER_PATH, "sha": GITHUB_BRANCH, "per_page": 1},
        timeout=20,
    )
    if commits.status_code == 200 and commits.json():
        iso_date = commits.json()[0].get("commit", {}).get("committer", {}).get("date")
        if iso_date:
            updated_at = (
                datetime.fromisoformat(iso_date.replace("Z", "+00:00"))
                .astimezone(BOGOTA_TZ)
            )

    return {
        "raw": raw_file,
        "sha": meta.get("sha"),
        "name": Path(GITHUB_MASTER_PATH).name,
        "source": "GitHub compartido",
        "updated_at": updated_at,
    }

def publish_master_to_github(raw_file: bytes, original_name: str):
    """Publica la nueva base como archivo maestro compartido."""
    if not GITHUB_SHARED_READY:
        return False, "La persistencia compartida de GitHub no está configurada."

    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_MASTER_PATH}"
    try:
        current = requests.get(
            url,
            headers=github_headers(),
            params={"ref": GITHUB_BRANCH},
            timeout=25,
        )

        payload = {
            "message": (
                "Actualiza base maestra VSE desde dashboard · "
                + datetime.now(BOGOTA_TZ).strftime("%Y-%m-%d %H:%M")
            ),
            "content": base64.b64encode(raw_file).decode("ascii"),
            "branch": GITHUB_BRANCH,
        }

        if current.status_code == 200:
            payload["sha"] = current.json().get("sha")
        elif current.status_code not in (404,):
            return False, f"GitHub respondió {current.status_code} al consultar la base vigente."

        update = requests.put(
            url,
            headers=github_headers(),
            json=payload,
            timeout=45,
        )

        if update.status_code not in (200, 201):
            detail = update.json().get("message", update.text[:200])
            return False, f"No fue posible publicar la base: {detail}"

        return True, "Base maestra publicada correctamente."
    except Exception as exc:
        return False, f"No fue posible conectar con GitHub: {exc}"


# 1. La app abre SIEMPRE con la última base vigente.
master_info = None

# En la sesión que acaba de publicar, usar inmediatamente la nueva versión.
if st.session_state.get("master_override_raw"):
    master_info = {
        "raw": st.session_state["master_override_raw"],
        "name": MASTER_FILE.name,
        "source": "Actualización recién publicada",
        "updated_at": st.session_state.get("master_override_updated_at"),
    }
else:
    # Si hay Secrets configurados, consultar GitHub en tiempo real.
    master_info = fetch_master_from_github()

# Respaldo 1: archivo maestro con nombre esperado en el repositorio/despliegue.
if master_info is None and MASTER_FILE.exists():
    master_info = {
        "raw": MASTER_FILE.read_bytes(),
        "name": MASTER_FILE.name,
        "source": "Base incluida en el despliegue",
        "updated_at": None,
    }

# Respaldo 2: autodetectar cualquier Excel local que contenga la hoja TODOS.
# Esto evita que el dashboard quede vacío si el archivo fue subido con otro nombre.
if master_info is None:
    local_candidates = sorted(
        [p for p in Path(".").glob("*.xlsx") if not p.name.startswith("~$")],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    for candidate in local_candidates:
        try:
            candidate_raw = candidate.read_bytes()
            # Validar estructura con la misma función oficial del dashboard.
            _candidate_df = load_excel(candidate_raw)
            if _candidate_df["FECHA_CARGA"].notna().sum() > 0:
                master_info = {
                    "raw": candidate_raw,
                    "name": candidate.name,
                    "source": "Base Excel autodetectada en el repositorio",
                    "updated_at": datetime.fromtimestamp(
                        candidate.stat().st_mtime, tz=BOGOTA_TZ
                    ),
                }
                break
        except Exception:
            continue

# Respaldo 3: si NO existe ninguna base persistente, permitir carga temporal
# en lugar de dejar una pantalla de error sin salida.
if master_info is None:
    st.warning(
        "⚠️ No se encontró una base maestra persistente en el despliegue. "
        "Puedes cargar una base ahora para visualizar el dashboard en esta sesión."
    )
    emergency_upload = st.file_uploader(
        "Cargar base temporal para iniciar",
        type=["xlsx"],
        accept_multiple_files=False,
        key="emergency_master_upload",
        help="La hoja debe llamarse TODOS y conservar las columnas requeridas.",
    )

    if emergency_upload is None:
        st.info(
            "Para que el dashboard abra automáticamente para todos, deja al menos un archivo Excel "
            "válido en el repositorio o configura la persistencia compartida desde Streamlit Secrets."
        )
        st.stop()

    emergency_raw = emergency_upload.getvalue()
    try:
        emergency_df = load_excel(emergency_raw)
        if emergency_df["FECHA_CARGA"].notna().sum() == 0:
            raise ValueError("La columna CARGA no contiene fechas válidas.")
        master_info = {
            "raw": emergency_raw,
            "name": emergency_upload.name,
            "source": "Base temporal cargada en esta sesión",
            "updated_at": datetime.now(BOGOTA_TZ),
        }
    except Exception as exc:
        st.error(f"❌ La base temporal no es válida: {exc}")
        st.stop()

raw = master_info["raw"]

try:
    df_all = load_excel(raw)
except ValueError as e:
    st.error(f"❌ La base vigente tiene una estructura no válida: {e}")
    st.stop()
except Exception as e:
    st.error(f"❌ No fue posible leer la base vigente: {e}")
    st.stop()

if df_all["FECHA_CARGA"].notna().sum() == 0:
    st.error("❌ La columna CARGA de la base vigente no contiene fechas válidas.")
    st.stop()

data_cutoff = df_all["FECHA_CARGA"].max()
data_start = df_all["FECHA_CARGA"].min()
valid_rows = len(df_all)
valid_dates = int(df_all["FECHA_CARGA"].notna().sum())

updated_label = (
    master_info["updated_at"].strftime("%d/%m/%Y %H:%M")
    if master_info.get("updated_at")
    else "versión del despliegue"
)

# Trazabilidad visible sin bloquear el dashboard.
st.markdown(
    f"""
    <div class="base-status">
        <strong>Base vigente:</strong> {master_info["name"]}
        &nbsp;·&nbsp; <strong>Registros:</strong> {fmt_int(valid_rows)}
        &nbsp;·&nbsp; <strong>Corte:</strong> {data_cutoff.strftime('%d/%m/%Y') if pd.notna(data_cutoff) else 'N/D'}
        &nbsp;·&nbsp; <strong>Actualización:</strong> {updated_label}
        &nbsp;·&nbsp; <strong>Origen:</strong> {master_info["source"]}
    </div>
    """,
    unsafe_allow_html=True,
)

# 2. Actualización opcional: no bloquea la visualización.
with st.sidebar:
    st.markdown("### 📁 Base vigente")
    st.success("✅ Dashboard disponible")
    st.caption(f"Corte actual: {data_cutoff.strftime('%d/%m/%Y') if pd.notna(data_cutoff) else 'N/D'}")
    st.caption(f"Registros: {fmt_int(valid_rows)}")

    with st.expander("🔄 Actualizar base vigente", expanded=False):
        st.caption(
            "Carga una nueva base únicamente cuando quieras reemplazar la información visible para todos."
        )

        uploaded_file = st.file_uploader(
            "Nueva base Excel",
            type=["xlsx"],
            accept_multiple_files=False,
            help="El nombre puede variar. La hoja debe llamarse TODOS.",
            key="shared_master_uploader",
        )

        # Plantilla vacía de estructura.
        template_cols = REQUIRED.copy()
        for extra_col in ["FECHA R", "F PRECUMP"]:
            if extra_col not in template_cols:
                template_cols.append(extra_col)

        template_buffer = BytesIO()
        with pd.ExcelWriter(template_buffer, engine="openpyxl") as writer:
            pd.DataFrame(columns=template_cols).to_excel(
                writer, sheet_name="TODOS", index=False
            )
        template_buffer.seek(0)

        st.download_button(
            "⬇️ Plantilla de estructura",
            data=template_buffer.getvalue(),
            file_name="PLANTILLA_BASE_VSE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        if uploaded_file is not None:
            candidate_raw = uploaded_file.getvalue()

            try:
                candidate_df = load_excel(candidate_raw)
                if candidate_df["FECHA_CARGA"].notna().sum() == 0:
                    raise ValueError("La columna CARGA no contiene fechas válidas.")

                candidate_cutoff = candidate_df["FECHA_CARGA"].max()
                candidate_start = candidate_df["FECHA_CARGA"].min()

                st.success("✓ Estructura validada")
                st.caption(f"Archivo: {uploaded_file.name}")
                st.caption(f"Registros: {fmt_int(len(candidate_df))}")
                if pd.notna(candidate_start) and pd.notna(candidate_cutoff):
                    st.caption(
                        "Periodo: "
                        f"{candidate_start.strftime('%d/%m/%Y')} — "
                        f"{candidate_cutoff.strftime('%d/%m/%Y')}"
                    )

                if not GITHUB_SHARED_READY:
                    st.warning(
                        "Para que esta actualización quede visible para todos, "
                        "configura GITHUB_TOKEN, GITHUB_REPO, GITHUB_BRANCH y GITHUB_MASTER_PATH "
                        "en Streamlit Secrets."
                    )

                publish_clicked = st.button(
                    "✅ Publicar como base vigente",
                    use_container_width=True,
                    disabled=not GITHUB_SHARED_READY,
                    key="publish_shared_master",
                )

                if publish_clicked:
                    with st.spinner("Publicando y actualizando el dashboard..."):
                        ok, message = publish_master_to_github(
                            candidate_raw, uploaded_file.name
                        )
                    if ok:
                        load_excel.clear()
                        st.session_state["master_override_raw"] = candidate_raw
                        st.session_state["master_override_updated_at"] = datetime.now(BOGOTA_TZ)
                        st.success("✅ Nueva base publicada. El dashboard se actualizará ahora.")
                        st.rerun()
                    else:
                        st.error(message)

            except Exception as e:
                st.error(f"❌ Base no válida: {e}")
                st.caption(
                    "No se publicará ningún cambio hasta que la estructura sea correcta."
                )


# =========================================================
# FILTROS
# =========================================================
with st.sidebar:
    st.markdown("### 🔎 Filtros principales")

    grain = st.selectbox(
        "Periodo de análisis",
        ["Mensual", "Bimestral", "Trimestral", "Semestral"],
        index=0,
    )

    years = sorted(df_all["AÑO"].dropna().astype(int).unique().tolist())
    selected_years = st.multiselect("Año", years, default=years)

    min_date = df_all["FECHA_CARGA"].min().date()
    max_date = df_all["FECHA_CARGA"].max().date()
    date_range = st.date_input(
        "Rango de fechas (CARGA)",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date, end_date = min_date, max_date

    # Meses en español, seleccionables de forma múltiple.
    # Siempre se calculan desde FECHA_CARGA.
    MESES_ES = {
        1: "ENERO",
        2: "FEBRERO",
        3: "MARZO",
        4: "ABRIL",
        5: "MAYO",
        6: "JUNIO",
        7: "JULIO",
        8: "AGOSTO",
        9: "SEPTIEMBRE",
        10: "OCTUBRE",
        11: "NOVIEMBRE",
        12: "DICIEMBRE",
    }
    meses_disponibles_num = sorted(
        df_all["FECHA_CARGA"].dropna().dt.month.unique().tolist()
    )
    meses_disponibles = [
        MESES_ES[int(m)] for m in meses_disponibles_num
    ]

    selected_months = st.multiselect(
        "Mes(es)",
        options=meses_disponibles,
        default=[],
        placeholder="Todos los meses",
        help="Selecciona uno o varios meses. El filtro usa la fecha CARGA.",
    )

    def multi_filter(label, col):
        opts = sorted([x for x in df_all[col].dropna().astype(str).unique().tolist() if x and x != "<NA>"])
        return st.multiselect(label, opts, default=[])

    f_client = multi_filter("Cliente", "CLIENTE")
    f_coord = multi_filter("Quién creó (Coordinador)", "Quien Creó")
    f_estado = multi_filter("Estado Operativo", "ESTADO OP N")
    f_flota = multi_filter("Tipo Flota", "Tipo Flota")
    f_linea = multi_filter("Línea de Negocio", "LINEA NEG")
    f_negocio = multi_filter("Tipo de Negocio", "T.NEGOCIO")
    f_tipologia = multi_filter("Tipología / Tipo Vehículo", "TIPOLOGIA")

    with st.expander("Filtros avanzados", expanded=False):
        f_placa = multi_filter("Placa", "PLACA")
        f_conductor = multi_filter("Conductor", "CONDUCTOR_NOMBRE")

    if st.button("↻ Limpiar filtros", use_container_width=True):
        st.rerun()


def apply_filters(df):
    x = df.copy()
    x = x[x["AÑO"].isin(selected_years)] if selected_years else x.iloc[0:0]
    x = x[
        x["FECHA_CARGA"].between(
            pd.Timestamp(start_date),
            pd.Timestamp(end_date) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1),
        )
    ]

    # Filtro múltiple de meses sobre CARGA.
    if selected_months:
        selected_month_numbers = [
            num for num, nombre in MESES_ES.items()
            if nombre in selected_months
        ]
        x = x[x["FECHA_CARGA"].dt.month.isin(selected_month_numbers)]

    mapping = [
        ("CLIENTE", f_client),
        ("Quien Creó", f_coord),
        ("ESTADO OP N", f_estado),
        ("Tipo Flota", f_flota),
        ("LINEA NEG", f_linea),
        ("T.NEGOCIO", f_negocio),
        ("TIPOLOGIA", f_tipologia),
        ("PLACA", f_placa),
        ("CONDUCTOR_NOMBRE", f_conductor),
    ]
    for col, values in mapping:
        if values:
            x = x[x[col].astype(str).isin(values)]
    return x


df = apply_filters(df_all)
valid = valid_services(df)

if valid.empty:
    st.warning("No hay registros válidos para los filtros seleccionados.")
    st.stop()


# =========================================================
# PÁGINA 2 · POSICIÓN FINANCIERA Y PROYECCIONES
# =========================================================
if page_mode == "POSICIÓN FINANCIERA Y PROYECCIONES":
    st.markdown(
        """
        <style>
            .stApp { background: #07111F !important; color: #FFFFFF !important; }
            .block-container { max-width: 1800px; }
            h1, h2, h3, h4 { color: #FFFFFF !important; }
            p, span, label { color: #E2E8F0; }
            [data-testid="stSidebar"] { background: #0B1626 !important; border-right: 1px solid #263B55 !important; }
            [data-testid="stSidebar"] * { color: #E8EEF7 !important; }
            .hero { background: linear-gradient(90deg, #0A2D59 0%, #064AA3 58%, #0B5ED7 100%) !important; box-shadow: 0 6px 22px rgba(0,0,0,.25) !important; }
            .section-title { background: #0F2D52 !important; color: #FFFFFF !important; border: 1px solid #263B55; }
            .fin-kpi { background: #0F1C2E; border: 1px solid #263B55; border-left: 5px solid #1769AA; border-radius: 12px; padding: 13px 14px; min-height: 118px; box-shadow: 0 3px 12px rgba(0,0,0,.20); }
            .fin-kpi .label { color: #CBD5E1; font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: .35px; }
            .fin-kpi .value { color: #FFFFFF; font-size: 23px; font-weight: 850; margin-top: 6px; }
            .fin-kpi .note { color: #94A3B8; font-size: 10.5px; margin-top: 5px; line-height: 1.25; }
            .fin-kpi.good { border-left-color: #2DA66F; }
            .fin-kpi.warn { border-left-color: #D98E04; }
            .fin-kpi.bad { border-left-color: #C94B4B; }
            .fin-panel { background: #0F1C2E; border: 1px solid #263B55; border-radius: 12px; padding: 14px 16px; margin: 6px 0 12px 0; }
            .fin-panel strong { color: #FFFFFF; }
            .fin-note { color:#CBD5E1; font-size: 11px; }
            div[data-testid="stDataFrame"] { border: 1px solid #263B55 !important; border-radius: 10px; overflow: hidden; }
            div[data-testid="stDownloadButton"] button { background:#0F2D52; color:white; border:1px solid #2E5A89; }
            div[data-testid="stDownloadButton"] button:hover { border-color:#64B5F6; color:white; }
            div[data-testid="stSelectbox"] label,
            div[data-testid="stMultiSelect"] label,
            div[data-testid="stDateInput"] label,
            div[data-testid="stRadio"] label { color:#E2E8F0 !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    def fin_card(label, value, note="", tone=""):
        tone_class = f" {tone}" if tone else ""
        html = (
            f'<div class="fin-kpi{tone_class}">'
            f'<div class="label">{label}</div>'
            f'<div class="value">{value}</div>'
            f'<div class="note">{note}</div>'
            '</div>'
        )
        st.markdown(html, unsafe_allow_html=True)

    def norm_fin_state(series: pd.Series) -> pd.Series:
        return norm_state(series.fillna(""))

    fin = valid.copy()
    fin["ESTADO_FA_N"] = norm_fin_state(fin["ESTADO FA"])
    fin["FAC_PROVE_N"] = norm_fin_state(fin["FAC PROVE"])

    fin["CLAS_CLIENTE"] = np.where(
        fin["ESTADO_FA_N"].eq("FACT DEFINITIVA"),
        "RECAUDADO",
        "PENDIENTE POR RECAUDAR",
    )

    # Compatibilidad histórica: la base local usa PAGADO; la nueva nomenclatura puede usar FACTURADO.
    paid_supplier_states = {"PAGADO", "FACTURADO"}
    fin["CLAS_PROVE"] = np.select(
        [fin["FAC_PROVE_N"].isin(paid_supplier_states), fin["FAC_PROVE_N"].eq("EN ESPERA")],
        ["PAGADO A TERCEROS", "PENDIENTE POR PAGAR"],
        default="POR CLASIFICAR",
    )

    servicios_fin = len(fin)
    produccion_fin = float(fin["V.CLIENTE"].fillna(0).sum())
    costo_fin = float(fin["V.CONDUCT"].fillna(0).sum())
    margen_fin = produccion_fin - costo_fin
    rent_fin = safe_div(margen_fin, produccion_fin)

    recaudado_cliente = float(fin.loc[fin["CLAS_CLIENTE"].eq("RECAUDADO"), "V.CLIENTE"].fillna(0).sum())
    pendiente_recaudar = float(fin.loc[fin["CLAS_CLIENTE"].eq("PENDIENTE POR RECAUDAR"), "V.CLIENTE"].fillna(0).sum())
    pct_recaudo = safe_div(recaudado_cliente, produccion_fin)

    pagado_terceros = float(fin.loc[fin["CLAS_PROVE"].eq("PAGADO A TERCEROS"), "V.CONDUCT"].fillna(0).sum())
    pendiente_pagar = float(fin.loc[fin["CLAS_PROVE"].eq("PENDIENTE POR PAGAR"), "V.CONDUCT"].fillna(0).sum())
    por_clasificar = float(fin.loc[fin["CLAS_PROVE"].eq("POR CLASIFICAR"), "V.CONDUCT"].fillna(0).sum())
    pct_pago_terceros = safe_div(pagado_terceros, costo_fin)
    pct_pendiente_pago = safe_div(pendiente_pagar, costo_fin)

    caja_operativa = recaudado_cliente - pagado_terceros
    brecha_pendiente = pendiente_recaudar - pendiente_pagar
    cobertura_obligaciones_caja = safe_div(caja_operativa, pendiente_pagar)

    tol_cop = 1.0
    tol_pct = 0.0001
    ctrl_margen = abs((produccion_fin - costo_fin) - margen_fin) < tol_cop
    ctrl_cliente = abs((recaudado_cliente + pendiente_recaudar) - produccion_fin) < tol_cop
    ctrl_prove = abs((pagado_terceros + pendiente_pagar + por_clasificar) - costo_fin) < tol_cop
    ctrl_rent = True if produccion_fin == 0 else abs((margen_fin / produccion_fin) - rent_fin) < tol_pct
    controles_ok = all([ctrl_margen, ctrl_cliente, ctrl_prove, ctrl_rent])

    st.markdown("## 1. Resultado económico")
    st.caption(
        "Primera lectura ejecutiva: cuánto se produjo, cuánto costó prestar los servicios "
        "y qué margen directo dejó la operación."
    )

    e1, e2, e3, e4 = st.columns(4)
    with e1:
        fin_card("Producción", fmt_money(produccion_fin), "Valor total de servicios · Σ V.CLIENTE")
    with e2:
        fin_card("Costo del servicio", fmt_money(costo_fin), "Costo directo asociado · Σ V.CONDUCT")
    with e3:
        fin_card("Margen directo", fmt_money(margen_fin), "Producción - Costo del Servicio", "good" if margen_fin >= 0 else "bad")
    with e4:
        fin_card("Rentabilidad directa", fmt_pct(rent_fin), "Margen directo / Producción", "good" if (pd.notna(rent_fin) and rent_fin >= 0) else "bad")

    st.markdown("## 2. Recaudo y obligaciones")
    st.caption(
        "Segunda lectura: qué parte de la producción ya fue recaudada y qué parte del costo "
        "ya fue pagada a terceros."
    )

    st.markdown('<div class="subsection-tag">CLIENTES</div>', unsafe_allow_html=True)
    rc1, rc2, rc3 = st.columns(3)
    with rc1:
        fin_card("Recaudado cliente", fmt_money(recaudado_cliente), "ESTADO FA = FACT DEFINITIVA", "good")
    with rc2:
        fin_card("Pendiente por recaudar", fmt_money(pendiente_recaudar), f"{fmt_pct(safe_div(pendiente_recaudar, produccion_fin))} de la producción", "warn" if pendiente_recaudar > 0 else "good")
    with rc3:
        fin_card("% Recaudo", fmt_pct(pct_recaudo), "Recaudado / Producción", "good")

    st.markdown('<div class="subsection-tag">TERCEROS</div>', unsafe_allow_html=True)
    pt1, pt2, pt3 = st.columns(3)
    with pt1:
        fin_card("Pagado a terceros", fmt_money(pagado_terceros), "FAC PROVE = PAGADO / FACTURADO", "good")
    with pt2:
        fin_card("Pendiente por pagar", fmt_money(pendiente_pagar), f"{fmt_pct(pct_pendiente_pago)} del costo", "warn" if pendiente_pagar > 0 else "good")
    with pt3:
        fin_card("% Pago a terceros", fmt_pct(pct_pago_terceros), "Pagado a terceros / Costo del Servicio")

    st.markdown("## 3. Posición de caja operativa")
    cx1, cx2, cx3 = st.columns(3)
    with cx1:
        fin_card("Caja operativa", fmt_money(caja_operativa), "Recaudado cliente - Pagado a terceros", "good" if caja_operativa >= 0 else "bad")
    with cx2:
        fin_card("Cobertura de obligaciones", fmt_pct(cobertura_obligaciones_caja), "Caja operativa / Pendiente por Pagar", "good" if (pd.notna(cobertura_obligaciones_caja) and cobertura_obligaciones_caja >= 1) else "warn")
    with cx3:
        fin_card("Brecha pendiente", fmt_money(brecha_pendiente), "Pendiente por recaudar - Pendiente por pagar", "good" if brecha_pendiente >= 0 else "bad")

    st.markdown(
        '<div class="fin-panel"><strong>Lectura ejecutiva:</strong>'
        '<span class="fin-note"> la caja operativa corresponde al recaudo registrado menos pagos a terceros '
        'del universo filtrado. No equivale al saldo bancario total ni a la utilidad contable.</span></div>',
        unsafe_allow_html=True,
    )

    with st.expander("✓ Control matemático y calidad de estados", expanded=False):
        if controles_ok:
            st.success(
                "Las identidades principales cuadran dentro de la tolerancia definida: "
                "Producción = Recaudado + Pendiente; Costo = Pagado + Pendiente + Por clasificar; "
                "Margen = Producción - Costo."
            )
        else:
            st.error(
                "REVISAR DATOS: al menos una identidad financiera no cuadra dentro de la tolerancia definida."
            )

        if por_clasificar > tol_cop:
            unknown_states = sorted([
                x for x in fin.loc[
                    fin["CLAS_PROVE"].eq("POR CLASIFICAR"), "FAC_PROVE_N"
                ].dropna().unique().tolist() if x
            ])
            st.warning(
                "FAC PROVE contiene valores no clasificados. "
                f"Costo por clasificar: {fmt_money(por_clasificar)}. "
                "Estados detectados: "
                + (", ".join(unknown_states) if unknown_states else "VACÍO")
            )


    def grouped_financial(data: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
        base = data.groupby(group_cols, dropna=False).agg(Servicios=("CLIENTE", "size"), Produccion=("V.CLIENTE", "sum"), Costo=("V.CONDUCT", "sum")).reset_index()

        def add_group_sum(mask, value_col, out_col):
            s = data.loc[mask].groupby(group_cols, dropna=False)[value_col].sum().rename(out_col).reset_index()
            return s

        for mask, value_col, out_col in [
            (data["CLAS_CLIENTE"].eq("RECAUDADO"), "V.CLIENTE", "Recaudado"),
            (data["CLAS_CLIENTE"].eq("PENDIENTE POR RECAUDAR"), "V.CLIENTE", "Pendiente Recaudar"),
            (data["CLAS_PROVE"].eq("PAGADO A TERCEROS"), "V.CONDUCT", "Pagado Terceros"),
            (data["CLAS_PROVE"].eq("PENDIENTE POR PAGAR"), "V.CONDUCT", "Pendiente Pagar"),
            (data["CLAS_PROVE"].eq("POR CLASIFICAR"), "V.CONDUCT", "Por Clasificar"),
        ]:
            base = base.merge(add_group_sum(mask, value_col, out_col), on=group_cols, how="left")

        for c in ["Recaudado", "Pendiente Recaudar", "Pagado Terceros", "Pendiente Pagar", "Por Clasificar"]:
            base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)
        base["Margen"] = base["Produccion"] - base["Costo"]
        base["Rentabilidad"] = np.where(base["Produccion"].ne(0), base["Margen"] / base["Produccion"], np.nan)
        base["% Recaudo"] = np.where(base["Produccion"].ne(0), base["Recaudado"] / base["Produccion"], np.nan)
        base["% Pago Terceros"] = np.where(base["Costo"].ne(0), base["Pagado Terceros"] / base["Costo"], np.nan)
        return base

    cliente_fin = grouped_financial(fin, ["CLIENTE"]).sort_values("Produccion", ascending=False).reset_index(drop=True)

    st.markdown("## 4. Evolución financiera mensual")
    st.caption(
        "Comportamiento mes a mes usando FECHA CUM como fecha de corte financiero-operativo. "
        "FACTURADO corresponde a servicios en estado CUMPLIDO; RECAUDADO corresponde a ESTADO FA = FACT DEFINITIVA; "
        "PAGADO A TERCEROS corresponde a FAC PROVE = PAGADO/FACTURADO. "
        "Mientras no exista una fecha bancaria independiente para recaudo/pago proveedor, estos movimientos se atribuyen al mes de FECHA CUM."
    )

    evo = df_all.copy()
    evo = valid_services(evo)

    # Mantener los filtros de negocio, pero permitir observar todo el histórico mensual.
    evo_filters = [
        ("CLIENTE", f_client),
        ("Quien Creó", f_coord),
        ("ESTADO OP N", f_estado),
        ("Tipo Flota", f_flota),
        ("LINEA NEG", f_linea),
        ("T.NEGOCIO", f_negocio),
        ("TIPOLOGIA", f_tipologia),
        ("PLACA", f_placa),
        ("CONDUCTOR_NOMBRE", f_conductor),
    ]
    for _col, _vals in evo_filters:
        if _vals:
            evo = evo[evo[_col].astype(str).isin(_vals)]

    evo["ESTADO_FA_N"] = norm_fin_state(evo["ESTADO FA"])
    evo["FAC_PROVE_N"] = norm_fin_state(evo["FAC PROVE"])

    # FECHA CUM ya viene normalizada a datetime desde el cargue de datos.
    evo = evo[evo["FECHA CUM"].notna()].copy()
    evo["PERIODO_CUM"] = evo["FECHA CUM"].dt.to_period("M")

    # FACTURADO: valor cliente de servicios que alcanzaron CUMPLIDO.
    fact_m = (
        evo.loc[evo["ESTADO OP N"].eq("CUMPLIDO")]
        .groupby("PERIODO_CUM")["V.CLIENTE"].sum()
        .rename("Facturado")
    )

    # RECAUDADO: regla de negocio definida por Vigía.
    rec_m = (
        evo.loc[evo["ESTADO_FA_N"].eq("FACT DEFINITIVA")]
        .groupby("PERIODO_CUM")["V.CLIENTE"].sum()
        .rename("Recaudado")
    )

    # PAGADO A TERCEROS: estados PAGADO o FACTURADO en FAC PROVE.
    pay_m = (
        evo.loc[evo["FAC_PROVE_N"].isin(paid_supplier_states)]
        .groupby("PERIODO_CUM")["V.CONDUCT"].sum()
        .rename("Pagado Terceros")
    )

    srv_m = (
        evo.loc[evo["ESTADO OP N"].eq("CUMPLIDO")]
        .groupby("PERIODO_CUM").size()
        .rename("Servicios Cumplidos")
    )

    if len(evo):
        min_per = evo["PERIODO_CUM"].min()
        max_per = evo["PERIODO_CUM"].max()
        all_periods = pd.period_range(min_per, max_per, freq="M")
        monthly_evo = pd.DataFrame(index=all_periods)
        monthly_evo.index.name = "Periodo"
        monthly_evo = monthly_evo.join([fact_m, rec_m, pay_m, srv_m]).fillna(0).reset_index()
    else:
        monthly_evo = pd.DataFrame(columns=["Periodo","Facturado","Recaudado","Pagado Terceros","Servicios Cumplidos"])

    if not monthly_evo.empty:
        monthly_evo["Caja Operativa"] = monthly_evo["Recaudado"] - monthly_evo["Pagado Terceros"]
        monthly_evo["% Recaudo sobre Facturado"] = np.where(
            monthly_evo["Facturado"].ne(0),
            monthly_evo["Recaudado"] / monthly_evo["Facturado"],
            np.nan,
        )
        monthly_evo["Mes"] = monthly_evo["Periodo"].apply(
            lambda p: f"{MESES_ES[int(p.month)]} {int(p.year)}"
        )

        # Mostrar por defecto los últimos 12 meses disponibles.
        monthly_view = monthly_evo.tail(12).copy()

        fig_evo = go.Figure()
        fig_evo.add_trace(go.Bar(
            x=monthly_view["Mes"], y=monthly_view["Facturado"],
            name="Facturado", marker_color="#1769AA",
            hovertemplate="<b>%{x}</b><br>Facturado: $%{y:,.0f}<extra></extra>"
        ))
        fig_evo.add_trace(go.Bar(
            x=monthly_view["Mes"], y=monthly_view["Recaudado"],
            name="Recaudado", marker_color="#2DA66F",
            hovertemplate="<b>%{x}</b><br>Recaudado: $%{y:,.0f}<extra></extra>"
        ))
        fig_evo.add_trace(go.Bar(
            x=monthly_view["Mes"], y=monthly_view["Pagado Terceros"],
            name="Pagado a terceros", marker_color="#64B5F6",
            hovertemplate="<b>%{x}</b><br>Pagado a terceros: $%{y:,.0f}<extra></extra>"
        ))
        fig_evo.update_layout(
            barmode="group",
            height=450,
            margin=dict(l=20, r=20, t=50, b=90),
            paper_bgcolor="#07111F",
            plot_bgcolor="#07111F",
            font=dict(color="#E2E8F0"),
            legend=dict(orientation="h", y=1.08, x=0),
            xaxis=dict(title="", tickangle=-30, gridcolor="#1E3652"),
            yaxis=dict(title="Valor COP", gridcolor="#1E3652", tickformat="~s"),
            title="FACTURADO · RECAUDADO · PAGADO A TERCEROS — EVOLUCIÓN MENSUAL",
        )
        st.plotly_chart(fig_evo, use_container_width=True, key="monthly_financial_evolution")

        # Tarjetas del último mes disponible.
        last_row = monthly_evo.iloc[-1]
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            fin_card("Facturado último mes", fmt_money(last_row["Facturado"]), str(last_row["Mes"]))
        with m2:
            fin_card("Recaudado último mes", fmt_money(last_row["Recaudado"]), f"{fmt_pct(last_row['% Recaudo sobre Facturado'])} del facturado", "good")
        with m3:
            fin_card("Pagado terceros último mes", fmt_money(last_row["Pagado Terceros"]), str(last_row["Mes"]))
        with m4:
            fin_card(
                "Caja operativa del mes",
                fmt_money(last_row["Caja Operativa"]),
                "Recaudado - Pagado a terceros",
                "good" if last_row["Caja Operativa"] >= 0 else "bad",
            )

        st.markdown('<div class="section-title">DETALLE MENSUAL</div>', unsafe_allow_html=True)
        evo_show = monthly_view[[
            "Mes","Servicios Cumplidos","Facturado","Recaudado","Pagado Terceros",
            "Caja Operativa","% Recaudo sobre Facturado"
        ]].copy()
        evo_show.columns = [
            "Mes","Servicios Cumplidos","Facturado","Recaudado","Pagado a Terceros",
            "Caja Operativa","% Recaudo"
        ]
        st.dataframe(
            evo_show.style.format({
                "Servicios Cumplidos": "{:,.0f}",
                "Facturado": lambda x: "$"+f"{x:,.0f}".replace(",", "."),
                "Recaudado": lambda x: "$"+f"{x:,.0f}".replace(",", "."),
                "Pagado a Terceros": lambda x: "$"+f"{x:,.0f}".replace(",", "."),
                "Caja Operativa": lambda x: "$"+f"{x:,.0f}".replace(",", "."),
                "% Recaudo": "{:.2%}",
            }),
            use_container_width=True,
            hide_index=True,
            height=min(470, 38 + 35 * max(1, len(evo_show))),
        )

        st.markdown(
            '<div class="fin-panel"><strong>Cómo leer esta gráfica:</strong>'
            '<span class="fin-note"> Facturado muestra el valor de los servicios que llegaron a CUMPLIDO en cada mes; '
            'Recaudado muestra cuánto de esos valores aparece como FACT DEFINITIVA; Pagado a Terceros muestra el costo '
            'asociado a registros marcados como PAGADO/FACTURADO. La diferencia Recaudado - Pagado a Terceros se muestra '
            'como Caja Operativa mensual del universo analizado.</span></div>',
            unsafe_allow_html=True,
        )
    else:
        st.info("No hay registros con FECHA CUM para construir la evolución financiera mensual.")


    st.markdown("## 5. Análisis por cliente y terceros")
    st.markdown("### 5.1 Recaudo por cliente")
    st.markdown('<div class="section-title">PRODUCCIÓN · RECAUDADO · PENDIENTE POR RECAUDAR POR CLIENTE</div>', unsafe_allow_html=True)
    top_cli = cliente_fin.head(15)
    fig_cli = go.Figure()
    for col, name, color in [("Produccion", "Producción", ROYAL_BLUE), ("Recaudado", "Recaudado", GREEN), ("Pendiente Recaudar", "Pendiente por Recaudar", AMBER)]:
        fig_cli.add_trace(go.Bar(x=top_cli["CLIENTE"], y=top_cli[col], name=name, marker_color=color, hovertemplate=f"<b>%{{x}}</b><br>{name}: $%{{y:,.0f}}<extra></extra>"))
    fig_cli.update_layout(barmode="group", height=440, margin=dict(l=20,r=20,t=45,b=140), paper_bgcolor="#07111F", plot_bgcolor="#0F1C2E", font=dict(color="#E8EEF7"), legend=dict(orientation="h", y=1.12, x=0), xaxis=dict(tickangle=-45, gridcolor="#263B55"), yaxis=dict(title="COP", gridcolor="#263B55"), bargap=.22, bargroupgap=.06)
    st.plotly_chart(fig_cli, use_container_width=True, key="fin_cliente_chart")

    cliente_show = cliente_fin[["CLIENTE", "Servicios", "Produccion", "Recaudado", "Pendiente Recaudar", "% Recaudo", "Costo", "Pagado Terceros", "Pendiente Pagar", "% Pago Terceros", "Margen", "Rentabilidad"]].copy()
    cliente_show.columns = ["Cliente", "Servicios", "Producción", "Recaudado", "Pendiente por Recaudar", "% Recaudo", "Costo del Servicio", "Pagado a Terceros", "Pendiente por Pagar", "% Pago Terceros", "Margen", "Rentabilidad"]
    st.dataframe(cliente_show.style.format({"Servicios":"{:,.0f}", "Producción":lambda x: "$"+f"{x:,.0f}".replace(",","."), "Recaudado":lambda x: "$"+f"{x:,.0f}".replace(",","."), "Pendiente por Recaudar":lambda x: "$"+f"{x:,.0f}".replace(",","."), "Costo del Servicio":lambda x: "$"+f"{x:,.0f}".replace(",","."), "Pagado a Terceros":lambda x: "$"+f"{x:,.0f}".replace(",","."), "Pendiente por Pagar":lambda x: "$"+f"{x:,.0f}".replace(",","."), "Margen":lambda x: "$"+f"{x:,.0f}".replace(",","."), "% Recaudo":"{:.2%}", "% Pago Terceros":"{:.2%}", "Rentabilidad":"{:.2%}"}), use_container_width=True, hide_index=True, height=460)

    st.markdown("### 5.2 Pagos a terceros")
    supplier_mode = st.radio("Analizar obligaciones por", ["CONDUCTOR", "PLACA", "CLIENTE"], horizontal=True, key="supplier_mode_fin")
    supplier_col = {"CONDUCTOR":"CONDUCTOR_NOMBRE", "PLACA":"PLACA", "CLIENTE":"CLIENTE"}[supplier_mode]
    supplier_fin = grouped_financial(fin, [supplier_col]).sort_values("Costo", ascending=False).reset_index(drop=True)
    top_sup = supplier_fin.head(15)
    fig_sup = go.Figure()
    for col, name, color in [("Costo", "Costo del Servicio", ROYAL_BLUE), ("Pagado Terceros", "Pagado a Terceros", GREEN), ("Pendiente Pagar", "Pendiente por Pagar", AMBER)]:
        fig_sup.add_trace(go.Bar(x=top_sup[supplier_col], y=top_sup[col], name=name, marker_color=color, hovertemplate=f"<b>%{{x}}</b><br>{name}: $%{{y:,.0f}}<extra></extra>"))
    fig_sup.update_layout(barmode="group", height=440, margin=dict(l=20,r=20,t=45,b=140), paper_bgcolor="#07111F", plot_bgcolor="#0F1C2E", font=dict(color="#E8EEF7"), legend=dict(orientation="h", y=1.12, x=0), xaxis=dict(tickangle=-45, gridcolor="#263B55"), yaxis=dict(title="COP", gridcolor="#263B55"))
    st.plotly_chart(fig_sup, use_container_width=True, key=f"fin_supplier_chart_{supplier_mode}")

    tercero_det = grouped_financial(fin, ["CONDUCTOR_NOMBRE", "PLACA", "CLIENTE"]).sort_values("Pendiente Pagar", ascending=False)
    tercero_show = tercero_det[["CONDUCTOR_NOMBRE","PLACA","CLIENTE","Servicios","Costo","Pagado Terceros","Pendiente Pagar","% Pago Terceros"]].copy()
    tercero_show.columns = ["Conductor","Placa","Cliente","Servicios","Costo del Servicio","Pagado a Terceros","Pendiente por Pagar","% Pago"]
    st.dataframe(tercero_show.style.format({"Servicios":"{:,.0f}", "Costo del Servicio":lambda x:"$"+f"{x:,.0f}".replace(",","."), "Pagado a Terceros":lambda x:"$"+f"{x:,.0f}".replace(",","."), "Pendiente por Pagar":lambda x:"$"+f"{x:,.0f}".replace(",","."), "% Pago":"{:.2%}"}), use_container_width=True, hide_index=True, height=470)

    st.markdown("### 5.3 Exposición prioritaria")
    ex1, ex2 = st.columns(2)
    with ex1:
        st.markdown('<div class="section-title">TOP CLIENTES · PENDIENTE POR RECAUDAR</div>', unsafe_allow_html=True)
        top_pending_client = cliente_show.sort_values("Pendiente por Recaudar", ascending=False).head(10)
        st.dataframe(top_pending_client[["Cliente","Producción","Recaudado","Pendiente por Recaudar","% Recaudo"]].style.format({"Producción":lambda x:"$"+f"{x:,.0f}".replace(",","."), "Recaudado":lambda x:"$"+f"{x:,.0f}".replace(",","."), "Pendiente por Recaudar":lambda x:"$"+f"{x:,.0f}".replace(",","."), "% Recaudo":"{:.2%}"}), use_container_width=True, hide_index=True, height=390)
    with ex2:
        st.markdown('<div class="section-title">TOP TERCEROS · PENDIENTE POR PAGAR</div>', unsafe_allow_html=True)
        top_pending_sup = supplier_fin.sort_values("Pendiente Pagar", ascending=False).head(10).copy()
        top_pending_sup["% Pendiente"] = np.where(top_pending_sup["Costo"].ne(0), top_pending_sup["Pendiente Pagar"] / top_pending_sup["Costo"], np.nan)
        top_pending_sup = top_pending_sup[[supplier_col,"Costo","Pagado Terceros","Pendiente Pagar","% Pendiente"]]
        top_pending_sup.columns = [supplier_mode.title(),"Costo del Servicio","Pagado","Pendiente por Pagar","% Pendiente"]
        st.dataframe(top_pending_sup.style.format({"Costo del Servicio":lambda x:"$"+f"{x:,.0f}".replace(",","."), "Pagado":lambda x:"$"+f"{x:,.0f}".replace(",","."), "Pendiente por Pagar":lambda x:"$"+f"{x:,.0f}".replace(",","."), "% Pendiente":"{:.2%}"}), use_container_width=True, hide_index=True, height=390)


    # =========================================================
    # 6. PROYECCIÓN ESTADÍSTICA — PRÓXIMO MES
    # =========================================================
    st.markdown("## 6. Proyección estadística — Próximo mes")
    st.caption("Proyección auditable construida con los últimos 6 meses completos. El escenario BASE usa tendencia lineal reciente; CONSERVADOR y OPTIMISTA reflejan la variabilidad histórica observada.")

    # ---------------------------------------------------------
    # Universo histórico para forecast
    # Se conservan filtros de negocio (cliente, coordinador, flota,
    # línea, tipo de negocio, tipología, placa y conductor), pero NO
    # se limita el histórico por Mes/Año/Rango de fecha seleccionado,
    # porque el modelo necesita meses anteriores completos.
    # ---------------------------------------------------------
    hist = df_all.copy()
    hist = valid_services(hist)
    hist_filters = [
        ("CLIENTE", f_client),
        ("Quien Creó", f_coord),
        ("ESTADO OP N", f_estado),
        ("Tipo Flota", f_flota),
        ("LINEA NEG", f_linea),
        ("T.NEGOCIO", f_negocio),
        ("TIPOLOGIA", f_tipologia),
        ("PLACA", f_placa),
        ("CONDUCTOR_NOMBRE", f_conductor),
    ]
    for _col, _vals in hist_filters:
        if _vals:
            hist = hist[hist[_col].astype(str).isin(_vals)]

    hist["ESTADO_FA_N"] = norm_fin_state(hist["ESTADO FA"])
    hist["FAC_PROVE_N"] = norm_fin_state(hist["FAC PROVE"])
    hist["CLAS_CLIENTE"] = np.where(hist["ESTADO_FA_N"].eq("FACT DEFINITIVA"), "RECAUDADO", "PENDIENTE POR RECAUDAR")
    hist["CLAS_PROVE"] = np.select(
        [hist["FAC_PROVE_N"].isin(paid_supplier_states), hist["FAC_PROVE_N"].eq("EN ESPERA")],
        ["PAGADO A TERCEROS", "PENDIENTE POR PAGAR"],
        default="POR CLASIFICAR",
    )

    today_fc = pd.Timestamp.now().normalize()
    current_month_start = today_fc.replace(day=1)
    last_complete_month_end = current_month_start - pd.Timedelta(days=1)
    forecast_month_start = current_month_start + pd.offsets.MonthBegin(1)
    forecast_month_name = f"{MESES_ES[int(forecast_month_start.month)]} {forecast_month_start.year}"

    # Últimos seis meses COMPLETOS anteriores al mes en curso.
    train_months = pd.period_range(end=last_complete_month_end.to_period("M"), periods=6, freq="M")
    hist_train = hist[hist["FECHA_CARGA"].dt.to_period("M").isin(train_months)].copy()
    hist_train["PERIODO_M"] = hist_train["FECHA_CARGA"].dt.to_period("M")

    def monthly_forecast_frame(data: pd.DataFrame) -> pd.DataFrame:
        rows = []
        for per in train_months:
            m = data[data["PERIODO_M"].eq(per)] if "PERIODO_M" in data.columns else data[data["FECHA_CARGA"].dt.to_period("M").eq(per)]
            prod_m = float(m["V.CLIENTE"].fillna(0).sum())
            cost_m = float(m["V.CONDUCT"].fillna(0).sum())
            rec_m = float(m.loc[m["CLAS_CLIENTE"].eq("RECAUDADO"), "V.CLIENTE"].fillna(0).sum()) if len(m) else 0.0
            pay_m = float(m.loc[m["CLAS_PROVE"].eq("PAGADO A TERCEROS"), "V.CONDUCT"].fillna(0).sum()) if len(m) else 0.0
            rows.append({"Periodo": per, "Producción": prod_m, "Costo": cost_m, "Recaudado": rec_m, "Pagado": pay_m, "Servicios": len(m)})
        out = pd.DataFrame(rows)
        out["Ratio Costo"] = np.where(out["Producción"].ne(0), out["Costo"] / out["Producción"], np.nan)
        out["Ratio Recaudo"] = np.where(out["Producción"].ne(0), out["Recaudado"] / out["Producción"], np.nan)
        out["Ratio Pago"] = np.where(out["Costo"].ne(0), out["Pagado"] / out["Costo"], np.nan)
        return out

    monthly_fc = monthly_forecast_frame(hist_train)
    nonzero_prod = monthly_fc[monthly_fc["Producción"].gt(0)].copy()

    def weighted_ratio(series: pd.Series, default=np.nan):
        s = pd.to_numeric(series, errors="coerce").dropna()
        if s.empty:
            return default
        w = np.arange(1, len(s) + 1, dtype=float)
        return float(np.average(s.to_numpy(dtype=float), weights=w))

    def trend_forecast(values: pd.Series):
        y = pd.to_numeric(values, errors="coerce").fillna(0).to_numpy(dtype=float)
        n = len(y)
        if n == 0:
            return 0.0, 0.0, "BAJA", 0.0
        if n == 1:
            return max(0.0, y[-1]), 0.0, "BAJA", 0.0
        x = np.arange(n, dtype=float)
        slope, intercept = np.polyfit(x, y, 1)
        fitted = intercept + slope * x
        base = max(0.0, float(intercept + slope * n))
        rmse = float(np.sqrt(np.mean((y - fitted) ** 2)))
        pct = pd.Series(y).replace(0, np.nan).pct_change(fill_method=None).replace([np.inf, -np.inf], np.nan).dropna()
        vol_amount = abs(base) * float(pct.std(ddof=1)) if len(pct) >= 2 and pd.notna(pct.std(ddof=1)) else 0.0
        uncertainty = max(rmse, vol_amount)
        rel = safe_div(uncertainty, base)
        if pd.isna(rel):
            confidence = "BAJA"
        elif rel <= 0.10:
            confidence = "ALTA"
        elif rel <= 0.20:
            confidence = "MEDIA"
        else:
            confidence = "BAJA"
        return base, uncertainty, confidence, float(slope)

    if len(nonzero_prod) < 3:
        st.warning("Histórico insuficiente para una proyección estadística robusta: se requieren al menos 3 meses completos con producción. El módulo mostrará la mejor estimación disponible y marcará confianza BAJA.")

    base_prod_fc, uncertainty_fc, confidence_fc, slope_fc = trend_forecast(monthly_fc["Producción"])
    cons_prod_fc = max(0.0, base_prod_fc - uncertainty_fc)
    opt_prod_fc = max(0.0, base_prod_fc + uncertainty_fc)

    ratio_cost_fc = weighted_ratio(monthly_fc["Ratio Costo"], safe_div(costo_fin, produccion_fin))
    ratio_rec_fc = weighted_ratio(monthly_fc["Ratio Recaudo"], safe_div(recaudado_cliente, produccion_fin))
    ratio_pay_fc = weighted_ratio(monthly_fc["Ratio Pago"], safe_div(pagado_terceros, costo_fin))
    ratio_cost_fc = float(np.clip(ratio_cost_fc if pd.notna(ratio_cost_fc) else 0, 0, 5))
    ratio_rec_fc = float(np.clip(ratio_rec_fc if pd.notna(ratio_rec_fc) else 0, 0, 1))
    ratio_pay_fc = float(np.clip(ratio_pay_fc if pd.notna(ratio_pay_fc) else 0, 0, 1))

    scenario = st.radio("Escenario", ["CONSERVADOR", "BASE", "OPTIMISTA"], index=1, horizontal=True, key="forecast_scenario")
    prod_by_scenario = {"CONSERVADOR": cons_prod_fc, "BASE": base_prod_fc, "OPTIMISTA": opt_prod_fc}
    proj_prod = prod_by_scenario[scenario]
    proj_recaudo = proj_prod * ratio_rec_fc
    proj_cost = proj_prod * ratio_cost_fc
    proj_paid = proj_cost * ratio_pay_fc
    proj_margin = proj_prod - proj_cost
    proj_rent = safe_div(proj_margin, proj_prod)
    proj_cash = proj_recaudo - proj_paid
    proj_pending_rec = proj_prod - proj_recaudo
    proj_pending_pay = proj_cost - proj_paid

    conf_tone = "good" if confidence_fc == "ALTA" else ("warn" if confidence_fc == "MEDIA" else "bad")
    st.markdown(f'<div class="fin-panel"><strong>Mes proyectado:</strong> {forecast_month_name} &nbsp;·&nbsp; <strong>Confianza del modelo:</strong> {confidence_fc} &nbsp;·&nbsp; <span class="fin-note">La confianza se determina por la variabilidad/residuo del histórico de producción. Los escenarios cambian el volumen de producción; los ratios financieros usan promedios históricos ponderados para evitar supuestos arbitrarios.</span></div>', unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns(4)
    with f1: fin_card("Producción proyectada", fmt_money(proj_prod), f"Escenario {scenario}", conf_tone)
    with f2: fin_card("Recaudo proyectado", fmt_money(proj_recaudo), f"Proporción histórica esperada: {fmt_pct(ratio_rec_fc)}", "good")
    with f3: fin_card("Costo proyectado", fmt_money(proj_cost), f"Proporción histórica costo/producción: {fmt_pct(ratio_cost_fc)}")
    with f4: fin_card("Pago proyectado terceros", fmt_money(proj_paid), f"Proporción histórica de pago: {fmt_pct(ratio_pay_fc)}")
    f5, f6, f7, f8 = st.columns(4)
    with f5: fin_card("Margen proyectado", fmt_money(proj_margin), "Producción proyectada - Costo proyectado", "good" if proj_margin >= 0 else "bad")
    with f6: fin_card("Rentabilidad proyectada", fmt_pct(proj_rent), "Margen proyectado / Producción proyectada", "good" if pd.notna(proj_rent) and proj_rent >= 0 else "bad")
    with f7: fin_card("Caja operativa proyectada", fmt_money(proj_cash), "Recaudo proyectado - Pago proyectado a terceros", "good" if proj_cash >= 0 else "bad")
    with f8: fin_card("Pendiente por recaudar proyectado", fmt_money(proj_pending_rec), "Producción proyectada - Recaudo proyectado", "warn" if proj_pending_rec > 0 else "good")

    # Comparativo contra el último mes completo real.
    last_row = monthly_fc.iloc[-1]
    last_label = str(last_row["Periodo"])
    comparison = pd.DataFrame([
        ["Producción", float(last_row["Producción"]), proj_prod],
        ["Recaudo", float(last_row["Recaudado"]), proj_recaudo],
        ["Costo del Servicio", float(last_row["Costo"]), proj_cost],
        ["Pago a Terceros", float(last_row["Pagado"]), proj_paid],
        ["Margen", float(last_row["Producción"] - last_row["Costo"]), proj_margin],
        ["Caja Operativa", float(last_row["Recaudado"] - last_row["Pagado"]), proj_cash],
    ], columns=["Indicador", f"Último mes completo ({last_label})", f"Proyección {forecast_month_name}"])
    comparison["Variación %"] = np.where(
        comparison[f"Último mes completo ({last_label})"].ne(0),
        comparison[f"Proyección {forecast_month_name}"] / comparison[f"Último mes completo ({last_label})"] - 1,
        np.nan,
    )
    last_rent = safe_div(float(last_row["Producción"] - last_row["Costo"]), float(last_row["Producción"]))
    comparison_rent = pd.DataFrame({
        "Indicador": ["Rentabilidad"],
        f"Último mes completo ({last_label})": [last_rent],
        f"Proyección {forecast_month_name}": [proj_rent],
        "Variación p.p.": [(proj_rent - last_rent) * 100 if pd.notna(proj_rent) and pd.notna(last_rent) else np.nan],
    })

    st.markdown('<div class="section-title">ÚLTIMO MES COMPLETO VS PRÓXIMO MES PROYECTADO</div>', unsafe_allow_html=True)
    st.dataframe(
        comparison.style.format({
            f"Último mes completo ({last_label})": lambda x: "$" + f"{x:,.0f}".replace(",", "."),
            f"Proyección {forecast_month_name}": lambda x: "$" + f"{x:,.0f}".replace(",", "."),
            "Variación %": "{:+.2%}",
        }),
        use_container_width=True, hide_index=True,
    )
    st.dataframe(
        comparison_rent.style.format({
            f"Último mes completo ({last_label})": "{:.2%}",
            f"Proyección {forecast_month_name}": "{:.2%}",
            "Variación p.p.": "{:+.2f}",
        }),
        use_container_width=True, hide_index=True,
    )

    # Histórico de 6 meses + forecast BASE para hacer visible la tendencia.
    fig_fc = go.Figure()
    x_hist = [str(p) for p in monthly_fc["Periodo"]]
    fig_fc.add_trace(go.Scatter(x=x_hist, y=monthly_fc["Producción"], mode="lines+markers", name="Producción histórica", line=dict(color="#64B5F6", width=3), marker=dict(size=8)))
    fig_fc.add_trace(go.Scatter(x=[x_hist[-1], forecast_month_name], y=[float(last_row["Producción"]), base_prod_fc], mode="lines+markers", name="Proyección BASE", line=dict(color="#2DA66F", width=3, dash="dash"), marker=dict(size=10)))
    fig_fc.add_trace(go.Scatter(x=[forecast_month_name, forecast_month_name], y=[cons_prod_fc, opt_prod_fc], mode="lines+markers", name="Rango escenarios", line=dict(color="#D98E04", width=7), marker=dict(size=10), hovertemplate="Rango: $%{y:,.0f}<extra></extra>"))
    fig_fc.update_layout(height=390, margin=dict(l=20,r=20,t=45,b=40), paper_bgcolor="#07111F", plot_bgcolor="#0F1C2E", font=dict(color="#E8EEF7"), legend=dict(orientation="h", y=1.12, x=0), yaxis=dict(title="Producción COP", gridcolor="#263B55"), xaxis=dict(gridcolor="#263B55"))
    st.plotly_chart(fig_fc, use_container_width=True, key="forecast_production_history")

    # Proyección por cliente: tendencia individual de producción + ratios históricos propios.
    st.markdown('<div class="section-title">PROYECCIÓN BASE POR CLIENTE — ¿QUIÉN EXPLICA EL PRÓXIMO MES?</div>', unsafe_allow_html=True)
    client_rows = []
    for cli in sorted(hist_train["CLIENTE"].dropna().astype(str).unique().tolist()):
        cd = hist_train[hist_train["CLIENTE"].astype(str).eq(cli)].copy()
        cm = monthly_forecast_frame(cd)
        cbase, _, cconf, _ = trend_forecast(cm["Producción"])
        cr_cost = weighted_ratio(cm["Ratio Costo"], ratio_cost_fc)
        cr_rec = weighted_ratio(cm["Ratio Recaudo"], ratio_rec_fc)
        cr_cost = float(np.clip(cr_cost if pd.notna(cr_cost) else ratio_cost_fc, 0, 5))
        cr_rec = float(np.clip(cr_rec if pd.notna(cr_rec) else ratio_rec_fc, 0, 1))
        ccost = cbase * cr_cost
        crec = cbase * cr_rec
        cmarg = cbase - ccost
        client_rows.append({"Cliente":cli, "Producción Proyectada":cbase, "Recaudo Proyectado":crec, "Costo Proyectado":ccost, "Margen Proyectado":cmarg, "Rentabilidad":safe_div(cmarg,cbase), "Confianza":cconf})
    client_forecast = pd.DataFrame(client_rows)
    if not client_forecast.empty:
        total_cli_proj = float(client_forecast["Producción Proyectada"].sum())
        client_forecast["Participación Proyectada %"] = np.where(total_cli_proj != 0, client_forecast["Producción Proyectada"] / total_cli_proj, np.nan)
        client_forecast = client_forecast.sort_values("Producción Proyectada", ascending=False).reset_index(drop=True)
        st.dataframe(client_forecast.head(20).style.format({
            "Producción Proyectada":lambda x:"$"+f"{x:,.0f}".replace(",","."),
            "Recaudo Proyectado":lambda x:"$"+f"{x:,.0f}".replace(",","."),
            "Costo Proyectado":lambda x:"$"+f"{x:,.0f}".replace(",","."),
            "Margen Proyectado":lambda x:"$"+f"{x:,.0f}".replace(",","."),
            "Rentabilidad":"{:.2%}", "Participación Proyectada %":"{:.2%}"
        }), use_container_width=True, hide_index=True, height=520)
    else:
        client_forecast = pd.DataFrame(columns=["Cliente","Producción Proyectada","Recaudo Proyectado","Costo Proyectado","Margen Proyectado","Rentabilidad","Confianza","Participación Proyectada %"])
        st.info("No hay histórico suficiente por cliente para construir el detalle de proyección.")

    with st.expander("🔎 Ver metodología de proyección", expanded=False):
        st.markdown(f"""
**Horizonte:** {forecast_month_name}.  
**Histórico utilizado:** últimos 6 meses completos anteriores al mes en curso: {', '.join(str(p) for p in train_months)}.  
**Producción BASE:** tendencia lineal por mínimos cuadrados sobre la producción mensual.  
**Rango CONSERVADOR/OPTIMISTA:** BASE ± incertidumbre observada; la incertidumbre toma el mayor valor entre RMSE de la tendencia y volatilidad mensual reciente aplicada al forecast. El escenario conservador nunca baja de cero.  
**Costo:** Producción proyectada × ratio histórico ponderado Costo/Producción.  
**Recaudo:** Producción proyectada × ratio histórico ponderado Recaudado/Producción.  
**Pago a terceros:** Costo proyectado × ratio histórico ponderado Pagado/Costo.  
**Margen:** Producción − Costo.  
**Rentabilidad:** Margen / Producción.  
**Caja operativa proyectada:** Recaudo proyectado − Pago proyectado a terceros.  
**Ponderación de proporciones históricas:** los meses más recientes reciben mayor peso (1,2,3,4,5,6).  
**Confianza:** ALTA si la incertidumbre relativa ≤10%; MEDIA si ≤20%; BAJA si >20% o el histórico es insuficiente.

**Advertencia metodológica:** sin una columna de fecha efectiva de recaudo/pago, el proyección de recaudo y pago estima la proporción esperada a partir de los estados financieros observados en la base; no modela todavía el desfase exacto en días de entrada o salida de caja.
        """)
        methodology_table = monthly_fc.copy()
        methodology_table["Periodo"] = methodology_table["Periodo"].astype(str)
        st.dataframe(methodology_table.style.format({"Producción":lambda x:"$"+f"{x:,.0f}".replace(",","."), "Costo":lambda x:"$"+f"{x:,.0f}".replace(",","."), "Recaudado":lambda x:"$"+f"{x:,.0f}".replace(",","."), "Pagado":lambda x:"$"+f"{x:,.0f}".replace(",","."), "Ratio Costo":"{:.2%}", "Ratio Recaudo":"{:.2%}", "Ratio Pago":"{:.2%}"}), use_container_width=True, hide_index=True)

    projection_table = pd.DataFrame([
        {"Escenario":"CONSERVADOR", "Producción":cons_prod_fc, "Recaudo":cons_prod_fc*ratio_rec_fc, "Costo del Servicio":cons_prod_fc*ratio_cost_fc, "Pago a Terceros":cons_prod_fc*ratio_cost_fc*ratio_pay_fc, "Margen":cons_prod_fc-(cons_prod_fc*ratio_cost_fc), "Rentabilidad":safe_div(cons_prod_fc-(cons_prod_fc*ratio_cost_fc),cons_prod_fc), "Caja Operativa":cons_prod_fc*ratio_rec_fc-(cons_prod_fc*ratio_cost_fc*ratio_pay_fc)},
        {"Escenario":"BASE", "Producción":base_prod_fc, "Recaudo":base_prod_fc*ratio_rec_fc, "Costo del Servicio":base_prod_fc*ratio_cost_fc, "Pago a Terceros":base_prod_fc*ratio_cost_fc*ratio_pay_fc, "Margen":base_prod_fc-(base_prod_fc*ratio_cost_fc), "Rentabilidad":safe_div(base_prod_fc-(base_prod_fc*ratio_cost_fc),base_prod_fc), "Caja Operativa":base_prod_fc*ratio_rec_fc-(base_prod_fc*ratio_cost_fc*ratio_pay_fc)},
        {"Escenario":"OPTIMISTA", "Producción":opt_prod_fc, "Recaudo":opt_prod_fc*ratio_rec_fc, "Costo del Servicio":opt_prod_fc*ratio_cost_fc, "Pago a Terceros":opt_prod_fc*ratio_cost_fc*ratio_pay_fc, "Margen":opt_prod_fc-(opt_prod_fc*ratio_cost_fc), "Rentabilidad":safe_div(opt_prod_fc-(opt_prod_fc*ratio_cost_fc),opt_prod_fc), "Caja Operativa":opt_prod_fc*ratio_rec_fc-(opt_prod_fc*ratio_cost_fc*ratio_pay_fc)},
    ])

    st.info("La proyección estadística es una estimación basada en comportamiento histórico; no constituye producción, recaudo ni pagos garantizados. El escenario BASE es la referencia central y los escenarios CONSERVADOR/OPTIMISTA expresan el rango derivado de la variabilidad observada.")

    st.markdown("## 7. Descarga ejecutiva")
    resumen_export = pd.DataFrame([
        ["Servicios", servicios_fin], ["Producción", produccion_fin], ["Recaudado cliente", recaudado_cliente], ["Pendiente por recaudar", pendiente_recaudar], ["% Recaudo", pct_recaudo], ["Costo del Servicio", costo_fin], ["Pagado a terceros", pagado_terceros], ["Pendiente por pagar", pendiente_pagar], ["Por clasificar", por_clasificar], ["% Pago a terceros", pct_pago_terceros], ["Margen", margen_fin], ["Rentabilidad", rent_fin], ["Caja operativa", caja_operativa], ["Brecha pendiente", brecha_pendiente], ["Cobertura de obligaciones con caja", cobertura_obligaciones_caja]
    ], columns=["Indicador","Valor"])

    fin_excel = BytesIO()
    with pd.ExcelWriter(fin_excel, engine="openpyxl") as writer:
        resumen_export.to_excel(writer, index=False, sheet_name="Resumen")
        cliente_show.to_excel(writer, index=False, sheet_name="Clientes")
        tercero_show.to_excel(writer, index=False, sheet_name="Terceros")
        projection_table.to_excel(writer, index=False, sheet_name="Forecast_Escenarios")
        comparison.to_excel(writer, index=False, sheet_name="Forecast_Comparativo")
        comparison_rent.to_excel(writer, index=False, sheet_name="Forecast_Rentabilidad")
        client_forecast.to_excel(writer, index=False, sheet_name="Forecast_Clientes")
        methodology_table.to_excel(writer, index=False, sheet_name="Forecast_Metodologia")
        pd.DataFrame({"Control":["Margen", "Cliente", "Terceros", "Rentabilidad"], "Resultado":[ctrl_margen, ctrl_cliente, ctrl_prove, ctrl_rent]}).to_excel(writer, index=False, sheet_name="Auditoria")
        for ws in writer.book.worksheets:
            fill = openpyxl.styles.PatternFill(fill_type="solid", fgColor="003B8E")
            font = openpyxl.styles.Font(color="FFFFFF", bold=True)
            for cell in ws[1]:
                cell.fill = fill
                cell.font = font
                cell.alignment = openpyxl.styles.Alignment(horizontal="center")
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col_cells in ws.columns:
                max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 12), 42)
    fin_excel.seek(0)

    st.download_button("⬇️ Descargar Posición Financiera y Proyecciones", data=fin_excel, file_name="Posicion_Financiera_y_Proyecciones_VSE.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.stop()


# =========================================================
# KPI Y COMPARATIVOS
# =========================================================
prod = len(valid)
fact = valid["V.CLIENTE"].sum()
cost = valid["V.CONDUCT"].sum()
margin = fact - cost
rent = safe_div(margin, fact)

cumplidos = int((valid["ESTADO OP N"] == "CUMPLIDO").sum())
cumpl_oper = int((valid["ESTADO OP N"] == "CUMPLIDO OPERATIVO").sum())
en_transito = int((valid["ESTADO OP N"] == "EN TRANSITO").sum())
en_programacion = int((valid["ESTADO OP N"] == "EN PROGRAMACION").sum())

# Regla de negocio confirmada:
# Pendiente operativo = EN TRANSITO + EN PROGRAMACION.
pend_oper = en_transito + en_programacion
closure = safe_div(cumplidos, prod)

valor_flota_propia = valid.loc[
    valid["Tipo Flota"].eq("FLOTA PROPIA"), "V.CLIENTE"
].sum()

valor_terceros = valid.loc[
    valid["Tipo Flota"].eq("TERCEROS"), "V.CLIENTE"
].sum()

trend = aggregate_period(valid, grain)
last = trend.iloc[-1]
prev = trend.iloc[-2] if len(trend) > 1 else None

ultimo_periodo_incompleto = bool(
    "Periodo_Incompleto" in trend.columns
    and len(trend)
    and trend.iloc[-1]["Periodo_Incompleto"]
)

d_prod = last["Var_Servicios"] if prev is not None else np.nan
d_fact = last["Var_Facturacion"] if prev is not None else np.nan
d_margin = last["Var_Margen"] if prev is not None else np.nan
d_rent_pp = last["Var_Rent_PP"] if prev is not None else np.nan

# Para pendientes y cierre, compara el último periodo con el inmediatamente anterior.
def period_tag(x, grain):
    if grain == "Mensual":
        return x["AÑO"] * 100 + x["MES_NUM"]
    if grain == "Bimestral":
        return x["AÑO"] * 10 + x["BIMESTRE_NUM"]
    if grain == "Trimestral":
        return x["AÑO"] * 10 + x["TRIMESTRE_NUM"]
    return x["AÑO"] * 10 + x["SEMESTRE_NUM"]

tmp = valid.copy()
tmp["PKEY"] = period_tag(tmp, grain)
keys = sorted(tmp["PKEY"].dropna().unique())
d_pend = np.nan
d_close = np.nan
if len(keys) >= 2:
    a = tmp[tmp["PKEY"] == keys[-1]]
    b = tmp[tmp["PKEY"] == keys[-2]]
    pa = ((a["ESTADO OP N"] == "EN TRANSITO") | (a["ESTADO OP N"] == "EN PROGRAMACION")).sum()
    pb = ((b["ESTADO OP N"] == "EN TRANSITO") | (b["ESTADO OP N"] == "EN PROGRAMACION")).sum()
    ca = safe_div((a["ESTADO OP N"] == "CUMPLIDO").sum(), len(a))
    cb = safe_div((b["ESTADO OP N"] == "CUMPLIDO").sum(), len(b))
    d_pend = safe_div(pa - pb, pb) if pb else (0 if pa == 0 else np.nan)
    d_close = ca - cb if not (pd.isna(ca) or pd.isna(cb)) else np.nan

k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
with k1:
    kpi_card("Total Servicios", fmt_int(prod), d_prod, True, "Servicios válidos")
with k2:
    kpi_card("Producción", fmt_money(fact), d_fact, True, "V.CLIENTE")
with k3:
    kpi_card(
        "Costo del Servicio",
        fmt_money(cost),
        note="Valor asociado a la prestación",
        show_delta=False,
    )
with k4:
    kpi_card("Margen", fmt_money(margin), d_margin, True, "Producción - Costo del Servicio")
with k5:
    kpi_card(
        "Rentabilidad", fmt_pct(rent),
        d_rent_pp, True, "Margen / Producción", suffix=" p.p."
    )
with k6:
    kpi_card(
        "Valor Flota Propia",
        fmt_money(valor_flota_propia),
        note="Producción · TRAY. PROP = 1",
        show_delta=False,
    )
with k7:
    kpi_card(
        "Valor Tercero",
        fmt_money(valor_terceros),
        note="Producción · TRAY. PROP = 0",
        show_delta=False,
    )



st.markdown("## 1. Resumen ejecutivo")
st.caption(
    "Vista general del periodo y filtros seleccionados: servicios, producción, "
    "rentabilidad, clientes y comportamiento por periodo."
)

if ultimo_periodo_incompleto:
    st.info(
        "📌 El último periodo está en curso. Para evitar una comparación estadísticamente "
        "sesgada contra un periodo completo, su variación porcentual no se calcula hasta "
        "el cierre del periodo."
    )

# =========================================================
# TENDENCIAS
# =========================================================
c1, c2 = st.columns(2, gap="large")
with c1:
    st.markdown('<div class="section-title">SERVICIOS POR PERIODO</div>', unsafe_allow_html=True)
    st.plotly_chart(
        trend_chart(trend, "Servicios", "Var_Servicios", "Servicios", money=False),
        use_container_width=True,
        key="trend_prod",
    )

with c2:
    st.markdown('<div class="section-title">PRODUCCIÓN POR PERIODO</div>', unsafe_allow_html=True)
    st.plotly_chart(
        trend_chart(trend, "Facturacion", "Var_Facturacion", "Producción", money=True),
        use_container_width=True,
        key="trend_fact",
    )


# =========================================================
# CLIENTES DESTACADOS + ESTADO OPERATIVO
# =========================================================
st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-title">CLIENTES DESTACADOS</div>',
    unsafe_allow_html=True,
)
st.caption(
    "Comparación de los clientes con mayor participación por volumen de servicios "
    "y por producción. Se separa del estado operativo para mejorar la lectura ejecutiva."
)
st.markdown(
    '<div class="exec-table-note">Tablas homologadas al sistema visual ejecutivo: encabezado azul, fondo Dark Navy y acentos semánticos discretos.</div>',
    unsafe_allow_html=True,
)

# Dos tablas amplias: evita comprimir nombres y valores.
a, b = st.columns(2, gap="large")

with a:
    st.markdown(
        '<div class="subsection-tag">TOP CLIENTES POR SERVICIOS</div>',
        unsafe_allow_html=True,
    )
    top_p = top_table(df, "CLIENTE", "Servicios", 7)
    total_p = max(prod, 1)
    show_serv = top_p[["CLIENTE", "Servicios"]].copy()
    show_serv["Participación %"] = show_serv["Servicios"] / total_p
    show_serv.columns = ["Cliente", "Servicios", "Participación %"]

    styled_serv = (
        show_serv.style
        .format({"Servicios": "{:,.0f}", "Participación %": "{:.2%}"})
        .set_properties(**{
            "background-color": "#101C2C",
            "color": "#F8FAFC",
            "border-color": "#29415E",
        })
        .set_table_styles([
            {
                "selector": "th",
                "props": [
                    ("background-color", "#12365E"),
                    ("color", "#FFFFFF"),
                    ("font-weight", "700"),
                    ("border-color", "#29415E"),
                ],
            }
        ])
    )
    st.dataframe(
        styled_serv,
        use_container_width=True,
        hide_index=True,
        height=300,
    )

with b:
    st.markdown(
        '<div class="subsection-tag">TOP CLIENTES POR PRODUCCIÓN</div>',
        unsafe_allow_html=True,
    )
    top_f = top_table(df, "CLIENTE", "Facturacion", 7)
    show_prod = top_f[["CLIENTE", "Facturacion", "Margen", "Rentabilidad %"]].copy()
    show_prod.columns = ["Cliente", "Producción", "Margen", "Rentabilidad %"]

    styled_prod = (
        show_prod.style
        .format({
            "Producción": lambda x: f"${x:,.0f}",
            "Margen": lambda x: f"${x:,.0f}",
            "Rentabilidad %": "{:.2%}",
        })
        .set_properties(**{
            "background-color": "#101C2C",
            "color": "#F8FAFC",
            "border-color": "#29415E",
        })
        .set_table_styles([
            {
                "selector": "th",
                "props": [
                    ("background-color", "#12365E"),
                    ("color", "#FFFFFF"),
                    ("font-weight", "700"),
                    ("border-color", "#29415E"),
                ],
            }
        ])
    )
    st.dataframe(
        styled_prod,
        use_container_width=True,
        hide_index=True,
        height=300,
    )

# Separación visual clara antes del estado operativo.
st.markdown('<div style="height:18px"></div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-title">ESTADO OPERATIVO</div>',
    unsafe_allow_html=True,
)
st.caption(
    "Distribución de los servicios válidos por estado. "
    "El gráfico y el detalle se muestran en una fila independiente para evitar saturación visual."
)

status = (
    valid["ESTADO OP N"]
    .value_counts()
    .rename_axis("Estado")
    .reset_index(name="Servicios")
)
status["%"] = status["Servicios"] / prod

s1, s2 = st.columns([0.9, 1.1], gap="large")

with s1:
    fig = go.Figure(
        go.Pie(
            labels=status["Estado"],
            values=status["Servicios"],
            hole=.60,
            textinfo="percent",
            marker=dict(
                colors=[
                    {
                        "CUMPLIDO": GREEN,
                        "CUMPLIDO OPERATIVO": SECONDARY_BLUE,
                        "EN PROGRAMACION": AMBER,
                        "EN TRANSITO": ORANGE,
                    }.get(str(e), "#64748B")
                    for e in status["Estado"]
                ]
            ),
            hovertemplate="%{label}<br>%{value:,} servicios<br>%{percent}<extra></extra>",
        )
    )
    fig.update_layout(
        height=300,
        margin=dict(l=10, r=10, t=15, b=15),
        showlegend=True,
        legend=dict(
            font=dict(color=TEXT, size=11),
            orientation="v",
            x=1.00,
            y=0.95,
        ),
        paper_bgcolor="#07111F",
        plot_bgcolor="#07111F",
        font=dict(color=TEXT),
    )
    st.plotly_chart(fig, use_container_width=True, key="status_donut")

with s2:
    st.markdown(
        '<div class="subsection-tag">DETALLE POR ESTADO</div>',
        unsafe_allow_html=True,
    )
    styled_status = (
        status.style
        .format({"Servicios": "{:,.0f}", "%": "{:.2%}"})
        .set_properties(**{
            "background-color": "#101C2C",
            "color": "#F8FAFC",
            "border-color": "#29415E",
        })
        .set_table_styles([
            {
                "selector": "th",
                "props": [
                    ("background-color", "#12365E"),
                    ("color", "#FFFFFF"),
                    ("font-weight", "700"),
                    ("border-color", "#29415E"),
                ],
            }
        ])
    )
    st.dataframe(
        styled_status,
        use_container_width=True,
        hide_index=True,
        height=300,
    )

st.markdown('<div style="height:16px"></div>', unsafe_allow_html=True)


# =========================================================
# SEGUIMIENTO DE ESTATUS OPERATIVO
# =========================================================
st.markdown("## 2. Seguimiento de estatus operativo")
st.caption(
    "Consulta por responsable o cliente. "
    "Pendiente Operativo = EN TRANSITO + EN PROGRAMACION. "
    "CUMPLIDO OPERATIVO se muestra separado."
)

status_cumplido = int((valid["ESTADO OP N"] == "CUMPLIDO").sum())
status_cumplido_operativo = int((valid["ESTADO OP N"] == "CUMPLIDO OPERATIVO").sum())
status_programacion = int((valid["ESTADO OP N"] == "EN PROGRAMACION").sum())
status_transito = int((valid["ESTADO OP N"] == "EN TRANSITO").sum())
status_pendientes = status_programacion + status_transito

# ICO histórico: F PRECUMP indica que el servicio pasó por CUMPLIDO OPERATIVO.
if "F PRECUMP" in valid.columns:
    servicios_con_cierre_operativo = int(valid["F PRECUMP"].notna().sum())
    ico_cierre_operativo = safe_div(servicios_con_cierre_operativo, prod)
else:
    ico_cierre_operativo = np.nan

status_total_servicios = len(valid)

sc0, sc1, sc2, sc3, sc4, sc5, sc6 = st.columns(7)
with sc0:
    kpi_card("Total Servicios", fmt_int(status_total_servicios), note="Servicios válidos · excluye ANULADO", show_delta=False)
with sc1:
    kpi_card("Cumplidos", fmt_int(status_cumplido), note="Cierre final / OC OK", show_delta=False)
with sc2:
    kpi_card("Cumplidos Operativos", fmt_int(status_cumplido_operativo), note="Servicio prestado · legalización", show_delta=False)
with sc3:
    kpi_card("En Programación", fmt_int(status_programacion), note="Pendiente operativo", show_delta=False)
with sc4:
    kpi_card("En Tránsito", fmt_int(status_transito), note="Pendiente operativo", show_delta=False)
with sc5:
    kpi_card("Pendientes Operativos", fmt_int(status_pendientes), note="Programación + Tránsito", show_delta=False)
with sc6:
    kpi_card("ICO", fmt_pct(ico_cierre_operativo), note="Índice de Cierre Operativo · F PRECUMP / Servicios", show_delta=False)

def build_status_table(data: pd.DataFrame, group_col: str, group_label: str) -> pd.DataFrame:
    x = data.copy()
    main_states = {"CUMPLIDO", "CUMPLIDO OPERATIVO", "EN PROGRAMACION", "EN TRANSITO"}

    # Conteos de estados válidos
    valid_x = x.loc[x["ESTADO OP N"].ne("ANULADO")].copy()
    counts = (
        valid_x.groupby(group_col, dropna=False)["ESTADO OP N"]
        .value_counts()
        .unstack(fill_value=0)
    )

    for c in ["CUMPLIDO", "CUMPLIDO OPERATIVO", "EN PROGRAMACION", "EN TRANSITO"]:
        if c not in counts.columns:
            counts[c] = 0

    other_mask = ~valid_x["ESTADO OP N"].isin(main_states)
    others = valid_x.loc[other_mask].groupby(group_col, dropna=False).size()
    counts["OTROS"] = others.reindex(counts.index, fill_value=0)

    counts["Total general"] = counts[
        ["CUMPLIDO", "CUMPLIDO OPERATIVO", "EN PROGRAMACION", "EN TRANSITO", "OTROS"]
    ].sum(axis=1)

    counts["Pendientes operativos"] = counts["EN PROGRAMACION"] + counts["EN TRANSITO"]

    denom = counts["Total general"].replace(0, np.nan)
    counts["% Cumplido"] = counts["CUMPLIDO"] / denom
    counts["% Cumplido Operativo"] = counts["CUMPLIDO OPERATIVO"] / denom
    counts["Pendientes %"] = counts["Pendientes operativos"] / denom

    out = counts[
        [
            "CUMPLIDO",
            "CUMPLIDO OPERATIVO",
            "EN PROGRAMACION",
            "EN TRANSITO",
            "OTROS",
            "Total general",
            "Pendientes operativos",
            "% Cumplido",
            "% Cumplido Operativo",
            "Pendientes %",
        ]
    ].sort_values("Total general", ascending=False)

    # Total correcto: suma cantidades y RECALCULA porcentajes.
    total_counts = out[
        [
            "CUMPLIDO",
            "CUMPLIDO OPERATIVO",
            "EN PROGRAMACION",
            "EN TRANSITO",
            "OTROS",
            "Total general",
            "Pendientes operativos",
        ]
    ].sum()

    total_general = total_counts["Total general"]
    out.loc["TOTAL"] = {
        "CUMPLIDO": total_counts["CUMPLIDO"],
        "CUMPLIDO OPERATIVO": total_counts["CUMPLIDO OPERATIVO"],
        "EN PROGRAMACION": total_counts["EN PROGRAMACION"],
        "EN TRANSITO": total_counts["EN TRANSITO"],
        "OTROS": total_counts["OTROS"],
        "Total general": total_general,
        "Pendientes operativos": total_counts["Pendientes operativos"],
        "% Cumplido": safe_div(total_counts["CUMPLIDO"], total_general),
        "% Cumplido Operativo": safe_div(total_counts["CUMPLIDO OPERATIVO"], total_general),
        "Pendientes %": safe_div(total_counts["Pendientes operativos"], total_general),
    }

    out.index.name = group_label
    return out


status_mode = st.radio(
    "Ver tabla de estatus por",
    ["QUIÉN CREÓ", "CLIENTE"],
    horizontal=True,
    key="status_tracking_mode",
)

if status_mode == "QUIÉN CREÓ":
    status_tbl = build_status_table(valid, "Quien Creó", "Quién creó")
else:
    status_tbl = build_status_table(valid, "CLIENTE", "Cliente")

st.markdown(
    f'<div class="section-title">ESTATUS OPERATIVO POR {status_mode}</div>',
    unsafe_allow_html=True,
)

st.dataframe(
    status_tbl.style.format(
        {
            "CUMPLIDO": "{:,.0f}",
            "CUMPLIDO OPERATIVO": "{:,.0f}",
            "EN PROGRAMACION": "{:,.0f}",
            "EN TRANSITO": "{:,.0f}",
            "OTROS": "{:,.0f}",
            "Total general": "{:,.0f}",
            "Pendientes operativos": "{:,.0f}",
            "% Cumplido": "{:.2%}",
            "% Cumplido Operativo": "{:.2%}",
            "Pendientes %": "{:.2%}",
        }
    ),
    use_container_width=True,
    height=520,
)





# =========================================================
# OPERACIÓN: COORDINADOR, TIPOLOGÍA, PLACA, CONDUCTOR
# =========================================================
st.markdown("## 3. Operación y recursos")

r1, r2 = st.columns(2)
with r1:
    st.markdown('<div class="section-title">SERVICIOS POR COORDINADOR</div>', unsafe_allow_html=True)
    coord = top_table(df, "Quien Creó", "Servicios", 12)
    fig = go.Figure(
        go.Bar(
            x=coord["Servicios"],
            y=coord["Quien Creó"],
            orientation="h",
            marker_color=ROYAL_BLUE,
            text=coord["Servicios"].map(fmt_int),
            textposition="outside",
        )
    )
    fig.update_layout(
        height=420, margin=dict(l=10, r=40, t=20, b=20),
        yaxis=dict(autorange="reversed"),
        xaxis=dict(gridcolor=GRID, title="Servicios"),
        plot_bgcolor="#101C2C", paper_bgcolor="#07111F", font=dict(color=TEXT),
    )
    st.plotly_chart(fig, use_container_width=True, key="coord_prod")

with r2:
    st.markdown('<div class="section-title">SERVICIOS POR TIPOLOGÍA / TIPO DE VEHÍCULO</div>', unsafe_allow_html=True)
    tipo = top_table(df, "TIPOLOGIA", "Servicios", 15)
    fig = go.Figure(
        go.Bar(
            x=tipo["Servicios"],
            y=tipo["TIPOLOGIA"],
            orientation="h",
            marker_color=SECONDARY_BLUE,
            text=tipo["Servicios"].map(fmt_int),
            textposition="outside",
        )
    )
    fig.update_layout(
        height=420, margin=dict(l=10, r=40, t=20, b=20),
        yaxis=dict(autorange="reversed"),
        xaxis=dict(gridcolor=GRID, title="Servicios"),
        plot_bgcolor="#101C2C", paper_bgcolor="#07111F", font=dict(color=TEXT),
    )
    st.plotly_chart(fig, use_container_width=True, key="tipo_prod")

r3, r4 = st.columns(2)
with r3:
    st.markdown('<div class="section-title">TOP PLACAS POR SERVICIOS/PRODUCCIÓN</div>', unsafe_allow_html=True)
    placa = top_table(df, "PLACA", "Servicios", 15)
    st.dataframe(
        placa[["PLACA", "Servicios", "Facturacion", "Costos", "Margen", "Rentabilidad %"]]
        .rename(columns={
            "Facturacion": "Producción",
            "Costos": "Costo del Servicio",
        })
        .style.format({
            "Servicios": "{:,.0f}",
            "Producción": lambda x: f"${x:,.0f}",
            "Costo del Servicio": lambda x: f"${x:,.0f}",
            "Margen": lambda x: f"${x:,.0f}",
            "Rentabilidad %": "{:.2%}",
        }),
        use_container_width=True,
        hide_index=True,
        height=430,
    )

with r4:
    st.markdown('<div class="section-title">TOP CONDUCTORES POR SERVICIOS/PRODUCCIÓN</div>', unsafe_allow_html=True)
    cond = top_table(df, "CONDUCTOR_NOMBRE", "Servicios", 15)
    st.dataframe(
        cond[["CONDUCTOR_NOMBRE", "Servicios", "Facturacion", "Costos", "Margen"]]
        .rename(columns={
            "CONDUCTOR_NOMBRE": "Conductor",
            "Facturacion": "Producción",
            "Costos": "Costo del Servicio",
        })
        .style.format({
            "Servicios": "{:,.0f}",
            "Producción": lambda x: f"${x:,.0f}",
            "Costo del Servicio": lambda x: f"${x:,.0f}",
            "Margen": lambda x: f"${x:,.0f}",
        }),
        use_container_width=True,
        hide_index=True,
        height=430,
    )



# =========================================================
# SEGUIMIENTO DE PRODUCCIÓN PAGADA Y PENDIENTE
# =========================================================
st.markdown("## 4. Seguimiento de recaudo asociado a la producción")
st.caption(
    "Control basado en ESTADO FA. "
    "FACT DEFINITIVA = RECAUDADO, según la regla de negocio definida. "
    "Todo estado diferente de FACT DEFINITIVA se considera PENDIENTE POR RECAUDAR "
    "hasta quedar definitivo. Los resultados respetan todos los filtros globales."
)

estado_fa_n = (
    valid["ESTADO FA"]
    .astype("string")
    .fillna("")
    .str.strip()
    .str.upper()
)

mask_facturado = estado_fa_n.eq("FACT DEFINITIVA")
mask_pendiente_pago = ~mask_facturado

pago_produccion = valid["V.CLIENTE"].sum()
pago_facturado = valid.loc[mask_facturado, "V.CLIENTE"].sum()
pago_pendiente = valid.loc[mask_pendiente_pago, "V.CLIENTE"].sum()

pago_pct_facturado = safe_div(pago_facturado, pago_produccion)
pago_pct_pendiente = safe_div(pago_pendiente, pago_produccion)
pago_servicios_pendientes = int(mask_pendiente_pago.sum())

pg1, pg2, pg3, pg4, pg5, pg6 = st.columns(6)
with pg1:
    kpi_card(
        "Producción",
        fmt_money(pago_produccion),
        note="Suma de V.CLIENTE",
        show_delta=False,
    )
with pg2:
    kpi_card(
        "Recaudado",
        fmt_money(pago_facturado),
        note="ESTADO FA = FACT DEFINITIVA",
        show_delta=False,
    )
with pg3:
    kpi_card(
        "Pendiente por Recaudar",
        fmt_money(pago_pendiente),
        note="Estado diferente de FACT DEFINITIVA",
        show_delta=False,
    )
with pg4:
    kpi_card(
        "% Recaudo",
        fmt_pct(pago_pct_facturado),
        note="Recaudado / Producción",
        show_delta=False,
    )
with pg5:
    kpi_card(
        "% Pendiente",
        fmt_pct(pago_pct_pendiente),
        note="Pendiente / Producción",
        show_delta=False,
    )
with pg6:
    kpi_card(
        "Servicios Pendientes",
        fmt_int(pago_servicios_pendientes),
        note="Servicios aún no definitivos",
        show_delta=False,
    )

pago_base = valid.copy()
pago_base["ESTADO_FA_N"] = estado_fa_n

pago_cliente = (
    pago_base.groupby("CLIENTE", dropna=False)
    .agg(
        Servicios=("CLIENTE", "size"),
        Produccion=("V.CLIENTE", "sum"),
    )
    .reset_index()
)

facturado_por_cliente = (
    pago_base.loc[pago_base["ESTADO_FA_N"].eq("FACT DEFINITIVA")]
    .groupby("CLIENTE")["V.CLIENTE"]
    .sum()
)

pendiente_por_cliente = (
    pago_base.loc[~pago_base["ESTADO_FA_N"].eq("FACT DEFINITIVA")]
    .groupby("CLIENTE")["V.CLIENTE"]
    .sum()
)

pago_cliente["Recaudado"] = pago_cliente["CLIENTE"].map(facturado_por_cliente).fillna(0)
pago_cliente["Pendiente por Recaudar"] = pago_cliente["CLIENTE"].map(pendiente_por_cliente).fillna(0)

pago_cliente["% Recaudo"] = np.where(
    pago_cliente["Produccion"].ne(0),
    pago_cliente["Recaudado"] / pago_cliente["Produccion"],
    np.nan,
)

pago_cliente["% Pendiente"] = np.where(
    pago_cliente["Produccion"].ne(0),
    pago_cliente["Pendiente por Recaudar"] / pago_cliente["Produccion"],
    np.nan,
)

pago_cliente = pago_cliente.sort_values(
    ["Pendiente por Recaudar", "Produccion"],
    ascending=[False, False],
).reset_index(drop=True)

st.markdown(
    '<div class="section-title">PRODUCCIÓN · RECAUDADO · PENDIENTE POR RECAUDAR POR CLIENTE</div>',
    unsafe_allow_html=True,
)

pago_chart = pago_cliente.head(15).copy()

fig_pago = go.Figure()
fig_pago.add_trace(
    go.Bar(
        x=pago_chart["CLIENTE"],
        y=pago_chart["Produccion"],
        name="Producción",
        marker_color=ROYAL_BLUE,
        hovertemplate="<b>%{x}</b><br>Producción: $%{y:,.0f}<extra></extra>",
    )
)
fig_pago.add_trace(
    go.Bar(
        x=pago_chart["CLIENTE"],
        y=pago_chart["Recaudado"],
        name="Recaudado",
        marker_color=GREEN,
        hovertemplate="<b>%{x}</b><br>Recaudado: $%{y:,.0f}<extra></extra>",
    )
)
fig_pago.add_trace(
    go.Bar(
        x=pago_chart["CLIENTE"],
        y=pago_chart["Pendiente por Recaudar"],
        name="Pendiente por Recaudar",
        marker_color=AMBER,
        hovertemplate="<b>%{x}</b><br>Pendiente por recaudar: $%{y:,.0f}<extra></extra>",
    )
)

fig_pago.update_layout(
    # Barras agrupadas para comparar FACTURADO vs PENDIENTE por cliente,
    # igual al estilo visual del análisis económico del punto 6.
    barmode="group",
    height=430,
    margin=dict(l=20, r=20, t=40, b=140),
    plot_bgcolor="#101C2C",
    paper_bgcolor="#07111F",
    font=dict(color=TEXT),
    legend=dict(orientation="h", y=1.10, x=0),
    xaxis=dict(title="Cliente", tickangle=-45),
    yaxis=dict(title="COP", gridcolor=GRID),
    bargap=0.22,
    bargroupgap=0.06,
)
fig_pago.update_layout(hovermode="x unified")
st.plotly_chart(
    fig_pago,
    use_container_width=True,
    key="pago_cliente_chart",
)

pago_show = pago_cliente[
    [
        "CLIENTE",
        "Servicios",
        "Produccion",
        "Recaudado",
        "Pendiente por Recaudar",
        "% Recaudo",
        "% Pendiente",
    ]
].copy()

pago_show.columns = [
    "Cliente",
    "Servicios",
    "Producción",
    "Recaudado",
    "Pendiente por Recaudar",
    "% Recaudo",
    "% Pendiente",
]

st.dataframe(
    pago_show.style.format(
        {
            "Servicios": "{:,.0f}",
            "Producción": lambda x: "$" + f"{x:,.0f}".replace(",", "."),
            "Recaudado": lambda x: "$" + f"{x:,.0f}".replace(",", "."),
            "Pendiente por Recaudar": lambda x: "$" + f"{x:,.0f}".replace(",", "."),
            "% Recaudo": "{:.2%}",
            "% Pendiente": "{:.2%}",
        }
    ),
    use_container_width=True,
    hide_index=True,
    height=480,
)

pago_excel = BytesIO()
with pd.ExcelWriter(pago_excel, engine="openpyxl") as writer:
    pago_show.to_excel(
        writer,
        index=False,
        sheet_name="Facturado_Pendiente",
    )

    ws = writer.book["Facturado_Pendiente"]

    blue_fill = openpyxl.styles.PatternFill(
        fill_type="solid",
        fgColor="003B8E",
    )
    white_bold = openpyxl.styles.Font(
        color="FFFFFF",
        bold=True,
    )
    centered = openpyxl.styles.Alignment(
        horizontal="center",
        vertical="center",
    )

    for cell in ws[1]:
        cell.fill = blue_fill
        cell.font = white_bold
        cell.alignment = centered

    for col in ("C", "D", "E"):
        for cell in ws[col][1:]:
            cell.number_format = '$#,##0'

    for col in ("F", "G"):
        for cell in ws[col][1:]:
            cell.number_format = '0.00%'

    widths = {
        "A": 42,
        "B": 14,
        "C": 20,
        "D": 20,
        "E": 22,
        "F": 16,
        "G": 18,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

pago_excel.seek(0)

st.download_button(
    "⬇️ Descargar seguimiento de pagos en Excel",
    data=pago_excel,
    file_name="Seguimiento_Produccion_Pagada_y_Pendiente.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)



# =========================================================
# MATRIZ CLIENTE X PERIODO
# =========================================================
st.markdown("## 5. Matriz Cliente × Periodo")
metric = st.radio(
    "Indicador de la matriz",
    ["Servicios", "Producción", "Margen", "Rentabilidad"],
    horizontal=True,
)

m = valid.copy()
if grain == "Mensual":
    m["P_ORD"] = m["AÑO"] * 100 + m["MES_NUM"]
    m["P_LABEL"] = m.apply(lambda r: f"{MONTHS_ES.get(r['MES_NUM'], '')} {int(r['AÑO'])}", axis=1)
elif grain == "Bimestral":
    m["P_ORD"] = m["AÑO"] * 10 + m["BIMESTRE_NUM"].astype(int)
    m["P_LABEL"] = m.apply(lambda r: f"B{int(r['BIMESTRE_NUM'])} {int(r['AÑO'])}", axis=1)
elif grain == "Trimestral":
    m["P_ORD"] = m["AÑO"] * 10 + m["TRIMESTRE_NUM"].astype(int)
    m["P_LABEL"] = m.apply(lambda r: f"T{int(r['TRIMESTRE_NUM'])} {int(r['AÑO'])}", axis=1)
else:
    m["P_ORD"] = m["AÑO"] * 10 + m["SEMESTRE_NUM"].astype(int)
    m["P_LABEL"] = m.apply(lambda r: f"S{int(r['SEMESTRE_NUM'])} {int(r['AÑO'])}", axis=1)

period_order = (
    m[["P_ORD", "P_LABEL"]].drop_duplicates().sort_values("P_ORD")["P_LABEL"].tolist()
)

base = (
    m.groupby(["CLIENTE", "P_LABEL"], dropna=False)
    .agg(Servicios=("CLIENTE", "size"), Facturacion=("V.CLIENTE", "sum"), Costos=("V.CONDUCT", "sum"))
    .reset_index()
)
base["Margen"] = base["Facturacion"] - base["Costos"]
base["Rentabilidad"] = np.where(base["Facturacion"].ne(0), base["Margen"] / base["Facturacion"], np.nan)

value_col = {"Servicios": "Servicios", "Producción": "Facturacion", "Margen": "Margen", "Rentabilidad": "Rentabilidad"}[metric]
pivot = base.pivot(index="CLIENTE", columns="P_LABEL", values=value_col).reindex(columns=period_order)

# Ordena clientes por total de la métrica.
if metric == "Rentabilidad":
    sort_key = base.groupby("CLIENTE")["Margen"].sum() / base.groupby("CLIENTE")["Facturacion"].sum()
else:
    sort_key = base.groupby("CLIENTE")[value_col].sum()
pivot = pivot.reindex(sort_key.sort_values(ascending=False).index)

if metric in ["Producción", "Margen"]:
    styled = pivot.style.format(lambda x: "" if pd.isna(x) else f"${x:,.0f}")
elif metric == "Rentabilidad":
    styled = pivot.style.format("{:.2%}")
else:
    styled = pivot.style.format(lambda x: "" if pd.isna(x) else f"{x:,.0f}")

st.dataframe(styled, use_container_width=True, height=480)



# =========================================================
# ANÁLISIS ECONÓMICO POR PLACA / CONDUCTOR
# =========================================================
st.markdown("## 6. Análisis económico por placa y conductor")
st.caption(
    "Complemento del tablero actual. Permite identificar con qué clientes trabajó "
    "una placa o conductor, cuántos trayectos realizó y cuánto generó económicamente."
)

analysis_mode = st.radio(
    "Analizar recurso por",
    ["PLACA", "CONDUCTOR"],
    horizontal=True,
    key="economic_resource_mode",
)

resource_col = "PLACA" if analysis_mode == "PLACA" else "CONDUCTOR_NOMBRE"
resource_label = "Placa" if analysis_mode == "PLACA" else "Conductor"

resource_options = (
    valid[resource_col]
    .dropna()
    .astype(str)
    .loc[lambda s: (s.str.strip() != "") & (s != "<NA>")]
    .value_counts()
    .index
    .tolist()
)

selected_resources = st.multiselect(
    f"Seleccionar {resource_label}(s)",
    resource_options,
    default=[],
    key="economic_resource_selection",
    help=(
        f"Puedes seleccionar uno o varios {resource_label.lower()}s. "
        "Si lo dejas vacío, toma todos los recursos dentro de los filtros globales."
    ),
)

economic_df = valid.copy()
if selected_resources:
    economic_df = economic_df[
        economic_df[resource_col].astype(str).isin(selected_resources)
    ].copy()

if economic_df.empty:
    st.warning("No hay información económica para la selección realizada.")
else:
    # KPIs
    eco_trayectos = len(economic_df)
    eco_facturacion = economic_df["V.CLIENTE"].sum()
    eco_valor_tercero = economic_df["V.CONDUCT"].sum()
    eco_margen = eco_facturacion - eco_valor_tercero
    eco_rentabilidad = safe_div(eco_margen, eco_facturacion)

    # Participación de producción:
    # Numerador = producción de la placa/conductor seleccionado.
    # Denominador = producción total del contexto GLOBAL ya filtrado
    #               (fecha, cliente, estado, flota, etc.), sin aplicar
    #               la selección interna de recurso de esta sección.
    eco_facturacion_base = valid["V.CLIENTE"].sum()
    eco_participacion_facturacion = safe_div(
        eco_facturacion,
        eco_facturacion_base,
    )

    e1, e2, e3, e4, e5, e6 = st.columns(6)
    with e1:
        kpi_card("Trayectos (Servicios)", fmt_int(eco_trayectos), note="Cantidad de servicios")
    with e2:
        kpi_card("Producción", fmt_money(eco_facturacion), note="Producción del recurso seleccionado")
    with e3:
        kpi_card("Costo del Servicio", fmt_money(eco_valor_tercero), note="Costo del Servicio")
    with e4:
        kpi_card("Margen", fmt_money(eco_margen), note="Producción - Costo del Servicio")
    with e5:
        kpi_card(
            "Participación Producción %",
            fmt_pct(eco_participacion_facturacion),
            note="Producción recurso / Producción total filtrada",
            show_delta=False,
        )
    with e6:
        kpi_card("Rentabilidad", fmt_pct(eco_rentabilidad), note="Margen / Producción")

    # Resumen económico por cliente
    eco_client = (
        economic_df
        .groupby("CLIENTE", dropna=False)
        .agg(
            Trayectos=("CLIENTE", "size"),
            Facturacion=("V.CLIENTE", "sum"),
            Valor_Tercero=("V.CONDUCT", "sum"),
        )
        .reset_index()
    )
    eco_client["Margen"] = eco_client["Facturacion"] - eco_client["Valor_Tercero"]
    eco_client["Rentabilidad"] = np.where(
        eco_client["Facturacion"].ne(0),
        eco_client["Margen"] / eco_client["Facturacion"],
        np.nan,
    )
    eco_client["Participación Producción %"] = np.where(
        eco_facturacion != 0,
        eco_client["Facturacion"] / eco_facturacion,
        np.nan,
    )
    eco_client = eco_client.sort_values(
        ["Facturacion", "Trayectos"],
        ascending=[False, False],
    )

    # Gráfica principal
    eco_chart = eco_client.head(15).copy()

    st.markdown(
        f'<div class="section-title">'
        f'PRODUCCIÓN · COSTO DEL SERVICIO · MARGEN POR CLIENTE — {analysis_mode}'
        f'</div>',
        unsafe_allow_html=True,
    )

    fig_eco = go.Figure()
    fig_eco.add_trace(
        go.Bar(
            x=eco_chart["CLIENTE"],
            y=eco_chart["Facturacion"],
            name="Producción",
            marker_color=ROYAL_BLUE,
            customdata=np.stack(
                [eco_chart["Trayectos"], eco_chart["Rentabilidad"]],
                axis=-1,
            ),
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Producción: $%{y:,.0f}<br>"
                "Trayectos: %{customdata[0]:,.0f}<br>"
                "Rentabilidad: %{customdata[1]:.2%}"
                "<extra></extra>"
            ),
        )
    )
    fig_eco.add_trace(
        go.Bar(
            x=eco_chart["CLIENTE"],
            y=eco_chart["Valor_Tercero"],
            name="Costo del Servicio",
            marker_color=SECONDARY_BLUE,
            hovertemplate="<b>%{x}</b><br>Costo del Servicio: $%{y:,.0f}<extra></extra>",
        )
    )
    fig_eco.add_trace(
        go.Bar(
            x=eco_chart["CLIENTE"],
            y=eco_chart["Margen"],
            name="Margen",
            marker_color=DARK_BLUE,
            hovertemplate="<b>%{x}</b><br>Margen: $%{y:,.0f}<extra></extra>",
        )
    )
    fig_eco.update_layout(
        barmode="group",
        height=470,
        margin=dict(l=20, r=20, t=45, b=150),
        plot_bgcolor="#101C2C",
        paper_bgcolor="#07111F",
        font=dict(color=TEXT),
        legend=dict(orientation="h", y=1.10, x=0),
        xaxis=dict(
            title="Cliente",
            tickangle=-50,
            tickfont=dict(size=10, color=TEXT),
        ),
        yaxis=dict(title="COP", gridcolor=GRID, zeroline=False),
        hovermode="x unified",
    )
    st.plotly_chart(
        fig_eco,
        use_container_width=True,
        key="economic_resource_client_chart",
    )

    # Tabla resumen por cliente
    st.markdown(
        '<div class="section-title">RESUMEN ECONÓMICO POR CLIENTE</div>',
        unsafe_allow_html=True,
    )
    eco_show = eco_client[
        [
            "CLIENTE",
            "Trayectos",
            "Facturacion",
            "Participación Producción %",
            "Valor_Tercero",
            "Margen",
            "Rentabilidad",
        ]
    ].copy()
    eco_show.columns = [
        "Cliente",
        "Trayectos",
        "Producción",
        "Participación Producción %",
        "Costo del Servicio",
        "Margen",
        "Rentabilidad %",
    ]
    st.dataframe(
        eco_show.style.format(
            {
                "Trayectos": "{:,.0f}",
                "Producción": lambda x: f"${x:,.0f}",
                "Participación Producción %": "{:.2%}",
                "Costo del Servicio": lambda x: f"${x:,.0f}",
                "Margen": lambda x: f"${x:,.0f}",
                "Rentabilidad %": "{:.2%}",
            }
        ),
        use_container_width=True,
        hide_index=True,
        height=420,
    )

    # Trazabilidad Placa-Conductor-Cliente
    st.markdown(
        '<div class="section-title">DETALLE PLACA · CONDUCTOR · CLIENTE</div>',
        unsafe_allow_html=True,
    )
    eco_detail = (
        economic_df
        .groupby(["PLACA", "CONDUCTOR_NOMBRE", "CLIENTE"], dropna=False)
        .agg(
            Trayectos=("CLIENTE", "size"),
            Facturacion=("V.CLIENTE", "sum"),
            Valor_Tercero=("V.CONDUCT", "sum"),
        )
        .reset_index()
    )
    eco_detail["Margen"] = eco_detail["Facturacion"] - eco_detail["Valor_Tercero"]
    eco_detail["Rentabilidad"] = np.where(
        eco_detail["Facturacion"].ne(0),
        eco_detail["Margen"] / eco_detail["Facturacion"],
        np.nan,
    )
    eco_detail = eco_detail.sort_values(
        ["Facturacion", "Trayectos"],
        ascending=[False, False],
    )
    eco_detail.columns = [
        "Placa",
        "Conductor",
        "Cliente",
        "Trayectos",
        "Producción",
        "Costo del Servicio",
        "Margen",
        "Rentabilidad %",
    ]

    st.dataframe(
        eco_detail.style.format(
            {
                "Trayectos": "{:,.0f}",
                "Producción": lambda x: f"${x:,.0f}",
                "Costo del Servicio": lambda x: f"${x:,.0f}",
                "Margen": lambda x: f"${x:,.0f}",
                "Rentabilidad %": "{:.2%}",
            }
        ),
        use_container_width=True,
        hide_index=True,
        height=500,
    )




# =========================================================
# SERVICIOS Y PRODUCCIÓN POR PLACA Y CLIENTE
# =========================================================
st.markdown("## 7. Servicios por placa y cliente")
st.caption(
    "La placa y el cliente se normalizan antes del cálculo para evitar duplicados "
    "por diferencias de mayúsculas/minúsculas o espacios. "
    "SERVICIOS = cantidad de registros válidos. "
    "PRODUCCIÓN = suma de V.CLIENTE. "
    "PARTICIPACIÓN EN % = producción del cliente dentro del total facturado por esa placa."
)

# Consolidación única por PLACA + CLIENTE ya normalizados.
servicios_pc = (
    valid.assign(
        PLACA_N=valid["PLACA"].astype("string").str.strip().str.upper(),
        CLIENTE_N=valid["CLIENTE"].astype("string").str.strip().str.upper(),
    )
    .groupby(["PLACA_N", "CLIENTE_N"], dropna=False)
    .agg(
        SERVICIOS=("CLIENTE", "size"),
        PRODUCCIÓN=("V.CLIENTE", "sum"),
    )
    .reset_index()
    .rename(columns={"PLACA_N": "PLACA", "CLIENTE_N": "CLIENTE"})
)

# Total facturado de cada placa.
totales_placa = (
    servicios_pc.groupby("PLACA", as_index=False)["PRODUCCIÓN"]
    .sum()
    .rename(columns={"PRODUCCIÓN": "TOTAL_FACTURADO_PLACA"})
)

servicios_pc = servicios_pc.merge(totales_placa, on="PLACA", how="left")

servicios_pc["PARTICIPACIÓN EN %"] = np.where(
    servicios_pc["TOTAL_FACTURADO_PLACA"].ne(0),
    servicios_pc["PRODUCCIÓN"] / servicios_pc["TOTAL_FACTURADO_PLACA"],
    np.nan,
)

# Ordenar por placa y luego por producción descendente.
servicios_pc = servicios_pc.sort_values(
    ["PLACA", "PRODUCCIÓN", "CLIENTE"],
    ascending=[True, False, True],
).reset_index(drop=True)

tabla_pc = servicios_pc[
    ["PLACA", "CLIENTE", "SERVICIOS", "PRODUCCIÓN", "PARTICIPACIÓN EN %"]
].copy()

st.dataframe(
    tabla_pc.style.format(
        {
            "SERVICIOS": "{:,.0f}",
            "PRODUCCIÓN": lambda x: "$" + f"{x:,.0f}".replace(",", "."),
            "PARTICIPACIÓN EN %": "{:.2%}",
        }
    ),
    use_container_width=True,
    hide_index=True,
    height=520,
)

# Descarga Excel con la misma lógica de consolidación.
excel_buffer = BytesIO()

with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    tabla_pc.to_excel(
        writer,
        index=False,
        sheet_name="Servicios_Placa_Cliente",
    )

    ws = writer.book["Servicios_Placa_Cliente"]

    header_fill = openpyxl.styles.PatternFill(
        fill_type="solid",
        fgColor="003B8E",
    )
    header_font = openpyxl.styles.Font(
        color="FFFFFF",
        bold=True,
    )
    header_alignment = openpyxl.styles.Alignment(
        horizontal="center",
        vertical="center",
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    for cell in ws["C"][1:]:
        cell.number_format = '#,##0'
    for cell in ws["D"][1:]:
        cell.number_format = '$#,##0'
    for cell in ws["E"][1:]:
        cell.number_format = '0.00%'

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

excel_buffer.seek(0)

st.download_button(
    label="⬇️ Descargar tabla en Excel",
    data=excel_buffer,
    file_name="Servicios_Produccion_Participacion_por_Placa_Cliente.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


st.markdown(
    f"""
    <div style="background:{ROYAL_BLUE};color:white;padding:8px 12px;border-radius:8px;margin-top:12px;font-size:12px;">
        Los resultados se calculan sobre <b>CARGA</b> como fecha maestra.
        Pendiente Operativo = <b>EN TRANSITO + EN PROGRAMACION</b>.
        Tipo Flota: 1 = FLOTA PROPIA · 0 = TERCEROS.
        Conductor: columna <b>CONDUCTO.1</b> (columna W del Excel).
        Tipología: <b>T. VEHICULO</b>.
    </div>
    """,
    unsafe_allow_html=True,
)
