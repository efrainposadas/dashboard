# ============================================================
#  DASHBOARD DE INVENTARIO 2026 — Script único completo
#  TAB 0 : RESUMEN GENERAL
#  TAB 1 : 🧮 LANDED COST
#  TAB 2 : 🚨 BACKORDER
#  TAB 3 : 🔎 DETALLE
#  INSTALACIÓN: pip install streamlit duckdb pandas plotly openpyxl pyarrow
#  USO: streamlit run app.py
# ============================================================

import streamlit as st
import streamlit.components.v1 as components
import duckdb
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io, os
from datetime import datetime
from detalle import render_detalle_tab

EXCEL_FILE   = "DATOS.xlsx"
PARQUET_FILE = "inventario_cache.parquet"

# ── Conexión DuckDB persistente ────────────────────────────
@st.cache_resource
def get_duckdb_con():
    """Conexión DuckDB reutilizable — se crea una sola vez."""
    return duckdb.connect()

def make_con(df_long):
    """Registra df_long en DuckDB (solo cuando cambia) y devuelve la conexión."""
    con = get_duckdb_con()
    # Evitar re-registrar si es el mismo objeto (ahorra overhead en cada llamada)
    _id = id(df_long)
    if st.session_state.get('_duckdb_df_id') != _id:
        con.register("datos", df_long)
        st.session_state['_duckdb_df_id'] = _id
    return con

MONTHS    = ['March','April','May','June','July','August',
             'September','October','November','December']
MONTHS_ES = ['Marzo','Abril','Mayo','Junio','Julio','Agosto',
             'Septiembre','Octubre','Noviembre','Diciembre']
MES_MAP   = dict(zip(MONTHS, MONTHS_ES))

# Meses extendidos — incluye Enero y Febrero para ventas históricas
MONTHS_ALL    = ['January','February'] + MONTHS
MONTHS_ALL_ES = ['Enero','Febrero'] + MONTHS_ES
MES_MAP_ALL   = dict(zip(MONTHS_ALL, MONTHS_ALL_ES))

DIM_COLS = ['FAMILIA','LINEA','GRUPO','SUPPLIER','STATUS','ITEM','ANIO',
            'NOMBRE_CORTO','POTENCIAL','GRUPO_PRODUCCION','PAIS_ORIGEN',
            'CURRENCY','SUPPLIER_NUMBER','FULL_NAME','ABC','XYZ','ABC/XYZ']
NUM_COLS = ['IGI','COSTO','INVENTARIO','PRICE_ORIGINAL','PRICE_USD','LT_TOTAL_DIAS']

# ── Configuración de monedas ──────────────────────────────
MONEDA_CONFIG = {
    'USD': {'label': '💵 Dólares (USD)',  'simbolo': '$',  'tc_default': 1.0,     'editable': False},
    'MXN': {'label': '🟢 Pesos (MXN)',    'simbolo': '$',  'tc_default': 17.15,   'editable': True},
    'EUR': {'label': '💶 Euros (EUR)',     'simbolo': '€',  'tc_default': 0.92,    'editable': True},
}


SECTIONS = [
    {
        "titulo":   "INVENTARIO (Piezas)",
        "color":    "#1F4E79",
        "alt":      "#D6E4F0",
        "es_fin":   False,
        "metricas": [
            ("Beginning Inventory Position (Inv Inicio Mes)", "Beginning Inventory Position"),
            ("Final Inventory Position (Inv Proyectado)",     "Final Inventory Position"),
            ("On Hand",                                      "On Hand"),
            ("On Order",                                     "On Order"),
            ("Projected Available Balance (PAB)",            "Projected Available Balance"),
            ("Safety Stock",                                 "Safety Stock"),
            # ── CAMBIO 1: Máximo y Mínimo ahora se muestran en MESES ──
            ("Máximo (Meses)",                               "__MAX_MOS__"),
            ("Mínimo (Meses)",                               "__MIN_MOS__"),
            ("Safety Stock (Meses)",                         "__SS_MOS__"),
        ],
    },
    {
        "titulo":   "COBERTURA & ROTACIÓN",
        "color":    "#1565A0",
        "alt":      "#D6E4F0",
        "es_fin":   False,
        "metricas": [
            ("Meses de Inventario (MOS)",    "__MOS__"),
            ("Vueltas de Inventario (prom)", "__VUELTAS__"),
        ],
    },
    {
        "titulo":   "DEMANDA & FORECAST (Piezas)",
        "color":    "#1A5C38",
        "alt":      "#D5F5E3",
        "es_fin":   False,
        "metricas": [
            ("Gross Forecast",                 "Gross Forecast"),
            ("Net Forecast",                   "Net Forecast"),
            ("Total Demand",                   "Total Demand"),
            ("Sales Orders",                   "Sales Orders"),
            ("Net Forecast $",                 "__NET_FC_USD__"),
        ],
    },
    {
        "titulo":   "REAPROVISIONAMIENTO (Piezas)",
        "color":    "#7B2C2C",
        "alt":      "#FADBD8",
        "es_fin":   False,
        "metricas": [
            ("OC Tránsito",                    "Purchase Orders"),
            ("Orden de Compra",                "Purchase Requisitions"),
            ("Planned Replen. by Order Date",  "Planned Replenishments by Order Date"),
            ("Planned Replen. by Due Date",    "Planned Replenishments by Due Date"),
            ("Cobertura de Reposición (Meses)", "__COB_REPOS__"),
        ],
    },
    {
        "titulo":   "COSTO DE PEDIDO",
        "color":    "#1A4A6B",
        "alt":      "#D6EAF8",
        "es_fin":   True,
        "metricas": [
            ("OC Tránsito",      "OC_USD"),
            ("Orden de Compra",  "REQUISICION_USD"),
            ("Pedido Sugerido",  "IMPORTE_PEDIDO_USD"),
            ("TOTAL COSTO PEDIDO", "__TOTAL_COSTO_PED__"),
        ],
    },
    {
        "titulo":   "FINANCIERO",
        "color":    "#3D3D3D",
        "alt":      "#F2F3F4",
        "es_fin":   True,
        "metricas": [
            ("Pedido Sugerido (CIF)",              "IMPORTE_PEDIDO_USD"),
            ("Costos Logísticos / Importación",    "__COSTOS_LOG__"),
            ("TOTAL LANDED COST",                  "__TOTAL_FIN__"),
        ],
    },
]

FIN_MEASURES = ["COSTO_TOTAL_USD","OC_USD","IMPORTE_PEDIDO_USD","REQUISICION_USD","COSTO_IGI",
                "VENTAS_IMP","PRESUPUESTO_IMP","NET_FC_USD"]

# ════════════════════════════════════════════════════════════
#  LANDED COST — Configuración y datos
# ════════════════════════════════════════════════════════════

LC_DEFAULT_PAISES = {
    "ARGENTINA": {
        "IGI": 7.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.0, "Transporte (MX)": 1.6,
        "Custodia": 0.25, "Seguro de Carga": 0.35,
        "Maniobras/Descarga": 0.5, "Almacenaje": 0.3,
    },
    "BRASIL": {
        "IGI": 9.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.0, "Transporte (MX)": 1.8,
        "Custodia": 0.25, "Seguro de Carga": 0.4,
        "Maniobras/Descarga": 0.5, "Almacenaje": 0.3,
    },
    "CHINA": {
        "IGI": 8.9, "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.2, "Transporte (MX)": 2.5,
        "Custodia": 0.3, "Seguro de Carga": 0.5,
        "Maniobras/Descarga": 0.6, "Almacenaje": 0.4,
    },
    "INDIA": {
        "IGI": 10.0, "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.2, "Transporte (MX)": 2.8,
        "Custodia": 0.3, "Seguro de Carga": 0.55,
        "Maniobras/Descarga": 0.6, "Almacenaje": 0.4,
    },
    "MALASIA": {
        "IGI": 8.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.2, "Transporte (MX)": 2.7,
        "Custodia": 0.3, "Seguro de Carga": 0.5,
        "Maniobras/Descarga": 0.6, "Almacenaje": 0.4,
    },
    "MEXICO": {
        "IGI": 0.0,  "DTA": 0.0, "Prevalidación": 0.05,
        "Honorarios AA": 0.6, "Transporte (MX)": 0.8,
        "Custodia": 0.1, "Seguro de Carga": 0.15,
        "Maniobras/Descarga": 0.3, "Almacenaje": 0.15,
    },
    "PERU": {
        "IGI": 7.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.0, "Transporte (MX)": 1.7,
        "Custodia": 0.25, "Seguro de Carga": 0.35,
        "Maniobras/Descarga": 0.5, "Almacenaje": 0.3,
    },
    "TURQUIA": {
        "IGI": 8.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.1, "Transporte (MX)": 3.0,
        "Custodia": 0.35, "Seguro de Carga": 0.55,
        "Maniobras/Descarga": 0.6, "Almacenaje": 0.4,
    },
    "USA": {
        "IGI": 5.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 0.9, "Transporte (MX)": 1.5,
        "Custodia": 0.2, "Seguro de Carga": 0.25,
        "Maniobras/Descarga": 0.4, "Almacenaje": 0.2,
    },
    "ALEMANIA": {
        "IGI": 7.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.1, "Transporte (MX)": 3.2,
        "Custodia": 0.35, "Seguro de Carga": 0.6,
        "Maniobras/Descarga": 0.6, "Almacenaje": 0.45,
    },
    "ITALIA": {
        "IGI": 7.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.1, "Transporte (MX)": 3.2,
        "Custodia": 0.35, "Seguro de Carga": 0.6,
        "Maniobras/Descarga": 0.6, "Almacenaje": 0.45,
    },
    "TAIWAN": {
        "IGI": 8.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.2, "Transporte (MX)": 2.6,
        "Custodia": 0.3, "Seguro de Carga": 0.5,
        "Maniobras/Descarga": 0.6, "Almacenaje": 0.4,
    },
    "COREA": {
        "IGI": 8.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.2, "Transporte (MX)": 2.6,
        "Custodia": 0.3, "Seguro de Carga": 0.5,
        "Maniobras/Descarga": 0.6, "Almacenaje": 0.4,
    },
    "JAPÓN": {
        "IGI": 6.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.1, "Transporte (MX)": 2.7,
        "Custodia": 0.3, "Seguro de Carga": 0.5,
        "Maniobras/Descarga": 0.6, "Almacenaje": 0.4,
    },
    "OTRO": {
        "IGI": 8.0,  "DTA": 0.8, "Prevalidación": 0.10,
        "Honorarios AA": 1.1, "Transporte (MX)": 2.5,
        "Custodia": 0.3, "Seguro de Carga": 0.5,
        "Maniobras/Descarga": 0.55, "Almacenaje": 0.35,
    },
}

LC_CONCEPTOS = list(list(LC_DEFAULT_PAISES.values())[0].keys())

LC_CONCEPTOS_COLOR = {
    "IGI":                "#1565C0",
    "DTA":                "#1976D2",
    "Prevalidación":      "#42A5F5",
    "Honorarios AA":      "#F57F17",
    "Transporte (MX)":    "#2E7D32",
    "Custodia":           "#C62828",
    "Seguro de Carga":    "#6A1B9A",
    "Maniobras/Descarga": "#00695C",
    "Almacenaje":         "#4E342E",
}

# ════════════════════════════════════════════════════════════
#  CSS GLOBAL
# ════════════════════════════════════════════════════════════
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
html,body,[class*="css"]{ font-family:'IBM Plex Sans',sans-serif; }

/* ── Eliminar espacio de iframes de components.html ── */
iframe[title="streamlit_components_v1.html"] {
  display:block !important;
  margin:0 !important;
  padding:0 !important;
}
.element-container:has(iframe) {
  margin-bottom: 0 !important;
  padding-bottom: 0 !important;
}
.stHtml, [data-testid="stHtml"] {
  margin: 0 !important;
  padding: 0 !important;
}

/* ── Tarjeta de moneda activa ── */
.moneda-activa{
  display:inline-flex; align-items:center; gap:.5rem;
  background:linear-gradient(135deg,#1A237E,#0D47A1);
  color:#fff; border-radius:10px; padding:.45rem 1.1rem;
  font-size:.8rem; font-weight:700; letter-spacing:.4px;
  box-shadow:0 3px 10px rgba(26,35,126,.35); margin-bottom:.8rem;
}
.moneda-activa .tc{ font-size:.68rem; opacity:.75; font-family:'IBM Plex Mono',monospace; }

/* ── Tarjetas KPI Premium ── */
.kpi-card{
  background:#fff; border-radius:16px; padding:1.1rem 1.3rem 1rem;
  box-shadow:0 2px 12px rgba(0,0,0,.07); position:relative;
  overflow:hidden; min-height:110px;
  border-top:4px solid #ccc;
}

/* ══ STICKY KPI BAR (aparece bajo los tabs cuando haces scroll) ══ */
#kpi-float {
  display: none;
  position: fixed;
  top: 0;
  left: 0; right: 0;
  z-index: 99999;
  background: linear-gradient(90deg,#040e1f 0%,#071830 60%,#040e1f 100%);
  border-bottom: 2px solid rgba(59,130,246,.4);
  box-shadow: 0 6px 28px rgba(0,0,0,.75);
  padding: .4rem 1rem .35rem 1rem;
}
#kpi-float.show { display: flex; gap:.5rem; align-items: stretch; }
#kpi-float .kpi-card {
  flex: 1;
  min-height: unset !important;
  padding: .5rem .75rem .45rem !important;
  border-radius: 10px !important;
  background: rgba(255,255,255,.07) !important;
  box-shadow: none !important;
  border-top-width: 3px !important;
}
#kpi-float .kpi-icon    { font-size:.85rem !important; margin-bottom:.05rem !important; }
#kpi-float .kpi-title   { font-size:.44rem !important; color:#94a3b8 !important; margin-bottom:.05rem !important; }
#kpi-float .kpi-currency{ font-size:.44rem !important; margin-bottom:.03rem !important; }
#kpi-float .kpi-value   { font-size:.92rem !important; font-weight:800 !important; line-height:1.1 !important; margin-bottom:.1rem !important; }
#kpi-float .kpi-pzas    { font-size:.54rem !important; color:#94a3b8 !important; }
#kpi-float .kpi-pzas span { color:#cbd5e1 !important; }
#kpi-float .kpi-sub-items { font-size:.5rem !important; margin-top:.1rem !important; }

/* ══════════════════════════════════════════════════════
   KPI STICKY BAR — fijo justo encima de los tabs
══════════════════════════════════════════════════════ */
.kpi-sticky-container {
  position: sticky;
  top: 0;
  z-index: 9999;
  background: linear-gradient(90deg,#040e1f 0%,#071830 55%,#040e1f 100%);
  border-top: 1px solid rgba(59,130,246,.2);
  border-bottom: 2px solid rgba(59,130,246,.4);
  box-shadow: 0 6px 28px rgba(0,0,0,.8);
  margin: 0 -1rem;
  padding: .4rem 1rem .35rem;
}
.kpi-sticky-inner {
  display: flex;
  gap: .5rem;
  align-items: stretch;
}
.kpi-s-card {
  flex: 1;
  display: flex;
  align-items: center;
  gap: .45rem;
  background: rgba(255,255,255,.055);
  border-radius: 10px;
  border-top: 3px solid rgba(59,130,246,.35);
  padding: .4rem .6rem;
  min-width: 0;
}
.kpi-s-icon { font-size: 1.1rem; flex-shrink: 0; }
.kpi-s-body { min-width: 0; }
.kpi-s-title {
  font-size: .47rem; font-weight: 800; letter-spacing: .7px;
  text-transform: uppercase; color: #94a3b8; white-space: nowrap;
  margin-bottom: .08rem;
}
.kpi-s-val {
  font-size: .95rem; font-weight: 800; line-height: 1.1;
  font-family: 'IBM Plex Mono', monospace;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.kpi-s-sub {
  font-size: .52rem; color: #64748b; white-space: nowrap;
  overflow: hidden; text-overflow: ellipsis;
}
.kpi-icon{ font-size:1.3rem; margin-bottom:.3rem; display:block; opacity:.85; }
.kpi-title{
  font-size:.6rem; font-weight:700; letter-spacing:.8px;
  text-transform:uppercase; color:#aaa; margin-bottom:.15rem;
}
.kpi-currency{
  font-size:.6rem; font-weight:600; letter-spacing:.5px;
  text-transform:uppercase; margin-bottom:.1rem;
}
.kpi-value{
  font-size:1.45rem; font-weight:800; line-height:1.15;
  font-family:'IBM Plex Mono',monospace; margin-bottom:.25rem;
}
.kpi-pzas{ font-size:.72rem; color:#888; font-weight:400; }
.kpi-pzas span{ font-weight:600; color:#555; }
.kpi-sub-items{ font-size:.68rem; font-weight:600; margin-top:.3rem; }

.kpi-inv   { border-top-color:#1A6FA8; }
.kpi-inv   .kpi-currency{ color:#1A6FA8; }
.kpi-inv   .kpi-value   { color:#1A6FA8; }
.kpi-trans { border-top-color:#2E7D53; }
.kpi-trans .kpi-currency{ color:#2E7D53; }
.kpi-trans .kpi-value   { color:#2E7D53; }
.kpi-items { border-top-color:#C0392B; }
.kpi-items .kpi-value   { color:#C0392B; }
.kpi-pedido{ border-top-color:#1A6FA8; }
.kpi-pedido .kpi-currency{ color:#1A6FA8; }
.kpi-pedido .kpi-value  { color:#1A6FA8; }
.kpi-cob   { border-top-color:#7B2FBE; }
.kpi-cob   .kpi-value   { color:#7B2FBE; font-size:2rem; }

.res-wrap{
  overflow-x:auto; border-radius:12px;
  box-shadow:0 4px 28px rgba(13,27,75,.15);
  font-family:'IBM Plex Sans',sans-serif; margin-bottom:.5rem;
}
.res-tbl{ border-collapse:collapse; width:100%; min-width:1160px; background:#fff; }
.res-tbl thead tr{ background:linear-gradient(180deg,#1A237E,#162082); }
.res-tbl thead th{
  color:#fff; font-weight:600; font-size:.71rem;
  letter-spacing:.3px; padding:10px 6px; text-align:center;
  border:none; white-space:nowrap;
}
.res-tbl thead th.th-tipo{ text-align:left; padding-left:32px; min-width:270px; }
.res-tbl thead th.th-total{ background:#c8960a; color:#2d1e00; min-width:108px; }
.sec-hdr td{
  color:#fff; font-weight:700; font-size:.78rem;
  letter-spacing:.5px; text-transform:uppercase;
  padding:9px 14px 9px 18px; white-space:nowrap;
}
.sec-hdr .bar{ width:6px; padding:0; }
.dr td{
  font-size:.77rem; padding:6px 8px;
  border-bottom:1px solid #e8eaf2;
  white-space:nowrap; vertical-align:middle;
  transition:filter .1s;
}
.dr:hover td{ filter:brightness(.96); }
.dr .bar{ width:6px; padding:0; }
.dr .lbl{ padding:6px 10px 6px 22px; text-align:left; font-size:.77rem; }
.dr .num{ text-align:right; padding:6px 10px; font-family:'IBM Plex Mono',monospace; font-size:.72rem; }
.row-tot td{ background:#FFF9C4 !important; color:#5D4037 !important; font-weight:700 !important; }
.col-tot{
  background:#FFF9C4 !important; color:#5D4037 !important;
  font-weight:700 !important; border-left:2px solid #c8960a !important;
  padding-right:14px !important;
}
.sep td{ height:5px; background:#E8EAF6; padding:0; border:none; }

/* ── Landed Cost ── */
.lc-wrap{
  overflow-x:auto; border-radius:12px;
  box-shadow:0 4px 24px rgba(13,27,75,.13); margin-bottom:1rem;
}
.lc-tbl{
  border-collapse:collapse; width:100%; min-width:900px;
  background:#fff; font-family:'IBM Plex Sans',sans-serif;
}
.lc-tbl thead tr{ background:linear-gradient(180deg,#1A237E,#162082); }
.lc-tbl thead th{
  color:#fff; font-weight:700; font-size:.68rem;
  letter-spacing:.4px; padding:10px 10px;
  text-align:center; border:none; white-space:nowrap;
}
.lc-tbl thead th.th-left{ text-align:left; padding-left:18px; min-width:200px; }
.lc-tbl thead th.th-pais{ background:#0D47A1; }
.lc-tbl thead th.th-tot { background:#c8960a; color:#2d1e00; }
.lc-tbl tbody tr:nth-child(even) td{ background:#F5F7FF; }
.lc-tbl tbody tr:hover td{ background:#EBF3FF !important; }
.lc-tbl td{
  font-size:.75rem; padding:7px 10px;
  border-bottom:1px solid #e8eaf2; white-space:nowrap;
  vertical-align:middle; color:#1a1a2e;
}
.lc-tbl td.td-left{ text-align:left; padding-left:18px; font-weight:600; }
.lc-tbl td.td-mxn{
  text-align:right; padding-right:14px;
  font-family:'IBM Plex Mono',monospace; font-size:.73rem; color:#1A5276;
}
.lc-tbl td.td-tot{
  text-align:right; padding-right:14px;
  font-family:'IBM Plex Mono',monospace; font-size:.78rem; font-weight:800;
  background:#FFF9C4 !important; color:#5D4037 !important;
  border-left:2px solid #c8960a;
}
.lc-tbl tr.row-lc-tot td{
  background:#E8F5E9 !important; color:#1B5E20 !important;
  font-weight:800 !important; font-size:.78rem;
  border-top:2px solid #388E3C;
}
.lc-tbl tr.row-lc-tot td.td-tot{
  background:#C8E6C9 !important; color:#1B5E20 !important;
}
.lc-kpi{
  background:#fff; border-radius:12px; padding:.8rem 1rem;
  box-shadow:0 2px 10px rgba(0,0,0,.07); border-top:3px solid #ccc;
}
.lc-kpi.blue { border-top-color:#1565C0; }
.lc-kpi.green{ border-top-color:#2E7D32; }
.lc-kpi.amber{ border-top-color:#F57F17; }
.lc-kpi.red  { border-top-color:#C62828; }
.lc-kpi-lbl{ font-size:.58rem; font-weight:700; letter-spacing:.7px; text-transform:uppercase; color:#aaa; }
.lc-kpi-val{ font-size:1.2rem; font-weight:800; font-family:'IBM Plex Mono',monospace; color:#1a1a2e; margin-top:.15rem; }
.lc-kpi-sub{ font-size:.62rem; color:#888; margin-top:.1rem; }

.stDownloadButton>button{
  background:#1A237E !important; color:#fff !important;
  border:none !important; border-radius:8px !important; font-weight:600 !important;
}

/* ══════════════════════════════════════════════
   TAB BACKORDER
   ══════════════════════════════════════════════ */
.bo-kpi-wrap{ display:flex; gap:.8rem; flex-wrap:wrap; margin-bottom:1rem; }
.bo-kpi{
  flex:1; min-width:140px; background:#fff; border-radius:14px;
  padding:.85rem 1.1rem; box-shadow:0 2px 12px rgba(0,0,0,.07);
  border-top:4px solid #ccc;
}
.bo-kpi.red   { border-top-color:#C0392B; }
.bo-kpi.amber { border-top-color:#E67E22; }
.bo-kpi.purple{ border-top-color:#6A1B9A; }
.bo-kpi-lbl{ font-size:.57rem; font-weight:700; letter-spacing:.8px; text-transform:uppercase; color:#aaa; margin-bottom:.2rem; }
.bo-kpi-val{ font-size:1.35rem; font-weight:800; font-family:'IBM Plex Mono',monospace; color:#1a1a2e; }
.bo-kpi-sub{ font-size:.62rem; color:#888; margin-top:.1rem; }
.bo-wrap{
  overflow-x:auto; border-radius:12px;
  box-shadow:0 4px 24px rgba(13,27,75,.13); margin-bottom:1rem;
}
.bo-tbl{
  border-collapse:collapse; width:100%; min-width:1100px;
  background:#fff; font-family:'IBM Plex Sans',sans-serif;
}
.bo-tbl thead tr{ background:linear-gradient(180deg,#1A237E,#162082); }
.bo-tbl thead th{
  color:#fff; font-weight:700; font-size:.68rem;
  letter-spacing:.4px; padding:10px 8px;
  text-align:center; border:none; white-space:nowrap;
}
.bo-tbl thead th.th-left{ text-align:left; padding-left:14px; min-width:130px; }
.bo-tbl thead th.th-semaforo{ background:#7B2C2C; min-width:90px; }
.bo-tbl thead th.th-gap{ background:#c8960a; color:#2d1e00; }
.bo-tbl tbody tr:nth-child(even) td{ background:#F8F9FF; }
.bo-tbl tbody tr:hover td{ background:#EBF3FF !important; }
.bo-tbl td{
  font-size:.74rem; padding:7px 8px;
  border-bottom:1px solid #E8EAF2; white-space:nowrap;
  vertical-align:middle; color:#1a1a2e;
}
.bo-tbl td.td-left { text-align:left; padding-left:14px; font-weight:600; font-size:.73rem; }
.bo-tbl td.td-num  { text-align:right; padding-right:12px; font-family:'IBM Plex Mono',monospace; font-size:.72rem; }
.bo-tbl td.td-gap  {
  text-align:right; padding-right:12px;
  font-family:'IBM Plex Mono',monospace; font-size:.75rem; font-weight:800;
  background:#FFF9C4 !important; color:#5D4037 !important;
  border-left:2px solid #c8960a;
}
.bo-tbl td.td-gap.gap-ok { background:#E8F5E9 !important; color:#1B5E20 !important; }
.bo-tbl td.td-gap.gap-bad{ background:#FDEDEC !important; color:#C0392B !important; }
.bo-sem{
  display:inline-block; border-radius:6px;
  padding:2px 8px; font-size:.66rem; font-weight:700;
  letter-spacing:.3px; white-space:nowrap;
}
.sem-verde   { background:#D5F5E3; color:#1A5C38; }
.sem-amarillo{ background:#FEF9E7; color:#9A6700; }
.sem-rojo    { background:#FADBD8; color:#922B21; }
.bo-tbl tr.bo-row-tot td{
  background:#FFF9C4 !important; color:#5D4037 !important;
  font-weight:800 !important; border-top:2px solid #c8960a;
}

/* ══════════════════════════════════════════════
   CUADRO EJECUTIVO — Resumen por dimensión
   ══════════════════════════════════════════════ */
.exec-wrap{
  overflow-x:auto; border-radius:14px;
  box-shadow:0 4px 24px rgba(13,27,75,.13);
  margin-top:.6rem; margin-bottom:.4rem;
}
.exec-tbl{
  border-collapse:collapse; width:100%;
  min-width:780px; background:#fff;
  font-family:'IBM Plex Sans',sans-serif;
}
.exec-tbl thead tr.exec-th-main{
  background:linear-gradient(180deg,#1A237E,#162082);
}
.exec-tbl thead tr.exec-th-main th{
  color:#fff; font-weight:700; font-size:.68rem;
  letter-spacing:.4px; padding:9px 10px;
  text-align:center; border:none; white-space:nowrap;
}
.exec-tbl thead tr.exec-th-main th.exec-th-dim{
  text-align:left; padding-left:18px; min-width:160px;
}
.exec-tbl thead tr.exec-th-sub th{
  font-size:.6rem; font-weight:700; letter-spacing:.3px;
  padding:5px 8px; text-align:right; white-space:nowrap;
  border-bottom:2px solid #E0E6F8;
}
.exec-tbl thead tr.exec-th-sub th.sub-inv { background:#D6E4F0; color:#1F4E79; }
.exec-tbl thead tr.exec-th-sub th.sub-tra { background:#D5F5E3; color:#1A5C38; }
.exec-tbl thead tr.exec-th-sub th.sub-ped { background:#FADBD8; color:#7B2C2C; }
.exec-tbl thead tr.exec-th-sub th.sub-dim { background:#ECEFF1; color:#455A64; text-align:left; padding-left:18px; }
.exec-tbl tbody tr:nth-child(even) td{ background:#F8F9FF; }
.exec-tbl tbody tr:hover td{ background:#EBF3FF !important; }
.exec-tbl td{
  font-size:.75rem; padding:7px 10px;
  border-bottom:1px solid #E8EAF2; white-space:nowrap;
  font-family:'IBM Plex Mono',monospace; color:#1a1a2e;
  text-align:right;
}
.exec-tbl td.exec-td-dim{
  text-align:left; padding-left:18px;
  font-family:'IBM Plex Sans',sans-serif;
  font-weight:700; font-size:.76rem; color:#1A237E;
}
.exec-tbl td.exec-td-inv { color:#1F4E79; }
.exec-tbl td.exec-td-tra { color:#1A5C38; }
.exec-tbl td.exec-td-ped { color:#7B2C2C; }
.exec-tbl td.exec-td-zero{ color:#C8D8E0; }
.exec-tbl tr.exec-row-tot td{
  background:#FFF9C4 !important; color:#5D4037 !important;
  font-weight:800 !important; border-top:2px solid #c8960a;
}
.exec-tbl tr.exec-row-tot td.exec-td-dim{
  color:#5D4037 !important;
}
.exec-tbl td.exec-sep, .exec-tbl th.exec-sep{
  border-left:2px solid #E0E6F8 !important;
}
        /* ── Sidebar: opciones de moneda y filtros bien visibles ── */
        section[data-testid="stSidebar"] .stRadio label,
        section[data-testid="stSidebar"] .stRadio label p,
        section[data-testid="stSidebar"] .stRadio label span,
        section[data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"] p {
            color: #E8F4FF !important;
            font-size: .75rem !important;
            font-weight: 700 !important;
            opacity: 1 !important;
        }
        section[data-testid="stSidebar"] .stRadio [data-baseweb="radio"] ~ div,
        section[data-testid="stSidebar"] .stRadio [data-baseweb="block"] > div {
            color: #E8F4FF !important;
        }
        section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="menu"] li,
        section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="menu"] span {
            color: #1a1a2e !important;
        }
        section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] span {
            color: #D0E8FF !important;
            font-weight: 600 !important;
        }
        section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] input::placeholder {
            color: #7EB8DA !important;
        }
        section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] input {
            color: #D0E8FF !important;
        }
        section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] span {
            color: #fff !important;
            font-weight: 700 !important;
        }
</style>
"""

# ════════════════════════════════════════════════════════════
#  CARGA DE DATOS — desde Excel
# ════════════════════════════════════════════════════════════
# Mapeo número de mes → nombre en inglés
_MES_NUM_EN = {1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',7:'July',
               8:'August',9:'September',10:'October',11:'November',12:'December'}

# Measures derivadas que DEBEN existir en el parquet — si falta alguna, se regenera
_REQUIRED_MEASURES = {"NET_FC_USD", "IMPORTE_PEDIDO_USD", "OC_USD", "REQUISICION_USD", "COSTO_TOTAL_USD"}

@st.cache_resource(show_spinner="⏳ Cargando datos desde Excel…")
def cargar_datos():
    # Si ya existe parquet, verificar que tenga todas las measures requeridas
    # Y que el parquet sea más reciente que el Excel (para forzar regeneración si cambió)
    if os.path.exists(PARQUET_FILE):
        try:
            # Invalidar caché si el Excel es más reciente que el parquet
            if os.path.exists(EXCEL_FILE):
                excel_mtime  = os.path.getmtime(EXCEL_FILE)
                parquet_mtime = os.path.getmtime(PARQUET_FILE)
                if excel_mtime > parquet_mtime:
                    os.remove(PARQUET_FILE)
                    raise FileNotFoundError("Excel más reciente que parquet — regenerando")
            df_check = pd.read_parquet(PARQUET_FILE)
            measures_en_cache = set(df_check['MEASURE'].unique()) if 'MEASURE' in df_check.columns else set()
            if _REQUIRED_MEASURES.issubset(measures_en_cache):
                # Leer hoja BO aunque venga del parquet
                _df_bo_par = pd.DataFrame()
                try:
                    _raw = pd.read_excel(EXCEL_FILE, sheet_name='BO')
                    _raw.columns = [str(c).strip().upper() for c in _raw.columns]
                    if {'CC','BO'}.issubset(_raw.columns):
                        _extra = [c for c in ['NOMBRE CORTO','COSTO TOTAL','COMENTARIO','COSTO'] if c in _raw.columns]
                        _df_bo_par = _raw[['CC','BO'] + _extra].copy()
                        _df_bo_par['CC'] = _df_bo_par['CC'].astype(str).str.strip()
                        _df_bo_par['BO'] = pd.to_numeric(_df_bo_par['BO'], errors='coerce').fillna(0)
                        for _c, _d in [('COSTO',0.0),('NOMBRE CORTO',''),('COSTO TOTAL',0.0),('COMENTARIO','')]:
                            if _c not in _df_bo_par.columns: _df_bo_par[_c] = _d
                        _df_bo_par['COSTO']        = pd.to_numeric(_df_bo_par['COSTO'],       errors='coerce').fillna(0)
                        _df_bo_par['COSTO TOTAL']  = pd.to_numeric(_df_bo_par['COSTO TOTAL'], errors='coerce').fillna(0)
                        _df_bo_par['NOMBRE CORTO'] = _df_bo_par['NOMBRE CORTO'].astype(str).str.strip().replace('nan','')
                        _df_bo_par['COMENTARIO']   = _df_bo_par['COMENTARIO'].astype(str).str.strip().replace('nan','')
                        _df_bo_par = _df_bo_par[_df_bo_par['BO'] > 0].reset_index(drop=True)
                except Exception:
                    _df_bo_par = pd.DataFrame()
                _df_prom_par = pd.DataFrame(columns=['CC','PROM_PZA','PROM_IMP'])
                try:
                    _dv2 = pd.read_excel(EXCEL_FILE, sheet_name='VENTAS')
                    _dv2['CC'] = _dv2['CC'].astype(str).str.strip()
                    _dv2r = _dv2[_dv2['TIPO'].astype(str).str.upper().str.strip().str.startswith('VENTA')].copy()
                    _dv2r['PIEZAS']   = pd.to_numeric(_dv2r['PIEZAS'],   errors='coerce').fillna(0)
                    _dv2r['IMPORTES'] = pd.to_numeric(_dv2r['IMPORTES'], errors='coerce').fillna(0)
                    _pp2 = _dv2r[_dv2r['PIEZAS']   > 0].groupby('CC')['PIEZAS'].mean().round(0)
                    _pi2 = _dv2r[_dv2r['IMPORTES'] > 0].groupby('CC')['IMPORTES'].mean().round(2)
                    _df_prom_par = pd.DataFrame({'PROM_PZA': _pp2, 'PROM_IMP': _pi2}).reset_index()
                except Exception:
                    pass
                # ── SIEMPRE actualizar COSTO desde ARTICULOS (parquet puede tener costo viejo) ──
                try:
                    _art_fresh = pd.read_excel(EXCEL_FILE, sheet_name='ARTICULOS')
                    _art_fresh['CC'] = _art_fresh['CC'].astype(str).str.strip()
                    _art_fresh.rename(columns={'CC':'ITEM'}, inplace=True)
                    _costo_map = _art_fresh.set_index('ITEM')['COSTO'].to_dict()
                    df_check['COSTO'] = df_check['ITEM'].astype(str).str.strip().map(_costo_map)
                except Exception:
                    pass
                try:
                    _dv_p = pd.read_excel(EXCEL_FILE, sheet_name='VENTAS')
                    _dv_p['CC'] = _dv_p['CC'].astype(str).str.strip()
                    _dv_p['PIEZAS']   = pd.to_numeric(_dv_p['PIEZAS'],   errors='coerce').fillna(0)
                    _dv_p['IMPORTES'] = pd.to_numeric(_dv_p['IMPORTES'], errors='coerce').fillna(0)
                    _dv_p = _dv_p[_dv_p['TIPO'].astype(str).str.upper().str.strip().str.startswith('VENTA')]
                    _df_vm_par = _dv_p[['CC','AÑO','MES','PIEZAS','IMPORTES']].copy()
                except Exception:
                    _df_vm_par = pd.DataFrame(columns=['CC','AÑO','MES','PIEZAS','IMPORTES'])
                return df_check, _df_bo_par, _df_prom_par, _df_vm_par, "parquet"
            else:
                # Faltan measures — borrar parquet y regenerar desde Excel
                os.remove(PARQUET_FILE)
        except Exception:
            if os.path.exists(PARQUET_FILE):
                os.remove(PARQUET_FILE)

    if not os.path.exists(EXCEL_FILE):
        st.error(f"❌ No se encontró el archivo {EXCEL_FILE} — colócalo en la misma carpeta que app.py")
        st.stop()

    try:
        # ── 1. PEDIDO ────────────────────────────────────────
        df_ped = pd.read_excel(EXCEL_FILE, sheet_name='PEDIDO')
        df_ped['CC'] = df_ped['CC'].astype(str).str.strip()
        # Renombrar columnas datetime → nombre mes inglés
        col_map = {c: _MES_NUM_EN[c.month] for c in df_ped.columns
                   if hasattr(c, 'month') and c.month in _MES_NUM_EN}
        df_ped.rename(columns=col_map, inplace=True)
        df_ped.rename(columns={'CC':'ITEM','Measure':'MEASURE'}, inplace=True)
        df_ped['ANIO'] = 2026

        # ── 2. ARTICULOS ─────────────────────────────────────
        df_art = pd.read_excel(EXCEL_FILE, sheet_name='ARTICULOS')
        df_art['CC'] = df_art['CC'].astype(str).str.strip()
        df_art.rename(columns={'CC':'ITEM'}, inplace=True)

        # ── 3. ACUERDOS ──────────────────────────────────────
        df_ac = pd.read_excel(EXCEL_FILE, sheet_name='ACUERDOS')
        df_ac['CC'] = df_ac['CC'].astype(str).str.strip()
        df_ac.rename(columns={'CC':'ITEM','PRICE':'PRICE_ORIGINAL'}, inplace=True)

        # ── Convertir PRICE_ORIGINAL a USD según CURRENCY ────
        # TC fijos de referencia para la conversión en carga
        # (el usuario puede editar en sidebar, pero aquí necesitamos un TC base para normalizar)
        _TC_MXN_REF = 17.6367   # 1 USD = X MXN  (referencia para normalizar)
        _TC_EUR_REF = 20.5505   # 1 EUR = X MXN → 1 EUR = X/TC_MXN USD

        # Vectorizado: sin apply fila por fila
        precio = pd.to_numeric(df_ac['PRICE_ORIGINAL'], errors='coerce').fillna(0.0)
        cur    = df_ac['CURRENCY'].astype(str).str.strip().str.upper() if 'CURRENCY' in df_ac.columns else 'USD'
        df_ac['PRICE_USD'] = precio  # default USD
        df_ac.loc[cur == 'MXN',         'PRICE_USD'] = precio[cur == 'MXN'] / _TC_MXN_REF
        df_ac.loc[cur.isin(['EUR','EURO']), 'PRICE_USD'] = precio[cur.isin(['EUR','EURO'])] * (_TC_EUR_REF / _TC_MXN_REF)

        # ── Si un ítem tiene varios proveedores → escoger el precio USD más bajo ──
        df_ac = (df_ac
                 .sort_values('PRICE_USD', ascending=True)          # menor precio primero
                 .drop_duplicates(subset='ITEM', keep='first')       # quedar con el más barato
                 .reset_index(drop=True))

        # ── 4. ABC ───────────────────────────────────────────
        df_abc = pd.read_excel(EXCEL_FILE, sheet_name='ABC')
        df_abc['CC'] = df_abc['CC'].astype(str).str.strip()
        df_abc.rename(columns={'CC':'ITEM'}, inplace=True)

        # ── 5. JOIN principal ────────────────────────────────
        df = df_ped.merge(df_art, on='ITEM', how='left')
        df = df.merge(
            df_ac[['ITEM','SUPPLIER_NUMBER','SUPPLIER','FULL_NAME',
                   'CURRENCY','PRICE_ORIGINAL','PRICE_USD','LT_TOTAL_DIAS','PAIS_ORIGEN']],
            on='ITEM', how='left')
        df = df.merge(df_abc[['ITEM','ABC','XYZ','ABC/XYZ']], on='ITEM', how='left')

        # ── 6. VENTAS y PRESUPUESTO → measures separadas ────
        df_v = pd.read_excel(EXCEL_FILE, sheet_name='VENTAS')
        df_v['CC']     = df_v['CC'].astype(str).str.strip()
        df_v['MES_EN'] = df_v['MES'].map(_MES_NUM_EN)
        df_v = df_v[df_v['MES_EN'].notna()].copy()

        # ── Ventas mensuales para Backorder (ANTES del rename CC→ITEM) ──
        _dv_vm = df_v[df_v['TIPO'].astype(str).str.upper().str.strip().str.startswith('VENTA')].copy()
        _dv_vm['PIEZAS']   = pd.to_numeric(_dv_vm['PIEZAS'],   errors='coerce').fillna(0)
        _dv_vm['IMPORTES'] = pd.to_numeric(_dv_vm['IMPORTES'], errors='coerce').fillna(0)
        df_ventas_mes = _dv_vm[['CC','AÑO','MES','PIEZAS','IMPORTES']].copy()

        df_v.rename(columns={'CC':'ITEM','AÑO':'ANIO'}, inplace=True)

        # ── Promedios ventas por CC (TIPO=VENTA*, valores > 0) ──
        _dv_real = df_v[df_v['TIPO'].astype(str).str.upper().str.strip().str.startswith('VENTA')].copy()
        _dv_real['PIEZAS']   = pd.to_numeric(_dv_real['PIEZAS'],   errors='coerce').fillna(0)
        _dv_real['IMPORTES'] = pd.to_numeric(_dv_real['IMPORTES'], errors='coerce').fillna(0)
        _pp = _dv_real[_dv_real['PIEZAS']   > 0].groupby('ITEM')['PIEZAS'].mean().round(0)
        _pi = _dv_real[_dv_real['IMPORTES'] > 0].groupby('ITEM')['IMPORTES'].mean().round(2)
        df_prom_ventas = pd.DataFrame({'PROM_PZA': _pp, 'PROM_IMP': _pi}).reset_index()
        df_prom_ventas.rename(columns={'ITEM': 'CC'}, inplace=True)
        df_prom_ventas['CC'] = df_prom_ventas['CC'].astype(str).str.strip()

        # Vectorizado: pivot en vez de triple loop + iterrows
        def _ventas_pivot(df_tipo, mea_pza, mea_imp):
            grp_cols = ['ITEM','ANIO','MES_EN']
            pza = df_tipo.pivot_table(index=['ITEM','ANIO'], columns='MES_EN', values='PIEZAS',
                                      aggfunc='sum').reset_index()
            imp = df_tipo.pivot_table(index=['ITEM','ANIO'], columns='MES_EN', values='IMPORTES',
                                      aggfunc='sum').reset_index()
            pza.columns.name = None; imp.columns.name = None
            pza['MEASURE'] = mea_pza; imp['MEASURE'] = mea_imp
            return pd.concat([pza, imp], ignore_index=True)

        partes_ventas = []
        for tipo, df_tipo in df_v.groupby('TIPO'):
            if str(tipo).upper() == 'PRESUPUESTO':
                mea_pza, mea_imp = 'PRESUPUESTO_PZA', 'PRESUPUESTO_IMP'
            else:
                mea_pza, mea_imp = 'VENTAS_PZA', 'VENTAS_IMP'
            partes_ventas.append(_ventas_pivot(df_tipo, mea_pza, mea_imp))

        df_ventas = pd.concat(partes_ventas, ignore_index=True) if partes_ventas else pd.DataFrame()
        df_ventas = df_ventas.merge(df_art, on='ITEM', how='left')
        df_ventas = df_ventas.merge(
            df_ac[['ITEM','SUPPLIER_NUMBER','SUPPLIER','FULL_NAME',
                   'CURRENCY','PRICE_ORIGINAL','PRICE_USD','LT_TOTAL_DIAS','PAIS_ORIGEN']],
            on='ITEM', how='left')
        df_ventas = df_ventas.merge(df_abc[['ITEM','ABC','XYZ','ABC/XYZ']], on='ITEM', how='left')

        # ── 7. Concatenar todo ───────────────────────────────
        df_final = pd.concat([df, df_ventas], ignore_index=True)

        # ── 8. Calcular measures financieras derivadas ───────
        # Estas medidas las calculaba el SQL antes; ahora las calculamos aquí
        # Base: tomar filas de Planned Replenishments, Purchase Orders, Purchase Requisitions
        df_base = df_final.copy()
        # Precio USD por ítem (lookup rápido)
        price_map = df_ac.set_index('ITEM')['PRICE_USD'].to_dict()
        costo_map = df_art.set_index('ITEM')['COSTO'].to_dict()
        igi_map   = df_art.set_index('ITEM')['IGI'].fillna(0).to_dict()

        import numpy as np

        def _calc_measure_vec(source_measure, new_measure, multiplier_map, factor_extra=None, anio_filter=None):
            """Calcula medida derivada 100% vectorizado — sin iterrows."""
            mask = df_final['MEASURE'] == source_measure
            if anio_filter is not None:
                mask = mask & (df_final['ANIO'] == anio_filter)
            src = df_final.loc[mask].copy()
            if src.empty: return None

            # Vector de multiplicadores por ítem
            mult_vec = src['ITEM'].map(multiplier_map).fillna(0).astype(float).values

            # Matriz de valores mensuales (filas=ítems, cols=meses)
            month_cols = [m for m in MONTHS if m in src.columns]
            val_mat = src[month_cols].to_numpy(dtype=float)
            np.nan_to_num(val_mat, copy=False)

            # Multiplicar cada fila por su factor
            result_mat = val_mat * mult_vec[:, None]

            if factor_extra is not None:
                extra_vec = src['ITEM'].map(factor_extra).fillna(0).astype(float).values / 100.0
                result_mat = result_mat * extra_vec[:, None]

            src = src.copy()
            src[month_cols] = result_mat
            src['MEASURE'] = new_measure
            return src

        derived = []
        d = _calc_measure_vec('Planned Replenishments by Order Date', 'IMPORTE_PEDIDO_USD', price_map)
        if d is not None: derived.append(d)
        d = _calc_measure_vec('Purchase Orders',        'OC_USD',          price_map)
        if d is not None: derived.append(d)
        d = _calc_measure_vec('Purchase Requisitions',  'REQUISICION_USD', price_map)
        if d is not None: derived.append(d)
        d = _calc_measure_vec('Planned Replenishments by Order Date', 'COSTO_IGI',       price_map, igi_map)
        if d is not None: derived.append(d)
        d = _calc_measure_vec('Planned Replenishments by Order Date', 'COSTO_TOTAL_USD', costo_map)
        if d is not None: derived.append(d)
        d = _calc_measure_vec('Net Forecast', 'NET_FC_USD', price_map, anio_filter=2026)
        if d is not None: derived.append(d)

        if derived:
            df_final = pd.concat([df_final] + derived, ignore_index=True)

        # ── 8b. MANUAL — OC Tránsito dinero adicional ───────
        # Lee hoja MANUAL (CC, Price, MONEDA, CANTIDAD, IMPORTE, MES)
        # MES acepta inglés (March) o español (Marzo) — sin MES usa Marzo por default
        # Solo suma IMPORTE a OC_USD (dinero); NO toca Purchase Orders (piezas)
        try:
            df_manual = pd.read_excel(EXCEL_FILE, sheet_name='MANUAL')
            df_manual.columns = [str(c).strip().upper() for c in df_manual.columns]
            if {'CC', 'IMPORTE'}.issubset(df_manual.columns):
                df_manual['CC'] = df_manual['CC'].astype(str).str.strip()
                df_manual['IMPORTE'] = pd.to_numeric(df_manual['IMPORTE'], errors='coerce').fillna(0)
                # Ignorar filas de comentario (CC empieza con #) o importe cero
                df_manual = df_manual[~df_manual['CC'].str.startswith('#') & (df_manual['IMPORTE'] > 0)].copy()

                _TC_MXN_REF_M = 17.6367
                _TC_EUR_REF_M = 20.5505
                if 'MONEDA' in df_manual.columns:
                    cur_m = df_manual['MONEDA'].astype(str).str.strip().str.upper()
                    imp   = df_manual['IMPORTE'].copy()
                    df_manual['IMPORTE_USD'] = imp
                    df_manual.loc[cur_m == 'MXN', 'IMPORTE_USD'] = imp[cur_m == 'MXN'] / _TC_MXN_REF_M
                    df_manual.loc[cur_m.isin(['EUR','EURO']), 'IMPORTE_USD'] = imp[cur_m.isin(['EUR','EURO'])] * (_TC_EUR_REF_M / _TC_MXN_REF_M)
                else:
                    df_manual['IMPORTE_USD'] = df_manual['IMPORTE']

                # Normalizar columna MES: acepta inglés o español
                _MES_ES_EN = {
                    'ENERO':'January','FEBRERO':'February','MARZO':'March',
                    'ABRIL':'April','MAYO':'May','JUNIO':'June',
                    'JULIO':'July','AGOSTO':'August','SEPTIEMBRE':'September',
                    'OCTUBRE':'October','NOVIEMBRE':'November','DICIEMBRE':'December',
                }
                _MONTHS_UPPER = {m.upper(): m for m in MONTHS_ALL}

                def _norm_mes(val):
                    s = str(val).strip().upper()
                    if s in _MONTHS_UPPER: return _MONTHS_UPPER[s]   # inglés
                    if s in _MES_ES_EN:    return _MES_ES_EN[s]       # español
                    return 'March'  # fallback

                df_manual['MES_EN'] = df_manual['MES'].apply(_norm_mes) if 'MES' in df_manual.columns else 'March'

                # Una fila wide por CC+MES (el importe va solo en su columna de mes)
                rows_manual = []
                for _, mrow in df_manual.iterrows():
                    cc      = mrow['CC']
                    imp_usd = float(mrow['IMPORTE_USD'])
                    mes_en  = mrow['MES_EN']
                    new_row = {'ITEM': cc, 'ANIO': 2026, 'MEASURE': 'OC_USD'}
                    for m in MONTHS_ALL:
                        new_row[m] = imp_usd if m == mes_en else 0.0
                    art_row = df_art[df_art['ITEM'] == cc]
                    if not art_row.empty:
                        for col in ['FAMILIA','LINEA','GRUPO','STATUS','NOMBRE_CORTO',
                                    'POTENCIAL','GRUPO_PRODUCCION','IGI','COSTO','INVENTARIO']:
                            if col in art_row.columns:
                                new_row[col] = art_row.iloc[0][col]
                    rows_manual.append(new_row)

                if rows_manual:
                    df_manual_rows = pd.DataFrame(rows_manual)
                    df_final = pd.concat([df_final, df_manual_rows], ignore_index=True)
        except Exception:
            pass  # Si no existe la hoja MANUAL, continúa sin error

        # ── 9. BACKORDER — hoja BO ──────────────────────────
        df_bo = pd.DataFrame()
        try:
            df_bo_raw = pd.read_excel(EXCEL_FILE, sheet_name='BO')
            df_bo_raw.columns = [str(c).strip().upper() for c in df_bo_raw.columns]
            if {'CC','BO'}.issubset(df_bo_raw.columns):
                extra_cols = [c for c in ['NOMBRE CORTO','COSTO TOTAL','COMENTARIO','COSTO'] if c in df_bo_raw.columns]
                base_cols  = ['CC','BO'] + extra_cols
                df_bo = df_bo_raw[base_cols].copy()
                df_bo['CC'] = df_bo['CC'].astype(str).str.strip()
                df_bo['BO'] = pd.to_numeric(df_bo['BO'], errors='coerce').fillna(0)
                if 'COSTO'        not in df_bo.columns: df_bo['COSTO']        = 0.0
                if 'NOMBRE CORTO' not in df_bo.columns: df_bo['NOMBRE CORTO'] = ''
                if 'COSTO TOTAL'  not in df_bo.columns: df_bo['COSTO TOTAL']  = 0.0
                if 'COMENTARIO'   not in df_bo.columns: df_bo['COMENTARIO']   = ''
                df_bo['COSTO']        = pd.to_numeric(df_bo['COSTO'],        errors='coerce').fillna(0)
                df_bo['COSTO TOTAL']  = pd.to_numeric(df_bo['COSTO TOTAL'],  errors='coerce').fillna(0)
                df_bo['NOMBRE CORTO'] = df_bo['NOMBRE CORTO'].astype(str).str.strip().replace('nan','')
                df_bo['COMENTARIO']   = df_bo['COMENTARIO'].astype(str).str.strip().replace('nan','')
                df_bo = df_bo[df_bo['BO'] > 0].reset_index(drop=True)
        except Exception:
            df_bo = pd.DataFrame()

        # ── NORMALIZAR COSTO desde ARTICULOS antes de guardar ──────────────
        # Garantiza que el parquet siempre tenga el COSTO correcto del Excel
        _costo_norm = df_art.set_index('ITEM')['COSTO'].to_dict()
        df_final['COSTO'] = df_final['ITEM'].astype(str).str.strip().map(_costo_norm)

        # Guardar parquet para próximas cargas
        df_final.to_parquet(PARQUET_FILE, index=False)
        return df_final, df_bo, df_prom_ventas, df_ventas_mes, "excel"

    except Exception as e:
        st.error(f"❌ Error al leer el Excel: {e}")
        st.stop()


@st.cache_data(show_spinner=False)
def calcular_mos(df):
    """Calcula MOS y Vueltas de forma 100% vectorizada — sin iterrows."""
    import numpy as np

    def get_m(name):
        s = df.loc[df['MEASURE'] == name, ['ITEM','ANIO'] + MONTHS]
        return s.reset_index(drop=True) if not s.empty else None

    df_pab = get_m('Projected Available Balance')
    if df_pab is None: return df

    df_fc  = get_m('Net Forecast')
    df_dem = get_m('Total Demand')
    if df_fc is None:
        df_fc = df_dem
    elif df_dem is not None:
        df_fc = df_fc.merge(df_dem, on=['ITEM','ANIO'], how='left', suffixes=('','_td'))
        for m in MONTHS:
            td = f'{m}_td'
            if td in df_fc.columns:
                df_fc[m] = df_fc[m].where(df_fc[m].fillna(0) != 0, df_fc[td])
        df_fc.drop(columns=[c for c in df_fc.columns if c.endswith('_td')], inplace=True)
    if df_fc is None: return df

    # ── Vectorizado: matrices NumPy de PAB y FC ──────────────
    base = df_pab[['ITEM','ANIO']].copy()
    pab_mat = df_pab[MONTHS].to_numpy(dtype=float, na_value=0.0)

    fc_aligned = df_fc.set_index(['ITEM','ANIO']).reindex(
        pd.MultiIndex.from_frame(base)
    )[MONTHS].to_numpy(dtype=float, na_value=0.0)
    fc_aligned = np.nan_to_num(fc_aligned, nan=0.0)

    with np.errstate(divide='ignore', invalid='ignore'):
        mos_mat  = np.where(fc_aligned > 0, np.round(pab_mat / fc_aligned, 2), np.nan)
        vuel_mat = np.where((~np.isnan(mos_mat)) & (mos_mat > 0),
                            np.round(12.0 / mos_mat, 2), np.nan)

    # ── Enriquecer con columnas de dimensión ─────────────────
    icols = [c for c in ['STATUS','NOMBRE_CORTO','FAMILIA','LINEA','GRUPO','POTENCIAL',
             'GRUPO_PRODUCCION','IGI','COSTO','INVENTARIO','SUPPLIER_NUMBER',
             'SUPPLIER','FULL_NAME','CURRENCY','PRICE_ORIGINAL','PRICE_USD',
             'LT_TOTAL_DIAS','PAIS_ORIGEN'] if c in df.columns]
    info_df = (df.drop_duplicates(subset=['ITEM','ANIO'])
                 .set_index(['ITEM','ANIO'])[icols])

    def _build(mat, measure_name):
        out = base.copy()
        out['MEASURE'] = measure_name
        for i, m in enumerate(MONTHS):
            out[m] = mat[:, i]
        # join dimensional info
        out = out.join(info_df, on=['ITEM','ANIO'], how='left')
        return out

    df_mos  = _build(mos_mat,  'Meses de Inventario (MOS)')
    df_vuel = _build(vuel_mat, 'Vueltas de Inventario (prom)')

    return pd.concat([df, df_mos, df_vuel], ignore_index=True)


@st.cache_data(show_spinner=False)
def a_largo(df):
    import numpy as np
    info_base = [c for c in DIM_COLS + NUM_COLS if c in df.columns]
    id_vars   = list(dict.fromkeys(['ITEM','ANIO','MEASURE'] + info_base))
    mcols     = [m for m in MONTHS_ALL if m in df.columns]

    # Convertir columnas mes a numérico antes del melt (evita conversión fila a fila después)
    df_m = df.copy()
    for m in mcols:
        df_m[m] = pd.to_numeric(df_m[m], errors='coerce').fillna(0)

    dl = df_m.melt(id_vars=id_vars, value_vars=mcols, var_name='MES_EN', value_name='VALOR')

    # Mapeos vectorizados
    _MES_MAP_ALL = MES_MAP_ALL
    _EN_TO_NUM   = {'January':1,'February':2,'March':3,'April':4,'May':5,'June':6,
                    'July':7,'August':8,'September':9,'October':10,'November':11,'December':12}
    _MES_NUM_IDX = {m: i+1 for i, m in enumerate(MONTHS_ALL)}

    dl['MES']      = dl['MES_EN'].map(_MES_MAP_ALL)
    dl['MES_NUM']  = dl['MES_EN'].map(_MES_NUM_IDX).fillna(0).astype(int)
    dl['MES_REAL'] = dl['MES_EN'].map(_EN_TO_NUM).fillna(0).astype(int)
    dl.drop(columns=['MES_EN'], inplace=True)

    # Eliminar filas con VALOR=0 que no sean dimensionales importantes
    # (reduce tamaño de df_long drásticamente → queries DuckDB más rápidas)
    _measures_keep_zeros = {'On Hand','Projected Available Balance','Safety Stock',
                            'Beginning Inventory Position','Final Inventory Position'}
    mask_keep = dl['MEASURE'].isin(_measures_keep_zeros) | (dl['VALOR'] != 0)
    dl = dl.loc[mask_keep].reset_index(drop=True)

    return dl


# ════════════════════════════════════════════════════════════
#  CAMBIO 2: calcular_mos_resumen — ahora incluye MAX y MIN MOS
# ════════════════════════════════════════════════════════════
def calcular_mos_resumen(pivot):
    pab_key = 'Projected Available Balance'
    fc_keys = ['Net Forecast', 'Total Demand', 'Gross Forecast']
    pab_row = pivot.loc[pab_key] if pab_key in pivot.index else None
    fc_row  = next((pivot.loc[k] for k in fc_keys if k in pivot.index), None)

    # ── Filas de Máximo y Mínimo (piezas → convertir a meses) ─
    max_row = pivot.loc['Final Maximum Quantity'] if 'Final Maximum Quantity' in pivot.index else None
    min_row = pivot.loc['Final Minimum Quantity'] if 'Final Minimum Quantity' in pivot.index else None

    mos_vals, vuel_vals, max_mos_vals, min_mos_vals = [], [], [], []

    for m in MONTHS:
        pab = float(pab_row[m]) if (pab_row is not None and pd.notna(pab_row.get(m))) else None
        fc  = float(fc_row[m])  if (fc_row  is not None and pd.notna(fc_row.get(m)))  else None

        if pab is not None and fc and fc > 0:
            mos  = round(pab / fc, 2)
            vuel = round(12 / mos, 2) if mos > 0 else None
        else:
            mos = vuel = None

        # Máximo en meses
        max_v   = float(max_row[m]) if (max_row is not None and pd.notna(max_row.get(m))) else None
        max_mos = round(max_v / fc, 2) if (max_v is not None and fc and fc > 0) else None

        # Mínimo en meses
        min_v   = float(min_row[m]) if (min_row is not None and pd.notna(min_row.get(m))) else None
        min_mos = round(min_v / fc, 2) if (min_v is not None and fc and fc > 0) else None

        mos_vals.append(mos)
        vuel_vals.append(vuel)
        max_mos_vals.append(max_mos)
        min_mos_vals.append(min_mos)

    # Totales anuales
    pab_tot = float(pab_row['TOTAL']) if (pab_row is not None and pd.notna(pab_row.get('TOTAL'))) else None
    fc_tot  = float(fc_row['TOTAL'])  if (fc_row  is not None and pd.notna(fc_row.get('TOTAL')))  else None

    if pab_tot and fc_tot and fc_tot > 0:
        mos_tot  = round(pab_tot / fc_tot, 2)
        vuel_tot = round(12 / mos_tot, 2) if mos_tot > 0 else None
    else:
        mos_tot = vuel_tot = None

    max_tot_v   = float(max_row['TOTAL']) if (max_row is not None and pd.notna(max_row.get('TOTAL'))) else None
    min_tot_v   = float(min_row['TOTAL']) if (min_row is not None and pd.notna(min_row.get('TOTAL'))) else None
    max_mos_tot = round(max_tot_v / fc_tot, 2) if (max_tot_v and fc_tot and fc_tot > 0) else None
    min_mos_tot = round(min_tot_v / fc_tot, 2) if (min_tot_v and fc_tot and fc_tot > 0) else None

    # ── Cobertura de Reposición = Planned Replen. by Order Date ÷ Forecast ──
    rpd_row = pivot.loc['Planned Replenishments by Order Date']               if 'Planned Replenishments by Order Date' in pivot.index else None

    cob_repos_vals = []
    for m in MONTHS:
        rpd_v = float(rpd_row[m]) if (rpd_row is not None and pd.notna(rpd_row.get(m))) else None
        fc_v  = float(fc_row[m])  if (fc_row  is not None and pd.notna(fc_row.get(m)))  else None
        cob   = round(rpd_v / fc_v, 2) if (rpd_v is not None and fc_v and fc_v > 0) else None
        cob_repos_vals.append(cob)

    rpd_tot_v   = float(rpd_row['TOTAL']) if (rpd_row is not None and pd.notna(rpd_row.get('TOTAL'))) else None
    cob_repos_tot = round(rpd_tot_v / fc_tot, 2) if (rpd_tot_v and fc_tot and fc_tot > 0) else None

    return {
        '__MOS__':        mos_vals       + [mos_tot],
        '__VUELTAS__':    vuel_vals      + [vuel_tot],
        '__MAX_MOS__':    max_mos_vals   + [max_mos_tot],
        '__MIN_MOS__':    min_mos_vals   + [min_mos_tot],
        '__COB_REPOS__':  cob_repos_vals + [cob_repos_tot],
    }


def calcular_landed_cost_mensual(df_long, filtros_key, filtros):
    """Calcula Landed Cost por mes: COSTO_TOTAL_USD por país × (1 + % costos aduanales)."""
    def _mapear_pais(p):
        p_up = str(p).upper().strip()
        if p_up in LC_DEFAULT_PAISES: return p_up
        _alias = {
            "TURKEY":"TURQUIA","MALAYSIA":"MALASIA","MALAYSIAN":"MALASIA",
            "BRAZIL":"BRASIL","MEX":"MEXICO","ESTADOS UNIDOS MEXICANOS":"MEXICO",
            "GERMANY":"ALEMANIA","DEUTSCHLAND":"ALEMANIA","ITALY":"ITALIA",
            "JAPON":"JAPON","KOREA":"COREA","SOUTH KOREA":"COREA",
            "UNITED STATES":"USA","EUA":"USA","ESTADOS UNIDOS":"USA",
        }
        if p_up in _alias: return _alias[p_up]
        for key in LC_DEFAULT_PAISES:
            if key in p_up or p_up in key: return key
        return "OTRO"

    _m_mask = df_long["MEASURE"] == "COSTO_TOTAL_USD"
    if not _m_mask.any():
        _m_mask = df_long["MEASURE"] == "IMPORTE_PEDIDO_USD"
    _f_mask = _m_mask.copy()
    for col, vals in filtros.items():
        if col in df_long.columns and vals:
            _f_mask &= df_long[col].isin(vals)
    df_base = df_long.loc[_f_mask]

    if df_base.empty:
        return {m: 0.0 for m in MONTHS + ["TOTAL"]}

    lc_config = st.session_state.get(
        "lc_config",
        {p: dict(LC_DEFAULT_PAISES[p]) for p in LC_DEFAULT_PAISES}
    )
    mes_map_inv = dict(zip(MONTHS_ES, MONTHS))

    if "PAIS_ORIGEN" in df_base.columns:
        pais_keys  = df_base["PAIS_ORIGEN"].map(_mapear_pais)
        extra_pcts = pais_keys.map(
            lambda pk: sum(lc_config.get(pk, LC_DEFAULT_PAISES.get(pk, LC_DEFAULT_PAISES["OTRO"])).values()) / 100.0
        )
        factor_arr = (1 + extra_pcts).values
    else:
        cfg_otro   = lc_config.get("OTRO", LC_DEFAULT_PAISES["OTRO"])
        factor_arr = 1 + sum(cfg_otro.values()) / 100.0

    mes_en_arr = df_base["MES"].map(mes_map_inv)
    val_arr    = pd.to_numeric(df_base["VALOR"], errors="coerce").fillna(0) * factor_arr
    valid      = mes_en_arr.notna()
    tmp = pd.Series(val_arr, index=df_base.index).loc[valid].groupby(mes_en_arr[valid]).sum().to_dict()
    result = {m: tmp.get(m, 0.0) for m in MONTHS}
    result["TOTAL"] = sum(result[m] for m in MONTHS)
    return result


def calcular_costos_logisticos_mensual(df_long, filtros_key, filtros):
    """Calcula costos logísticos adicionales sobre IMPORTE_PEDIDO_USD."""
    def _mapear_pais(p):
        p_up = str(p).upper().strip()
        if p_up in LC_DEFAULT_PAISES: return p_up
        _alias = {
            "TURKEY":"TURQUIA","MALAYSIA":"MALASIA","MALAYSIAN":"MALASIA",
            "BRAZIL":"BRASIL","MEX":"MEXICO","ESTADOS UNIDOS MEXICANOS":"MEXICO",
            "GERMANY":"ALEMANIA","DEUTSCHLAND":"ALEMANIA","ITALY":"ITALIA",
            "JAPON":"JAPON","KOREA":"COREA","SOUTH KOREA":"COREA",
            "UNITED STATES":"USA","EUA":"USA","ESTADOS UNIDOS":"USA",
        }
        if p_up in _alias: return _alias[p_up]
        for key in LC_DEFAULT_PAISES:
            if key in p_up or p_up in key: return key
        return "OTRO"

    _m_mask = df_long["MEASURE"] == "IMPORTE_PEDIDO_USD"
    _f_mask = _m_mask.copy()
    for col, vals in filtros.items():
        if col in df_long.columns and vals:
            _f_mask &= df_long[col].isin(vals)
    df_base = df_long.loc[_f_mask]

    if df_base.empty:
        return {m: 0.0 for m in MONTHS + ["TOTAL"]}

    lc_config = st.session_state.get(
        "lc_config",
        {p: dict(LC_DEFAULT_PAISES[p]) for p in LC_DEFAULT_PAISES}
    )
    mes_map_inv = dict(zip(MONTHS_ES, MONTHS))

    if "PAIS_ORIGEN" in df_base.columns:
        pais_keys  = df_base["PAIS_ORIGEN"].map(_mapear_pais)
        extra_pcts = pais_keys.map(
            lambda pk: sum(lc_config.get(pk, LC_DEFAULT_PAISES.get(pk, LC_DEFAULT_PAISES["OTRO"])).values()) / 100.0
        )
        factor_arr = extra_pcts.values
    else:
        cfg_otro   = lc_config.get("OTRO", LC_DEFAULT_PAISES["OTRO"])
        factor_arr = sum(cfg_otro.values()) / 100.0

    mes_en_arr = df_base["MES"].map(mes_map_inv)
    val_arr    = pd.to_numeric(df_base["VALOR"], errors="coerce").fillna(0) * factor_arr
    valid      = mes_en_arr.notna()
    tmp = pd.Series(val_arr, index=df_base.index).loc[valid].groupby(mes_en_arr[valid]).sum().to_dict()
    result = {m: tmp.get(m, 0.0) for m in MONTHS}
    result["TOTAL"] = sum(result[m] for m in MONTHS)
    return result



def _ventas_data_cached(df_long, filtros_key, filtros):
    """Agrega datos de Ventas/Presupuesto — cacheado en session_state por filtros_key."""
    _cache_key = ('ventas_data', filtros_key)
    if st.session_state.get('_ventas_key') == _cache_key:
        return st.session_state['_ventas_data']

    _MES_ES = {3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',7:'Julio',
               8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre',
               1:'Enero',2:'Febrero'}

    # Construir máscara sin copiar el dataframe completo
    mask = pd.Series(True, index=df_long.index)
    for col, vals in filtros.items():
        if col in df_long.columns and vals and col not in ('MES', 'MEASURE'):
            mask &= df_long[col].isin(vals)
    df_f = df_long.loc[mask]

    def _agg_measure(mea):
        d = df_f.loc[df_f['MEASURE'] == mea, ['ANIO','MES_REAL','VALOR']]
        if d.empty: return {}
        return d.groupby(['ANIO','MES_REAL'])['VALOR'].sum().to_dict()

    def _meses(mea):
        d = df_f.loc[(df_f['MEASURE']==mea)&(df_f['VALOR']>0)&(df_f['MES_REAL']>0), ['ANIO','MES_REAL']]
        if d.empty: return []
        return sorted(d.drop_duplicates().apply(tuple,axis=1).tolist())

    result = {
        'vta_pza':    _agg_measure('VENTAS_PZA'),
        'vta_imp':    _agg_measure('VENTAS_IMP'),
        'pre_pza':    _agg_measure('PRESUPUESTO_PZA'),
        'pre_imp':    _agg_measure('PRESUPUESTO_IMP'),
        'meses_venta': _meses('VENTAS_PZA'),
        'meses_pres':  _meses('PRESUPUESTO_PZA'),
        'mes_es':      _MES_ES,
    }
    st.session_state['_ventas_key']  = _cache_key
    st.session_state['_ventas_data'] = result
    return result

def render_ventas_presupuesto(df_long, filtros_key, filtros, factor_mon, simbolo_mon, moneda_sel, tc_usd_mxn=17.6367):
    """Cuadros independientes de VENTAS HISTÓRICAS y PRESUPUESTO con meses/años correctos."""

    tc_usd_mxn_ref = tc_usd_mxn  # tipo de cambio USD→MXN para convertir importes

    # ── Datos pesados ya cacheados — solo aplica factor de moneda ──
    _d   = _ventas_data_cached(df_long, filtros_key, filtros)

    # Mapeo mes número → nombre español
    _MES_ES    = _d['mes_es']
    _MES_EN_MAP = {3:'March',4:'April',5:'May',6:'June',7:'July',
                   8:'August',9:'September',10:'October',11:'November',12:'December',
                   1:'January',2:'February'}

    _tabla_counter = [0]  # contador para IDs únicos

    def _build_tabla(titulo, color, alt_color, mea_pza, mea_imp, meses_config, pza_vals_in=None, imp_vals_in=None):
        pza_vals = pza_vals_in if pza_vals_in is not None else {}
        imp_vals = imp_vals_in if imp_vals_in is not None else {}
        tid = f"vp_{_tabla_counter[0]}"
        _tabla_counter[0] += 1

        def fp(v):
            if v == 0: return '<span style="color:#CBD5E1;">—</span>'
            return f"{v:,.0f}"

        if moneda_sel == 'MXN':
            conv_imp = 1.0
        elif moneda_sel == 'EUR':
            conv_imp = factor_mon / tc_usd_mxn_ref
        else:
            conv_imp = 1.0 / tc_usd_mxn_ref

        def fi(v):
            if v == 0: return '<span style="color:#CBD5E1;">—</span>'
            vm = v * conv_imp
            return f"{simbolo_mon}{vm:,.0f}"

        # Encabezados meses en la misma fila del título
        ths = ""
        for (anio, mes_num) in meses_config:
            mes_es = _MES_ES.get(mes_num, str(mes_num))
            ths += f'<th style="min-width:90px;text-align:center;padding:9px 6px;color:#fff;font-size:.65rem;font-weight:700;letter-spacing:.3px;border:none;white-space:nowrap;">{mes_es}<br><span style="font-size:.58rem;opacity:.7;font-weight:400;">{anio}</span></th>'
        ths += f'<th style="min-width:110px;text-align:center;padding:9px 6px;color:#FFD700;font-size:.65rem;font-weight:700;background:rgba(0,0,0,.25);border:none;white-space:nowrap;letter-spacing:.5px;">TOTAL ANUAL</th>'

        def _row(label, vals_dict, fmt_fn, row_color):
            total = 0
            cells = ""
            for (anio, mes_num) in meses_config:
                v = float(vals_dict.get((anio, mes_num), 0) or 0)
                total += v
                cells += f'<td class="{tid}_row" style="text-align:right;padding:7px 10px;font-family:\'IBM Plex Mono\',monospace;font-size:.72rem;background:{row_color};border-bottom:1px solid #E8EAF2;">{fmt_fn(v)}</td>'
            cells += f'<td class="{tid}_row" style="text-align:right;padding:7px 10px;font-family:\'IBM Plex Mono\',monospace;font-size:.75rem;font-weight:800;background:#FFF9C4;color:#5D4037;border-left:2px solid #c8960a;">{fmt_fn(total)}</td>'
            return (f'<tr class="{tid}_row">'
                    f'<td style="padding:7px 14px;font-size:.76rem;font-weight:600;color:#1E3A5F;background:{row_color};border-bottom:1px solid #E8EAF2;border-left:3px solid {color};white-space:nowrap;">{label}</td>'
                    f'{cells}</tr>')

        row_pza = _row("Piezas",          pza_vals, fp, "#F8FAFF")
        row_imp = _row(f"Importe ({moneda_sel})", imp_vals, fi, "#EFF6FF")

        return f"""
        <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=IBM+Plex+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
        <div style="margin-bottom:4px;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(15,23,42,.1);border:1px solid #E2E8F0;">
          <div style="overflow-x:auto;">
            <table style="border-collapse:collapse;width:100%;background:#fff;">
              <thead>
                <tr style="background:linear-gradient(135deg,{color},{color}CC);cursor:pointer;"
                    onclick="(function(){{
                      var rows=document.querySelectorAll('.{tid}_row');
                      var arrow=document.getElementById('{tid}_arrow');
                      var vis=rows.length>0&&rows[0].style.display!=='none';
                      rows.forEach(function(r){{r.style.display=vis?'none':''}});
                      if(arrow)arrow.style.transform=vis?'rotate(-90deg)':'rotate(0deg)';
                    }})()">
                  <th style="text-align:left;padding:12px 18px;color:#fff;font-size:.82rem;font-weight:800;min-width:220px;letter-spacing:.5px;border:none;white-space:nowrap;text-transform:uppercase;">
                    <span id="{tid}_arrow" style="display:inline-block;transition:transform .25s ease;font-size:.9rem;margin-right:6px;">▾</span>{titulo}
                  </th>
                  {ths}
                </tr>
              </thead>
              <tbody>
                {row_pza}
                {row_imp}
              </tbody>
            </table>
          </div>
        </div>"""

    # ── Meses ya calculados en caché ─────────────────────────
    meses_venta = _d['meses_venta']
    meses_pres  = _d['meses_pres']
    html = ""
    if meses_venta:
        html += _build_tabla(
            titulo      = "📈 VENTAS HISTÓRICAS",
            color       = "#1F4E79",
            alt_color   = "#D6E4F0",
            mea_pza     = "VENTAS_PZA",
            mea_imp     = "VENTAS_IMP",
            meses_config= meses_venta,
            pza_vals_in = _d['vta_pza'],
            imp_vals_in = _d['vta_imp'],
        )
    if meses_pres:
        html += _build_tabla(
            titulo      = "🎯 PRESUPUESTO DE VENTAS",
            color       = "#1F4E79",
            alt_color   = "#D6E4F0",
            mea_pza     = "PRESUPUESTO_PZA",
            mea_imp     = "PRESUPUESTO_IMP",
            meses_config= meses_pres,
            pza_vals_in = _d['pre_pza'],
            imp_vals_in = _d['pre_imp'],
        )

    if html:
        n_tables = (1 if meses_venta else 0) + (1 if meses_pres else 0)
        # header (40px) + 2 filas (33px cada una) + 4px margen entre tablas
        _h_por_tabla = 40 + 2 * 33
        altura = n_tables * _h_por_tabla + max(0, n_tables - 1) * 4 + 2
        components.html(html, height=altura, scrolling=False)


def render_resumen(pivot, filtros_label, moneda='USD', factor=1.0, simbolo='$', landed_cost_monthly=None, costos_log_monthly=None):
    mos_data = calcular_mos_resumen(pivot)

    def gv(measure, col):
        # CAMBIO 3: incluir __MAX_MOS__ y __MIN_MOS__ en el mismo bloque
        if measure in ('__MOS__', '__VUELTAS__', '__MAX_MOS__', '__MIN_MOS__', '__COB_REPOS__'):
            arr = mos_data[measure]
            if col == 'TOTAL': return arr[-1]
            if col in MONTHS:
                idx = MONTHS.index(col)
                return arr[idx] if idx < len(arr) else None
            return None
        if measure == '__SS_MOS__':
            # Safety Stock (Meses) = Safety Stock / Net Forecast  (misma lógica que MOS)
            if 'Safety Stock' not in pivot.index: return None
            ss_v = pivot.loc['Safety Stock', col] if col in pivot.columns else None
            if ss_v is None or (isinstance(ss_v, float) and pd.isna(ss_v)): return None
            # buscar forecast igual que calcular_mos_resumen
            fc_keys = ['Net Forecast', 'Total Demand', 'Gross Forecast']
            fc_row  = next((pivot.loc[k] for k in fc_keys if k in pivot.index), None)
            fc_v    = float(fc_row[col]) if (fc_row is not None and col in fc_row.index and pd.notna(fc_row.get(col))) else None
            if fc_v and fc_v > 0:
                return round(float(ss_v) / fc_v, 2)
            return None
        if measure == "__TOTAL_COSTO_PED__":
            _cped = ["OC_USD", "REQUISICION_USD", "IMPORTE_PEDIDO_USD"]
            vals  = [float(pivot.loc[m, col]) for m in _cped
                     if m in pivot.index and col in pivot.columns and pd.notna(pivot.loc[m, col])]
            return sum(vals) * factor if vals else 0.0
        if measure == "__COSTOS_LOG__":
            # Costos logísticos en USD → convertir a moneda seleccionada
            if costos_log_monthly is None:
                return 0.0
            if col == "TOTAL":
                return costos_log_monthly.get("TOTAL", 0.0) * factor
            if col in MONTHS:
                return costos_log_monthly.get(col, 0.0) * factor
            return 0.0
        if measure == "__TOTAL_FIN__":
            # TOTAL LANDED COST = (Pedido Sugerido CIF + Costos Logísticos) × factor
            imp = float(pivot.loc["IMPORTE_PEDIDO_USD", col]) \
                  if ("IMPORTE_PEDIDO_USD" in pivot.index and col in pivot.columns
                      and pd.notna(pivot.loc["IMPORTE_PEDIDO_USD", col])) else 0.0
            log = costos_log_monthly.get(col if col in MONTHS else "TOTAL", 0.0) \
                  if costos_log_monthly is not None else 0.0
            return (imp + log) * factor
        if measure == "__NET_FC_USD__":
            # Net Forecast × PRICE_USD (precio de proveedor hoja ACUERDOS, ya calculado en carga)
            if "NET_FC_USD" not in pivot.index or col not in pivot.columns:
                return 0.0
            v = pivot.loc["NET_FC_USD", col]
            return float(v) * factor if pd.notna(v) else 0.0
        if measure in pivot.index and col in pivot.columns:
            v = pivot.loc[measure, col]
            if not pd.notna(v):
                return 0.0
            # Medidas financieras (USD) → aplicar factor de conversión igual que el total
            if measure in FIN_MEASURES:
                return float(v) * factor
            return float(v)
        return 0.0

    # ── CAMBIO 4: fmt acepta measure para mostrar " m" en métricas de meses ──
    _MOS_MESES  = {'__MOS__', '__MAX_MOS__', '__MIN_MOS__', '__COB_REPOS__', '__SS_MOS__'}
    _FIN_EXTRA  = {'__NET_FC_USD__'}   # medidas $ dentro de secciones no-financieras

    def fmt(v, es_fin, measure=""):
        if v is None or (isinstance(v, float) and pd.isna(v)): return "—"
        try:
            n = float(v)
            if es_fin or measure in _FIN_EXTRA:
                return f"{simbolo}{n:,.0f}"
            if measure in _MOS_MESES:
                return f"{n:,.2f} m"
            if abs(n) < 100_000 and n != int(n):
                return f"{n:,.2f}"
            return f"{n:,.0f}"
        except:
            return str(v)

    body = ""
    alt  = True

    for sec_idx, sec in enumerate(SECTIONS):
        bg    = sec["color"]
        alt_c = sec["alt"]
        es_f  = sec["es_fin"]
        titulo = f"{sec['titulo']} ({moneda})" if es_f else sec['titulo']
        sec_id = f"sec_{sec_idx}"

        ths_hdr = "".join(
            f'<th style="min-width:82px;text-align:center;padding:9px 6px;color:#fff;font-size:.67rem;font-weight:700;letter-spacing:.3px;border:none;white-space:nowrap;">'
            f'{m}<br><span style="font-size:.56rem;font-weight:400;opacity:.7;">2026</span></th>'
            for m in MONTHS_ES
        )
        body += f"""
        <tr class="sec-hdr" style="background:{bg};cursor:pointer;" onclick="toggleSection('{sec_id}')">
          <td style="width:6px;background:{bg};border:none;"></td>
          <th style="text-align:left;padding:10px 18px;color:#fff;font-size:.80rem;font-weight:800;min-width:270px;letter-spacing:.3px;white-space:nowrap;background:{bg};border:none;">
            <span class="arrow" id="{sec_id}_arrow" style="transform:rotate(0deg);display:inline-block;transition:transform .25s ease;font-size:.9rem;">▾</span>
            &nbsp;{titulo}
          </th>
          {ths_hdr}
          <th style="min-width:108px;text-align:center;padding:9px 6px;color:#2d1e00;font-size:.67rem;font-weight:700;background:#c8960a;border:none;white-space:nowrap;">TOTAL ANUAL</th>
        </tr>"""

        for label, measure in sec["metricas"]:
            is_tot    = measure in ("__TOTAL_FIN__", "__TOTAL_COSTO_PED__")
            is_net_usd = measure == "__NET_FC_USD__"
            fill   = "#FFF9C4" if is_tot else ("#E8F5E9" if is_net_usd else (alt_c if alt else "#FFFFFF"))
            fw     = "700"     if (is_tot or is_net_usd) else "400"
            tc     = "#5D4037" if is_tot else ("#1B5E20" if is_net_usd else "#1a1a2e")
            lbl_d  = f"{label} ({simbolo})" if (es_f or is_net_usd) else label

            cells = ""
            for month_en in MONTHS:
                v = gv(measure, month_en)
                cells += (f'<td class="num" style="background:{fill};color:{tc};font-weight:{fw};">'
                          f'{fmt(v, es_f, measure)}</td>')

            v_tot = gv(measure, 'TOTAL')
            cells += f'<td class="num col-tot">{fmt(v_tot, es_f, measure)}</td>'

            row_class = "dr row-tot" if is_tot else "dr"
            body += f"""
            <tr class="{row_class} {sec_id}_row" style="display:table-row;transition:all .2s;">
              <td class="bar" style="background:{bg};border-right:2px solid {bg};"></td>
              <td class="lbl" style="background:{fill};color:{tc};font-weight:{fw};">{lbl_d}</td>
              {cells}
            </tr>"""
            alt = not alt

        body += f'<tr class="sep {sec_id}_row" style="display:table-row;"><td colspan="{len(MONTHS_ES)+3}"></td></tr>'

    return f"""
    <!DOCTYPE html><html><head>
    <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
    *{{box-sizing:border-box;margin:0;padding:0;}}
    body{{font-family:'IBM Plex Sans',sans-serif;background:transparent;}}
    .res-wrap{{overflow-x:auto;border-radius:12px;box-shadow:0 4px 28px rgba(13,27,75,.15);margin-bottom:.5rem;}}
    .res-tbl{{border-collapse:collapse;width:100%;min-width:1160px;background:#fff;}}
    .res-tbl thead tr{{background:linear-gradient(180deg,#1A237E,#162082);}}
    .res-tbl thead th{{color:#fff;font-weight:600;font-size:.71rem;letter-spacing:.3px;padding:8px 6px;text-align:center;border:none;white-space:normal;line-height:1.4;}}
    .res-tbl thead th.th-tipo{{text-align:left;padding-left:32px;min-width:270px;}}
    .res-tbl thead th.th-total{{background:#c8960a;color:#2d1e00;min-width:108px;}}
    .sec-hdr td{{color:#fff;font-weight:700;font-size:.78rem;letter-spacing:.5px;text-transform:uppercase;padding:9px 14px 9px 18px;white-space:nowrap;cursor:pointer;user-select:none;}}
    .sec-hdr:hover td{{filter:brightness(1.1);}}
    .sec-hdr .bar{{width:6px;padding:0;}}
    .dr td{{font-size:.77rem;padding:6px 8px;border-bottom:1px solid #e8eaf2;white-space:nowrap;vertical-align:middle;}}
    .dr .bar{{width:6px;padding:0;}}
    .dr .lbl{{padding:6px 10px 6px 22px;text-align:left;font-size:.77rem;}}
    .dr .num{{text-align:right;padding:6px 10px;font-family:'IBM Plex Mono',monospace;font-size:.72rem;}}
    .row-tot td{{background:#FFF9C4!important;color:#5D4037!important;font-weight:700!important;}}
    .col-tot{{background:#FFF9C4!important;color:#5D4037!important;font-weight:700!important;border-left:2px solid #c8960a!important;padding-right:14px!important;}}
    .sep td{{height:5px;background:#E8EAF6;padding:0;border:none;}}
    .arrow{{display:inline-block;transition:transform .25s ease;font-size:.9rem;}}
    p.footer{{font-size:.64rem;color:#bbb;text-align:right;margin-top:4px;margin-bottom:0;}}
    html,body{{margin:0;padding:0;overflow-x:hidden;}}
    </style></head><body style="margin:0;padding:0;">
    <div class="res-wrap">
      <table class="res-tbl">
        <tbody>{body}</tbody>
      </table>
    </div>
    <p class="footer">Filtros: {filtros_label} &nbsp;·&nbsp; Moneda: {moneda} &nbsp;·&nbsp; {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
    <script>
    function toggleSection(secId) {{
        var rows = document.querySelectorAll('.' + secId + '_row');
        var arrow = document.getElementById(secId + '_arrow');
        var isVisible = rows.length > 0 && rows[0].style.display !== 'none';
        rows.forEach(function(r) {{ r.style.display = isVisible ? 'none' : ''; }});
        if (arrow) arrow.style.transform = isVisible ? 'rotate(-90deg)' : 'rotate(0deg)';
    }}
    </script>
    </body></html>"""


def duck_pivot(df_long, filas, columnas, agg, filtros_key, filtros):
    con = make_con(df_long)
    where = []
    for col, vals in filtros.items():
        if vals:
            lista = ", ".join(f"'{v}'" for v in vals)
            where.append(f'"{col}" IN ({lista})')
    wc = "WHERE " + " AND ".join(where) if where else ""

    if columnas == 'MES':
        col_vals = [m for m in MONTHS_ES if m in df_long['MES'].unique()]
    else:
        col_vals = sorted([str(v) for v in df_long[columnas].dropna().unique()])[:60]

    if not col_vals or not filas: return pd.DataFrame()

    pcols = ", ".join([
        f'{agg}(CASE WHEN "{columnas}"=\'{v}\' THEN VALOR ELSE 0 END) AS "{v}"'
        for v in col_vals
    ])
    gcols = ", ".join(f'"{c}"' for c in filas)
    q = f"SELECT {gcols},{pcols},{agg}(VALOR) AS TOTAL FROM datos {wc} GROUP BY {gcols} ORDER BY {gcols}"
    try:
        r = con.execute(q).df(); return r
    except Exception as e:
        st.error(f"Error pivot: {e}"); return pd.DataFrame()


def duck_items_pedido(df_long, filtros_key, filtros):
    """Cuenta ítems con Planned Replenishments (total) y con Marzo > 0 — 1 sola query."""
    con = make_con(df_long)
    where = ["MEASURE = 'Planned Replenishments by Order Date'"]
    for col, vals in filtros.items():
        if vals and col != 'MES':
            lista = ", ".join(f"'{v}'" for v in vals)
            where.append(f'"{col}" IN ({lista})')
    wc = "WHERE " + " AND ".join(where)
    r = con.execute(f"""
        SELECT
            COUNT(DISTINCT ITEM) AS total,
            COUNT(DISTINCT CASE WHEN MES = 'Marzo' AND VALOR > 0 THEN ITEM END) AS marzo
        FROM datos {wc}
    """).fetchone()
    return int(r[0]), int(r[1])


def duck_kpis(df_long, filtros_key, filtros):
    """Stats básicos — combinados en la misma query para no hacer un scan extra."""
    con = make_con(df_long)
    where = []
    for col, vals in filtros.items():
        if vals:
            lista = ", ".join(f"'{v}'" for v in vals)
            where.append(f'"{col}" IN ({lista})')
    wc = "WHERE " + " AND ".join(where) if where else ""
    r = con.execute(f"""
        SELECT COUNT(DISTINCT ITEM)     AS items,
               COUNT(DISTINCT MEASURE)  AS metricas,
               COUNT(DISTINCT FAMILIA)  AS familias,
               COUNT(DISTINCT SUPPLIER) AS suppliers,
               COUNT(*)                 AS total_filas
        FROM datos {wc}
    """).df().iloc[0]
    return r


def duck_kpis_rich(df_long, filtros_key, filtros):
    con = make_con(df_long)
    where = ["ANIO = 2026"]
    for col, vals in filtros.items():
        if vals:
            lista = ", ".join(f"'{v}'" for v in vals)
            where.append(f'"{col}" IN ({lista})')
    wc = "WHERE " + " AND ".join(where)

    # ── Una sola query: todos los KPIs en un solo escaneo ─────────────────
    r = con.execute(f"""
        SELECT
            COALESCE(SUM(CASE WHEN MEASURE='OC_USD'                                              THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='Purchase Orders'                                     THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='REQUISICION_USD'                                     THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='Purchase Requisitions'                               THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='IMPORTE_PEDIDO_USD'                                  THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='Planned Replenishments by Order Date'                THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='Projected Available Balance'                         THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='Net Forecast'                                        THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='Total Demand'                                        THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='On Hand' THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='On Hand' THEN VALOR * COSTO ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='IMPORTE_PEDIDO_USD'               AND MES='Marzo'   THEN VALOR ELSE 0 END),0),
            COALESCE(SUM(CASE WHEN MEASURE='Planned Replenishments by Order Date' AND MES='Marzo' THEN VALOR ELSE 0 END),0)
        FROM datos {wc}
    """).fetchone()

    (oct_usd, oct_pzas, oc_usd, oc_pzas, pedido_usd, pedido_pzas,
     pab_tot, fc_net, fc_dem, inv_pzas, inv_mxn,
     pedido_marzo_usd, pedido_marzo_pzas) = [float(x) for x in r]

    fc_tot = fc_net if fc_net != 0 else fc_dem
    mos    = round(pab_tot / fc_tot, 1) if fc_tot > 0 else None
    vuel   = round(12 / mos, 1) if (mos and mos > 0) else None

    return {
        'inv_mxn':    inv_mxn,    'inv_pzas':    inv_pzas,
        'oct_usd':    oct_usd,    'oct_pzas':    oct_pzas,
        'oc_usd':     oc_usd,     'oc_pzas':     oc_pzas,
        'pedido_usd': pedido_usd, 'pedido_pzas': pedido_pzas,
        'pedido_marzo_usd':  pedido_marzo_usd,
        'pedido_marzo_pzas': pedido_marzo_pzas,
        'mos': mos, 'vuel': vuel,
    }


def get_graf_data(df_long, g_meas, g_dim, filtros_key, filtros):
    con = make_con(df_long)
    wg = f"WHERE MEASURE = '{g_meas}'"
    for col, vals in filtros.items():
        if vals:
            lista = ", ".join(f"'{v}'" for v in vals)
            wg += f' AND "{col}" IN ({lista})'
    df_graf = con.execute(f"""
        SELECT "{g_dim}", MES_NUM, MES, SUM(VALOR) AS VALOR
        FROM datos {wg}
        GROUP BY "{g_dim}", MES_NUM, MES
        ORDER BY "{g_dim}", MES_NUM
    """).df()
    df_top = con.execute(f"""
        SELECT ITEM, SUM(VALOR) AS TOTAL FROM datos {wg}
        GROUP BY ITEM ORDER BY TOTAL DESC LIMIT 20
    """).df()
    return df_graf, df_top, wg


def get_pivot_resumen(df, filtros_key, filtros):
    """Pivot de resumen — vectorizado sin copy innecesario."""
    # Filtrar solo ANIO=2026
    mask = (df['ANIO'] == 2026) if 'ANIO' in df.columns else pd.Series(True, index=df.index)
    for col, vals in filtros.items():
        if col in df.columns and vals and col != 'MES':
            mask = mask & df[col].isin(vals)

    # Seleccionar solo columnas necesarias antes del groupby
    cols_needed = ['MEASURE'] + [m for m in MONTHS if m in df.columns]
    df_fil = df.loc[mask, cols_needed]

    if 'MES' in filtros and filtros['MES']:
        meses_en = [k for k, v in MES_MAP.items() if v in filtros['MES']]
        cols_zero = [m for m in MONTHS if m not in meses_en and m in df_fil.columns]
        if cols_zero:
            df_fil = df_fil.copy()
            df_fil[cols_zero] = 0

    pivot = df_fil.groupby('MEASURE', sort=False)[MONTHS].sum()
    pivot['TOTAL'] = pivot.sum(axis=1)
    return pivot


@st.cache_data(show_spinner=False)
def get_workbench(df_long, item_sel, anio_sel):
    con = make_con(df_long)
    where_parts = [f"ITEM = '{item_sel}'", f"ANIO = {anio_sel}"]
    wc = "WHERE " + " AND ".join(where_parts)

    months_q = ", ".join([
        f'ROUND(SUM(CASE WHEN MES=\'{mes_es}\' THEN VALOR ELSE 0 END),2) AS "{mes_es}"'
        for m, mes_es in zip(MONTHS, MONTHS_ES)
    ])
    q = f"""
        SELECT MEASURE,
               {months_q},
               ROUND(SUM(VALOR),2) AS TOTAL
        FROM datos {wc}
        GROUP BY MEASURE
        ORDER BY MEASURE
    """
    df = con.execute(q).df()

    ORDER = [
        'Gross Forecast','Net Forecast','Sales Orders','Total Demand',
        'Final Minimum Quantity','Final Maximum Quantity',
        'Safety Stock','Safety Stock Days',
        'On Hand','Purchase Orders','Purchase Requisitions',
        'Planned Replenishments by Order Date','Planned Replenishments by Due Date',
        'On Order','Total Supply',
        'Projected Available Balance',
        'Beginning Inventory Position','Final Inventory Position',
        'COSTO_TOTAL_USD','OC_USD','IMPORTE_PEDIDO_USD','REQUISICION_USD','COSTO_IGI',
    ]
    df['_ord'] = df['MEASURE'].apply(lambda x: ORDER.index(x) if x in ORDER else 999)
    df = df.sort_values('_ord').drop(columns='_ord').reset_index(drop=True)
    return df


# ════════════════════════════════════════════════════════════
#  LANDED COST — función de renderizado
# ════════════════════════════════════════════════════════════
def render_landed_cost_tab(df_long, tc_usd_mxn, moneda_sel, factor_mon, simbolo_mon, filtros_key=(), filtros=None):
    moneda_label = MONEDA_CONFIG[moneda_sel]["label"]
    if moneda_sel == "USD":
        tc_label = "Base USD — sin conversión"
    elif moneda_sel == "MXN":
        tc_label = f"1 USD = {tc_usd_mxn:,.4f} MXN"
    else:
        tc_eur = tc_usd_mxn / (st.session_state.get("tc_eur_mxn", 20.55))
        tc_label = f"1 USD = {tc_eur:,.4f} EUR"

    st.markdown("### 🧮 Landed Cost — Costo Total de Importación")
    st.markdown(
        f'<div class="moneda-activa">'
        f'{moneda_label}'
        f'<span class="tc"> · {tc_label} · TC editable en sidebar</span>'
        f'</div>',
        unsafe_allow_html=True
    )

    def mapear_pais(p):
        p_up = str(p).upper().strip()
        if p_up in LC_DEFAULT_PAISES:
            return p_up
        _alias = {
            "TURKEY": "TURQUIA",
            "MALAYSIA": "MALASIA", "MALAYSIAN": "MALASIA",
            "BRAZIL": "BRASIL",
            "MEX": "MEXICO", "ESTADOS UNIDOS MEXICANOS": "MEXICO",
            "GERMANY": "ALEMANIA", "DEUTSCHLAND": "ALEMANIA",
            "ITALY": "ITALIA",
            "JAPON": "JAPON",
            "KOREA": "COREA", "SOUTH KOREA": "COREA",
            "UNITED STATES": "USA", "EUA": "USA", "ESTADOS UNIDOS": "USA",
        }
        if p_up in _alias:
            return _alias[p_up]
        for key in LC_DEFAULT_PAISES:
            if key in p_up or p_up in key:
                return key
        return "OTRO"

    paises_en_datos = []
    if "PAIS_ORIGEN" in df_long.columns:
        for p in df_long["PAIS_ORIGEN"].dropna().unique().tolist():
            m = mapear_pais(p)
            if m not in paises_en_datos:
                paises_en_datos.append(m)

    default_sel = paises_en_datos[:8] if paises_en_datos else ["CHINA", "USA"]

    st.markdown("---")
    st.markdown("#### ⚙️ Configuración de % por País de Origen")
    st.caption(f"Porcentajes sobre el **Valor CIF en USD**. El resultado se muestra en **{moneda_sel}** ({simbolo_mon}).")

    paises_opts = list(LC_DEFAULT_PAISES.keys())
    paises_sel  = st.multiselect(
        "🌍 Países a configurar",
        paises_opts,
        default=default_sel,
        key="lc_paises_sel"
    )
    if not paises_sel:
        st.info("Selecciona al menos un país para continuar.")
        return

    if "lc_config" not in st.session_state:
        st.session_state["lc_config"] = {p: dict(LC_DEFAULT_PAISES[p]) for p in LC_DEFAULT_PAISES}

    n_cols    = min(3, len(paises_sel))
    cols_edit = st.columns(n_cols)
    for idx, pais in enumerate(paises_sel):
        with cols_edit[idx % n_cols]:
            with st.expander(f"🌐 {pais}", expanded=(idx == 0)):
                cfg     = st.session_state["lc_config"].get(pais, dict(LC_DEFAULT_PAISES.get(pais, LC_DEFAULT_PAISES["OTRO"])))
                new_cfg = {}
                for concepto in LC_CONCEPTOS:
                    new_cfg[concepto] = st.number_input(
                        f"{concepto} %",
                        value=float(cfg.get(concepto, 0.0)),
                        min_value=0.0, max_value=50.0,
                        step=0.05, format="%.2f",
                        key=f"lc_{pais}_{concepto}"
                    )
                st.session_state["lc_config"][pais] = new_cfg
                total_pct = sum(new_cfg.values())
                st.markdown(
                    f'<div style="background:#E3F2FD;border-radius:7px;padding:.35rem .6rem;'
                    f'font-size:.68rem;font-weight:700;color:#1565C0;text-align:center;">'
                    f'Σ Costos adicionales: {total_pct:.2f}%</div>',
                    unsafe_allow_html=True
                )

    st.markdown("---")
    st.markdown(f"#### 📊 Resumen Landed Cost — en **{moneda_sel}** ({simbolo_mon})")

    modo = st.radio(
        "Fuente del Valor CIF:",
        ["📦 Desde datos (COSTO_TOTAL_USD por país)", "✏️ Ingresar valor CIF manualmente"],
        horizontal=True,
        key="lc_modo"
    )

    if "manual" in modo:
        cols_m = st.columns(min(4, len(paises_sel)))
        valores_cif = {}
        for idx, pais in enumerate(paises_sel):
            with cols_m[idx % len(cols_m)]:
                valores_cif[pais] = st.number_input(
                    f"CIF {pais} (USD)",
                    value=100_000.0, min_value=0.0,
                    step=1000.0, format="%.2f",
                    key=f"lc_cif_{pais}"
                )
    else:
        df_base = df_long[df_long["MEASURE"] == "COSTO_TOTAL_USD"].copy()
        if df_base.empty:
            df_base = df_long[df_long["MEASURE"] == "IMPORTE_PEDIDO_USD"].copy()

        valores_cif = {p: 0.0 for p in paises_sel}

        if not df_base.empty and "PAIS_ORIGEN" in df_base.columns:
            df_pais_val = df_base.groupby("PAIS_ORIGEN")["VALOR"].sum().reset_index()
            df_pais_val["PAIS_KEY"] = df_pais_val["PAIS_ORIGEN"].apply(mapear_pais)
            df_agg = df_pais_val.groupby("PAIS_KEY")["VALOR"].sum()
            for p in paises_sel:
                valores_cif[p] = float(df_agg.get(p, 0.0))

            df_show = pd.DataFrame([
                {
                    "País": p,
                    f"Valor CIF (USD)": f"${v:,.2f}",
                    f"Valor CIF ({moneda_sel})": f"{simbolo_mon}{v * factor_mon:,.2f}",
                }
                for p, v in valores_cif.items() if v > 0
            ])
            if not df_show.empty:
                st.caption(f"Valores CIF detectados — mostrados en {moneda_sel}:")
                st.dataframe(df_show, use_container_width=False, hide_index=True)

    filas_tabla      = []
    totales_concepto = {c: 0.0 for c in LC_CONCEPTOS}
    total_cif_usd    = 0.0
    total_costos_usd = 0.0

    for pais in paises_sel:
        cif_usd = valores_cif.get(pais, 0.0)
        cfg_p   = st.session_state["lc_config"].get(pais, LC_DEFAULT_PAISES.get(pais, LC_DEFAULT_PAISES["OTRO"]))
        fila    = {
            "País": pais,
            "Valor CIF (USD)":    cif_usd,
            f"Valor CIF ({moneda_sel})": cif_usd * factor_mon,
        }

        suma_costos_usd = 0.0
        for concepto in LC_CONCEPTOS:
            pct     = cfg_p.get(concepto, 0.0) / 100.0
            val_usd = cif_usd * pct
            fila[f"{concepto}_USD"] = val_usd
            fila[f"{concepto}_MON"] = val_usd * factor_mon
            fila[f"{concepto}_PCT"] = cfg_p.get(concepto, 0.0)
            totales_concepto[concepto] += val_usd * factor_mon
            suma_costos_usd += val_usd

        fila["Total Costos (USD)"]          = suma_costos_usd
        fila[f"Total Costos ({moneda_sel})"] = suma_costos_usd * factor_mon
        fila["Landed Cost (USD)"]            = cif_usd + suma_costos_usd
        fila[f"Landed Cost ({moneda_sel})"]  = (cif_usd + suma_costos_usd) * factor_mon
        fila["% Adicional Total"]            = (suma_costos_usd / cif_usd * 100) if cif_usd > 0 else 0.0

        total_cif_usd    += cif_usd
        total_costos_usd += suma_costos_usd
        filas_tabla.append(fila)

    df_lc = pd.DataFrame(filas_tabla)
    if df_lc.empty:
        st.warning("Sin datos para calcular.")
        return

    total_lc_usd     = total_cif_usd + total_costos_usd
    total_lc_mon     = total_lc_usd  * factor_mon
    total_cif_mon    = total_cif_usd * factor_mon
    total_costos_mon = total_costos_usd * factor_mon
    pct_ad_prom      = (total_costos_usd / total_cif_usd * 100) if total_cif_usd > 0 else 0.0

    def fmt_k(v):
        return f"{simbolo_mon}{v:,.0f}"

    kc1, kc2, kc3, kc4 = st.columns(4)
    kc1.markdown(f"""<div class="lc-kpi blue">
        <div class="lc-kpi-lbl">Valor CIF Total</div>
        <div class="lc-kpi-val">{fmt_k(total_cif_mon)}</div>
        <div class="lc-kpi-sub">{moneda_sel} · ${total_cif_usd:,.0f} USD base</div>
    </div>""", unsafe_allow_html=True)

    kc2.markdown(f"""<div class="lc-kpi amber">
        <div class="lc-kpi-lbl">Total Costos Aduanales</div>
        <div class="lc-kpi-val">{fmt_k(total_costos_mon)}</div>
        <div class="lc-kpi-sub">{moneda_sel} · ${total_costos_usd:,.0f} USD base</div>
    </div>""", unsafe_allow_html=True)

    kc3.markdown(f"""<div class="lc-kpi green">
        <div class="lc-kpi-lbl">Landed Cost Total</div>
        <div class="lc-kpi-val">{fmt_k(total_lc_mon)}</div>
        <div class="lc-kpi-sub">{moneda_sel} · ${total_lc_usd:,.0f} USD base</div>
    </div>""", unsafe_allow_html=True)

    kc4.markdown(f"""<div class="lc-kpi red">
        <div class="lc-kpi-lbl">% Adicional Promedio</div>
        <div class="lc-kpi-val">{pct_ad_prom:.1f}%</div>
        <div class="lc-kpi-sub">sobre Valor CIF</div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(f"##### 📋 Desglose por País y Concepto ({moneda_sel} — {simbolo_mon})")

    ths_c = "".join(
        f'<th style="min-width:110px;">{c}<br>'
        f'<span style="font-size:.58rem;opacity:.7;">% / {simbolo_mon}</span></th>'
        for c in LC_CONCEPTOS
    )
    body_lc = ""
    for _, row in df_lc.iterrows():
        cif_mon = row[f"Valor CIF ({moneda_sel})"]
        cells_c = ""
        for concepto in LC_CONCEPTOS:
            pct_v = row.get(f"{concepto}_PCT", 0.0)
            mon_v = row.get(f"{concepto}_MON", 0.0)
            cells_c += (
                f'<td class="td-mxn">'
                f'<span style="font-size:.62rem;color:#888;">{pct_v:.2f}%</span><br>'
                f'{simbolo_mon}{mon_v:,.0f}</td>'
            )
        lc_mon  = row.get(f"Landed Cost ({moneda_sel})", 0.0)
        cos_mon = row.get(f"Total Costos ({moneda_sel})", 0.0)
        body_lc += f"""
        <tr>
          <td class="td-left">🌐 {row['País']}</td>
          <td class="td-mxn">{simbolo_mon}{cif_mon:,.0f}</td>
          {cells_c}
          <td class="td-mxn">{simbolo_mon}{cos_mon:,.0f}</td>
          <td class="td-tot">{simbolo_mon}{lc_mon:,.0f}</td>
          <td class="td-mxn" style="text-align:center;color:#C62828;font-weight:700;">
            {row['% Adicional Total']:.1f}%</td>
        </tr>"""

    tot_cells = "".join(
        f'<td class="td-mxn" style="font-weight:700;">{simbolo_mon}{totales_concepto[c]:,.0f}</td>'
        for c in LC_CONCEPTOS
    )
    body_lc += f"""
    <tr class="row-lc-tot">
      <td class="td-left">∑ TOTAL</td>
      <td class="td-mxn" style="font-weight:700;">{simbolo_mon}{total_cif_mon:,.0f}</td>
      {tot_cells}
      <td class="td-mxn" style="font-weight:700;">{simbolo_mon}{total_costos_mon:,.0f}</td>
      <td class="td-tot">{simbolo_mon}{total_lc_mon:,.0f}</td>
      <td class="td-mxn" style="text-align:center;color:#1B5E20;font-weight:800;">
        {pct_ad_prom:.1f}%</td>
    </tr>"""

    st.markdown(f"""
    <div class="lc-wrap">
    <table class="lc-tbl">
      <thead><tr>
        <th class="th-left">País Origen</th>
        <th class="th-pais" style="min-width:130px;">Valor CIF ({moneda_sel})</th>
        {ths_c}
        <th style="min-width:130px;">Total Costos ({moneda_sel})</th>
        <th class="th-tot" style="min-width:140px;">LANDED COST ({moneda_sel})</th>
        <th style="min-width:90px;">% Adicional</th>
      </tr></thead>
      <tbody>{body_lc}</tbody>
    </table></div>
    <p style="font-size:.62rem;color:#bbb;text-align:right;margin-top:4px;">
      Moneda: {moneda_sel} · Factor: {factor_mon:,.4f} · {tc_label}
      &nbsp;·&nbsp; {datetime.now().strftime('%d/%m/%Y %H:%M')}
    </p>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown(f"##### 📈 Visualización ({moneda_sel})")

    gc1, gc2 = st.columns(2)

    with gc1:
        conceptos_vals = {c: totales_concepto[c] for c in LC_CONCEPTOS if totales_concepto[c] > 0}
        if conceptos_vals:
            fig_pie = go.Figure(go.Pie(
                labels=list(conceptos_vals.keys()),
                values=list(conceptos_vals.values()),
                hole=.45,
                marker_colors=[LC_CONCEPTOS_COLOR.get(c, "#999") for c in conceptos_vals],
                textinfo="label+percent",
                textfont_size=10,
            ))
            fig_pie.update_layout(
                title=f"Distribución Costos Aduanales ({moneda_sel})",
                template="plotly_white", height=380,
                font=dict(family="IBM Plex Sans", size=11),
                margin=dict(l=10, r=10, t=50, b=10),
                showlegend=False,
            )
            st.plotly_chart(fig_pie, use_container_width=True)

    with gc2:
        if not df_lc.empty:
            bar_data = []
            for _, row in df_lc.iterrows():
                bar_data.append({"País": row["País"], "Concepto": "Valor CIF", "Valor": row[f"Valor CIF ({moneda_sel})"]})
                for concepto in LC_CONCEPTOS:
                    v = row.get(f"{concepto}_MON", 0.0)
                    if v > 0:
                        bar_data.append({"País": row["País"], "Concepto": concepto, "Valor": v})
            df_bar = pd.DataFrame(bar_data)
            if not df_bar.empty:
                color_map = {"Valor CIF": "#CFD8DC"}
                color_map.update(LC_CONCEPTOS_COLOR)
                fig_bar = px.bar(
                    df_bar, x="País", y="Valor", color="Concepto",
                    color_discrete_map=color_map,
                    title=f"Landed Cost por País ({moneda_sel})",
                    barmode="stack",
                    labels={"Valor": f"{moneda_sel} ({simbolo_mon})"},
                )
                fig_bar.update_layout(
                    template="plotly_white", height=380,
                    font=dict(family="IBM Plex Sans", size=11),
                    margin=dict(l=10, r=10, t=50, b=10),
                    legend=dict(font=dict(size=9)),
                )
                st.plotly_chart(fig_bar, use_container_width=True)

    cols_exp = ["País", "Valor CIF (USD)", f"Valor CIF ({moneda_sel})"]
    for c in LC_CONCEPTOS:
        cols_exp += [f"{c}_PCT", f"{c}_MON"]
    cols_exp += [
        "Total Costos (USD)", f"Total Costos ({moneda_sel})",
        "Landed Cost (USD)",  f"Landed Cost ({moneda_sel})",
        "% Adicional Total"
    ]
    df_exp = df_lc[[c for c in cols_exp if c in df_lc.columns]].copy()
    buf_lc = io.BytesIO()
    df_exp.to_excel(buf_lc, index=False, sheet_name="Landed Cost")
    buf_lc.seek(0)
    st.download_button(
        f"⬇️ Exportar Landed Cost ({moneda_sel}) .xlsx", buf_lc,
        f"LandedCost_{moneda_sel}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ════════════════════════════════════════════════════════════
#  CUADRO EJECUTIVO — Una sola query DuckDB (ultrarrápido)
# ════════════════════════════════════════════════════════════
def get_cuadro_ejecutivo(df_long, dimension, filtros_key, filtros, mes_filtro=None):
    con = make_con(df_long)

    where = ["ANIO = 2026"]  # solo datos 2026
    for col, vals in filtros.items():
        if vals:
            lista = ", ".join(f"'{v}'" for v in vals)
            where.append(f'"{col}" IN ({lista})')
    if mes_filtro and mes_filtro != "📅 Anual (Total)":
        where.append(f"MES = '{mes_filtro}'")
    wc = "WHERE " + " AND ".join(where)

    # ── UNA SOLA QUERY con CASE WHEN — 10x más rápido ─────────
    q = f"""
    SELECT
        "{dimension}" AS DIM,
        COALESCE(SUM(CASE WHEN MEASURE='On Hand'
                     THEN VALOR ELSE 0 END), 0)                         AS INV_PZA,
        COALESCE(SUM(CASE WHEN MEASURE='On Hand'
                     THEN VALOR * COSTO ELSE 0 END), 0)                 AS INV_DIN,
        COALESCE(SUM(CASE WHEN MEASURE='Purchase Orders'
                     THEN VALOR ELSE 0 END), 0)                         AS OCT_PZA,
        COALESCE(SUM(CASE WHEN MEASURE='OC_USD'
                     THEN VALOR ELSE 0 END), 0)                         AS OCT_DIN,
        COALESCE(SUM(CASE WHEN MEASURE='Purchase Requisitions'
                     THEN VALOR ELSE 0 END), 0)                         AS OC_PZA,
        COALESCE(SUM(CASE WHEN MEASURE='REQUISICION_USD'
                     THEN VALOR ELSE 0 END), 0)                         AS OC_DIN,
        COALESCE(SUM(CASE WHEN MEASURE='Planned Replenishments by Order Date'
                     THEN VALOR ELSE 0 END), 0)                         AS PED_PZA,
        COALESCE(SUM(CASE WHEN MEASURE='IMPORTE_PEDIDO_USD'
                     THEN VALOR ELSE 0 END), 0)                         AS PED_DIN
    FROM datos
    {wc}
    GROUP BY "{dimension}"
    HAVING (INV_PZA + OCT_PZA + OC_PZA + PED_PZA + INV_DIN + OCT_DIN + OC_DIN + PED_DIN) > 0
    """
    try:
        base = con.execute(q).df()
    except Exception as e:
        return pd.DataFrame()

    base["DIM"] = base["DIM"].astype(str).str.strip()
    base = base[base["DIM"].str.upper() != "NAN"].reset_index(drop=True)
    return base


@st.fragment
def render_cuadro_ejecutivo(df_long, filtros_key, filtros, factor_mon, simbolo_mon, moneda_sel, tc_usd_mxn=17.6367, tc_eur_mxn=20.5505):
    st.markdown("---")

    dim_opts = {"GRUPO":"Grupo","LINEA":"Línea","FAMILIA":"Familia","POTENCIAL":"Potencial","ABC":"ABC","XYZ":"XYZ","ABC/XYZ":"ABC/XYZ"}
    dim_disponibles = {k: v for k, v in dim_opts.items() if k in df_long.columns}

    meses_es_disp = [m for m in MONTHS_ES if m in df_long["MES"].unique()] if "MES" in df_long.columns else MONTHS_ES
    opciones_mes  = ["📅 Anual (Total)"] + meses_es_disp

    # ── Controles: solo Agrupar y Ver mes (sort va a JS) ─────
    ctrl1, ctrl_mes = st.columns([3, 1.5])
    with ctrl1:
        sel_dim = st.radio("Agrupar por:", list(dim_disponibles.keys()),
            format_func=lambda x: dim_disponibles[x], horizontal=True, key="exec_dim_radio")
    with ctrl_mes:
        sel_mes = st.selectbox("📅 Ver mes:", opciones_mes, index=0, key="exec_mes_sel")

    mes_filtro = None if sel_mes == "📅 Anual (Total)" else sel_mes
    df_ce = get_cuadro_ejecutivo(df_long, sel_dim, filtros_key, filtros, mes_filtro=mes_filtro)
    if df_ce.empty:
        st.info("Sin datos para la selección actual."); return

    # Ordenar por defecto por PED_DIN desc (se cambia luego vía JS)
    df_ce = df_ce.sort_values("PED_DIN", ascending=False).reset_index(drop=True)
    tot = {c: df_ce[c].sum() for c in ["INV_PZA","OCT_PZA","OC_PZA","PED_PZA","INV_DIN","OCT_DIN","OC_DIN","PED_DIN"]}

    # ── Factores de conversión ────────────────────────────────
    _fi = {'USD': 1/tc_usd_mxn, 'MXN': 1.0, 'EUR': 1/tc_eur_mxn}.get(moneda_sel, 1.0)
    _fd = factor_mon

    def _fmt(v, factor, with_sym=True):
        if v == 0: return '—'
        n = v * factor
        sym = simbolo_mon if with_sym else ''
        return f"{sym}{int(round(n)):,}"

    def fp(v):  return _fmt(v, 1.0, False)
    def fi(v):  return _fmt(v, _fi)
    def fd(v):  return _fmt(v, _fd)

    max_inv_p = df_ce["INV_PZA"].max() or 1
    max_ped_p = df_ce["PED_PZA"].max() or 1
    max_inv_d = (df_ce["INV_DIN"] * _fi).max() or 1
    max_ped_d = (df_ce["PED_DIN"] * _fd).max() or 1

    def spark(v, maxv, color, factor=1.0):
        pct = min(int(v * factor / maxv * 100), 100) if maxv > 0 and v > 0 else 0
        if pct == 0: return ''
        return (f'<div style="height:3px;border-radius:2px;background:rgba(0,0,0,.06);margin-top:4px;">'
                f'<div style="width:{pct}%;height:100%;border-radius:2px;'
                f'background:linear-gradient(90deg,{color},{color}99);"></div></div>')

    tot_inv_d = tot["INV_DIN"] * _fi or 1
    tot_ped_d = tot["PED_DIN"] * _fd or 1
    dim_label = dim_disponibles.get(sel_dim, sel_dim)

    # ── Construir filas con data-* para sort JS ───────────────
    rows = []
    for i, row in df_ce.iterrows():
        dim_v   = str(row["DIM"]) or "—"
        i_pza   = row["INV_PZA"]; oct_pza = row["OCT_PZA"]
        oc_pza  = row["OC_PZA"];  ped_pza = row["PED_PZA"]
        i_din   = row["INV_DIN"]; oct_din = row["OCT_DIN"]
        oc_din  = row["OC_DIN"];  ped_din = row["PED_DIN"]
        pct_inv = i_din * _fi / tot_inv_d * 100 if tot_inv_d > 0 else 0
        pct_ped = ped_din * _fd / tot_ped_d * 100 if tot_ped_d > 0 else 0
        bg  = "#F8FAFF" if i % 2 == 0 else "#FFFFFF"
        bg2 = "rgba(239,246,255,.6)" if i % 2 == 0 else "rgba(240,249,255,.4)"

        # data-* attrs guardan valores numéricos para el sort JS
        rows.append(f"""
        <tr class="erow" style="background:{bg};"
            data-inv_pza="{i_pza}" data-oct_pza="{oct_pza}"
            data-oc_pza="{oc_pza}"  data-ped_pza="{ped_pza}"
            data-inv_din="{i_din*_fi:.2f}" data-oct_din="{oct_din*_fd:.2f}"
            data-oc_din="{oc_din*_fd:.2f}" data-ped_din="{ped_din*_fd:.2f}"
            data-dim="{dim_v}">
          <td class="edim erow-num">{i+1}</td>
          <td class="edim-name">{dim_v}</td>
          <td class="epza" style="color:#1D4ED8;">{fp(i_pza)}{spark(i_pza,max_inv_p,'#3B82F6')}</td>
          <td class="epza" style="color:#059669;">{fp(oct_pza)}{spark(oct_pza,max_inv_p,'#10B981')}</td>
          <td class="epza" style="color:#7C3AED;">{fp(oc_pza)}{spark(oc_pza,max_inv_p,'#8B5CF6')}</td>
          <td class="epza eped" style="color:#DC2626;">{fp(ped_pza)}{spark(ped_pza,max_ped_p,'#EF4444')}</td>
          <td class="edin" style="background:{bg2};color:#1D4ED8;font-weight:700;">
            {fi(i_din)}{spark(i_din,max_inv_d,'#3B82F6',_fi)}
            <div style="font-size:.55rem;color:#93C5FD;margin-top:1px;">{pct_inv:.1f}%</div>
          </td>
          <td class="edin" style="background:{bg2};color:#059669;">{fd(oct_din)}</td>
          <td class="edin" style="background:{bg2};color:#7C3AED;">{fd(oc_din)}</td>
          <td class="edin eped" style="background:{bg2};color:#DC2626;font-weight:700;">
            {fd(ped_din)}{spark(ped_din,max_ped_d,'#EF4444',_fd)}
            <div style="font-size:.55rem;color:#FCA5A5;margin-top:1px;">{pct_ped:.1f}%</div>
          </td>
        </tr>""")

    rows_html = "".join(rows)

    # Fila totales (fija, no participa en sort)
    rows_html += f"""
    <tr class="etot" id="etot-row">
      <td></td>
      <td class="edim-name" style="color:#92400E;">∑ TOTAL GENERAL</td>
      <td class="epza" style="color:#1D4ED8;font-weight:800;">{fp(tot['INV_PZA'])}</td>
      <td class="epza" style="color:#059669;font-weight:800;">{fp(tot['OCT_PZA'])}</td>
      <td class="epza" style="color:#7C3AED;font-weight:800;">{fp(tot['OC_PZA'])}</td>
      <td class="epza eped" style="color:#DC2626;font-weight:800;">{fp(tot['PED_PZA'])}</td>
      <td class="edin" style="color:#1D4ED8;font-weight:800;">{fi(tot['INV_DIN'])}</td>
      <td class="edin" style="color:#059669;font-weight:800;">{fd(tot['OCT_DIN'])}</td>
      <td class="edin" style="color:#7C3AED;font-weight:800;">{fd(tot['OC_DIN'])}</td>
      <td class="edin eped" style="color:#DC2626;font-weight:800;">{fd(tot['PED_DIN'])}</td>
    </tr>"""

    html = f"""
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
.exec-premium{{font-family:'Inter',sans-serif;}}
.exec-header{{
  background:linear-gradient(135deg,#0F172A 0%,#1E3A5F 50%,#1E40AF 100%);
  border-radius:16px 16px 0 0;padding:1.1rem 1.4rem .9rem;
  display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:.8rem;
}}
.exec-title{{font-size:1.05rem;font-weight:800;color:#fff;letter-spacing:-.3px;}}
.exec-sub{{font-size:.6rem;color:#93C5FD;margin-top:.2rem;letter-spacing:.2px;}}
.exec-kpis{{display:flex;gap:.55rem;flex-wrap:wrap;}}
.exec-kpi{{
  background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12);
  border-radius:10px;padding:.4rem .75rem;text-align:center;min-width:90px;
  backdrop-filter:blur(4px);
}}
.exec-kpi-lbl{{font-size:.52rem;font-weight:700;letter-spacing:.6px;text-transform:uppercase;opacity:.65;color:#fff;}}
.exec-kpi-val{{font-size:.9rem;font-weight:800;font-family:'IBM Plex Mono',monospace;color:#fff;margin-top:.1rem;}}
.exec-kpi.inv .exec-kpi-val{{color:#93C5FD;}}
.exec-kpi.tra .exec-kpi-val{{color:#6EE7B7;}}
.exec-kpi.oc  .exec-kpi-val{{color:#C4B5FD;}}
.exec-kpi.ped .exec-kpi-val{{color:#FCA5A5;}}
.exec-wrap{{overflow-x:auto;border-radius:0 0 16px 16px;
  box-shadow:0 8px 40px rgba(15,23,42,.18);border:1px solid #E2E8F0;border-top:none;}}
.etbl{{border-collapse:collapse;width:100%;min-width:980px;background:#fff;}}
.etbl thead tr.th1{{background:linear-gradient(135deg,#1E3A5F,#2563EB);}}
.etbl thead tr.th2{{background:linear-gradient(135deg,#1E3A5F,#1E40AF);}}
.etbl thead th{{color:#fff;font-size:.6rem;font-weight:700;letter-spacing:.5px;
  padding:8px 10px;border:none;white-space:nowrap;text-transform:uppercase;}}
.etbl thead th.th-grp-pza{{text-align:center;color:#93C5FD;padding:8px 10px 6px;
  border-bottom:1px solid rgba(255,255,255,.1);}}
.etbl thead th.th-grp-din{{text-align:center;color:#6EE7B7;padding:8px 10px 6px;
  border-bottom:1px solid rgba(255,255,255,.1);border-left:2px solid rgba(255,255,255,.15);}}
.etbl thead th.th-num{{text-align:right;cursor:pointer;user-select:none;transition:opacity .15s;}}
.etbl thead th.th-num:hover{{opacity:.75;}}
.etbl thead th.th-dim{{text-align:left;padding-left:14px;cursor:pointer;user-select:none;}}
.etbl thead th.th-dim:hover{{opacity:.75;}}
.etbl thead th.sep{{border-left:2px solid rgba(255,255,255,.15);}}
.sort-active{{background:rgba(255,255,255,.15) !important;border-radius:4px;}}
.sort-arrow{{margin-left:4px;font-size:.7rem;opacity:.9;}}
.erow:hover td{{background:#EFF6FF !important;transition:background .15s;}}
.edim{{font-size:.65rem;color:#CBD5E1;font-weight:600;padding:9px 6px 9px 14px;
  border-bottom:1px solid #F1F5F9;width:28px;text-align:center;}}
.edim-name{{font-size:.76rem;font-weight:700;color:#0F172A;padding:9px 14px;
  border-bottom:1px solid #F1F5F9;white-space:nowrap;border-left:3px solid #3B82F6;}}
.epza{{font-size:.73rem;padding:9px 12px;text-align:right;
  font-family:'IBM Plex Mono',monospace;border-bottom:1px solid #F1F5F9;
  white-space:nowrap;min-width:90px;}}
.epza.eped{{border-right:2px solid #DBEAFE;}}
.edin{{font-size:.73rem;padding:9px 12px;text-align:right;
  font-family:'IBM Plex Mono',monospace;border-bottom:1px solid #F1F5F9;
  white-space:nowrap;min-width:100px;}}
.edin.eped{{border-right:none;}}
.etot td{{background:linear-gradient(90deg,#FFFBEB,#FEF9C3) !important;
  border-top:2px solid #F59E0B !important;font-size:.74rem;
  padding:10px 12px;font-family:'IBM Plex Mono',monospace;
  border-bottom:none !important;}}
.etot .edim-name{{color:#92400E;border-left-color:#F59E0B;}}
.exec-footer{{font-size:.58rem;color:#94A3B8;text-align:right;margin-top:.4rem;padding-right:.2rem;}}
</style>
<div class="exec-premium">
  <div class="exec-header">
    <div>
      <div class="exec-title">📊 Cuadro Ejecutivo &nbsp;<span style="opacity:.5;font-weight:500;font-size:.75rem;">por {dim_label} · {moneda_sel} · {sel_mes}</span></div>
      <div class="exec-sub">On Hand = Inventario × Costo (MXN) &nbsp;·&nbsp; OC Tránsito = Purchase Orders &nbsp;·&nbsp; Orden Compra = Requisitions &nbsp;·&nbsp; Pedido = Planned Replen.</div>
    </div>
    <div class="exec-kpis">
      <div class="exec-kpi inv"><div class="exec-kpi-lbl">📦 On Hand</div><div class="exec-kpi-val">{fi(tot['INV_DIN'])}</div></div>
      <div class="exec-kpi tra"><div class="exec-kpi-lbl">🚢 OC Tránsito</div><div class="exec-kpi-val">{fd(tot['OCT_DIN'])}</div></div>
      <div class="exec-kpi oc"><div class="exec-kpi-lbl">📋 Ord. Compra</div><div class="exec-kpi-val">{fd(tot['OC_DIN'])}</div></div>
      <div class="exec-kpi ped"><div class="exec-kpi-lbl">🛒 Pedido</div><div class="exec-kpi-val">{fd(tot['PED_DIN'])}</div></div>
    </div>
  </div>
  <div class="exec-wrap">
    <table class="etbl" id="exec-tbl">
      <thead>
        <tr class="th1">
          <th rowspan="2" style="width:28px;padding:8px 6px;text-align:center;">#</th>
          <th rowspan="2" class="th-dim" data-key="dim" onclick="sortTable(this)">{dim_label.upper()} <span class="sort-arrow" id="arrow-dim"></span></th>
          <th colspan="4" class="th-grp-pza">📦 &nbsp; PIEZAS</th>
          <th colspan="4" class="th-grp-din sep">💰 &nbsp; DINERO ({moneda_sel})</th>
        </tr>
        <tr class="th2">
          <th class="th-num" data-key="inv_pza" onclick="sortTable(this)" style="color:#BFDBFE;min-width:88px;">ON HAND <span class="sort-arrow" id="arrow-inv_pza"></span></th>
          <th class="th-num" data-key="oct_pza" onclick="sortTable(this)" style="color:#6EE7B7;min-width:88px;">OC TRÁNSITO <span class="sort-arrow" id="arrow-oct_pza"></span></th>
          <th class="th-num" data-key="oc_pza"  onclick="sortTable(this)" style="color:#C4B5FD;min-width:95px;">ORD. COMPRA <span class="sort-arrow" id="arrow-oc_pza"></span></th>
          <th class="th-num" data-key="ped_pza" onclick="sortTable(this)" style="color:#FCA5A5;min-width:88px;border-right:2px solid rgba(255,255,255,.15);">PEDIDO <span class="sort-arrow" id="arrow-ped_pza"></span></th>
          <th class="th-num sep" data-key="inv_din" onclick="sortTable(this)" style="color:#BFDBFE;min-width:100px;background:rgba(37,99,235,.2);">ON HAND <span class="sort-arrow" id="arrow-inv_din"></span></th>
          <th class="th-num" data-key="oct_din" onclick="sortTable(this)" style="color:#6EE7B7;min-width:100px;background:rgba(37,99,235,.2);">OC TRÁNSITO <span class="sort-arrow" id="arrow-oct_din"></span></th>
          <th class="th-num" data-key="oc_din"  onclick="sortTable(this)" style="color:#C4B5FD;min-width:105px;background:rgba(37,99,235,.2);">ORD. COMPRA <span class="sort-arrow" id="arrow-oc_din"></span></th>
          <th class="th-num" data-key="ped_din" onclick="sortTable(this)" style="color:#FCA5A5;min-width:100px;background:rgba(37,99,235,.2);">PEDIDO <span class="sort-arrow" id="arrow-ped_din"></span></th>
        </tr>
      </thead>
      <tbody id="exec-tbody">{rows_html}</tbody>
    </table>
  </div>
  <p class="exec-footer">Agrupado por {sel_dim} &nbsp;·&nbsp; {moneda_sel} &nbsp;·&nbsp; {datetime.now().strftime('%d/%m/%Y %H:%M')} &nbsp;·&nbsp; ⚡ DuckDB · 🔀 Clic en columna para ordenar</p>
</div>
<script>
(function(){{
  var _sortKey = 'ped_din';
  var _asc     = false;

  // Ordenar al cargar (por Pedido $ desc)
  _doSort('ped_din', false);

  function sortTable(th) {{
    var key = th.getAttribute('data-key');
    if (_sortKey === key) {{ _asc = !_asc; }}
    else {{ _sortKey = key; _asc = false; }}
    _doSort(key, _asc);
  }}
  window.sortTable = sortTable;

  function _doSort(key, asc) {{
    var tbody  = document.getElementById('exec-tbody');
    if (!tbody) return;
    var totRow = document.getElementById('etot-row');
    var rows   = Array.from(tbody.querySelectorAll('tr.erow'));

    rows.sort(function(a, b) {{
      var va = a.getAttribute('data-' + key) || '';
      var vb = b.getAttribute('data-' + key) || '';
      // texto vs número
      var na = parseFloat(va);
      var nb = parseFloat(vb);
      if (!isNaN(na) && !isNaN(nb)) {{
        return asc ? na - nb : nb - na;
      }}
      return asc ? va.localeCompare(vb) : vb.localeCompare(va);
    }});

    // Re-numerar y re-insertar
    rows.forEach(function(r, idx) {{
      var numCell = r.querySelector('.erow-num');
      if (numCell) numCell.textContent = idx + 1;
      // alternar colores de fila
      var even = idx % 2 === 0;
      r.style.background = even ? '#F8FAFF' : '#FFFFFF';
      var dinCells = r.querySelectorAll('.edin');
      dinCells.forEach(function(c) {{
        c.style.background = even ? 'rgba(239,246,255,.6)' : 'rgba(240,249,255,.4)';
      }});
      tbody.appendChild(r);
    }});
    if (totRow) tbody.appendChild(totRow);

    // Actualizar flechas en todos los headers
    document.querySelectorAll('.sort-arrow').forEach(function(s) {{ s.textContent = ''; }});
    var arrowEl = document.getElementById('arrow-' + key);
    if (arrowEl) arrowEl.textContent = asc ? ' ▲' : ' ▼';

    // Resaltar columna activa
    document.querySelectorAll('th.sort-active').forEach(function(t) {{ t.classList.remove('sort-active'); }});
    document.querySelectorAll('[data-key="' + key + '"]').forEach(function(t) {{ t.classList.add('sort-active'); }});
  }}
}})();
</script>"""

    import streamlit.components.v1 as components
    components.html(html, height=800, scrolling=True)

    # ── Exportar ──────────────────────────────────────────────
    df_exp = df_ce.rename(columns={
        "DIM": sel_dim,
        "INV_PZA":"ON HAND PIEZAS","OCT_PZA":"OC TRANSITO PIEZAS",
        "OC_PZA":"ORDEN COMPRA PIEZAS","PED_PZA":"PEDIDO PIEZAS",
        "INV_DIN":"ON HAND DINERO","OCT_DIN":"OC TRANSITO DINERO",
        "OC_DIN":"ORDEN COMPRA DINERO","PED_DIN":"PEDIDO DINERO"
    }).copy()
    df_exp["ON HAND DINERO"]       = df_exp["ON HAND DINERO"] * _fi
    df_exp["OC TRANSITO DINERO"]   = df_exp["OC TRANSITO DINERO"] * _fd
    df_exp["ORDEN COMPRA DINERO"]  = df_exp["ORDEN COMPRA DINERO"] * _fd
    df_exp["PEDIDO DINERO"]        = df_exp["PEDIDO DINERO"] * _fd
    buf = io.BytesIO()
    _sheet = f"Ejecutivo_{sel_dim}".replace("/","-")[:31]
    df_exp.to_excel(buf, index=False, sheet_name=_sheet)
    buf.seek(0)
    st.download_button(
        f"⬇️ Exportar Excel ({sel_dim})", buf,
        f"Ejecutivo_{sel_dim}_{moneda_sel}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════
# ════════════════════════════════════════════════════════════
#  TAB 4 — BACKORDER ANALYSIS
# ════════════════════════════════════════════════════════════
def render_backorder_tab(df_long, factor_mon, simbolo_mon, moneda_sel, df_bo=None, df_prom_ventas=None, df_ventas_mes=None, filtros_key=(), filtros=None, tc_usd_mxn=17.6367):
    _df = df_long
    st.markdown("### 📦 Análisis de Backorder")

    if df_bo is None or df_bo.empty:
        st.info("ℹ️ No hay datos de Backorder. Asegúrate de que el archivo DATOS.xlsx tenga una hoja llamada **BO** con columnas **CC** y **BO**.")
        return

    prom_pza_map = {}
    prom_imp_map = {}
    if df_prom_ventas is not None and not df_prom_ventas.empty:
        prom_pza_map = df_prom_ventas.set_index('CC')['PROM_PZA'].to_dict()
        prom_imp_map = df_prom_ventas.set_index('CC')['PROM_IMP'].to_dict()

    # ── Ventas mensuales desde hoja VENTAS (CC, AÑO, MES, PIEZAS, IMPORTES) ─
    _MES_ES = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
               7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'}
    ventas_pza_mes = {}
    ventas_imp_mes = {}
    _meses_orden  = []
    _meses_labels = []
    if df_ventas_mes is not None and not df_ventas_mes.empty:
        _dvm = df_ventas_mes.copy()
        _dvm['CC']       = _dvm['CC'].astype(str).str.strip()
        _dvm['PIEZAS']   = pd.to_numeric(_dvm['PIEZAS'],   errors='coerce').fillna(0)
        _dvm['IMPORTES'] = pd.to_numeric(_dvm['IMPORTES'], errors='coerce').fillna(0)
        _grp = _dvm.groupby(['CC','AÑO','MES'])[['PIEZAS','IMPORTES']].sum().reset_index()
        _meses_uniq  = sorted(set(zip(_grp['AÑO'].tolist(), _grp['MES'].tolist())))
        _meses_orden = _meses_uniq[-6:]
        _meses_labels = [f"{_MES_ES.get(m,str(m))} {str(a)[2:]}" for a,m in _meses_orden]
        for _, _vrow in _grp.iterrows():
            _cc = str(_vrow["CC"])
            _k  = (int(_vrow["AÑO"]), int(_vrow["MES"]))
            if _cc not in ventas_pza_mes: ventas_pza_mes[_cc] = {}
            if _cc not in ventas_imp_mes: ventas_imp_mes[_cc] = {}
            ventas_pza_mes[_cc][_k] = float(_vrow["PIEZAS"])
            ventas_imp_mes[_cc][_k] = float(_vrow["IMPORTES"])

    _mostrar_ventas = st.checkbox(
        "👁 Mostrar ventas por mes (últimos 6 meses)",
        value=False, key="bo_mostrar_ventas"
    )

    # ── Leyenda semáforo ─────────────────────────────────────
    st.markdown(
        '<div style="font-size:.65rem;color:#888;padding:.2rem 0 .5rem;display:flex;flex-wrap:wrap;gap:.4rem;align-items:center;">'
        '<span class="bo-sem sem-verde">✅ DISPONIBLE</span>&nbsp; On Hand ≥ BO &nbsp;·&nbsp;'
        '<span class="bo-sem sem-amarillo">🟡 EN TRÁNSITO</span>&nbsp; Supply total ≥ BO &nbsp;·&nbsp;'
        '<span class="bo-sem sem-amarillo">🟠 PARCIAL</span>&nbsp; Hay supply pero insuficiente &nbsp;·&nbsp;'
        '<span class="bo-sem sem-rojo">🔴 SIN ABASTO</span>&nbsp; Sin OC, Req ni Pedido activo &nbsp;·&nbsp;'
        '<b>🚨</b> = Doble Riesgo (urgente)'
        '</div>',
        unsafe_allow_html=True,
    )

    def get_measure_sum(measure_name):
        sub = _df[_df["MEASURE"] == measure_name]
        if sub.empty: return pd.Series(dtype=float)
        return sub.groupby("ITEM")["VALOR"].sum()

    inv_series = get_measure_sum("On Hand")
    oc_series  = get_measure_sum("Purchase Orders")
    req_series = get_measure_sum("Purchase Requisitions")
    ped_series = get_measure_sum("Planned Replenishments by Order Date")
    dem_series = get_measure_sum("Net Forecast")
    if dem_series.empty: dem_series = get_measure_sum("Total Demand")

    costo_unitario = (
        _df[["ITEM","COSTO"]].drop_duplicates("ITEM").set_index("ITEM")["COSTO"]
    )

    rows = []
    for _, r in df_bo.iterrows():
        cc        = r["CC"]
        bo        = r["BO"]
        costo_exc = r["COSTO"]
        costo_uni = costo_exc if costo_exc > 0 else float(costo_unitario.get(cc, 0))
        inv_oh    = float(inv_series.get(cc, 0))
        oc_val    = float(oc_series.get(cc, 0))
        req_val   = float(req_series.get(cc, 0))
        ped_val   = float(ped_series.get(cc, 0))
        dem_val   = float(dem_series.get(cc, 0))
        supply_total = inv_oh + oc_val + req_val + ped_val
        gap          = supply_total - bo
        pct_fulfil   = min(inv_oh / bo * 100, 100) if bo > 0 else 0
        costo_bo     = bo * costo_uni
        doble_riesgo = (bo > 0 and oc_val == 0 and req_val == 0 and ped_val == 0)

        if inv_oh >= bo:
            semaforo = ("✅ DISPONIBLE",   "sem-verde")
        elif supply_total >= bo:
            semaforo = ("🟡 EN TRÁNSITO",  "sem-amarillo")
        elif oc_val > 0 or req_val > 0 or ped_val > 0:
            semaforo = ("🟠 PARCIAL",      "sem-amarillo")
        else:
            semaforo = ("🔴 SIN ABASTO",   "sem-rojo")

        _vpza = ventas_pza_mes.get(cc, {})
        _vimp = ventas_imp_mes.get(cc, {})
        _rd = {
            "CC": cc,
            "NOMBRE_CORTO":   str(r.get("NOMBRE CORTO", "")),
            "BO":             bo,
            "COSTO_UNI":      costo_uni,
            "COSTO_TOTAL_BO": float(r.get("COSTO TOTAL", 0) or 0),
            "COSTO_BO":       costo_bo,
            "INV_OH":         inv_oh,
            "OC_TRANSITO":    oc_val,
            "REQUISICION":    req_val,
            "PEDIDO_SUG":     ped_val,
            "SUPPLY_TOTAL":   supply_total,
            "GAP":            gap,
            "PCT_FULFIL":     pct_fulfil,
            "SEM_TXT":        semaforo[0],
            "SEM_CLS":        semaforo[1],
            "DOBLE_RIESGO":   doble_riesgo,
            "COMENTARIO":     str(r.get("COMENTARIO", "")),
            "PROM_PZA":       float(prom_pza_map.get(cc, 0) or 0),
            "PROM_IMP":       float(prom_imp_map.get(cc, 0) or 0),
        }
        if _mostrar_ventas:
            for _k in _meses_orden:
                _rd[f"VTA_PZA_{_k}"] = float(_vpza.get(_k, 0))
                _rd[f"VTA_IMP_{_k}"] = float(_vimp.get(_k, 0))
        rows.append(_rd)
    df_ana = pd.DataFrame(rows)

    total_bo_pzas = df_ana["BO"].sum()
    total_bo_usd  = df_ana["COSTO_BO"].sum()
    total_riesgo  = df_ana["DOBLE_RIESGO"].sum()
    total_gap_neg = df_ana[df_ana["GAP"] < 0]["GAP"].sum()

    # factor_mon convierte USD -> moneda seleccionada
    # COSTO_TOTAL_BO y PROM_IMP vienen del Excel en MXN -> hay que convertir
    # _tc_ref = cuántos MXN vale 1 USD (para convertir valores MXN -> moneda seleccionada)
    _tc_ref = tc_usd_mxn if tc_usd_mxn > 0 else 17.6367

    def fmon(v):      return f"{simbolo_mon}{v * factor_mon:,.0f}"            # v en USD
    def fmon_mxn(v):  return f"{simbolo_mon}{v / _tc_ref * factor_mon:,.0f}"  # v en MXN
    def fpz(v):       return f"{v:,.0f}"

    st.markdown(
        f'<div class="bo-kpi-wrap">'
        f'<div class="bo-kpi red"><div class="bo-kpi-lbl">📋 Líneas Backorder</div>'
        f'<div class="bo-kpi-val">{len(df_ana):,}</div>'
        f'<div class="bo-kpi-sub">{total_bo_pzas:,.0f} piezas totales</div></div>'

        f'<div class="bo-kpi amber"><div class="bo-kpi-lbl">💰 Valor en Riesgo</div>'
        f'<div class="bo-kpi-val">{fmon(total_bo_usd)}</div>'
        f'<div class="bo-kpi-sub">BO × Costo unitario ({moneda_sel})</div></div>'

        f'<div class="bo-kpi red"><div class="bo-kpi-lbl">🚨 Doble Riesgo</div>'
        f'<div class="bo-kpi-val">{int(total_riesgo)}</div>'
        f'<div class="bo-kpi-sub">BO sin OC, tránsito ni pedido</div></div>'

        f'<div class="bo-kpi purple"><div class="bo-kpi-lbl">⚠️ GAP Neto (faltante)</div>'
        f'<div class="bo-kpi-val">{abs(total_gap_neg):,.0f}</div>'
        f'<div class="bo-kpi-sub">Piezas que no alcanza a cubrir el supply</div></div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Garantizar columnas aunque df_prom_ventas venga vacío (caché viejo)
    if "PROM_PZA" not in df_ana.columns: df_ana["PROM_PZA"] = 0.0
    if "PROM_IMP" not in df_ana.columns: df_ana["PROM_IMP"] = 0.0
    df_show = df_ana.sort_values("BO", ascending=False).reset_index(drop=True)

    # ── Tabla HTML ────────────────────────────────────────────
    def td(val, cls="td-num", data_val=None):
        dv = f' data-val="{data_val}"' if data_val is not None else ''
        return f'<td class="{cls}"{dv}>{val}</td>'

    # columnas ventas: primero todas las Pzas, luego todos los Importes
    # El botón toggle en JS oculta/muestra las columnas individuales dejando solo el promedio
    _n_meses = len(_meses_orden) if (_mostrar_ventas and _meses_orden) else 0
    # cols: pzas mes1..N, imp mes1..N, prom_pza, prom_imp  = N*2 + 2 total
    # índices de columnas oculables: 2..(2+N-1) para pzas, (2+N)..(2+2N-1) para imp
    _mes_hdrs_after_nombre = ''
    if _mostrar_ventas and _meses_orden:
        # Grupo Pzas
        _mes_hdrs_after_nombre += '<th class="th-mes-grp" colspan="' + str(_n_meses) + '" style="background:#0D47A1;text-align:center;font-size:.62rem;letter-spacing:.5px;">📦 VENTAS PIEZAS <button onclick=\'toggleMeses()\' style=\'background:rgba(255,255,255,.2);border:none;color:#fff;border-radius:4px;padding:1px 6px;cursor:pointer;font-size:.6rem;margin-left:6px;\' id=\'btn-toggle\'>▼ ocultar</button></th>'
        # Grupo Importes
        _mes_hdrs_after_nombre += '<th class="th-mes-grp" colspan="' + str(_n_meses) + '" style="background:#1B5E20;text-align:center;font-size:.62rem;letter-spacing:.5px;">💰 VENTAS IMPORTE</th>'
        # Promedios (siempre visibles)
        _mes_hdrs_after_nombre += '<th class="th-mes-pza" style="background:#0D3B5E;" rowspan="1">Prom Pzas<br><small>6m</small></th>'
        _mes_hdrs_after_nombre += '<th class="th-mes-imp" style="background:#0A2E4A;" rowspan="1">Prom Imp $<br><small>6m</small></th>'

    _mes_subhdrs = ''
    if _mostrar_ventas and _meses_orden:
        _mes_subhdrs  = "".join(f'<th class="th-mes-pza mes-col">{_meses_labels[i]}</th>' for i in range(_n_meses))
        _mes_subhdrs += "".join(f'<th class="th-mes-imp mes-col">{_meses_labels[i]}</th>' for i in range(_n_meses))
        _mes_subhdrs += '<th class="th-mes-pza" style="background:#0D3B5E;"></th><th class="th-mes-imp" style="background:#0A2E4A;"></th>'

    _static_hdrs = (
        '<th class="th-left sortable" data-col="0" onclick="boSort(this)">CC / Item <span class="bo-arrow"></span></th>'
        '<th class="th-left sortable" data-col="1" onclick="boSort(this)">Nombre Corto <span class="bo-arrow"></span></th>'
    )
    _static_subhdrs = (
        '<th class="th-left" style="background:#1A237E;"></th>'
        '<th class="th-left" style="background:#1A237E;"></th>'
    )

    header = (
        '<thead>'
        '<tr>' + _static_hdrs + _mes_hdrs_after_nombre +
        '<th class="sortable" data-col="2" onclick="boSort(this)" rowspan="2">BO<br>(Pzas)</th>'
        '<th class="sortable" data-col="3" onclick="boSort(this)" rowspan="2">Costo<br>Total</th>'
        '<th class="sortable" data-col="6" onclick="boSort(this)" rowspan="2">On Hand</th>'
        '<th class="sortable" data-col="7" onclick="boSort(this)" rowspan="2">OC<br>Transito</th>'
        '<th class="sortable" data-col="8" onclick="boSort(this)" rowspan="2">Req.</th>'
        '<th class="sortable" data-col="9" onclick="boSort(this)" rowspan="2">Pedido<br>Sug.</th>'
        '<th class="sortable" data-col="10" onclick="boSort(this)" rowspan="2">Supply<br>Total</th>'
        '<th class="th-gap sortable" data-col="11" onclick="boSort(this)" rowspan="2">GAP<br>(Supply-BO)</th>'
        '<th class="th-semaforo sortable" data-col="12" onclick="boSort(this)" rowspan="2">Semaforo</th>'
        '</tr>'
        '<tr>' + _static_subhdrs + _mes_subhdrs + '</tr>'
        '</thead>'
    ) if (_mostrar_ventas and _meses_orden) else (
        '<thead><tr>'
        '<th class="th-left sortable" data-col="0" onclick="boSort(this)">CC / Item <span class="bo-arrow"></span></th>'
        '<th class="th-left sortable" data-col="1" onclick="boSort(this)">Nombre Corto <span class="bo-arrow"></span></th>'
        '<th class="sortable" data-col="2" onclick="boSort(this)">BO (Pzas)</th>'
        '<th class="sortable" data-col="3" onclick="boSort(this)">Costo Total</th>'
        '<th class="sortable" data-col="6" onclick="boSort(this)">On Hand</th>'
        '<th class="sortable" data-col="7" onclick="boSort(this)">OC Transito</th>'
        '<th class="sortable" data-col="8" onclick="boSort(this)">Req.</th>'
        '<th class="sortable" data-col="9" onclick="boSort(this)">Pedido Sug.</th>'
        '<th class="sortable" data-col="10" onclick="boSort(this)">Supply Total</th>'
        '<th class="th-gap sortable" data-col="11" onclick="boSort(this)">GAP (Supply-BO)</th>'
        '<th class="th-semaforo sortable" data-col="12" onclick="boSort(this)">Semaforo</th>'
        '</tr></thead>'
    )

    body_rows = ""
    _n_vcols = (_n_meses * 2 + 2) if (_mostrar_ventas and _meses_orden) else 0
    _min_w = 1100 + _n_vcols * 75
    for _, r in df_show.iterrows():
        gap_v   = r["GAP"]
        gap_cls = "td-gap gap-ok" if gap_v >= 0 else "td-gap gap-bad"
        gap_str = f"+{fpz(gap_v)}" if gap_v >= 0 else fpz(gap_v)
        dr_icon = " 🚨" if r["DOBLE_RIESGO"] else ""
        nombre     = r["NOMBRE_CORTO"] if r["NOMBRE_CORTO"] and r["NOMBRE_CORTO"] != "nan" else "—"
        comentario = r["COMENTARIO"]   if r["COMENTARIO"]   and r["COMENTARIO"]   != "nan" else "—"
        costo_tot  = r["COSTO_TOTAL_BO"] if r["COSTO_TOTAL_BO"] > 0 else r["COSTO_BO"]

        _vc = ''
        if _mostrar_ventas and _meses_orden:
            # Primero todas las Pzas, luego todos los Importes
            _vc  = "".join(td(fpz(r.get(f'VTA_PZA_{k}',0)) if r.get(f'VTA_PZA_{k}',0)>0 else '—','td-num mes-col',r.get(f'VTA_PZA_{k}',0)) for k in _meses_orden)
            _vc += "".join(td(fmon_mxn(r.get(f'VTA_IMP_{k}',0)) if r.get(f'VTA_IMP_{k}',0)>0 else '—','td-num mes-col',r.get(f'VTA_IMP_{k}',0)) for k in _meses_orden)
            _pv=[r.get(f'VTA_PZA_{k}',0) for k in _meses_orden if r.get(f'VTA_PZA_{k}',0)>0]
            _iv=[r.get(f'VTA_IMP_{k}',0) for k in _meses_orden if r.get(f'VTA_IMP_{k}',0)>0]
            _ap=sum(_pv)/len(_pv) if _pv else 0
            _ai=sum(_iv)/len(_iv) if _iv else 0
            _vc += td(fpz(_ap) if _ap>0 else '—','td-num td-prom',_ap)+td(fmon_mxn(_ai) if _ai>0 else '—','td-num td-prom',_ai)
        body_rows += (
            f'<tr>'
            f'<td class="td-left" data-val="{r["CC"]}">{r["CC"]}{dr_icon}</td>'
            f'<td class="td-left" data-val="{nombre}" style="max-width:160px;overflow:hidden;text-overflow:ellipsis;" title="{nombre}">{nombre}</td>'
            + _vc
            + td(fpz(r["BO"]), "td-num", r["BO"])
            + td(fmon_mxn(costo_tot), "td-num", costo_tot)
            + td(fpz(r["INV_OH"]), "td-num", r["INV_OH"])
            + td(fpz(r["OC_TRANSITO"]), "td-num", r["OC_TRANSITO"])
            + td(fpz(r["REQUISICION"]), "td-num", r["REQUISICION"])
            + td(fpz(r["PEDIDO_SUG"]), "td-num", r["PEDIDO_SUG"])
            + td(fpz(r["SUPPLY_TOTAL"]), "td-num", r["SUPPLY_TOTAL"])
            + f'<td class="{gap_cls}" data-val="{gap_v}">{gap_str}</td>'
            + f'<td class="td-num" data-val="{r["SEM_TXT"]}"><span class="bo-sem {r["SEM_CLS"]}">{r["SEM_TXT"]}</span></td>'
            + '</tr>\n'
        )

    tot_supply  = df_show["SUPPLY_TOTAL"].sum()
    tot_gap     = df_show["GAP"].sum()
    tot_gap_cls = "td-gap gap-ok" if tot_gap >= 0 else "td-gap gap-bad"
    tot_gap_str = f"+{fpz(tot_gap)}" if tot_gap >= 0 else fpz(tot_gap)

    # Totales de columnas de ventas
    _tot_vc = ''
    if _mostrar_ventas and _meses_orden:
        for _k in _meses_orden:
            _c = f'VTA_PZA_{_k}'
            _s = df_show[_c].sum() if _c in df_show.columns else 0
            _tot_vc += td(fpz(_s) if _s > 0 else '—', 'td-num mes-col', _s)
        for _k in _meses_orden:
            _c = f'VTA_IMP_{_k}'
            _s = df_show[_c].sum() if _c in df_show.columns else 0
            _tot_vc += td(fmon_mxn(_s) if _s > 0 else '—', 'td-num mes-col', _s)
        _ap_t = [df_show[f'VTA_PZA_{k}'].sum() for k in _meses_orden if f'VTA_PZA_{k}' in df_show.columns]
        _ai_t = [df_show[f'VTA_IMP_{k}'].sum() for k in _meses_orden if f'VTA_IMP_{k}' in df_show.columns]
        _tp = sum(_ap_t)/len(_ap_t) if _ap_t else 0
        _ti = sum(_ai_t)/len(_ai_t) if _ai_t else 0
        _tot_vc += td(fpz(_tp) if _tp > 0 else '—', 'td-num td-prom', _tp)
        _tot_vc += td(fmon_mxn(_ti) if _ti > 0 else '—', 'td-num td-prom', _ti)

    tot_row = (
        f'<tr class="bo-row-tot">'
        f'<td class="td-left">∑ TOTAL ({len(df_show)} líneas)</td>'
        + td("—")
        + _tot_vc
        + td(fpz(df_show["BO"].sum()))
        + td(fmon(df_show["COSTO_BO"].sum()))
        + td(fpz(df_show["INV_OH"].sum()))
        + td(fpz(df_show["OC_TRANSITO"].sum()))
        + td(fpz(df_show["REQUISICION"].sum()))
        + td(fpz(df_show["PEDIDO_SUG"].sum()))
        + td(fpz(tot_supply))
        + f'<td class="{tot_gap_cls}">{tot_gap_str}</td>'
        + td("—")
        + '</tr>'
    )
    body_rows = tot_row + body_rows

    # ── Número de filas para calcular altura del iframe ──────
    _n_rows   = len(df_show)
    _row_h    = 32   # px por fila aprox
    _head_h   = 48
    _foot_h   = 40
    _iframe_h = _head_h + _n_rows * _row_h + _foot_h + 30

    bo_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'IBM Plex Sans',sans-serif; background:transparent;}}
.bo-wrap{{overflow-x:auto;border-radius:12px;box-shadow:0 4px 24px rgba(13,27,75,.13);margin-bottom:1rem;}}
.bo-tbl{{border-collapse:collapse;width:100%;min-width:{_min_w}px;background:#fff;font-family:'IBM Plex Sans',sans-serif;}}
.bo-tbl thead tr{{background:linear-gradient(180deg,#1A237E,#162082);}}
.bo-tbl thead th{{color:#fff;font-weight:700;font-size:.68rem;letter-spacing:.4px;padding:10px 8px;text-align:center;border:none;white-space:nowrap;}}
.bo-tbl thead th.th-left{{text-align:left;padding-left:14px;min-width:130px;}}
.bo-tbl thead th.th-semaforo{{background:#7B2C2C;min-width:90px;}}
.bo-tbl thead th.th-gap{{background:#c8960a;color:#2d1e00;}}
.bo-tbl tbody tr:nth-child(even) td{{background:#F8F9FF;}}
.bo-tbl tbody tr:hover td{{background:#EBF3FF !important;}}
.bo-tbl td{{font-size:.74rem;padding:7px 8px;border-bottom:1px solid #E8EAF2;white-space:nowrap;vertical-align:middle;color:#1a1a2e;}}
.bo-tbl td.td-left{{text-align:left;padding-left:14px;font-weight:600;font-size:.73rem;}}
.bo-tbl td.td-num{{text-align:right;padding-right:12px;font-family:'IBM Plex Mono',monospace;font-size:.72rem;}}
.bo-tbl td.td-gap{{text-align:right;padding-right:12px;font-family:'IBM Plex Mono',monospace;font-size:.75rem;font-weight:800;background:#FFF9C4 !important;color:#5D4037 !important;border-left:2px solid #c8960a;}}
.bo-tbl td.td-gap.gap-ok{{background:#E8F5E9 !important;color:#1B5E20 !important;}}
.bo-tbl td.td-gap.gap-bad{{background:#FDEDEC !important;color:#C0392B !important;}}
.bo-sem{{display:inline-block;border-radius:6px;padding:2px 8px;font-size:.66rem;font-weight:700;letter-spacing:.3px;white-space:nowrap;}}
.sem-verde{{background:#D5F5E3;color:#1A5C38;}}
.sem-amarillo{{background:#FEF9E7;color:#9A6700;}}
.sem-rojo{{background:#FADBD8;color:#922B21;}}
.bo-tbl tr.bo-row-tot td{{background:#FFF9C4 !important;color:#5D4037 !important;font-weight:800 !important;border-top:2px solid #c8960a;}}
/* ── SORT ── */
.bo-tbl thead th.sortable{{cursor:pointer;user-select:none;transition:background .15s;}}
.bo-tbl thead th.sortable:hover{{background:rgba(255,255,255,.12);}}
.bo-tbl thead th.sort-asc{{background:rgba(59,130,246,.35) !important;}}
.bo-tbl thead th.sort-desc{{background:rgba(239,68,68,.3) !important;}}
.bo-arrow{{font-size:.7rem;margin-left:4px;}}
.bo-tbl thead th .bo-arrow::after{{content:'⇅';opacity:.5;}}
.bo-tbl thead th.sort-asc  .bo-arrow::after{{content:' ▲';opacity:1;color:#90caf9;}}
.bo-tbl thead th.sort-desc .bo-arrow::after{{content:' ▼';opacity:1;color:#ffb3b3;}}
.bo-tbl thead th.th-mes-pza{{background:#1A5276;min-width:68px;font-size:.60rem;line-height:1.3;text-align:center;}}
.bo-tbl thead th.th-mes-imp{{background:#154360;min-width:78px;font-size:.60rem;line-height:1.3;text-align:center;}}
.bo-tbl thead th.th-mes-grp{{color:#fff;font-weight:700;font-size:.62rem;letter-spacing:.4px;padding:7px 8px;border:none;white-space:nowrap;}}
.td-prom{{background:#EBF5FB !important;font-weight:800 !important;color:#0D3B5E !important;border-left:3px solid #1A5276 !important;font-size:.74rem !important;}}
.mes-col{{transition:all .2s;}}
.mes-col.hidden{{display:none;}}
</style>
</head><body>
<div class="bo-wrap"><table class="bo-tbl">
{header}
<tbody>{body_rows}</tbody>
</table></div>
<script>
(function(){{
  var sortCol = 2;
  var sortAsc = false;

  function boSort(th) {{
    var col = parseInt(th.getAttribute('data-col'));
    if (sortCol === col) {{ sortAsc = !sortAsc; }}
    else {{ sortCol = col; sortAsc = false; }}

    document.querySelectorAll('.bo-tbl thead th.sortable').forEach(function(h){{
      h.classList.remove('sort-asc','sort-desc');
    }});
    th.classList.add(sortAsc ? 'sort-asc' : 'sort-desc');

    var tbody = th.closest('table').querySelector('tbody');
    var rows  = Array.from(tbody.querySelectorAll('tr:not(.bo-row-tot)'));
    var tot   = tbody.querySelector('tr.bo-row-tot');

    rows.sort(function(a, b){{
      var tdA = a.querySelectorAll('td')[col];
      var tdB = b.querySelectorAll('td')[col];
      var vA  = tdA ? (tdA.getAttribute('data-val') || tdA.textContent.trim()) : '';
      var vB  = tdB ? (tdB.getAttribute('data-val') || tdB.textContent.trim()) : '';
      var nA  = parseFloat(String(vA).replace(/[^0-9.+-]/g,''));
      var nB  = parseFloat(String(vB).replace(/[^0-9.+-]/g,''));
      if (!isNaN(nA) && !isNaN(nB)) {{
        return sortAsc ? nA - nB : nB - nA;
      }}
      return sortAsc
        ? String(vA).localeCompare(String(vB))
        : String(vB).localeCompare(String(vA));
    }});

    rows.forEach(function(r){{ tbody.appendChild(r); }});
    if (tot) tbody.insertBefore(tot, tbody.firstChild);
  }}

  // exponer globalmente para onclick
  window.boSort = boSort;

  // Toggle columnas de meses
  var mesVisible = true;
  window.toggleMeses = function() {{
    mesVisible = !mesVisible;
    document.querySelectorAll('.mes-col').forEach(function(c) {{
      c.style.display = mesVisible ? '' : 'none';
    }});
    var btn = document.getElementById('btn-toggle');
    if (btn) btn.textContent = mesVisible ? '\u25bc ocultar' : '\u25b6 ver meses';
  }};

  // auto-sort BO desc al cargar
  var th2 = document.querySelector('th[data-col="2"]');
  if (th2) {{ sortAsc = true; boSort(th2); }}
}})();
</script>
</body></html>"""

    # ── Exportar ──────────────────────────────────────────────
    export_df = df_ana[[
        "CC","NOMBRE_CORTO","BO","COSTO_UNI","COSTO_TOTAL_BO","COSTO_BO",
        "INV_OH","OC_TRANSITO","REQUISICION","PEDIDO_SUG",
        "SUPPLY_TOTAL","GAP","PCT_FULFIL","SEM_TXT","DOBLE_RIESGO","COMENTARIO"
    ]].rename(columns={
        "NOMBRE_CORTO":"NOMBRE CORTO","COSTO_UNI":"COSTO UNIT (USD)",
        "COSTO_TOTAL_BO":"COSTO TOTAL (EXCEL)","COSTO_BO":"VALOR BO CALC (USD)",
        "INV_OH":"ON HAND","OC_TRANSITO":"OC TRÁNSITO","REQUISICION":"REQUISICIÓN",
        "PEDIDO_SUG":"PEDIDO SUGERIDO","SUPPLY_TOTAL":"SUPPLY TOTAL",
        "GAP":"GAP (Supply−BO)","PCT_FULFIL":"% FULFILLMENT",
        "SEM_TXT":"SEMÁFORO","DOBLE_RIESGO":"DOBLE RIESGO",
    })
    buf_bo = io.BytesIO()
    export_df.to_excel(buf_bo, index=False, sheet_name="Backorder_Analisis")
    buf_bo.seek(0)
    col_exp1, col_exp2, _ = st.columns([2, 2, 6])
    with col_exp1:
        st.download_button(
            "⬇️ Exportar Análisis Backorder", buf_bo,
            f"Backorder_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_exp2:
        st.download_button(
            "⬇️ CSV", export_df.to_csv(index=False).encode("utf-8"),
            f"Backorder_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", "text/csv",
        )

    components.html(bo_html, height=_iframe_h, scrolling=True)



def main():
    st.set_page_config(
        page_title="Demand & Replenishment",
        page_icon="📦", layout="wide",
        initial_sidebar_state="collapsed"
    )
    st.markdown(CSS, unsafe_allow_html=True)
    # ── JS: eliminar botón sidebar ──
    import streamlit.components.v1 as _c
    _c.html("""
    <script>
    (function(){
        function kill(){
            document.querySelectorAll(
                '[data-testid="stSidebarCollapsedControl"]'
            ).forEach(function(el){
                el.style.cssText='display:none!important;width:0!important;height:0!important;';
            });
        }
        kill();
        new MutationObserver(kill).observe(document.body,{childList:true,subtree:true});
    })();
    </script>
    """, height=0, scrolling=False)

    # ══════════════════════════════════════════════════════════
    #  SISTEMA DE LOGIN — módulo externo login.py
    # ══════════════════════════════════════════════════════════
    from login import render_login
    if not render_login():
        return  # No continuar si no está logueado


    # ══════════════════════════════════════════════════════════
    #  DASHBOARD PRINCIPAL
    # ══════════════════════════════════════════════════════════
    df_raw, df_bo, df_prom_ventas, df_ventas_mes, fuente = cargar_datos()

    # ── Cachear df y df_long en session_state para no recalcular en cada rerun ──
    _data_hash = str(len(df_raw)) + str(df_raw.columns.tolist())
    if st.session_state.get('_data_hash') != _data_hash:
        with st.spinner("🔢 Calculando MOS..."):
            df = calcular_mos(df_raw)
        with st.spinner("🔄 Preparando datos..."):
            df_long = a_largo(df)
        st.session_state['_data_hash'] = _data_hash
        st.session_state['_df']        = df
        st.session_state['_df_long']   = df_long
    else:
        df      = st.session_state['_df']
        df_long = st.session_state['_df_long']
        fuente  = fuente


    # ══════════════════════════════════════════════════════════
    #  BARRA SUPERIOR — Filtros + Moneda
    # ══════════════════════════════════════════════════════════
    st.markdown("""
    <style>
    .block-container { padding-top: 0.3rem !important; }
    .topbar-wrap {
        background: linear-gradient(90deg, #040e1f 0%, #071830 40%, #0a1f3a 70%, #040e1f 100%);
        border-bottom: 1.5px solid rgba(59,130,246,.25);
        border-radius: 0 0 14px 14px;
        padding: .55rem 1.1rem .45rem 1.1rem;
        margin-bottom: .7rem;
        box-shadow: 0 4px 24px rgba(0,0,0,.45);

    }


    .topbar-wrap .stMultiSelect label,
    .topbar-wrap .stMultiSelect [data-testid="stWidgetLabel"] {
        color: #5B9BD5 !important; font-size: .58rem !important; font-weight: 800 !important;
        text-transform: uppercase !important; letter-spacing: .6px !important;
    }
    .topbar-wrap .stMultiSelect [data-baseweb="select"] > div:first-child {
        background: rgba(7,18,38,.9) !important; border: 1px solid rgba(59,130,246,.28) !important;
        border-radius: 8px !important; min-height: 30px !important; font-size: .68rem !important;
        transition: border-color .2s, box-shadow .2s !important;
    }
    /* ── FILTRO ACTIVO: borde amarillo + glow cuando hay tags seleccionados ── */
    .topbar-wrap .stMultiSelect [data-baseweb="select"]:has([data-baseweb="tag"]) > div:first-child {
        border: 1.5px solid #FACC15 !important;
        box-shadow: 0 0 10px rgba(250,204,21,.35), inset 0 0 6px rgba(250,204,21,.08) !important;
        background: rgba(30,22,5,.95) !important;
    }
    /* Label amarillo cuando hay filtro activo */
    .topbar-wrap .stMultiSelect:has([data-baseweb="tag"]) [data-testid="stWidgetLabel"],
    .topbar-wrap .stMultiSelect:has([data-baseweb="tag"]) label {
        color: #FACC15 !important;
    }
    .topbar-wrap .stMultiSelect [data-baseweb="tag"] {
        background: linear-gradient(135deg,rgba(234,179,8,.4),rgba(161,122,2,.35)) !important;
        color: #FEF08A !important; border-radius: 5px !important; font-size: .62rem !important;
        font-weight: 700 !important; border: 1px solid rgba(250,204,21,.5) !important;
    }
    .topbar-wrap .stRadio [data-testid="stWidgetLabel"] p {
        color: #5B9BD5 !important; font-size:.58rem !important; font-weight:800 !important;
        text-transform:uppercase !important; letter-spacing:.6px !important;
    }
    .topbar-wrap .stRadio label { color:#BAD8FF !important; font-size:.68rem !important; font-weight:700 !important; }
    .topbar-wrap .stNumberInput [data-testid="stWidgetLabel"] {
        color: #5B9BD5 !important; font-size:.58rem !important; font-weight:800 !important;
        text-transform:uppercase !important; letter-spacing:.6px !important;
    }
    .topbar-wrap .stNumberInput input {
        background: rgba(7,18,38,.9) !important; color: #5CE0FF !important;
        border: 1px solid rgba(59,130,246,.28) !important; border-radius: 8px !important;
        font-size: .72rem !important; font-weight: 700 !important;
    }
    .topbar-wrap button[kind="primary"] {
        background: linear-gradient(135deg,#1D4ED8,#3B82F6) !important;
        border: none !important; border-radius: 9px !important;
        font-size: .68rem !important; font-weight: 800 !important;
        box-shadow: 0 3px 12px rgba(59,130,246,.4) !important;
    }
    .topbar-wrap button[kind="secondary"] {
        background: rgba(30,50,90,.6) !important; border: 1px solid rgba(59,130,246,.25) !important;
        border-radius: 9px !important; color: #94A3B8 !important;
        font-size: .68rem !important; font-weight: 700 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="topbar-wrap">', unsafe_allow_html=True)

        def mf_top(col, label, container):
            opts = sorted(df_long[col].dropna().unique().tolist()) if col in df_long.columns else []
            return container.multiselect(label, opts, default=[], key=f"f_{col}")

        # FILA 1: Filtros
        c1, c2, c3, c4, c5, c6, c7 = st.columns([2.0, 2.0, 2.0, 1.6, 2.5, 1.6, 0.85])
        with c1: f_lin = mf_top("LINEA",     "📦  Línea",    c1)
        with c2: f_grp = mf_top("GRUPO",     "📁  Grupo",    c2)
        with c3: f_fam = mf_top("FAMILIA",   "🏷️  Familia",  c3)
        with c4: f_pot = mf_top("POTENCIAL", "⚡  Potencial", c4)
        with c5: f_sup = mf_top("SUPPLIER",  "🏭  Supplier", c5)
        with c6:
            _mes_opts = sorted(df_long["MES"].dropna().unique().tolist()) if "MES" in df_long.columns else []
            _mes_sel_opts = ["📅 TOTAL"] + _mes_opts
            _mes_sel = st.selectbox("📅  Mes", options=_mes_sel_opts, index=0, key="f_MES_sel")
            f_mes = [] if _mes_sel == "📅 TOTAL" else [_mes_sel]
        with c7:
            st.markdown('<div style="height:20px"></div>', unsafe_allow_html=True)
            btn_clear = st.button("✖  Limpiar", use_container_width=True, key="btn_clear")

        # FILA 2: Moneda + TC + Mes + Badge
        st.markdown('<div style="height:4px;border-top:1px solid rgba(59,130,246,.12);margin:.3rem 0 .25rem 0"></div>', unsafe_allow_html=True)
        cm1, cm2, cm3, cm5 = st.columns([2.2, 1.3, 1.3, 6.6])
        with cm1:
            moneda_sel = st.radio(
                "💱  Moneda", options=list(MONEDA_CONFIG.keys()),
                format_func=lambda x: MONEDA_CONFIG[x]['label'],
                index=1, key="moneda_radio", horizontal=True,
            )
        with cm2:
            tc_usd_mxn = st.number_input("💵  USD→MXN", value=17.6367, min_value=0.01, step=0.01, format="%.4f", key="tc_mxn")
        with cm3:
            tc_eur_mxn = st.number_input("💶  EUR→MXN", value=20.5505, min_value=0.01, step=0.01, format="%.4f", key="tc_eur_mxn")
        with cm5:
            _f_draft = {k: v for k, v in {
                'LINEA': f_lin, 'GRUPO': f_grp, 'FAMILIA': f_fam,
                'POTENCIAL': f_pot, 'SUPPLIER': f_sup, 'MES': f_mes
            }.items() if v}
            n_act = len(_f_draft)
            _f_key_draft = tuple(sorted((k, tuple(sorted(v))) for k, v in _f_draft.items()))
            if n_act:
                chips = "".join(
                    f'<span style="background:rgba(37,99,235,.3);border:1px solid rgba(59,130,246,.4);'
                    f'border-radius:6px;padding:.2rem .55rem;font-size:.6rem;color:#93C5FD;font-weight:700;white-space:nowrap;">'
                    f'{"📦" if k=="LINEA" else "📁" if k=="GRUPO" else "🏷️" if k=="FAMILIA" else "⚡" if k=="POTENCIAL" else "🏭" if k=="SUPPLIER" else "📅"}'
                    f' {k}: {", ".join(str(x) for x in v[:2])}{"…" if len(v)>2 else ""}</span>'
                    for k, v in _f_draft.items()
                )
                badge_html = f'<div style="display:flex;align-items:center;flex-wrap:wrap;gap:.35rem;padding-top:1.2rem">{chips}</div>'
            else:
                badge_html = ('<div style="padding-top:1.3rem">'
                    '<span style="background:rgba(100,116,139,.1);border:1px solid rgba(100,116,139,.25);'
                    'border-radius:7px;padding:.25rem .7rem;font-size:.62rem;color:#64748B;font-weight:600;">'
                    '— Sin filtros activos</span></div>')
            st.markdown(badge_html, unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

    # ── Botón Limpiar ──────────────────────────────────────
    if btn_clear:
        for _wk in ['f_LINEA', 'f_GRUPO', 'f_FAMILIA', 'f_POTENCIAL', 'f_SUPPLIER', 'f_MES', 'f_MES_sel']:
            st.session_state.pop(_wk, None)
        for _ck in ['_ck_base', '_tab0_key', '_res_key', '_ventas_key']:
            st.session_state.pop(_ck, None)
        st.rerun()

    # ── Calcular TC y moneda ──────────────────────────────────
    tc_usd_eur  = tc_usd_mxn / tc_eur_mxn if tc_eur_mxn > 0 else 0.86
    tc_map      = {'USD': 1.0, 'MXN': tc_usd_mxn, 'EUR': tc_usd_eur}
    factor_mon  = tc_map[moneda_sel]
    simbolo_mon = MONEDA_CONFIG[moneda_sel]['simbolo']
    if moneda_sel == 'USD':
        badge_txt = "Base — sin conversión"; badge_color = "#60D4F8"
    elif moneda_sel == 'MXN':
        badge_txt = f"1 USD = {tc_usd_mxn:,.4f} MXN"; badge_color = "#4ADE80"
    else:
        badge_txt = f"1 USD = {tc_usd_eur:,.4f} EUR"; badge_color = "#FBBF24"

    # ── Sidebar oculto ──────────────────────────────────────
    st.markdown('''<style>
    [data-testid="stSidebar"],
    [data-testid="stSidebarNav"],
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="stSidebarCollapsedControl"] * {
        display: none !important;
        visibility: hidden !important;
        width: 0 !important;
        height: 0 !important;
        pointer-events: none !important;
    }
</style>''', unsafe_allow_html=True)

    # ── Filtros instantáneos ──
    filtros = _f_draft

    # ── KPIs Premium ──────────────────────────────────────────
    items_total  = int(df_long['ITEM'].nunique())
    filtros_key  = tuple(sorted((k, tuple(sorted(v))) for k,v in filtros.items()))
    hay_filtros  = bool(filtros)

    # ── PRE-FILTRAR df_long UNA SOLA VEZ ─────────────────────
    # Todas las funciones usan df_filt para evitar escaneos duplicados
    _fkey_vacia     = ()
    _filtros_vacios = {}
    _filt_ck = ('df_filt', filtros_key)
    if st.session_state.get('_filt_ck') != _filt_ck or '_df_filt' not in st.session_state:
        if filtros:
            mask = None
            for col, vals in filtros.items():
                if col in df_long.columns and vals and col != 'MES':
                    m = df_long[col].isin(vals)
                    mask = m if mask is None else (mask & m)
            if 'MES' in filtros and filtros['MES']:
                m_mes = df_long['MES'].isin(filtros['MES'])
                mask = m_mes if mask is None else (mask & m_mes)
            df_filt = df_long.loc[mask].reset_index(drop=True) if mask is not None else df_long
        else:
            df_filt = df_long
        st.session_state['_filt_ck']  = _filt_ck
        st.session_state['_df_filt']  = df_filt
        con = get_duckdb_con()
        con.register("datos", df_filt)
        st.session_state['_duckdb_df_id'] = id(df_filt)
    else:
        df_filt = st.session_state['_df_filt']

    # Caché manual por filtros_key — invalida cuando cambian los filtros
    _ck = ('duck_base', filtros_key)
    if st.session_state.get('_ck_base') != _ck:
        st.session_state['_ck_base']    = _ck
        st.session_state['_kpi_base']   = duck_kpis(df_filt, _fkey_vacia, _filtros_vacios)
        st.session_state['_items_ped']  = duck_items_pedido(df_filt, _fkey_vacia, _filtros_vacios)
        st.session_state['_kpi_rich']   = duck_kpis_rich(df_filt, _fkey_vacia, _filtros_vacios)
    kpi_base                       = st.session_state['_kpi_base']
    items_con_pedido, items_activos = st.session_state['_items_ped']
    kpi                            = st.session_state['_kpi_rich']
    items_filtrados                = int(kpi_base['items'])

    def fmt_usd(v, factor=1.0, sim='$'):
        n = v * factor
        return f"{sim}{n:,.0f}"

    def fmt_pzas(v):
        return f"{v:,.0f} pzas"

    sim = MONEDA_CONFIG[moneda_sel]['simbolo']

    ck  = st.columns(6)

    ck[0].markdown(f"""
    <div class="kpi-card kpi-inv">
      <span class="kpi-icon">📦</span>
      <div class="kpi-title">On Hand</div>
      <div class="kpi-currency">{sim} {moneda_sel}</div>
      <div class="kpi-value">{fmt_usd(kpi['inv_mxn'], {'USD': 1/tc_usd_mxn, 'MXN': 1.0, 'EUR': 1/tc_eur_mxn}.get(moneda_sel, 1.0), sim)}</div>
      <div class="kpi-pzas">PIEZAS &nbsp;<span>{fmt_pzas(kpi['inv_pzas'])}</span></div>
    </div>""", unsafe_allow_html=True)

    ck[1].markdown(f"""
    <div class="kpi-card kpi-trans">
      <span class="kpi-icon">🚢</span>
      <div class="kpi-title">OC Tránsito</div>
      <div class="kpi-currency">{sim} {moneda_sel}</div>
      <div class="kpi-value">{fmt_usd(kpi['oct_usd'], factor_mon, sim)}</div>
      <div class="kpi-pzas">PIEZAS &nbsp;<span>{fmt_pzas(kpi['oct_pzas'])}</span></div>
    </div>""", unsafe_allow_html=True)

    ck[2].markdown(f"""
    <div class="kpi-card kpi-trans" style="border-top-color:#2E86C1;">
      <span class="kpi-icon">📋</span>
      <div class="kpi-title">Orden de Compra</div>
      <div class="kpi-currency" style="color:#2E86C1;">{sim} {moneda_sel}</div>
      <div class="kpi-value" style="color:#2E86C1;">{fmt_usd(kpi['oc_usd'], factor_mon, sim)}</div>
      <div class="kpi-pzas">PIEZAS &nbsp;<span>{fmt_pzas(kpi['oc_pzas'])}</span></div>
    </div>""", unsafe_allow_html=True)

    ck[3].markdown(f"""
    <div class="kpi-card kpi-pedido" style="border-top-color:#E67E22;">
      <span class="kpi-icon">📅</span>
      <div class="kpi-title">Pedido Marzo</div>
      <div class="kpi-currency" style="color:#E67E22;">{sim} {moneda_sel}</div>
      <div class="kpi-value" style="color:#E67E22;">{fmt_usd(kpi['pedido_marzo_usd'], factor_mon, sim)}</div>
      <div class="kpi-pzas">PIEZAS &nbsp;<span>{fmt_pzas(kpi['pedido_marzo_pzas'])}</span></div>
    </div>""", unsafe_allow_html=True)

    ck[4].markdown(f"""
    <div class="kpi-card kpi-pedido">
      <span class="kpi-icon">🛒</span>
      <div class="kpi-title">Pedido Proyectado</div>
      <div class="kpi-currency">{sim} {moneda_sel}</div>
      <div class="kpi-value">{fmt_usd(kpi['pedido_usd'], factor_mon, sim)}</div>
      <div class="kpi-pzas">PIEZAS &nbsp;<span>{fmt_pzas(kpi['pedido_pzas'])}</span></div>
    </div>""", unsafe_allow_html=True)



    mos_str  = f"{kpi['mos']} m"  if kpi['mos']  is not None else "—"
    vuel_str = f"≈ {kpi['vuel']} vueltas / año" if kpi['vuel'] is not None else ""
    ck[5].markdown(f"""
    <div class="kpi-card kpi-cob">
      <span class="kpi-icon">📊</span>
      <div class="kpi-title">Cobertura Promedio</div>
      <div class="kpi-currency" style="color:#7B2FBE;">— —</div>
      <div class="kpi-value">{mos_str}</div>
      <div class="kpi-pzas">{vuel_str}</div>
    </div>""", unsafe_allow_html=True)


    # ── Tabs con persistencia via query_params ───────────────
    _MAIN_TABS = ["📋 Resumen General", "🔎 Detalle"]
    _mtab_idx = 0
    if "main_tab" in st.query_params:
        try:
            _mtab_idx = int(st.query_params["main_tab"])
        except Exception:
            pass
    _mtab_sel = st.radio(
        "Sección",
        options=[0, 1],
        format_func=lambda i: _MAIN_TABS[i],
        index=_mtab_idx,
        horizontal=True,
        key="main_tab_radio",
        label_visibility="collapsed",
    )
    st.query_params["main_tab"] = str(_mtab_sel)
    st.markdown("<hr style='margin:.4rem 0 1rem'>", unsafe_allow_html=True)

    # ── TAB 0 — RESUMEN GENERAL ───────────────────────────────
    if _mtab_sel == 0:
        st.markdown("### 📋 Demand & Replenishment Control — 2026")
        tc_label = f"1 USD = {factor_mon:,.4f} {moneda_sel}" if moneda_sel != 'USD' else "Base USD"
        st.markdown(
            f'<div class="moneda-activa">{MONEDA_CONFIG[moneda_sel]["label"]}'
            f'<span class="tc">· {tc_label}</span></div>',
            unsafe_allow_html=True
        )
        filtros_label = " | ".join(
            f"{k}: {', '.join(str(x) for x in v)}" for k, v in filtros.items() if v
        ) or "Sin filtros"

        # ── Cache pesado por filtros_key (pivot, ventas, landed cost) ──────
        _tab0_key = ('tab0_data', filtros_key)
        if st.session_state.get('_tab0_key') != _tab0_key:
            st.session_state['_tab0_key']    = _tab0_key
            st.session_state['_pivot']       = get_pivot_resumen(df, filtros_key, filtros)
            st.session_state['_lc_monthly']  = calcular_landed_cost_mensual(df_filt, filtros_key, filtros)
            st.session_state['_cl_monthly']  = calcular_costos_logisticos_mensual(df_filt, filtros_key, filtros)
        pivot               = st.session_state['_pivot']
        landed_cost_monthly = st.session_state['_lc_monthly']
        costos_log_monthly  = st.session_state['_cl_monthly']

        # ── Cuadros de VENTAS y PRESUPUESTO ──────────────────
        render_ventas_presupuesto(
            df_long     = df_filt,
            filtros_key = filtros_key,
            filtros     = filtros,
            factor_mon  = factor_mon,
            simbolo_mon = simbolo_mon,
            moneda_sel  = moneda_sel,
            tc_usd_mxn  = tc_usd_mxn,
        )

        # ── Cache HTML del resumen en session_state para que cambiar moneda no recalcule ──
        _res_key = ('html_resumen', filtros_key, moneda_sel, factor_mon)
        if st.session_state.get('_res_key') != _res_key:
            st.session_state['_res_key']     = _res_key
            st.session_state['_res_html']    = render_resumen(
                pivot, filtros_label, moneda_sel, factor_mon, simbolo_mon,
                landed_cost_monthly=landed_cost_monthly,
                costos_log_monthly=costos_log_monthly)
        # Altura exacta: 1 fila header (40px) + N métricas (31px) + 1 sep (5px) por sección + footer (22px)
        _altura_resumen = (
            sum(40 + len(s["metricas"]) * 31 + 5 for s in SECTIONS)
            + 22  # footer
        )
        components.html(st.session_state['_res_html'], height=_altura_resumen, scrolling=False)

        render_cuadro_ejecutivo(
            df_long     = df_filt,
            filtros_key = _fkey_vacia,
            filtros     = _filtros_vacios,
            factor_mon  = factor_mon,
            simbolo_mon = simbolo_mon,
            moneda_sel  = moneda_sel,
            tc_usd_mxn  = tc_usd_mxn,
            tc_eur_mxn  = tc_eur_mxn,
        )

    # ── TAB 1 — DETALLE ──────────────────────────────────────
    if _mtab_sel == 1:
        render_detalle_tab(
            df_long        = df_filt,
            con            = make_con(df_filt),
            factor_mon     = factor_mon,
            simbolo_mon    = simbolo_mon,
            moneda_sel     = moneda_sel,
            tc_usd_mxn     = tc_usd_mxn,
            tc_eur_mxn     = tc_eur_mxn,
            filtros        = filtros,
            df_bo          = df_bo,
            df_prom_ventas = df_prom_ventas,
            df_ventas_mes  = df_ventas_mes,
        )
    # ── Footer ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown(
        f'<p style="text-align:center;color:#bbb;font-size:.67rem;">'
        f'Dashboard Inventario 2026 &nbsp;·&nbsp; DuckDB {duckdb.__version__} &nbsp;·&nbsp; '
        f'{len(df_filt):,} filas &nbsp;·&nbsp; {moneda_sel} &nbsp;·&nbsp; '
        f'{datetime.now().strftime("%d/%m/%Y %H:%M")}</p>',
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()