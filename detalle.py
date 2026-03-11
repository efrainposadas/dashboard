# ============================================================
#  DETALLE.PY  — Pestaña de Análisis Ejecutivo
#  Bloque 1 : 💰 Pedido por Dimensión
#  Bloque 2 : 📦 Dónde Está el Dinero (Inventario)
#  Bloque 3 : 📊 Meses de Inventario (MOS)
#  VERSION   : 2026-03-07 FIX-INVENTARIO
# ============================================================

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ── Paleta corporativa ────────────────────────────────────────
COLORES = [
    "#1A6FA8","#2E7D53","#7B2C2C","#E67E22","#6A1B9A",
    "#1565C0","#00695C","#C62828","#F57F17","#283593",
]

DIMS = {
    "GRUPO":    "Grupo",
    "LINEA":    "Línea",
    "FAMILIA":  "Familia",
    "POTENCIAL":"Potencial",
    "ABC":      "ABC",
}

# ── Helpers de formato ────────────────────────────────────────
def _fmt_mon(v, factor, simbolo):
    n = v * factor
    return f"{simbolo}{n:,.0f}"

def _fmt_pzas(v):
    return f"{v:,.0f} pzas"

# ── Query helper ──────────────────────────────────────────────
def _q_group(con, dimension, filtros, extra_where=""):
    """Agrega las principales métricas por dimensión en una sola pasada DuckDB."""
    where = ["ANIO = 2026"]
    for col, vals in (filtros or {}).items():
        if vals:
            lista = ", ".join(f"'{v}'" for v in vals)
            where.append(f'"{col}" IN ({lista})')
    if extra_where:
        where.append(extra_where)
    wc = "WHERE " + " AND ".join(where)

    return con.execute(f"""
        SELECT
            "{dimension}"                                                                      AS DIM,
            COUNT(DISTINCT CASE WHEN MEASURE='Planned Replenishments by Order Date' AND VALOR > 0 THEN ITEM END) AS N_ITEMS,
            COALESCE(SUM(CASE WHEN MEASURE='On Hand' AND MES='Marzo'
                         THEN VALOR       ELSE 0 END),0)                                      AS INV_PZA,
            COALESCE(SUM(CASE WHEN MEASURE='On Hand' AND MES='Marzo'
                         THEN VALOR*COSTO ELSE 0 END),0)                                      AS INV_USD,
            COALESCE(SUM(CASE WHEN MEASURE='IMPORTE_PEDIDO_USD'   THEN VALOR ELSE 0 END),0)   AS PED_USD,
            COALESCE(SUM(CASE WHEN MEASURE='Planned Replenishments by Order Date'
                         THEN VALOR ELSE 0 END),0)                                            AS PED_PZA,
            COALESCE(SUM(CASE WHEN MEASURE='OC_USD'               THEN VALOR ELSE 0 END),0)   AS OCT_USD,
            COALESCE(SUM(CASE WHEN MEASURE='Purchase Orders'      THEN VALOR ELSE 0 END),0)   AS OCT_PZA,
            COALESCE(SUM(CASE WHEN MEASURE='Net Forecast'         THEN VALOR ELSE 0 END),0)   AS FC_PZA,
            COALESCE(SUM(CASE WHEN MEASURE='Net Forecast'         THEN VALOR ELSE 0 END),0)   AS DEM_PZA
        FROM datos
        {wc}
        GROUP BY "{dimension}"
        HAVING (INV_PZA + PED_USD + OCT_USD) > 0
        ORDER BY PED_USD DESC
    """).df()


def _q_detail(con, dimension, dim_val, filtros):
    """Detalle de ítems para un valor de dimensión dado."""
    where = ["ANIO = 2026", f'"{dimension}" = \'{dim_val}\'']
    for col, vals in (filtros or {}).items():
        if vals:
            lista = ", ".join(f"'{v}'" for v in vals)
            where.append(f'"{col}" IN ({lista})')
    wc = "WHERE " + " AND ".join(where)

    return con.execute(f"""
        SELECT
            ITEM                                                                                          AS CC,
            MAX(NOMBRE_CORTO)                                                                             AS NOMBRE,
            MAX(GRUPO)      AS GRUPO,
            MAX(LINEA)      AS LINEA,
            MAX(ABC)        AS ABC,
            -- ── On Hand ──
            COALESCE(SUM(CASE WHEN MEASURE='On Hand' AND MES='Marzo' THEN VALOR*COSTO ELSE 0 END),0)     AS INV_USD,
            COALESCE(SUM(CASE WHEN MEASURE='On Hand' AND MES='Marzo' THEN VALOR       ELSE 0 END),0)     AS INV_PZA,
            -- ── Componentes PAB ──
            COALESCE(SUM(CASE WHEN MEASURE='Purchase Orders'        AND MES='Marzo' THEN VALOR ELSE 0 END),0) AS OCT_PZA,
            COALESCE(SUM(CASE WHEN MEASURE='Purchase Requisitions'  AND MES='Marzo' THEN VALOR ELSE 0 END),0) AS REQ_PZA,
            COALESCE(SUM(CASE WHEN MEASURE='Planned Replenishments by Order Date' AND MES='Marzo' THEN VALOR ELSE 0 END),0) AS PLAN_PZA,
            COALESCE(SUM(CASE WHEN MEASURE='Sales Orders'           AND MES='Marzo' THEN VALOR ELSE 0 END),0) AS SO_PZA,
            -- ── Pedido sugerido ──
            COALESCE(SUM(CASE WHEN MEASURE='IMPORTE_PEDIDO_USD'     THEN VALOR ELSE 0 END),0)             AS PED_USD,
            COALESCE(SUM(CASE WHEN MEASURE='Planned Replenishments by Order Date' THEN VALOR ELSE 0 END),0) AS PED_PZA,
            -- ── OC Tránsito $ ──
            COALESCE(SUM(CASE WHEN MEASURE='OC_USD'                 THEN VALOR ELSE 0 END),0)             AS OCT_USD,
            -- ── Forecast anual ──
            COALESCE(SUM(CASE WHEN MEASURE='Net Forecast'           THEN VALOR ELSE 0 END),0)             AS FC_PZA,
            MAX(COSTO)                                                                                     AS COSTO_UNI,
            -- ── Lead Time ──
            COALESCE(MAX(LT_TOTAL_DIAS), 150)                                                              AS LT_DIAS
        FROM datos
        {wc}
        GROUP BY ITEM
        HAVING (INV_USD + PED_USD + OCT_USD) > 0
        ORDER BY INV_USD DESC
    """).df()


def _q_mos(con, dimension, filtros, n_meses=10):
    """
    MOS por dimensión usando PAB real (no solo On Hand):
      PAB = On Hand + OC Tránsito + Requisiciones + Planned Replen - Sales Orders
      MOS = PAB / (Net Forecast promedio mensual)
    Semáforo DINÁMICO por LT_TOTAL_DIAS de cada ítem:
      🔴 Crítico → MOS < LT_meses
      ✅ Normal  → LT_meses ≤ MOS ≤ LT_meses + 2
      ⚠️ Exceso  → MOS > LT_meses + 2
    """
    where = ["ANIO = 2026"]
    for col, vals in (filtros or {}).items():
        if vals:
            lista = ", ".join(f"'{v}'" for v in vals)
            where.append(f'"{col}" IN ({lista})')
    wc = "WHERE " + " AND ".join(where)

    return con.execute(f"""
        WITH base AS (
            SELECT "{dimension}" AS DIM, ITEM,
                -- ── PAB = todo el pipeline de abastecimiento ──
                COALESCE(SUM(CASE WHEN MEASURE='On Hand'                             AND MES='Marzo' THEN VALOR ELSE 0 END),0)
                + COALESCE(SUM(CASE WHEN MEASURE='Purchase Orders'                  AND MES='Marzo' THEN VALOR ELSE 0 END),0)
                + COALESCE(SUM(CASE WHEN MEASURE='Purchase Requisitions'            AND MES='Marzo' THEN VALOR ELSE 0 END),0)
                + COALESCE(SUM(CASE WHEN MEASURE='Planned Replenishments by Order Date' AND MES='Marzo' THEN VALOR ELSE 0 END),0)
                - COALESCE(SUM(CASE WHEN MEASURE='Sales Orders'                     AND MES='Marzo' THEN VALOR ELSE 0 END),0)
                                                                                               AS pab,
                -- ── Forecast promedio mensual ──
                COALESCE(SUM(CASE WHEN MEASURE='Net Forecast' THEN VALOR ELSE 0 END),0) / {n_meses}.0  AS fc_mes,
                -- ── Valor On Hand en MXN ──
                COALESCE(SUM(CASE WHEN MEASURE='On Hand' AND MES='Marzo' THEN VALOR*COSTO ELSE 0 END),0) AS inv_usd,
                -- ── Lead Time en meses (default 5m si no hay dato) ──
                COALESCE(MAX(LT_TOTAL_DIAS), 150) / 30.0                                       AS lt_meses
            FROM datos {wc}
            GROUP BY "{dimension}", ITEM
        )
        SELECT DIM,
               COUNT(ITEM)                                                                          AS N_ITEMS,
               ROUND(AVG(CASE WHEN fc_mes>0 THEN pab/fc_mes ELSE NULL END),1)                      AS MOS_PROM,
               SUM(inv_usd)                                                                         AS INV_USD,
               -- ── Semáforo dinámico por LT ──
               COUNT(CASE WHEN fc_mes>0 AND pab/fc_mes < lt_meses                     THEN 1 END)  AS CRITICOS,
               COUNT(CASE WHEN fc_mes>0 AND pab/fc_mes >= lt_meses
                                       AND pab/fc_mes <= lt_meses + 2                 THEN 1 END)  AS NORMALES,
               COUNT(CASE WHEN fc_mes>0 AND pab/fc_mes > lt_meses + 2                 THEN 1 END)  AS EXCESO,
               COUNT(CASE WHEN fc_mes=0  OR fc_mes IS NULL                             THEN 1 END)  AS SIN_FC
        FROM base
        GROUP BY DIM
        ORDER BY INV_USD DESC
    """).df()


# ═══════════════════════════════════════════════════════════════
#  BLOQUE 1 — PEDIDO POR DIMENSIÓN
# ═══════════════════════════════════════════════════════════════
def _bloque_pedido(con, df_long, dimension, filtros, factor_mon, simbolo_mon, moneda_sel, tc_usd_mxn):
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#020c1e 0%,#0d2248 50%,#071630 100%);
    border-radius:16px;padding:1.1rem 1.6rem;margin-bottom:1.2rem;
    border:1px solid rgba(59,130,246,.2);position:relative;overflow:hidden;">
    <div style="position:absolute;top:0;left:0;right:0;height:2px;
    background:linear-gradient(90deg,transparent,#1d4ed8,#60a5fa,transparent);"></div>
    <div style="position:absolute;top:-40px;right:-30px;width:180px;height:180px;
    background:radial-gradient(circle,rgba(59,130,246,.1) 0%,transparent 70%);pointer-events:none;"></div>
    <div style="font-family:'Syne',sans-serif;font-size:1.1rem;font-weight:800;color:#dbeafe;letter-spacing:.4px;">
    💰 PEDIDO SUGERIDO — Análisis por Dimensión</div>
    <div style="font-size:.7rem;color:#64748b;margin-top:.15rem;">
    Planned Replenishments by Order Date · Anual 2026 · Agrupar: <b style="color:#93c5fd;">{DIMS.get(dimension,dimension)}</b></div>
    </div>
    """, unsafe_allow_html=True)

    df = _q_group(con, dimension, filtros)
    df = df[df["DIM"].notna() & (df["DIM"].astype(str).str.lower() != "nan")].copy()
    if df.empty:
        st.info("Sin datos de pedido para la selección actual.")
        return

    tot_ped  = df["PED_USD"].sum()
    tot_pza  = df["PED_PZA"].sum()
    tot_oct  = df["OCT_USD"].sum()
    tot_items = df["N_ITEMS"].sum()

    # ── Gráfico + Tabla ejecutiva lado a lado ─────────────────────────────
    df["PCT"]     = df["PED_USD"] / tot_ped * 100 if tot_ped > 0 else 0
    df["PCT_PZA"] = df["PED_PZA"] / tot_pza  * 100 if tot_pza  > 0 else 0
    df["PED_MON"] = df["PED_USD"] * factor_mon
    # Mayor a menor
    df_plot = df.sort_values("PED_MON", ascending=True)
    df_sort = df.sort_values("PED_USD", ascending=False)

    col_graf, col_tbl = st.columns([1.0, 1], gap='small')

    with col_graf:
        n = len(df_plot)
        if n == 0:
            st.info("Sin datos")
        else:
            df_plot2 = df_plot.sort_values("PED_MON", ascending=True).reset_index(drop=True)

            # Paleta degradada azul oscuro → azul claro
            n_bars = len(df_plot2)
            colors = [
                f"rgb({int(13 + 56*(i/max(n_bars-1,1)))},"
                f"{int(31 + 99*(i/max(n_bars-1,1)))},"
                f"{int(60 + 120*(i/max(n_bars-1,1)))})"
                for i in range(n_bars)
            ]

            # Hover con número completo
            hover_texts = [
                f"<b>{row['DIM']}</b><br>"
                f"{simbolo_mon}{row['PED_MON'] * factor_mon:,.0f}<br>"
                f"{row['PCT']:.1f}%"
                for _, row in df_plot2.iterrows()
            ]

            fig = go.Figure()
            fig.add_trace(go.Bar(
                y=df_plot2["DIM"],
                x=df_plot2["PED_MON"] * factor_mon,
                orientation='h',
                marker=dict(color=colors, line=dict(width=0)),
                text=[f'{row["PCT"]:.1f}%' for _, row in df_plot2.iterrows()],
                textposition='outside',
                textfont=dict(size=11, color='#334155', family='IBM Plex Sans'),
                hovertemplate='%{customdata}<extra></extra>',
                customdata=hover_texts,
                showlegend=False,
            ))

            dim_label_g = DIMS.get(dimension, dimension)
            fig.update_layout(
                title=dict(
                    text=f"Pedido Sugerido por {dim_label_g}",
                    font=dict(size=13, color='#1e293b', family='IBM Plex Sans'),
                    x=0, xanchor='left',
                ),
                height=max(320, n_bars * 54 + 80),
                margin=dict(l=0, r=60, t=40, b=10),
                plot_bgcolor='white',
                paper_bgcolor='white',
                showlegend=False,
                bargap=0.28,
                xaxis=dict(
                    showgrid=True, gridcolor='#f0f4f8',
                    zeroline=False, showticklabels=True,
                    tickfont=dict(size=10, color='#94a3b8', family='IBM Plex Sans'),
                    tickformat=',.0f',
                    tickprefix=simbolo_mon,
                ),
                yaxis=dict(
                    showgrid=False,
                    tickfont=dict(size=11, color='#1e293b', family='IBM Plex Sans'),
                    zeroline=False,
                ),
                font=dict(family='IBM Plex Sans', size=11),
                hoverlabel=dict(bgcolor='#1e293b', font_color='white', font_size=12),
            )
            st.plotly_chart(fig, use_container_width=True, key=f"graf_ped_{dimension}")

    with col_tbl:
        dim_label   = DIMS.get(dimension, dimension)
        tot_items_n = int(df["N_ITEMS"].sum())
        tot_pza_n   = df["PED_PZA"].sum()
        tot_ped_mon = df["PED_USD"].sum() * factor_mon

        filas_html = ""
        for _, row in df_sort.iterrows():
            pct       = row["PCT"]
            pct_pza   = row["PCT_PZA"]
            pct_items = row["N_ITEMS"]/tot_items_n*100 if tot_items_n else 0
            # Barra mini para Pedido
            pbar_ped = (f'<div style="height:3px;border-radius:2px;'
                        f'background:linear-gradient(90deg,#1A6FA8,#60a5fa);'
                        f'width:{min(pct,100):.0f}%;margin-top:2px;"></div>')
            # Barra mini para Pzas
            pbar_pza = (f'<div style="height:3px;border-radius:2px;'
                        f'background:linear-gradient(90deg,#059669,#34d399);'
                        f'width:{min(pct_pza,100):.0f}%;margin-top:2px;"></div>')
            filas_html += (
                f'<tr>'
                f'<td style="font-weight:600;color:#1e293b;padding:.4rem .55rem;text-align:left;">{row["DIM"]}</td>'
                f'<td style="text-align:left;color:#475569;padding:.4rem .55rem;">{int(row["N_ITEMS"]):,}'
                f'<br><span style="font-size:.57rem;color:#94a3b8;">{pct_items:.1f}%</span></td>'
                f'<td style="text-align:left;font-weight:700;color:#1A6FA8;padding:.4rem .55rem;">'
                f'{_fmt_mon(row["PED_USD"], factor_mon, simbolo_mon)}'
                f'<br><span style="font-size:.57rem;color:#64748b;">{pct:.1f}%</span>{pbar_ped}</td>'
                f'<td style="text-align:left;color:#374151;padding:.4rem .55rem;">{_fmt_pzas(row["PED_PZA"])}'
                f'<br><span style="font-size:.57rem;color:#64748b;">{pct_pza:.1f}%</span>{pbar_pza}</td>'
                f'</tr>'
            )
        totales_html = (
            f'<tr style="background:linear-gradient(90deg,#0f172a,#1e3a5f);border-top:2px solid #1A6FA8;">'
            f'<td style="font-weight:800;color:#f1f5f9;padding:.5rem .55rem;text-align:left;">TOTAL</td>'
            f'<td style="text-align:left;font-weight:800;color:#93c5fd;padding:.5rem .55rem;">{tot_items_n:,}'
            f'<br><span style="font-size:.57rem;color:#60a5fa;">100%</span></td>'
            f'<td style="text-align:left;font-weight:800;color:#60a5fa;padding:.5rem .55rem;">{_fmt_mon(tot_ped_mon,1,simbolo_mon)}'
            f'<br><span style="font-size:.57rem;color:#93c5fd;">100%</span></td>'
            f'<td style="text-align:left;font-weight:800;color:#93c5fd;padding:.5rem .55rem;">{_fmt_pzas(tot_pza_n)}'
            f'<br><span style="font-size:.57rem;color:#93c5fd;">100%</span></td>'
            f'</tr>'
        )
        tabla_html = (
            '<style>'
            '.exec-tbl{width:100%;border-collapse:collapse;font-family:"IBM Plex Sans",sans-serif;'
            'font-size:.72rem;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.08);}'
            '.exec-tbl thead tr{background:linear-gradient(90deg,#0d1f3c,#1a3a6e);}'
            '.exec-tbl thead th{color:#93c5fd;font-size:.58rem;font-weight:800;text-transform:uppercase;'
            'letter-spacing:.5px;padding:.5rem .55rem;text-align:left;}'
            
            '.exec-tbl tbody tr{background:#fff;border-bottom:1px solid #f1f5f9;transition:background .12s;}'
            '.exec-tbl tbody tr:nth-child(even){background:#f8faff;}'
            '.exec-tbl tbody tr:hover{background:#eff6ff;}'
            '</style>'
            f'<table class="exec-tbl"><thead><tr>'
            f'<th style="text-align:left;">{dim_label}</th>'
            f'<th>Ítems</th><th>Pedido</th><th>Pzas</th>'
            f'</tr></thead><tbody>{filas_html}{totales_html}</tbody></table>'
        )
        st.markdown(tabla_html, unsafe_allow_html=True)

        # ── Detalle por ítem ──────────────────────────────────────
    st.markdown("---")
    st.markdown(f"#### 🔍 Detalle de Ítems — selecciona un {DIMS.get(dimension,'Grupo')}")
    grupos_lista = sorted([x for x in df["DIM"].unique().tolist() if x is not None and str(x).lower() != "nan"])
    dims_disponibles = ["— Seleccionar —", "⭐ TODOS"] + grupos_lista

    sel_key = f"det_ped_sel_{dimension}"
    sel = st.selectbox(f"Ver detalle de:", dims_disponibles, key=sel_key)

    if sel and sel not in ("— Seleccionar —",):
        if sel == "⭐ TODOS":
            # Traer todos los grupos concatenados
            frames = []
            for g in grupos_lista:
                df_g = _q_detail(con, dimension, g, filtros)
                if not df_g.empty:
                    frames.append(df_g)
            df_det = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
            titulo_render = f"TODOS — {DIMS.get(dimension, dimension)}"
        else:
            df_det = _q_detail(con, dimension, sel, filtros)
            titulo_render = sel

        if df_det.empty:
            st.info("Sin ítems para esta selección.")
        else:
            _render_detalle_items(df_det, titulo_render, factor_mon, simbolo_mon, tc_usd_mxn, modo="pedido")


# ═══════════════════════════════════════════════════════════════
#  BLOQUE 2 — DÓNDE ESTÁ EL DINERO (INVENTARIO)
# ═══════════════════════════════════════════════════════════════
def _bloque_inventario(con, df_long, dimension, filtros, factor_mon, simbolo_mon, moneda_sel, tc_usd_mxn):
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#021a0e 0%,#064e2a 50%,#021a0e 100%);
    border-radius:16px;padding:1.1rem 1.6rem;margin-bottom:1.2rem;
    border:1px solid rgba(16,185,129,.2);position:relative;overflow:hidden;">
    <div style="position:absolute;top:0;left:0;right:0;height:2px;
    background:linear-gradient(90deg,transparent,#059669,#6ee7b7,transparent);"></div>
    <div style="position:absolute;top:-40px;right:-30px;width:180px;height:180px;
    background:radial-gradient(circle,rgba(16,185,129,.1) 0%,transparent 70%);pointer-events:none;"></div>
    <div style="font-family:'Syne',sans-serif;font-size:1.1rem;font-weight:800;color:#d1fae5;letter-spacing:.4px;">
    📦 ¿DÓNDE ESTÁ EL DINERO? — Inventario On Hand</div>
    <div style="font-size:.7rem;color:#64748b;margin-top:.15rem;">
    On Hand × Costo Unitario · Valor en libros · Agrupar: <b style="color:#6ee7b7;">{DIMS.get(dimension,dimension)}</b></div>
    </div>
    """, unsafe_allow_html=True)

    df = _q_group(con, dimension, filtros)
    df = df[df["DIM"].notna() & (df["DIM"].astype(str).str.lower() != "nan")].copy()
    if df.empty:
        st.info("Sin datos de inventario para la selección actual.")
        return

    # INV_USD viene del SQL como VALOR*COSTO → resultado en MXN (costo está en MXN)
    # Normalizar a USD real dividiendo entre tc_usd_mxn para que factor_mon funcione correctamente
    _tc = tc_usd_mxn if tc_usd_mxn and tc_usd_mxn > 0 else 17.6367
    df["INV_USD"] = df["INV_USD"] / _tc
    tot_inv_usd  = df["INV_USD"].sum()
    tot_inv_pzas = df["INV_PZA"].sum()

    # ── KPIs ─────────────────────────────────────────────────
    k1, k2, k3 = st.columns(3)
    kpi_style = (
        "background:linear-gradient(135deg,#f0fdf8,#ecfdf5);"
        "border-radius:14px;padding:.85rem 1.1rem;"
        "border-left:5px solid {c};"
        "box-shadow:0 4px 20px rgba(0,0,0,.07),0 1px 3px rgba(0,0,0,.05);"
    )
    k1.markdown(f'<div style="{kpi_style.format(c="#2E7D53")}">'
                f'<div style="font-size:.6rem;color:#64748b;font-weight:700;text-transform:uppercase;">Inventario Total</div>'
                f'<div style="font-size:1.3rem;font-weight:800;color:#2E7D53;">{_fmt_mon(tot_inv_usd, factor_mon, simbolo_mon)}</div>'
                f'<div style="font-size:.65rem;color:#888;">{_fmt_pzas(tot_inv_pzas)}</div></div>', unsafe_allow_html=True)
    top1 = df.iloc[0] if not df.empty else None
    if top1 is not None:
        pct1 = top1["INV_USD"] / tot_inv_usd * 100 if tot_inv_usd > 0 else 0
        k2.markdown(f'<div style="{kpi_style.format(c="#1A6FA8")}">'
                    f'<div style="font-size:.6rem;color:#64748b;font-weight:700;text-transform:uppercase;">Mayor Concentración</div>'
                    f'<div style="font-size:1.1rem;font-weight:800;color:#1A6FA8;">{top1["DIM"]}</div>'
                    f'<div style="font-size:.65rem;color:#888;">{pct1:.1f}% del total · {_fmt_mon(top1["INV_USD"], factor_mon, simbolo_mon)}</div></div>', unsafe_allow_html=True)
    # Pareto 80%
    df_s = df.sort_values("INV_USD", ascending=False).copy()
    df_s["CUM_PCT"] = df_s["INV_USD"].cumsum() / tot_inv_usd * 100 if tot_inv_usd > 0 else 0
    n80 = (df_s["CUM_PCT"] <= 80).sum() + 1
    k3.markdown(f'<div style="{kpi_style.format(c="#7B2C2C")}">'
                f'<div style="font-size:.6rem;color:#64748b;font-weight:700;text-transform:uppercase;">Regla 80/20</div>'
                f'<div style="font-size:1.3rem;font-weight:800;color:#7B2C2C;">{n80} de {len(df)}</div>'
                f'<div style="font-size:.65rem;color:#888;">{DIMS.get(dimension,"grupos")} concentran el 80% del inventario</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='height:.5rem'></div>", unsafe_allow_html=True)

    col_dona, col_bar = st.columns([1, 1.3])

    with col_dona:
        df_pie = df.sort_values("INV_USD", ascending=False).head(8).copy()
        otros  = df[~df["DIM"].isin(df_pie["DIM"])]["INV_USD"].sum()
        if otros > 0:
            df_pie = pd.concat([df_pie, pd.DataFrame([{"DIM":"Otros","INV_USD":otros}])], ignore_index=True)
        df_pie["INV_MON"] = df_pie["INV_USD"] * factor_mon

        fig_pie = go.Figure(go.Pie(
            labels=df_pie["DIM"],
            values=df_pie["INV_MON"],
            hole=0.52,
            marker_colors=COLORES[:len(df_pie)],
            textinfo='label+percent',
            textfont_size=11,
            hovertemplate='<b>%{label}</b><br>%{customdata}<br>%{percent}<extra></extra>',
            customdata=[_fmt_mon(v, 1, simbolo_mon) for v in df_pie["INV_MON"]],
        ))
        fig_pie.add_annotation(
            text=f"<b>{_fmt_mon(tot_inv_usd, factor_mon, simbolo_mon)}</b>",
            x=0.5, y=0.5, font_size=14, showarrow=False,
            font_color="#1a1a2e",
        )
        fig_pie.update_layout(
            height=320, margin=dict(l=0,r=0,t=30,b=0),
            showlegend=False, paper_bgcolor='white',
            font=dict(family='IBM Plex Sans'),
            title=dict(text="Distribución del Inventario", font_size=13, x=0.5),
        )
        st.plotly_chart(fig_pie, use_container_width=True, key=f"dona_inv_{dimension}")

    with col_bar:
        df_bar = df.sort_values("INV_USD", ascending=True).copy()
        df_bar["INV_MON"] = df_bar["INV_USD"] * factor_mon
        df_bar["PCT"] = df_bar["INV_USD"] / tot_inv_usd * 100 if tot_inv_usd > 0 else 0

        fig_bar = go.Figure(go.Bar(
            y=df_bar["DIM"], x=df_bar["INV_MON"],
            orientation='h',
            marker=dict(
                color=df_bar["PCT"],
                colorscale=[[0,"#d4edda"],[0.5,"#1A6FA8"],[1,"#0d2b5e"]],
                showscale=False,
            ),
            text=[f'{p:.1f}%' for p in df_bar["PCT"]],
            textposition='outside',
            textfont_size=10,
            hovertemplate='<b>%{y}</b><br>%{customdata}<br>%{text} del total<extra></extra>',
            customdata=[_fmt_mon(v, 1, simbolo_mon) for v in df_bar["INV_MON"]],
        ))
        fig_bar.update_layout(
            height=max(280, len(df_bar)*46),
            margin=dict(l=0,r=60,t=30,b=20),
            plot_bgcolor='white', paper_bgcolor='white',
            xaxis=dict(showgrid=True, gridcolor='#f0f0f0', zeroline=False),
            yaxis=dict(showgrid=False),
            font=dict(family='IBM Plex Sans', size=11),
            title=dict(text="Valor On Hand por grupo", font_size=13),
            hoverlabel=dict(bgcolor='#1a1a2e', font_color='white'),
        )
        st.plotly_chart(fig_bar, use_container_width=True, key=f"bar_inv_{dimension}")

    # ── Detalle ───────────────────────────────────────────────
    st.markdown("---")
    st.markdown(f"#### 🔍 Detalle de Ítems — selecciona un {DIMS.get(dimension,'Grupo')}")
    dims_disponibles = ["— Seleccionar —", "⭐ TODOS"] + sorted([x for x in df["DIM"].unique().tolist() if x is not None and str(x).lower() != "nan"])
    sel = st.selectbox("Ver detalle de:", dims_disponibles, key=f"det_inv_sel_{dimension}")

    if sel and sel != "— Seleccionar —":
        if sel == "⭐ TODOS":
            _dims_inv = sorted([x for x in df["DIM"].unique().tolist() if x is not None and str(x).lower() != "nan"])
            frames_inv = [_q_detail(con, dimension, g, filtros) for g in _dims_inv]
            df_det = pd.concat([f for f in frames_inv if not f.empty], ignore_index=True) if frames_inv else pd.DataFrame()
            titulo_inv = f"TODOS ({DIMS.get(dimension, dimension)})"
        else:
            df_det = _q_detail(con, dimension, sel, filtros)
            titulo_inv = sel
        if df_det.empty:
            st.info("Sin ítems para esta selección.")
        else:
            _render_detalle_items(df_det, titulo_inv, factor_mon, simbolo_mon, tc_usd_mxn, modo="inventario")


# ═══════════════════════════════════════════════════════════════
#  BLOQUE 3 — MESES DE INVENTARIO (MOS)  [REHECHO]
# ═══════════════════════════════════════════════════════════════

N_MESES_FC = 10  # meses de forecast en el horizonte

def _q_mos_items(con, filtros):
    """
    Devuelve un DataFrame con un ítem por fila, con todos los
    componentes del PAB y los 4 MOS calculados en SQL.
    Fuente 100%: hoja PEDIDO (tabla `datos` en DuckDB).
    """
    where = ["ANIO = 2026"]
    for col, vals in (filtros or {}).items():
        if vals:
            lista = ", ".join(f"'{v}'" for v in vals)
            where.append(f'"{col}" IN ({lista})')
    wc = "WHERE " + " AND ".join(where)

    return con.execute(f"""
        WITH base AS (
            SELECT
                ITEM,
                MAX(GRUPO)         AS GRUPO,
                MAX(LINEA)         AS LINEA,
                MAX(FAMILIA)       AS FAMILIA,
                MAX(POTENCIAL)     AS POTENCIAL,
                MAX(ABC)           AS ABC,
                MAX(NOMBRE_CORTO)  AS NOMBRE,
                MAX(COSTO)         AS COSTO,
                COALESCE(MAX(LT_TOTAL_DIAS), 150)  AS LT_DIAS,
                -- ── Componentes ────────────────────────────────────
                COALESCE(SUM(CASE WHEN MEASURE='On Hand'
                                   AND MES='Marzo'  THEN VALOR ELSE 0 END),0)  AS OH,
                COALESCE(SUM(CASE WHEN MEASURE='Purchase Orders'
                                   AND MES='Marzo'  THEN VALOR ELSE 0 END),0)  AS OCT,
                COALESCE(SUM(CASE WHEN MEASURE='Purchase Requisitions'
                                   AND MES='Marzo'  THEN VALOR ELSE 0 END),0)  AS REQ,
                COALESCE(SUM(CASE WHEN MEASURE='Planned Replenishments by Order Date'
                                   AND MES='Marzo'  THEN VALOR ELSE 0 END),0)  AS PLAN,
                COALESCE(SUM(CASE WHEN MEASURE='Sales Orders'
                                   AND MES='Marzo'  THEN VALOR ELSE 0 END),0)  AS SO,
                -- ── Forecast promedio mensual ───────────────────────
                COALESCE(SUM(CASE WHEN MEASURE='Net Forecast'
                                   THEN VALOR ELSE 0 END),0) / {N_MESES_FC}.0 AS FC_MES,
                -- ── Valor $ On Hand ─────────────────────────────────
                COALESCE(SUM(CASE WHEN MEASURE='On Hand'
                                   AND MES='Marzo'
                                   THEN VALOR*COSTO ELSE 0 END),0)             AS INV_MXN
            FROM datos {wc}
            GROUP BY ITEM
        ),
        calcs AS (
            SELECT *,
                GREATEST(OH + OCT + REQ + PLAN - SO, 0)  AS PAB,
                -- 4 MOS
                CASE WHEN FC_MES>0 THEN ROUND(OH/FC_MES,1)                           ELSE NULL END AS MOS_OH,
                CASE WHEN FC_MES>0 THEN ROUND((OH+OCT)/FC_MES,1)                     ELSE NULL END AS MOS_CONF,
                CASE WHEN FC_MES>0 THEN ROUND((OH+OCT+PLAN)/FC_MES,1)                ELSE NULL END AS MOS_PLAN,
                CASE WHEN FC_MES>0 THEN ROUND(GREATEST(OH+OCT+REQ+PLAN-SO,0)/FC_MES,1) ELSE NULL END AS MOS_PAB,
                LT_DIAS / 30.0  AS LT_MESES
            FROM base
        )
        SELECT * FROM calcs
        WHERE (OH > 0 OR OCT > 0 OR PLAN > 0 OR FC_MES > 0)
        ORDER BY INV_MXN DESC
    """).df()


def _q_mos_resumen(con, dimension, filtros):
    """Agrupado por dimensión: MOS promedio + semáforos dinámico y fijo."""
    where = ["ANIO = 2026"]
    for col, vals in (filtros or {}).items():
        if vals:
            lista = ", ".join(f"'{v}'" for v in vals)
            where.append(f'"{col}" IN ({lista})')
    wc = "WHERE " + " AND ".join(where)

    return con.execute(f"""
        WITH base AS (
            SELECT
                "{dimension}"  AS DIM,
                ITEM,
                COALESCE(SUM(CASE WHEN MEASURE='On Hand'
                                   AND MES='Marzo' THEN VALOR ELSE 0 END),0)           AS OH,
                COALESCE(SUM(CASE WHEN MEASURE='Purchase Orders'
                                   AND MES='Marzo' THEN VALOR ELSE 0 END),0)           AS OCT,
                COALESCE(SUM(CASE WHEN MEASURE='Purchase Requisitions'
                                   AND MES='Marzo' THEN VALOR ELSE 0 END),0)           AS REQ,
                COALESCE(SUM(CASE WHEN MEASURE='Planned Replenishments by Order Date'
                                   AND MES='Marzo' THEN VALOR ELSE 0 END),0)           AS PLAN,
                COALESCE(SUM(CASE WHEN MEASURE='Sales Orders'
                                   AND MES='Marzo' THEN VALOR ELSE 0 END),0)           AS SO,
                COALESCE(SUM(CASE WHEN MEASURE='Net Forecast'
                                   THEN VALOR ELSE 0 END),0) / {N_MESES_FC}.0         AS FC_MES,
                COALESCE(SUM(CASE WHEN MEASURE='On Hand' AND MES='Marzo'
                                   THEN VALOR*COSTO ELSE 0 END),0)                     AS INV_MXN,
                COALESCE(MAX(LT_TOTAL_DIAS), 150) / 30.0                               AS LT_MESES
            FROM datos {wc}
            GROUP BY "{dimension}", ITEM
        ),
        calcs AS (
            SELECT *,
                GREATEST(OH+OCT+REQ+PLAN-SO, 0) AS PAB,
                CASE WHEN FC_MES>0 THEN ROUND(GREATEST(OH+OCT+REQ+PLAN-SO,0)/FC_MES,1) ELSE NULL END AS MOS_PAB
            FROM base
        )
        SELECT
            DIM,
            COUNT(ITEM)                                                              AS N_ITEMS,
            ROUND(AVG(MOS_PAB), 1)                                                   AS MOS_PROM,
            SUM(INV_MXN)                                                             AS INV_MXN,
            -- Semáforo DINÁMICO (LT por ítem)
            COUNT(CASE WHEN MOS_PAB IS NOT NULL AND MOS_PAB < LT_MESES            THEN 1 END) AS DIN_CRITICO,
            COUNT(CASE WHEN MOS_PAB IS NOT NULL AND MOS_PAB >= LT_MESES
                                               AND MOS_PAB <= LT_MESES+2          THEN 1 END) AS DIN_NORMAL,
            COUNT(CASE WHEN MOS_PAB IS NOT NULL AND MOS_PAB > LT_MESES+2
                                               AND MOS_PAB <= LT_MESES+6          THEN 1 END) AS DIN_ALTO,
            COUNT(CASE WHEN MOS_PAB IS NOT NULL AND MOS_PAB > LT_MESES+6          THEN 1 END) AS DIN_EXCESO,
            -- Semáforo FIJO
            COUNT(CASE WHEN MOS_PAB IS NOT NULL AND MOS_PAB < 3                   THEN 1 END) AS FIJ_CRITICO,
            COUNT(CASE WHEN MOS_PAB IS NOT NULL AND MOS_PAB BETWEEN 3 AND 6       THEN 1 END) AS FIJ_NORMAL,
            COUNT(CASE WHEN MOS_PAB IS NOT NULL AND MOS_PAB > 6                   THEN 1 END) AS FIJ_EXCESO,
            COUNT(CASE WHEN MOS_PAB IS NULL OR FC_MES=0                           THEN 1 END) AS SIN_FC
        FROM calcs
        GROUP BY DIM
        ORDER BY INV_MXN DESC
    """).df()


# ── helpers semáforo ────────────────────────────────────────────
def _pill_mos(v, lt_m=None, modo="dinamico"):
    """Devuelve (html_pill, sort_val)."""
    if v is None or (hasattr(v, '__float__') and v != v):  # nan
        return '<span style="color:#94a3b8;font-size:.65rem;">—</span>', -1
    v = float(v)
    if modo == "dinamico" and lt_m is not None:
        lt_m = float(lt_m)
        if v < lt_m:
            cls = "background:linear-gradient(90deg,#dc2626,#b91c1c);color:#fff;"
        elif v <= lt_m + 2:
            cls = "background:linear-gradient(90deg,#d1fae5,#a7f3d0);color:#065f46;"
        elif v <= lt_m + 6:
            cls = "background:linear-gradient(90deg,#fef3c7,#fde68a);color:#92400e;"
        else:
            cls = "background:linear-gradient(90deg,#7c2d12,#9a3412);color:#fff;"
    else:  # fijo
        if v < 3:
            cls = "background:linear-gradient(90deg,#dc2626,#b91c1c);color:#fff;"
        elif v <= 6:
            cls = "background:linear-gradient(90deg,#d1fae5,#a7f3d0);color:#065f46;"
        elif v <= 9:
            cls = "background:linear-gradient(90deg,#fef3c7,#fde68a);color:#92400e;"
        else:
            cls = "background:linear-gradient(90deg,#7c2d12,#9a3412);color:#fff;"
    html = (f'<span style="{cls}border-radius:6px;padding:2px 7px;'
            f'font-weight:800;font-size:.68rem;font-family:\'IBM Plex Mono\',monospace;">'
            f'{v:.1f}m</span>')
    return html, v


def _clf_pill(mos_pab, lt_m):
    """Clasificación de riesgo basada en MOS PAB dinámico."""
    if mos_pab is None or (hasattr(mos_pab, '__float__') and mos_pab != mos_pab):
        return '<span style="color:#94a3b8;font-size:.64rem;">Sin FC</span>', -999
    v, lt = float(mos_pab), float(lt_m)
    if v < lt:
        return ('<span style="background:#dc2626;color:#fff;border-radius:6px;'
                'padding:2px 8px;font-weight:800;font-size:.63rem;">🚨 CRÍTICO</span>'), 0
    if v <= lt + 2:
        return ('<span style="background:#16a34a;color:#fff;border-radius:6px;'
                'padding:2px 8px;font-weight:800;font-size:.63rem;">✅ SALUDABLE</span>'), 1
    if v <= lt + 6:
        return ('<span style="background:#d97706;color:#fff;border-radius:6px;'
                'padding:2px 8px;font-weight:800;font-size:.63rem;">🟡 ALTO</span>'), 2
    return ('<span style="background:#7c2d12;color:#fff;border-radius:6px;'
            'padding:2px 8px;font-weight:800;font-size:.63rem;">⚠️ EXCESO</span>'), 3


_MOS_TAB_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&family=Syne:wght@700;800&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'IBM Plex Sans',sans-serif;background:#f8fafc;}
/* ── Tabla ── */
.mt{border-collapse:collapse;width:100%;background:#fff;}
.mt thead tr{background:linear-gradient(135deg,#0d1b3e,#1a3a6e);}
.mt thead th{
  color:#cbd5e1;font-weight:700;font-size:.61rem;letter-spacing:.4px;
  padding:9px 7px;text-align:center;white-space:nowrap;
  cursor:pointer;user-select:none;border-right:1px solid rgba(255,255,255,.05);
}
.mt thead th.thl{text-align:left;padding-left:11px;}
.mt thead th:last-child{border-right:none;}
.mt thead th:hover{background:rgba(255,255,255,.08);}
.mt thead th.sort-asc{background:rgba(59,130,246,.3)!important;}
.mt thead th.sort-desc{background:rgba(239,68,68,.2)!important;}
.mt thead th .arr::after{content:'⇅';opacity:.3;margin-left:3px;font-size:.58rem;}
.mt thead th.sort-asc  .arr::after{content:' ▲';opacity:1;color:#93c5fd;}
.mt thead th.sort-desc .arr::after{content:' ▼';opacity:1;color:#fca5a5;}
/* Grupos de cabecera */
.mt thead th.hpab{background:rgba(16,185,129,.18)!important;color:#6ee7b7!important;}
.mt thead th.hdin{background:rgba(251,191,36,.18)!important;color:#fcd34d!important;}
.mt thead th.hfij{background:rgba(99,102,241,.2)!important;color:#c7d2fe!important;}
.mt thead th.hrsk{background:rgba(139,92,246,.2)!important;color:#e9d5ff!important;}
/* Celdas */
.mt tbody tr:nth-child(even) td{background:#f8faff;}
.mt tbody tr:hover td{background:linear-gradient(90deg,#eff6ff,#f0f9ff)!important;}
.mt td{font-size:.7rem;padding:5px 7px;border-bottom:1px solid #eef2f8;
        white-space:nowrap;vertical-align:middle;color:#1e293b;}
.mt td.tl{text-align:left;padding-left:11px;font-weight:700;color:#0f172a;}
.mt td.tr{text-align:right;padding-right:9px;font-family:'IBM Plex Mono',monospace;font-size:.68rem;}
.mt td.cpab{background:rgba(16,185,129,.04);}
.mt td.cdin{background:rgba(251,191,36,.04);}
.mt td.cfij{background:rgba(99,102,241,.04);}
/* Fila total */
.mt tr.tot td{background:linear-gradient(90deg,#1e3a5f,#1a3a6e)!important;
              color:#e2e8f0!important;font-weight:800!important;
              border-top:2px solid #3b82f6;font-family:'IBM Plex Mono',monospace;}
.mt tr.tot td.tl{color:#93c5fd!important;}
</style>"""

_MOS_TAB_JS = """
<script>
(function(){
  var sc=0,sa=false;
  window.msSort=function(th){
    var c=parseInt(th.getAttribute('data-col'));
    if(sc===c){sa=!sa;}else{sc=c;sa=false;}
    document.querySelectorAll('.mt thead th').forEach(function(h){h.classList.remove('sort-asc','sort-desc');});
    th.classList.add(sa?'sort-asc':'sort-desc');
    var tb=th.closest('table').querySelector('tbody');
    var rows=Array.from(tb.querySelectorAll('tr:not(.tot)'));
    var tot=tb.querySelector('tr.tot');
    rows.sort(function(a,b){
      var A=a.querySelectorAll('td')[c],B=b.querySelectorAll('td')[c];
      var vA=A?(A.getAttribute('data-val')||A.textContent.trim()):'';
      var vB=B?(B.getAttribute('data-val')||B.textContent.trim()):'';
      var nA=parseFloat(String(vA).replace(/[^0-9.-]/g,''));
      var nB=parseFloat(String(vB).replace(/[^0-9.-]/g,''));
      if(!isNaN(nA)&&!isNaN(nB)) return sa?nA-nB:nB-nA;
      return sa?String(vA).localeCompare(String(vB)):String(vB).localeCompare(String(vA));
    });
    rows.forEach(function(r){tb.appendChild(r);});
    if(tot)tb.appendChild(tot);
  };
})();
</script>"""


def _render_mos_items(df, titulo, factor_mon, simbolo_mon, tc_usd_mxn):
    """Tabla de ítems con 4 MOS + semáforo dinámico + fijo."""
    df = df.copy()
    tc = tc_usd_mxn if tc_usd_mxn and tc_usd_mxn > 0 else 17.6367
    df["INV_USD"] = df["INV_MXN"] / tc

    n   = len(df)
    rh  = 34
    hh  = 56   # header 2 filas aprox
    h   = max(200, 32 + hh + n * rh + 40 + 20)

    cols = [
        ("num","",    "#"),
        ("1", "thl",  "CC / Ítem"),
        ("2", "thl",  "Nombre"),
        ("3", "",     "LT<br>(días)"),
        # PAB
        ("4", "hpab", "On Hand<br>(pzas)"),
        ("5", "hpab", "OC Tránsito<br>(pzas)"),
        ("6", "hpab", "Requisición<br>(pzas)"),
        ("7", "hpab", "Planned<br>(pzas)"),
        ("8", "hpab", "Sales Orders<br>(pzas)"),
        ("9", "hpab", "PAB<br>(pzas)"),
        ("10","hpab", "FC / mes<br>(pzas)"),
        # 4 MOS dinámico
        ("11","hdin", "MOS On Hand<br>(dinámico)"),
        ("12","hdin", "MOS Confirmado<br>(dinámico)"),
        ("13","hdin", "MOS Planeado<br>(dinámico)"),
        ("14","hdin", "MOS PAB<br>(dinámico)"),
        # 4 MOS fijo
        ("15","hfij", "MOS On Hand<br>(fijo)"),
        ("16","hfij", "MOS Confirmado<br>(fijo)"),
        ("17","hfij", "MOS Planeado<br>(fijo)"),
        ("18","hfij", "MOS PAB<br>(fijo)"),
        # Clasificación + $
        ("19","hrsk", "Clasificación<br>(dinámica)"),
        ("20","",     f"Inv. On Hand<br>({simbolo_mon})"),
    ]

    thead = "<thead><tr>"
    for idx, cls, lbl in cols:
        if idx == "num":
            thead += f'<th class="{cls}" style="width:28px;">{lbl}</th>'
        else:
            thead += (f'<th class="{cls} sortable" data-col="{idx}" onclick="msSort(this)">'
                      f'{lbl} <span class="arr"></span></th>')
    thead += "</tr></thead>"

    body = ""
    tot_oh = tot_oct = tot_plan = tot_pab = tot_fc = tot_inv = 0.0
    rn = 0

    for _, r in df.iterrows():
        rn += 1
        oh   = float(r.get("OH",   0) or 0)
        oct  = float(r.get("OCT",  0) or 0)
        req  = float(r.get("REQ",  0) or 0)
        plan = float(r.get("PLAN", 0) or 0)
        so   = float(r.get("SO",   0) or 0)
        pab  = float(r.get("PAB",  0) or 0)
        fc   = float(r.get("FC_MES", 0) or 0)
        lt_d = float(r.get("LT_DIAS", 150) or 150)
        lt_m = lt_d / 30.0
        inv  = float(r.get("INV_USD", 0) or 0) * factor_mon
        nombre = str(r.get("NOMBRE","") or "—")
        if nombre in ("","nan","None"): nombre = "—"

        # 4 valores MOS
        def mv(num): return round(num/fc, 1) if fc > 0 else None
        m_oh   = mv(oh)
        m_conf = mv(oh + oct)
        m_plan = mv(oh + oct + plan)
        m_pab  = mv(pab)

        # Pills dinámicos
        p_oh_d,  s_oh_d  = _pill_mos(m_oh,   lt_m, "dinamico")
        p_cf_d,  s_cf_d  = _pill_mos(m_conf, lt_m, "dinamico")
        p_pl_d,  s_pl_d  = _pill_mos(m_plan, lt_m, "dinamico")
        p_pb_d,  s_pb_d  = _pill_mos(m_pab,  lt_m, "dinamico")
        # Pills fijos
        p_oh_f,  s_oh_f  = _pill_mos(m_oh,   lt_m, "fijo")
        p_cf_f,  s_cf_f  = _pill_mos(m_conf, lt_m, "fijo")
        p_pl_f,  s_pl_f  = _pill_mos(m_plan, lt_m, "fijo")
        p_pb_f,  s_pb_f  = _pill_mos(m_pab,  lt_m, "fijo")

        clf_html, clf_sort = _clf_pill(m_pab, lt_m)

        tot_oh  += oh; tot_oct += oct; tot_plan += plan
        tot_pab += pab; tot_fc += fc; tot_inv += inv

        body += (
            f'<tr>'
            f'<td class="tr" data-val="{rn}" style="color:#94a3b8;font-size:.6rem;">{rn}</td>'
            f'<td class="tl" data-val="{r["ITEM"]}">{r["ITEM"]}</td>'
            f'<td class="tl" data-val="{nombre}" style="max-width:155px;overflow:hidden;text-overflow:ellipsis;" title="{nombre}">{nombre}</td>'
            f'<td class="tr" data-val="{lt_d:.0f}" style="color:#64748b;">{lt_d:.0f}d</td>'
            # PAB
            f'<td class="tr cpab" data-val="{oh:.0f}">{oh:,.0f}</td>'
            f'<td class="tr cpab" data-val="{oct:.0f}">{oct:,.0f}</td>'
            f'<td class="tr cpab" data-val="{req:.0f}">{req:,.0f}</td>'
            f'<td class="tr cpab" data-val="{plan:.0f}">{plan:,.0f}</td>'
            f'<td class="tr cpab" data-val="{so:.0f}" style="color:#dc2626;">{so:,.0f}</td>'
            f'<td class="tr cpab" data-val="{pab:.0f}" style="font-weight:800;">{pab:,.0f}</td>'
            f'<td class="tr cpab" data-val="{fc:.1f}">{fc:,.1f}</td>'
            # MOS dinámico
            f'<td class="tr cdin" data-val="{s_oh_d}">{p_oh_d}</td>'
            f'<td class="tr cdin" data-val="{s_cf_d}">{p_cf_d}</td>'
            f'<td class="tr cdin" data-val="{s_pl_d}">{p_pl_d}</td>'
            f'<td class="tr cdin" data-val="{s_pb_d}">{p_pb_d}</td>'
            # MOS fijo
            f'<td class="tr cfij" data-val="{s_oh_f}">{p_oh_f}</td>'
            f'<td class="tr cfij" data-val="{s_cf_f}">{p_cf_f}</td>'
            f'<td class="tr cfij" data-val="{s_pl_f}">{p_pl_f}</td>'
            f'<td class="tr cfij" data-val="{s_pb_f}">{p_pb_f}</td>'
            # Clasificación + $
            f'<td class="tr" data-val="{clf_sort}" style="text-align:center;">{clf_html}</td>'
            f'<td class="tr" data-val="{inv:.2f}">{_fmt_mon(inv, 1, simbolo_mon)}</td>'
            f'</tr>\n'
        )

    # Fila total
    fc_tot_mes = tot_fc / n if n > 0 else 0
    body += (
        f'<tr class="tot">'
        f'<td></td><td class="tl">∑ TOTAL ({n} ítems)</td><td></td><td></td>'
        f'<td class="tr cpab">{tot_oh:,.0f}</td>'
        f'<td class="tr cpab">{tot_oct:,.0f}</td>'
        f'<td class="tr cpab">—</td>'
        f'<td class="tr cpab">{tot_plan:,.0f}</td>'
        f'<td class="tr cpab">—</td>'
        f'<td class="tr cpab" style="font-weight:900;">{tot_pab:,.0f}</td>'
        f'<td class="tr cpab">{fc_tot_mes:,.1f}</td>'
        f'<td></td><td></td><td></td><td></td>'
        f'<td></td><td></td><td></td><td></td>'
        f'<td></td>'
        f'<td class="tr">{_fmt_mon(tot_inv, 1, simbolo_mon)}</td>'
        f'</tr>'
    )

    html = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap" rel="stylesheet">
{_MOS_TAB_CSS}
</head><body>
<div style="padding:.25rem .1rem .4rem;font-size:.78rem;color:#334155;font-family:'IBM Plex Sans';">
  <b>{n}</b> ítems · <b>{titulo}</b>
  <span style="margin-left:1rem;font-size:.63rem;color:#94a3b8;">
    🟡 Dinámico = umbral por LT del ítem &nbsp;|&nbsp; 🟣 Fijo = &lt;3m crítico · 3–6m normal · &gt;6m exceso
  </span>
</div>
<div style="overflow-x:auto;border-radius:12px;box-shadow:0 6px 30px rgba(13,27,75,.12);">
  <table class="mt">{thead}<tbody>{body}</tbody></table>
</div>
{_MOS_TAB_JS}
</body></html>"""

    components.html(html, height=h, scrolling=True)


def _bloque_mos(con, df_long, dimension, filtros, factor_mon, simbolo_mon, moneda_sel, tc_usd_mxn):
    tc = tc_usd_mxn if tc_usd_mxn and tc_usd_mxn > 0 else 17.6367

    # ── Header ────────────────────────────────────────────────
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#1a1000 0%,#4a3000 50%,#1a1000 100%);
    border-radius:16px;padding:1.1rem 1.6rem;margin-bottom:1rem;
    border:1px solid rgba(251,191,36,.25);position:relative;overflow:hidden;">
    <div style="position:absolute;top:0;left:0;right:0;height:2px;
    background:linear-gradient(90deg,transparent,#d97706,#fcd34d,transparent);"></div>
    <div style="font-family:'Syne',sans-serif;font-size:1.1rem;font-weight:800;color:#fef3c7;letter-spacing:.4px;">
    📊 MESES DE INVENTARIO (MOS) — Cobertura & Riesgo</div>
    <div style="font-size:.7rem;color:#94a3b8;margin-top:.2rem;">
    Fuente: <b style="color:#fcd34d;">Hoja PEDIDO</b> &nbsp;·&nbsp;
    PAB = On Hand + OC Tránsito + Requisición + Planned − Sales Orders &nbsp;·&nbsp;
    MOS = PAB ÷ FC Promedio Mensual &nbsp;·&nbsp;
    Agrupar: <b style="color:#fcd34d;">{DIMS.get(dimension,dimension)}</b>
    </div>
    </div>""", unsafe_allow_html=True)

    # ── Panel de fórmulas ─────────────────────────────────────
    with st.expander("📐 ¿Cómo se calcula el MOS? — Ver fórmulas", expanded=False):
        st.markdown("""
<style>
.fbox{background:linear-gradient(135deg,#0f172a,#1e293b);border-radius:12px;
      padding:1rem 1.3rem;margin-bottom:.5rem;border:1px solid rgba(251,191,36,.2);}
.ftit{font-size:.7rem;font-weight:800;text-transform:uppercase;letter-spacing:.8px;color:#fcd34d;margin-bottom:.4rem;}
.feq{font-family:'IBM Plex Mono',monospace;font-size:.85rem;font-weight:700;color:#f1f5f9;
     background:rgba(255,255,255,.06);border-radius:7px;padding:.4rem .85rem;display:inline-block;margin:.2rem 0;}
.fsub{font-size:.67rem;color:#94a3b8;margin-top:.25rem;line-height:1.6;}
.fsub b{color:#e2e8f0;}
.ftag{display:inline-block;border-radius:20px;padding:.1rem .5rem;font-size:.59rem;font-weight:800;margin:.1rem .1rem;}
.tgb{background:rgba(59,130,246,.15);border:1px solid rgba(59,130,246,.3);color:#93c5fd;}
.tgr{background:rgba(239,68,68,.15);border:1px solid rgba(239,68,68,.3);color:#fca5a5;}
.sem-row{display:flex;gap:.5rem;flex-wrap:wrap;margin-top:.45rem;}
.schip{border-radius:7px;padding:.3rem .6rem;font-size:.63rem;font-weight:700;
       display:flex;flex-direction:column;align-items:center;min-width:100px;}
</style>
<div class="fbox">
  <div class="ftit">① 4 tipos de MOS que verás en la tabla</div>
  <div class="feq">MOS On Hand     = On Hand ÷ FC/mes</div><br>
  <div class="feq">MOS Confirmado  = (On Hand + OC Tránsito) ÷ FC/mes</div><br>
  <div class="feq">MOS Planeado    = (On Hand + OC Tránsito + Planned) ÷ FC/mes</div><br>
  <div class="feq">MOS PAB         = (On Hand + OC Tránsito + Requisición + Planned − Sales Orders) ÷ FC/mes</div>
  <div class="fsub">
    <b>FC/mes</b> = Forecast anual total ÷ 10 meses de horizonte.<br>
    El <b>MOS PAB</b> es el más completo: incluye todo el pipeline de abastecimiento.
  </div>
</div>
<div class="fbox">
  <div class="ftit">② Dos semáforos en paralelo para cada MOS</div>
  <div class="fsub" style="margin-bottom:.4rem;">
    <b style="color:#fcd34d;">🟡 Dinámico</b> — el umbral varía según el <b>Lead Time (LT_TOTAL_DIAS)</b> de cada ítem.<br>
    <b style="color:#c7d2fe;">🟣 Fijo</b> — umbrales estándar iguales para todos los ítems.
  </div>
  <div class="sem-row">
    <div class="schip" style="background:rgba(220,38,38,.15);border:1px solid rgba(220,38,38,.4);">
      <span>🔴 CRÍTICO</span>
      <span style="color:#f87171;font-family:'IBM Plex Mono';font-size:.62rem;">Dinámico: &lt; LT_meses</span>
      <span style="color:#94a3b8;font-family:'IBM Plex Mono';font-size:.62rem;">Fijo: &lt; 3 meses</span>
    </div>
    <div class="schip" style="background:rgba(22,163,74,.12);border:1px solid rgba(22,163,74,.4);">
      <span>✅ SALUDABLE</span>
      <span style="color:#4ade80;font-family:'IBM Plex Mono';font-size:.62rem;">Dinámico: LT a LT+2</span>
      <span style="color:#94a3b8;font-family:'IBM Plex Mono';font-size:.62rem;">Fijo: 3 – 6 meses</span>
    </div>
    <div class="schip" style="background:rgba(217,119,6,.12);border:1px solid rgba(217,119,6,.4);">
      <span>🟡 ALTO</span>
      <span style="color:#fbbf24;font-family:'IBM Plex Mono';font-size:.62rem;">Dinámico: LT+2 a LT+6</span>
      <span style="color:#94a3b8;font-family:'IBM Plex Mono';font-size:.62rem;">Fijo: 6 – 9 meses</span>
    </div>
    <div class="schip" style="background:rgba(124,45,18,.2);border:1px solid rgba(124,45,18,.5);">
      <span>⚠️ EXCESO</span>
      <span style="color:#fb923c;font-family:'IBM Plex Mono';font-size:.62rem;">Dinámico: &gt; LT+6</span>
      <span style="color:#94a3b8;font-family:'IBM Plex Mono';font-size:.62rem;">Fijo: &gt; 9 meses</span>
    </div>
  </div>
  <div class="fsub" style="margin-top:.5rem;">
    <b>Ejemplo China</b> LT = 150 d → LT_meses = 5.0 → Saludable si MOS entre 5.0 y 7.0 m<br>
    <b>Ejemplo USA</b>   LT = 30 d  → LT_meses = 1.0 → Saludable si MOS entre 1.0 y 3.0 m
  </div>
</div>""", unsafe_allow_html=True)

    # ── Datos ─────────────────────────────────────────────────
    df_res = _q_mos_resumen(con, dimension, filtros)
    df_res = df_res[df_res["DIM"].notna() & (df_res["DIM"].astype(str).str.lower() != "nan")].copy()
    if df_res.empty:
        st.info("Sin datos MOS para la selección actual.")
        return

    df_res["INV_USD"] = df_res["INV_MXN"] / tc

    # ── KPIs globales ─────────────────────────────────────────
    tot_items  = int(df_res["N_ITEMS"].sum())
    tot_crit_d = int(df_res["DIN_CRITICO"].sum())
    tot_norm_d = int(df_res["DIN_NORMAL"].sum())
    tot_alto_d = int(df_res["DIN_ALTO"].sum())
    tot_exc_d  = int(df_res["DIN_EXCESO"].sum())
    tot_crit_f = int(df_res["FIJ_CRITICO"].sum())
    tot_norm_f = int(df_res["FIJ_NORMAL"].sum())
    tot_exc_f  = int(df_res["FIJ_EXCESO"].sum())
    tot_sinfc  = int(df_res["SIN_FC"].sum())
    tot_inv    = df_res["INV_USD"].sum()

    def _kpi(col, icon, label, val, sub, bg, bc, vc):
        col.markdown(
            f'<div style="background:{bg};border-radius:10px;padding:.65rem .9rem;'
            f'border-left:4px solid {bc};text-align:center;">'
            f'<div style="font-size:.58rem;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:.5px;">{icon} {label}</div>'
            f'<div style="font-size:1.45rem;font-weight:900;color:{vc};line-height:1.2;">{val}</div>'
            f'<div style="font-size:.6rem;color:#888;margin-top:.1rem;">{sub}</div></div>',
            unsafe_allow_html=True)

    st.markdown("##### 🟡 Semáforo Dinámico (por Lead Time)")
    c1,c2,c3,c4,c5 = st.columns(5)
    pct = lambda x: f"{x/tot_items*100:.0f}% de ítems" if tot_items else "—"
    _kpi(c1,"🔴","Crítico (< LT)",    tot_crit_d, pct(tot_crit_d), "#fff5f5","#dc2626","#dc2626")
    _kpi(c2,"✅","Saludable (LT–LT+2)",tot_norm_d, pct(tot_norm_d), "#f0fff8","#16a34a","#16a34a")
    _kpi(c3,"🟡","Alto (LT+2–LT+6)",  tot_alto_d, pct(tot_alto_d), "#fffbeb","#d97706","#d97706")
    _kpi(c4,"⚠️","Exceso (> LT+6)",   tot_exc_d,  pct(tot_exc_d),  "#fff7ed","#7c2d12","#c2410c")
    _kpi(c5,"📭","Sin Forecast",       tot_sinfc,  "No calculable",  "#f8f9fa","#7c3aed","#7c3aed")

    st.markdown("<div style='height:.4rem'></div>", unsafe_allow_html=True)
    st.markdown("##### 🟣 Semáforo Fijo (&lt;3m · 3–6m · &gt;6m)")
    f1,f2,f3,f4 = st.columns(4)
    _kpi(f1,"🔴","Crítico (< 3 m)",   tot_crit_f, pct(tot_crit_f), "#fff5f5","#dc2626","#dc2626")
    _kpi(f2,"✅","Normal (3 – 6 m)",  tot_norm_f, pct(tot_norm_f), "#f0fff8","#16a34a","#16a34a")
    _kpi(f3,"⚠️","Exceso (> 6 m)",   tot_exc_f,  pct(tot_exc_f),  "#fff7ed","#7c2d12","#c2410c")
    _kpi(f4,"📭","Sin Forecast",       tot_sinfc,  "No calculable",  "#f8f9fa","#7c3aed","#7c3aed")

    st.markdown("<div style='height:.6rem'></div>", unsafe_allow_html=True)

    # ── Gráfica + Tabla resumen ───────────────────────────────
    col_bar, col_tbl = st.columns([1.3, 1])

    with col_bar:
        df_plot = df_res.sort_values("MOS_PROM", ascending=True).copy()
        df_plot["MOS_PROM"] = df_plot["MOS_PROM"].fillna(0)

        def _bar_color(v):
            if v < 3:  return "#dc2626"
            if v <= 6: return "#16a34a"
            if v <= 9: return "#d97706"
            return "#7c2d12"

        fig = go.Figure(go.Bar(
            y=df_plot["DIM"], x=df_plot["MOS_PROM"],
            orientation='h',
            marker_color=[_bar_color(v) for v in df_plot["MOS_PROM"]],
            text=[f"{v:.1f} m" for v in df_plot["MOS_PROM"]],
            textposition='outside', textfont_size=11,
            hovertemplate='<b>%{y}</b><br>MOS PAB Prom: %{x:.1f} m<extra></extra>',
        ))
        for xv, lbl, clr in [(3,"Crítico","#dc2626"),(6,"Normal","#16a34a"),(9,"Exceso","#d97706")]:
            fig.add_vline(x=xv, line_dash="dash", line_color=clr, line_width=1.5,
                          annotation_text=lbl, annotation_position="top",
                          annotation_font_size=9, annotation_font_color=clr)
        fig.update_layout(
            height=max(280, len(df_plot)*52),
            margin=dict(l=0,r=65,t=35,b=15),
            plot_bgcolor='white', paper_bgcolor='white',
            xaxis=dict(showgrid=True, gridcolor='#f0f0f0', zeroline=False,
                       title="MOS PAB Promedio (meses)"),
            yaxis=dict(showgrid=False),
            font=dict(family='IBM Plex Sans', size=11),
            title=dict(text=f"MOS PAB Promedio por {DIMS.get(dimension,dimension)}", font_size=13),
            hoverlabel=dict(bgcolor='#1a1a2e', font_color='white'),
        )
        st.plotly_chart(fig, use_container_width=True, key=f"mos_bar_{dimension}")

    with col_tbl:
        st.markdown(f"**Resumen por {DIMS.get(dimension,dimension)}**")
        df_t = df_res[["DIM","MOS_PROM","DIN_CRITICO","DIN_NORMAL","DIN_EXCESO",
                        "FIJ_CRITICO","FIJ_NORMAL","FIJ_EXCESO","SIN_FC","INV_USD"]].copy()
        df_t["MOS_PROM"] = df_t["MOS_PROM"].apply(lambda v: f"{v:.1f} m" if pd.notna(v) else "—")
        df_t["INV_USD"]  = df_t["INV_USD"].apply(lambda v: _fmt_mon(v, factor_mon, simbolo_mon))
        df_t.columns = [DIMS.get(dimension,'Grupo'),
                        "MOS Prom",
                        "🔴 Crit (D)","✅ OK (D)","⚠️ Exc (D)",
                        "🔴 Crit (F)","✅ OK (F)","⚠️ Exc (F)",
                        "📭 Sin FC", f"Inv. {simbolo_mon}"]
        st.dataframe(df_t, hide_index=True, use_container_width=True,
                     height=min(400, 40 + len(df_t)*35))
        st.caption("(D) = Semáforo Dinámico por LT · (F) = Semáforo Fijo")

    # ── Detalle por ítem ──────────────────────────────────────
    st.markdown("---")
    st.markdown(f"#### 🔍 Detalle por Ítem — selecciona un {DIMS.get(dimension,'Grupo')}")

    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        dims_disp = ["— Seleccionar —", "⭐ TODOS"] + sorted(
            [x for x in df_res["DIM"].unique().tolist()
             if x is not None and str(x).lower() != "nan"])
        sel = st.selectbox("Grupo / Dimensión:", dims_disp, key=f"mos_sel_{dimension}")
    with col_b:
        zona = st.selectbox("Zona MOS (Dinámico):",
                            ["Todos","🔴 Crítico","✅ Saludable","🟡 Alto","⚠️ Exceso"],
                            key=f"mos_zona_{dimension}")
    with col_c:
        orden = st.selectbox("Ordenar por:",
                             ["Inv. On Hand $","MOS PAB ↑","MOS PAB ↓","On Hand pzas"],
                             key=f"mos_ord_{dimension}")

    if sel and sel != "— Seleccionar —":
        # Obtener ítems
        df_items = _q_mos_items(con, filtros)

        if sel != "⭐ TODOS":
            df_items = df_items[df_items[dimension] == sel].copy()
            titulo_det = sel
        else:
            titulo_det = f"TODOS ({DIMS.get(dimension,dimension)})"

        if df_items.empty:
            st.info("Sin ítems para esta selección.")
        else:
            # Filtro zona
            if zona != "Todos":
                lt_s  = (df_items["LT_DIAS"].fillna(150) / 30.0)
                mp    = df_items["MOS_PAB"].fillna(float('nan'))
                if   "Crítico"   in zona: mask = mp < lt_s
                elif "Saludable" in zona: mask = (mp >= lt_s) & (mp <= lt_s + 2)
                elif "Alto"      in zona: mask = (mp > lt_s + 2) & (mp <= lt_s + 6)
                elif "Exceso"    in zona: mask = mp > lt_s + 6
                else: mask = pd.Series([True]*len(df_items), index=df_items.index)
                df_items = df_items[mask].copy()

            # Orden
            if   orden == "MOS PAB ↑":      df_items = df_items.sort_values("MOS_PAB", ascending=True)
            elif orden == "MOS PAB ↓":      df_items = df_items.sort_values("MOS_PAB", ascending=False)
            elif orden == "On Hand pzas":   df_items = df_items.sort_values("OH", ascending=False)
            else:                           df_items = df_items.sort_values("INV_MXN", ascending=False)

            if df_items.empty:
                st.info("No hay ítems en esa zona para este grupo.")
            else:
                _render_mos_items(df_items, titulo_det, factor_mon, simbolo_mon, tc_usd_mxn)
# ═══════════════════════════════════════════════════════════════
#  TABLA DE DETALLE DE ÍTEMS (compartida entre bloques)
# ═══════════════════════════════════════════════════════════════
def _render_detalle_items(df_det, titulo, factor_mon, simbolo_mon, tc_usd_mxn, modo="pedido"):
    df_det = df_det.copy()

    # Normalizar INV_USD de MXN → USD real
    _tc_det = tc_usd_mxn if tc_usd_mxn and tc_usd_mxn > 0 else 17.6367
    df_det["INV_USD"] = df_det["INV_USD"] / _tc_det

    # ── Calcular PAB y MOS con semáforo DINÁMICO por LT ──────
    # Asegurar columnas nuevas (compatibilidad con modo=pedido donde no vienen)
    for col in ["OCT_PZA","REQ_PZA","PLAN_PZA","SO_PZA","LT_DIAS"]:
        if col not in df_det.columns:
            df_det[col] = 0 if col != "LT_DIAS" else 150

    df_det["PAB"] = (
        df_det["INV_PZA"]
        + df_det["OCT_PZA"]
        + df_det["REQ_PZA"]
        + df_det["PLAN_PZA"]
        - df_det["SO_PZA"]
    ).clip(lower=0)

    fc_mes = (df_det["FC_PZA"] / 10.0).replace(0, float("nan"))
    df_det["MOS"] = (df_det["PAB"] / fc_mes).round(1)
    df_det["LT_MESES"] = (df_det["LT_DIAS"].fillna(150) / 30.0).round(1)

    def _sem(mos_v, lt_m):
        """Semáforo dinámico basado en Lead Time del ítem."""
        if pd.isna(mos_v):       return "—",        "#888",     ""
        if mos_v < lt_m:         return f"🔴 {mos_v:.1f}m", "#C0392B", "mos-crit"
        if mos_v <= lt_m + 2:    return f"✅ {mos_v:.1f}m", "#1A5C38", "mos-ok"
        if mos_v <= lt_m + 4:    return f"🟡 {mos_v:.1f}m", "#9A6700", "mos-warn"
        return                          f"⚠️ {mos_v:.1f}m", "#7B2C2C", "mos-exc"

    # Filtrar según modo
    if modo == "pedido":
        df_det = df_det[df_det["PED_PZA"] > 0].reset_index(drop=True)
    else:
        df_det = df_det[
            (df_det["PED_PZA"] > 0) | (df_det["INV_PZA"] > 0) | (df_det["OCT_USD"] > 0)
        ].reset_index(drop=True)

    n_rows    = len(df_det)
    # Tabla MOS es más ancha → altura por fila un poco mayor
    row_h     = 36
    header_h  = 52
    total_h   = 40
    label_h   = 30
    padding_h = 20
    h_iframe  = label_h + header_h + (n_rows * row_h) + total_h + padding_h
    h_iframe  = max(h_iframe, 180)

    css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'IBM Plex Sans',sans-serif;background:#f8fafc;}
.det-wrap{
  overflow-x:auto;border-radius:14px;
  box-shadow:0 8px 40px rgba(13,27,75,.14),0 2px 8px rgba(0,0,0,.06);
  margin-bottom:.5rem;
}
.det-tbl{border-collapse:collapse;width:100%;min-width:1600px;background:#ffffff;}
.det-tbl thead tr{
  background:linear-gradient(135deg,#0d1b3e 0%,#1a3a6e 60%,#0f2550 100%);
}
.det-tbl thead th{
  color:#cbd5e1;font-weight:700;font-size:.62rem;letter-spacing:.5px;
  padding:10px 8px;text-align:center;white-space:nowrap;
  cursor:pointer;user-select:none;transition:background .15s;
  border-right:1px solid rgba(255,255,255,.05);
}
.det-tbl thead th:last-child{border-right:none;}
.det-tbl thead th:hover{background:rgba(255,255,255,.08);}
.det-tbl thead th.sort-asc{background:rgba(59,130,246,.3)!important;}
.det-tbl thead th.sort-desc{background:rgba(239,68,68,.25)!important;}
.det-tbl thead th .arr::after{content:'⇅';opacity:.35;margin-left:4px;font-size:.6rem;}
.det-tbl thead th.sort-asc  .arr::after{content:' ▲';opacity:1;color:#93c5fd;}
.det-tbl thead th.sort-desc .arr::after{content:' ▼';opacity:1;color:#fca5a5;}
.det-tbl thead th.th-left{text-align:left;padding-left:12px;}
/* Grupos de columnas */
.det-tbl thead th.th-pab{background:rgba(16,185,129,.15)!important;color:#6ee7b7!important;}
.det-tbl thead th.th-mos{background:rgba(251,191,36,.15)!important;color:#fcd34d!important;}
.det-tbl thead th.th-rsk{background:rgba(139,92,246,.18)!important;color:#c4b5fd!important;}
.det-tbl tbody tr{transition:background .1s;}
.det-tbl tbody tr:nth-child(even) td{background:#f8faff;}
.det-tbl tbody tr:hover td{
  background:linear-gradient(90deg,#eff6ff,#f0f9ff)!important;
}
.det-tbl td{
  font-size:.71rem;padding:6px 8px;
  border-bottom:1px solid #eef2f8;
  white-space:nowrap;vertical-align:middle;color:#1e293b;
}
.det-tbl td.tl{text-align:left;padding-left:12px;font-weight:700;color:#0f172a;}
.det-tbl td.tr{
  text-align:right;padding-right:10px;
  font-family:'IBM Plex Mono',monospace;font-size:.69rem;
}
.det-tbl td.td-pab{background:rgba(16,185,129,.05);}
.det-tbl td.td-mos{background:rgba(251,191,36,.05);}
.mos-crit{
  color:#fff!important;font-weight:800!important;
  background:linear-gradient(90deg,#dc2626,#b91c1c)!important;
  border-radius:6px;padding:2px 7px;
}
.mos-ok{
  color:#065f46!important;font-weight:800!important;
  background:linear-gradient(90deg,#d1fae5,#a7f3d0)!important;
  border-radius:6px;padding:2px 7px;
}
.mos-warn{
  color:#92400e!important;font-weight:800!important;
  background:linear-gradient(90deg,#fef3c7,#fde68a)!important;
  border-radius:6px;padding:2px 7px;
}
.mos-exc{
  color:#fff!important;font-weight:800!important;
  background:linear-gradient(90deg,#9a3412,#7c2d12)!important;
  border-radius:6px;padding:2px 7px;
}
.det-tbl tr.tot-row td{
  background:linear-gradient(90deg,#1e3a5f,#1a3a6e)!important;
  color:#e2e8f0!important;font-weight:800!important;
  border-top:2px solid #3b82f6;font-family:'IBM Plex Mono',monospace;
}
.det-tbl tr.tot-row td.tl{color:#93c5fd!important;}
</style>"""

    # ── Definición de columnas según modo ─────────────────────
    if modo == "mos":
        cols_def = [
            ("num", "",         "#"),
            ("1",  "th-left",   "CC / Ítem"),
            ("2",  "th-left",   "Nombre"),
            ("3",  "",          "LT\n(días)"),
            ("4",  "th-pab",    "On Hand\n(pzas)"),
            ("5",  "th-pab",    "OC Tránsito\n(pzas)"),
            ("6",  "th-pab",    "Requisición\n(pzas)"),
            ("7",  "th-pab",    "Planned\n(pzas)"),
            ("8",  "th-pab",    "Sales Orders\n(pzas)"),
            ("9",  "th-pab",    "PAB\n(pzas)"),
            ("10", "th-pab",    "FC / mes\n(pzas)"),
            ("11", "th-mos",    "MOS\nOn Hand"),
            ("12", "th-mos",    "MOS\nConfirmado"),
            ("13", "th-mos",    "MOS\nPlaneado"),
            ("14", "th-mos",    "MOS\nPAB"),
            ("15", "th-rsk",    "Clasificación"),
            ("16", "",          f"Inv. ({simbolo_mon})"),
        ]
    else:
        cols_def = [
            ("num", "",        "#"),
            ("1",  "th-left",  "CC / Ítem"),
            ("2",  "th-left",  "Nombre"),
            ("3",  "",         f"Inv. On Hand\n({simbolo_mon})"),
            ("4",  "",         "Pzas On Hand"),
            ("5",  "",         f"Pedido ({simbolo_mon})"),
            ("6",  "",         "Pzas Pedido"),
            ("7",  "",         f"OC Tránsito ({simbolo_mon})"),
            ("8",  "th-mos",   "MOS"),
            ("9",  "",         "Forecast Pzas"),
        ]

    thead_parts = []
    for idx, cls, lbl in cols_def:
        label_html = lbl.replace("\n", "<br>")
        if idx == "num":
            thead_parts.append(f'<th class="{cls}" style="width:30px;text-align:center;">{label_html}</th>')
        else:
            thead_parts.append(
                f'<th class="{cls} sortable" data-col="{idx}" onclick="dtSort(this)">'
                f'{label_html} <span class="arr"></span></th>'
            )
    thead = "<thead><tr>" + "".join(thead_parts) + "</tr></thead>"

    # ── Filas ─────────────────────────────────────────────────
    body = ""
    tot_inv_usd = tot_inv_pza = tot_ped_usd = tot_ped_pza = tot_oct_usd = 0.0
    tot_pab = tot_fc = 0.0
    row_num = 0

    for _, r in df_det.iterrows():
        row_num += 1
        inv_usd = r["INV_USD"] * factor_mon
        inv_pza = r["INV_PZA"]
        ped_usd = r["PED_USD"] * factor_mon
        ped_pza = r["PED_PZA"]
        oct_usd = r["OCT_USD"] * factor_mon
        fc_pza  = r["FC_PZA"]
        mos_v   = r["MOS"]
        lt_m    = r["LT_MESES"]
        mos_str, _, mos_pill = _sem(mos_v, lt_m)
        mos_raw = mos_v if pd.notna(mos_v) else -1

        nombre = str(r.get("NOMBRE","")) if str(r.get("NOMBRE","")) not in ("","nan") else "—"
        tot_inv_usd += inv_usd; tot_inv_pza += inv_pza
        tot_ped_usd += ped_usd; tot_ped_pza += ped_pza
        tot_oct_usd += oct_usd

        if modo == "mos":
            oct_pza  = r.get("OCT_PZA",  0)
            req_pza  = r.get("REQ_PZA",  0)
            plan_pza = r.get("PLAN_PZA", 0)
            so_pza   = r.get("SO_PZA",   0)
            pab_pza  = r.get("PAB",      0)
            fc_mes_v = fc_pza / 10.0
            tot_pab += pab_pza; tot_fc += fc_pza
            body += (
                f'<tr>'
                f'<td class="tr" data-val="{row_num}" style="color:#94a3b8;font-size:.63rem;width:28px;">{row_num}</td>'
                f'<td class="tl" data-val="{r["CC"]}">{r["CC"]}</td>'
                f'<td class="tl" data-val="{nombre}" style="max-width:170px;overflow:hidden;text-overflow:ellipsis;" title="{nombre}">{nombre}</td>'
                f'<td class="tr" data-val="{inv_pza:.0f}">{inv_pza:,.0f}</td>'
                f'<td class="tr td-pab" data-val="{oct_pza:.0f}">{oct_pza:,.0f}</td>'
                f'<td class="tr td-pab" data-val="{req_pza:.0f}">{req_pza:,.0f}</td>'
                f'<td class="tr td-pab" data-val="{plan_pza:.0f}">{plan_pza:,.0f}</td>'
                f'<td class="tr td-pab" data-val="{so_pza:.0f}" style="color:#C0392B;">{so_pza:,.0f}</td>'
                f'<td class="tr td-pab" data-val="{pab_pza:.0f}" style="font-weight:800;">{pab_pza:,.0f}</td>'
                f'<td class="tr td-mos" data-val="{fc_mes_v:.1f}">{fc_mes_v:,.1f}</td>'
                f'<td class="tr td-mos" data-val="{lt_m:.1f}">{lt_m:.1f} m</td>'
                f'<td class="tr td-mos" data-val="{mos_raw}"><span class="{mos_pill}">{mos_str}</span></td>'
                f'<td class="tr" data-val="{inv_usd:.2f}">{_fmt_mon(inv_usd,1,simbolo_mon)}</td>'
                f'</tr>\n'
            )
        else:
            body += (
                f'<tr>'
                f'<td class="tr" data-val="{row_num}" style="color:#94a3b8;font-size:.65rem;width:32px;">{row_num}</td>'
                f'<td class="tl" data-val="{r["CC"]}">{r["CC"]}</td>'
                f'<td class="tl" data-val="{nombre}" style="max-width:180px;overflow:hidden;text-overflow:ellipsis;" title="{nombre}">{nombre}</td>'
                f'<td class="tr" data-val="{inv_usd:.2f}">{_fmt_mon(inv_usd,1,simbolo_mon)}</td>'
                f'<td class="tr" data-val="{inv_pza:.0f}">{inv_pza:,.0f}</td>'
                f'<td class="tr" data-val="{ped_usd:.2f}">{_fmt_mon(ped_usd,1,simbolo_mon)}</td>'
                f'<td class="tr" data-val="{ped_pza:.0f}">{ped_pza:,.0f}</td>'
                f'<td class="tr" data-val="{oct_usd:.2f}">{_fmt_mon(oct_usd,1,simbolo_mon)}</td>'
                f'<td class="tr td-mos" data-val="{mos_raw}"><span class="{mos_pill}">{mos_str}</span></td>'
                f'<td class="tr" data-val="{fc_pza:.0f}">{fc_pza:,.0f}</td>'
                f'</tr>\n'
            )

    # ── Fila TOTAL ─────────────────────────────────────────────
    if modo == "mos":
        fc_mes_tot = (tot_fc / 10.0) if tot_fc else 0
        body += (
            f'<tr class="tot-row">'
            f'<td></td>'
            f'<td class="tl">∑ TOTAL ({n_rows} ítems)</td><td></td>'
            f'<td class="tr">{tot_inv_pza:,.0f}</td>'
            f'<td class="tr td-pab"></td><td class="tr td-pab"></td>'
            f'<td class="tr td-pab"></td><td class="tr td-pab"></td>'
            f'<td class="tr td-pab" style="font-weight:800;">{tot_pab:,.0f}</td>'
            f'<td class="tr td-mos">{fc_mes_tot:,.1f}</td>'
            f'<td class="tr td-mos"></td><td class="tr td-mos"></td>'
            f'<td class="tr">{_fmt_mon(tot_inv_usd,1,simbolo_mon)}</td>'
            f'</tr>'
        )
    else:
        body += (
            f'<tr class="tot-row">'
            f'<td></td>'
            f'<td class="tl">∑ TOTAL ({n_rows} ítems)</td><td></td>'
            f'<td class="tr">{_fmt_mon(tot_inv_usd,1,simbolo_mon)}</td>'
            f'<td class="tr">{tot_inv_pza:,.0f}</td>'
            f'<td class="tr">{_fmt_mon(tot_ped_usd,1,simbolo_mon)}</td>'
            f'<td class="tr">{tot_ped_pza:,.0f}</td>'
            f'<td class="tr">{_fmt_mon(tot_oct_usd,1,simbolo_mon)}</td>'
            f'<td></td><td></td></tr>'
        )

    js = """
<script>
(function(){
  var sc=0, sa=false;
  window.dtSort=function(th){
    var c=parseInt(th.getAttribute('data-col'));
    if(sc===c){sa=!sa;}else{sc=c;sa=false;}
    document.querySelectorAll('.det-tbl thead th.sortable').forEach(function(h){
      h.classList.remove('sort-asc','sort-desc');
    });
    th.classList.add(sa?'sort-asc':'sort-desc');
    var tb=th.closest('table').querySelector('tbody');
    var rows=Array.from(tb.querySelectorAll('tr:not(.tot-row)'));
    var tot=tb.querySelector('tr.tot-row');
    rows.sort(function(a,b){
      var A=a.querySelectorAll('td')[c],B=b.querySelectorAll('td')[c];
      var vA=A?(A.getAttribute('data-val')||A.textContent.trim()):'';
      var vB=B?(B.getAttribute('data-val')||B.textContent.trim()):'';
      var nA=parseFloat(String(vA).replace(/[^0-9.+-]/g,''));
      var nB=parseFloat(String(vB).replace(/[^0-9.+-]/g,''));
      if(!isNaN(nA)&&!isNaN(nB)) return sa?nA-nB:nB-nA;
      return sa?String(vA).localeCompare(String(vB)):String(vB).localeCompare(String(vA));
    });
    rows.forEach(function(r){tb.appendChild(r);});
    if(tot) tb.appendChild(tot);
  };
})();
</script>"""

    titulo_limpio = str(titulo).replace("'", "")
    html = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap" rel="stylesheet">
{css}
</head><body>
<div style="padding:.3rem .1rem .5rem;font-family:'IBM Plex Sans';font-size:.8rem;color:#334155;">
<b>{n_rows}</b> ítems en <b>{titulo_limpio}</b>
</div>
<div class="det-wrap"><table class="det-tbl">
{thead}<tbody>{body}</tbody>
</table></div>
{js}
</body></html>"""

    components.html(html, height=h_iframe, scrolling=True)


# ═══════════════════════════════════════════════════════════════
#  BLOQUE 4 — LANDED COST (wrapper con header premium)
# ═══════════════════════════════════════════════════════════════
def _bloque_landed_cost(df_long, tc_usd_mxn, moneda_sel, factor_mon, simbolo_mon):
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#0a1628 0%,#1a237e 50%,#0d1b3e 100%);
    border-radius:16px;padding:1.1rem 1.6rem;margin-bottom:1.2rem;
    border:1px solid rgba(100,120,255,.2);position:relative;overflow:hidden;">
    <div style="position:absolute;top:0;left:0;right:0;height:2px;
    background:linear-gradient(90deg,transparent,#7c3aed,#3b82f6,transparent);"></div>
    <div style="position:absolute;top:-50px;right:-40px;width:200px;height:200px;
    background:radial-gradient(circle,rgba(124,58,237,.12) 0%,transparent 70%);pointer-events:none;"></div>
    <div style="font-family:'Syne',sans-serif;font-size:1.1rem;font-weight:800;
    color:#e2e8f0;letter-spacing:.4px;">🧮 COSTO TOTAL DE IMPORTACIÓN — Landed Cost</div>
    <div style="font-size:.7rem;color:#64748b;margin-top:.2rem;">
    Valor CIF + Aranceles + Logística · Calcula el costo real de traer cada producto a México</div>
    <div style="display:flex;gap:.5rem;margin-top:.6rem;flex-wrap:wrap;">
    <span style="background:rgba(124,58,237,.15);border:1px solid rgba(124,58,237,.3);
    border-radius:20px;padding:.15rem .6rem;font-size:.58rem;font-weight:700;color:#a78bfa;letter-spacing:.4px;">
    IGI · DTA · Prevalidación</span>
    <span style="background:rgba(59,130,246,.12);border:1px solid rgba(59,130,246,.25);
    border-radius:20px;padding:.15rem .6rem;font-size:.58rem;font-weight:700;color:#93c5fd;letter-spacing:.4px;">
    Transporte · Custodia · Seguro</span>
    <span style="background:rgba(16,185,129,.1);border:1px solid rgba(16,185,129,.2);
    border-radius:20px;padding:.15rem .6rem;font-size:.58rem;font-weight:700;color:#6ee7b7;letter-spacing:.4px;">
    Moneda: {moneda_sel}</span>
    </div></div>
    """, unsafe_allow_html=True)
    from app_final import render_landed_cost_tab
    render_landed_cost_tab(
        df_long     = df_long,
        tc_usd_mxn  = tc_usd_mxn,
        moneda_sel  = moneda_sel,
        factor_mon  = factor_mon,
        simbolo_mon = simbolo_mon,
        filtros_key = (),
        filtros     = {},
    )


# ═══════════════════════════════════════════════════════════════
#  BLOQUE 5 — BACKORDER (wrapper con header premium)
# ═══════════════════════════════════════════════════════════════
def _bloque_backorder(df_long, factor_mon, simbolo_mon, moneda_sel,
                      df_bo, df_prom_ventas, tc_usd_mxn, df_ventas_mes=None):
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#1a0505 0%,#7b1a1a 40%,#1a0505 100%);
    border-radius:16px;padding:1.1rem 1.6rem;margin-bottom:1.2rem;
    border:1px solid rgba(239,68,68,.2);position:relative;overflow:hidden;">
    <div style="position:absolute;top:0;left:0;right:0;height:2px;
    background:linear-gradient(90deg,transparent,#ef4444,#f97316,transparent);"></div>
    <div style="position:absolute;top:-50px;right:-40px;width:200px;height:200px;
    background:radial-gradient(circle,rgba(239,68,68,.1) 0%,transparent 70%);pointer-events:none;"></div>
    <div style="font-family:'Syne',sans-serif;font-size:1.1rem;font-weight:800;
    color:#fef2f2;letter-spacing:.4px;">🚨 ANÁLISIS DE BACKORDER — Riesgo de Abasto</div>
    <div style="font-size:.7rem;color:#94a3b8;margin-top:.2rem;">
    Demanda no cubierta · Semáforo de disponibilidad · Ítems en doble riesgo</div>
    <div style="display:flex;gap:.5rem;margin-top:.6rem;flex-wrap:wrap;">
    <span style="background:rgba(239,68,68,.15);border:1px solid rgba(239,68,68,.3);
    border-radius:20px;padding:.15rem .6rem;font-size:.58rem;font-weight:700;color:#fca5a5;letter-spacing:.4px;">
    🔴 Sin Abasto</span>
    <span style="background:rgba(234,179,8,.12);border:1px solid rgba(234,179,8,.25);
    border-radius:20px;padding:.15rem .6rem;font-size:.58rem;font-weight:700;color:#fde68a;letter-spacing:.4px;">
    🟡 En Tránsito / Parcial</span>
    <span style="background:rgba(16,185,129,.1);border:1px solid rgba(16,185,129,.2);
    border-radius:20px;padding:.15rem .6rem;font-size:.58rem;font-weight:700;color:#6ee7b7;letter-spacing:.4px;">
    ✅ Disponible</span>
    <span style="background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.2);
    border-radius:20px;padding:.15rem .6rem;font-size:.58rem;font-weight:700;color:#fca5a5;letter-spacing:.4px;">
    Moneda: {moneda_sel}</span>
    </div></div>
    """, unsafe_allow_html=True)
    from app_final import render_backorder_tab
    render_backorder_tab(
        df_long        = df_long,
        factor_mon     = factor_mon,
        simbolo_mon    = simbolo_mon,
        moneda_sel     = moneda_sel,
        df_bo          = df_bo,
        df_prom_ventas = df_prom_ventas,
        df_ventas_mes  = df_ventas_mes,
        filtros_key    = (),
        filtros        = {},
        tc_usd_mxn     = tc_usd_mxn,
    )



# ═══════════════════════════════════════════════════════════════
#  FUNCIÓN PRINCIPAL — render_detalle_tab
# ═══════════════════════════════════════════════════════════════
def render_detalle_tab(df_long, con, factor_mon, simbolo_mon, moneda_sel,
                       tc_usd_mxn=17.6367, tc_eur_mxn=20.5505,
                       filtros=None, df_bo=None, df_prom_ventas=None, df_ventas_mes=None):
    """
    Pestaña DETALLE — 5 sub-tabs ejecutivos.
    Recibe `con` ya registrada con la tabla `datos`.
    """
    from app_final import render_landed_cost_tab, render_backorder_tab
    filtros = filtros or {}

    # ══════════════════════════════════════════════════════════
    #  HEADER PREMIUM — Impacto inmediato para directivos
    # ══════════════════════════════════════════════════════════
    st.markdown("""
    <style>
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="stSidebarCollapsedControl"] * {
        display: none !important; visibility: hidden !important;
        width: 0 !important; height: 0 !important;
        pointer-events: none !important;
    }

    @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=IBM+Plex+Sans:wght@300;400;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');
    .det-dim-bar {
        background: linear-gradient(90deg,rgba(15,30,60,.8),rgba(10,20,45,.6));
        border-radius: 12px; padding: .55rem 1.1rem;
        margin-bottom: 1rem;
        border: 1px solid rgba(59,130,246,.12);
        display: flex; align-items: center; gap: 1rem; flex-wrap: wrap;
    }
    .det-dim-label {
        font-size: .62rem; font-weight: 800; color: #475569;
        text-transform: uppercase; letter-spacing: .8px; white-space: nowrap;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Selector de dimensión (aplica a Pedido, Inventario y MOS) ───────
    dims_disp = {k: v for k, v in DIMS.items() if k in df_long.columns}

    st.markdown('<div class="det-dim-bar">'
                '<span class="det-dim-label">📐 Agrupar análisis por:</span>',
                unsafe_allow_html=True)
    dim_sel = st.radio("Dimensión", list(dims_disp.keys()),
                       format_func=lambda x: dims_disp[x],
                       horizontal=True, key="detalle_dim_radio",
                       label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    # ── 5 Sub-tabs con persistencia de selección ─────────────
    _SUBTAB_LABELS = [
        "💰 Pedido Sugerido",
        "📦 ¿Dónde está el dinero?",
        "🧮 Landed Cost",
        "🚨 Backorder",
    ]
    # Leer sub-tab guardada en query_params (sobrevive reruns por radio)
    _qp = st.query_params
    _subtab_idx = 0
    if "det_sub" in _qp:
        try:
            _subtab_idx = int(_qp["det_sub"])
        except Exception:
            pass

    _subtab_sel = st.radio(
        "Sub-sección",
        options=list(range(len(_SUBTAB_LABELS))),
        format_func=lambda i: _SUBTAB_LABELS[i],
        index=min(_subtab_idx, len(_SUBTAB_LABELS)-1),
        horizontal=True,
        key="det_subtab_radio",
        label_visibility="collapsed",
    )
    st.query_params["det_sub"] = str(_subtab_sel)

    st.markdown("---")

    if _subtab_sel == 0:
        _bloque_pedido(con, df_long, dim_sel, filtros,
                       factor_mon, simbolo_mon, moneda_sel, tc_usd_mxn)
    elif _subtab_sel == 1:
        _bloque_inventario(con, df_long, dim_sel, filtros,
                           factor_mon, simbolo_mon, moneda_sel, tc_usd_mxn)
    elif _subtab_sel == 2:
        _bloque_landed_cost(df_long, tc_usd_mxn, moneda_sel, factor_mon, simbolo_mon)
    elif _subtab_sel == 3:
        _bloque_backorder(df_long, factor_mon, simbolo_mon, moneda_sel,
                          df_bo, df_prom_ventas, tc_usd_mxn,
                          df_ventas_mes=df_ventas_mes)