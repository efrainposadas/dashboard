# ============================================================
#  LOGIN.PY — InterAuto BI
#  Diseño visual idéntico al Flask + submit 100% Streamlit
#  SIN iframe para el submit — directo con st.form
# ============================================================

import streamlit as st
import streamlit.components.v1 as components

USUARIOS = {
    "DIRECCIÓN GENERAL":   "521",
    "DIRECCION COMPRAS":   "521",
    "PLANEACIÓN":          "521",
    "DIRECCION FINANZAS":  "521",
    "DIRECCION COMERCIAL": "521",
}


def render_login() -> bool:
    if "logged_in"   not in st.session_state: st.session_state.logged_in   = False
    if "login_error" not in st.session_state: st.session_state.login_error = False
    if "usuario"     not in st.session_state: st.session_state.usuario     = ""

    if st.session_state.logged_in:
        return True

    # ── CSS global: fondo + ocultar chrome + DISEÑO DE TARJETA ──
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;600;700&display=swap');

    html, body,
    [data-testid="stAppViewContainer"],
    [data-testid="stMain"],
    [data-testid="stAppViewBlockContainer"] {
        background: linear-gradient(135deg, #051625 0%, #0B2C4A 50%, #0e4976 100%) !important;
        min-height: 100vh !important;
    }
    /* Patrón de puntos animado */
    [data-testid="stAppViewContainer"]::before {
        content: '';
        position: fixed; inset: 0; z-index: 0; pointer-events: none;
        background-image:
            radial-gradient(circle, rgba(14,165,233,0.08) 1px, transparent 1px),
            radial-gradient(circle, rgba(255,204,0,0.04) 1px, transparent 1px);
        background-size: 50px 50px, 80px 80px;
        background-position: 0 0, 40px 40px;
        animation: floatbg 20s linear infinite;
    }
    @keyframes floatbg {
        0%   { background-position: 0 0, 40px 40px; }
        100% { background-position: -50px -50px, -10px -10px; }
    }

    /* Ocultar todo el chrome de Streamlit */
    [data-testid="stSidebar"], [data-testid="stHeader"],
    [data-testid="stSidebarCollapsedControl"], [data-testid="stToolbar"],
    [data-testid="stDecoration"], [data-testid="stStatusWidget"],
    button[kind="header"], #MainMenu, footer, header {
        display: none !important; visibility: hidden !important;
    }

    /* Centrar contenido */
    .block-container {
        max-width: 460px !important;
        padding: 8vh 0 0 0 !important;
        margin: 0 auto !important;
        position: relative; z-index: 1;
    }
    [data-testid="stVerticalBlock"] {
        gap: 0 !important;
    }
    .element-container, [data-testid="stMarkdownContainer"] {
        margin: 0 !important; padding: 0 !important;
    }

    /* ── TARJETA BLANCA ── */
    [data-testid="stForm"] {
        background: rgba(255,255,255,0.98) !important;
        border-radius: 20px !important;
        border: none !important;
        box-shadow:
            0 20px 60px rgba(0,0,0,0.5),
            0 0 100px rgba(14,165,233,0.15),
            inset 0 1px 0 rgba(255,255,255,0.8) !important;
        padding: 0 !important;
        overflow: hidden !important;
        animation: slideIn .6s ease-out !important;
    }
    @keyframes slideIn {
        from { opacity:0; transform:translateY(-20px); }
        to   { opacity:1; transform:translateY(0); }
    }
    [data-testid="stForm"] > div {
        padding: 0 !important; gap: 0 !important;
        background: transparent !important;
    }

    /* ── HEADER AZUL OSCURO (logo section) ── */
    .ia-header-wrap {
        background: transparent;
        padding: 40px 40px 30px;
        text-align: center;
    }
    .ia-logo-box {
        display: inline-block;
        background: linear-gradient(135deg, #051625 0%, #0B2C4A 100%);
        padding: 18px 40px;
        border-radius: 15px;
        margin-bottom: 18px;
        box-shadow: 0 8px 25px rgba(5,22,37,0.35);
        position: relative; overflow: hidden;
    }
    .ia-logo-box::before {
        content: '';
        position: absolute; top:-50%; left:-50%; width:200%; height:200%;
        background: linear-gradient(45deg,transparent,rgba(14,165,233,.12),transparent);
        animation: shine 3s infinite;
    }
    @keyframes shine { 0%{transform:rotate(0deg)} 100%{transform:rotate(360deg)} }
    .ia-logo-box h1 {
        color: white; font-size: 26px; font-weight: 700;
        letter-spacing: 2px; margin: 0;
        position: relative; z-index: 1;
        font-family: 'Segoe UI', sans-serif;
    }
    .ia-gold-bar {
        display: block; width: 60px; height: 4px;
        background: linear-gradient(90deg, #FFCC00, #FFD700);
        margin: 8px auto 0; border-radius: 2px;
        box-shadow: 0 0 10px rgba(255,204,0,0.5);
    }
    .ia-subtitle {
        color: #64748b; font-size: 14px; font-weight: 500; margin-top: 10px;
    }

    /* ── CUERPO DEL FORM ── */
    .ia-body { padding: 0 40px 35px; }
    .ia-divider {
        display: flex; align-items: center; gap: 12px;
        margin: 0 0 22px;
    }
    .ia-divider-line {
        flex: 1; height: 1px;
        background: linear-gradient(90deg, transparent, #d0e0f0, transparent);
    }

    /* Labels */
    .ia-label {
        display: block;
        color: #334155; font-size: 12px; font-weight: 700;
        text-transform: uppercase; letter-spacing: .5px;
        margin: 0 0 8px;
        font-family: 'Segoe UI', sans-serif;
    }

    /* Error banner */
    .ia-error {
        background: linear-gradient(135deg, #fee2e2, #fecaca);
        color: #dc2626; padding: 12px 16px; border-radius: 10px;
        margin-bottom: 20px; font-size: 13px; font-weight: 500;
        border: 1px solid #fca5a5;
        animation: shake .5s;
    }
    @keyframes shake {
        0%,100%{transform:translateX(0)} 25%{transform:translateX(-4px)} 75%{transform:translateX(4px)}
    }

    /* ── SELECTBOX ── */
    div[data-testid="stSelectbox"] > div > div {
        background: white !important;
        border: 2px solid #e2e8f0 !important;
        border-radius: 12px !important;
        transition: all .3s !important;
    }
    div[data-testid="stSelectbox"] > div > div:focus-within {
        border-color: #0ea5e9 !important;
        box-shadow: 0 0 0 4px rgba(14,165,233,0.1) !important;
        transform: translateY(-2px) !important;
    }
    div[data-testid="stSelectbox"] * {
        color: #1e293b !important; font-size: 15px !important;
        font-family: 'Segoe UI', sans-serif !important;
    }
    div[data-testid="stSelectbox"] svg path { stroke: #0ea5e9 !important; }
    [data-baseweb="popover"] ul {
        background: #fff !important;
        border: 1px solid #d0e0f8 !important;
        border-radius: 12px !important;
        box-shadow: 0 16px 50px rgba(15,40,100,.2) !important;
    }
    [data-baseweb="option"]:hover,
    [data-baseweb="option"][aria-selected="true"] {
        background: #eff6ff !important; color: #1d4ed8 !important;
    }

    /* ── TEXT INPUT ── */
    div[data-testid="stTextInput"] > div {
        background: white !important;
        border: 2px solid #e2e8f0 !important;
        border-radius: 12px !important;
        overflow: hidden !important;
        transition: all .3s !important;
    }
    div[data-testid="stTextInput"] > div:focus-within {
        border-color: #0ea5e9 !important;
        box-shadow: 0 0 0 4px rgba(14,165,233,0.1) !important;
        transform: translateY(-2px) !important;
    }
    div[data-testid="stTextInput"] input {
        background: transparent !important; border: none !important;
        box-shadow: none !important; font-size: 15px !important;
        color: #1e293b !important; padding: 15px 18px !important;
        font-family: 'Segoe UI', sans-serif !important;
    }
    div[data-testid="stTextInput"] input::placeholder { color: #94a3b8 !important; }
    /* Ocultar botón ojo */
    div[data-testid="stTextInput"] button {
        background: transparent !important; border: none !important;
        box-shadow: none !important; color: #94a3b8 !important;
    }

    /* ── BOTÓN SUBMIT ── */
    [data-testid="stForm"] button[kind="primaryFormSubmit"] {
        background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%) !important;
        color: white !important; border: none !important;
        border-radius: 12px !important;
        font-size: 15px !important; font-weight: 700 !important;
        text-transform: uppercase !important; letter-spacing: 1px !important;
        padding: 16px !important; width: 100% !important;
        box-shadow: 0 6px 20px rgba(14,165,233,0.4) !important;
        transition: all .3s ease !important;
        cursor: pointer !important;
    }
    [data-testid="stForm"] button[kind="primaryFormSubmit"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 25px rgba(14,165,233,0.5) !important;
    }

    /* Footer */
    .ia-footer {
        text-align: center; padding: 18px 0 0;
        border-top: 1px solid #f1f5f9;
        color: #94a3b8; font-size: 12px; line-height: 1.9;
    }
    .ia-footer strong { color: #0ea5e9; font-weight: 600; }

    /* Quitar padding extra del form interno */
    [data-testid="stForm"] [data-testid="stVerticalBlock"] {
        padding: 0 40px 35px !important;
        gap: 0 !important;
    }
    [data-testid="stForm"] .element-container { margin-bottom: 20px !important; }
    [data-testid="stForm"] .element-container:last-child { margin-bottom: 0 !important; }
    </style>
    """, unsafe_allow_html=True)

    # ── HEADER (logo) — fuera del form para que sea parte de la tarjeta visualmente ──
    # Usamos un st.container para que quede dentro del mismo bloque

    with st.form("ia_login", clear_on_submit=False):

        # Logo header
        st.markdown("""
        <div style="text-align:center; padding:40px 40px 24px; background:white; margin:-1px -1px 0;">
          <div class="ia-logo-box">
            <h1>InterAuto BI</h1>
            <span class="ia-gold-bar"></span>
          </div>
          <p class="ia-subtitle">Panel de Demand Planning &amp; Replenishment</p>
        </div>
        """, unsafe_allow_html=True)

        # Error si lo hay
        if st.session_state.login_error:
            st.markdown("""
            <div class="ia-error">
              ⚠️ &nbsp; Usuario o contraseña incorrectos
            </div>
            """, unsafe_allow_html=True)

        # Label Usuario
        st.markdown('<span class="ia-label">👤 &nbsp; USUARIO</span>', unsafe_allow_html=True)
        usuario = st.selectbox(
            "", options=list(USUARIOS.keys()),
            label_visibility="collapsed", key="ia_sel_usuario"
        )

        # Separador visual
        st.markdown('<div style="height:4px"></div>', unsafe_allow_html=True)

        # Label Contraseña
        st.markdown('<span class="ia-label">🔒 &nbsp; CONTRASEÑA</span>', unsafe_allow_html=True)
        password = st.text_input(
            "Contraseña", type="password",
            placeholder="Ingrese su contraseña",
            label_visibility="collapsed",
            key="ia_inp_password"
        )

        st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)

        submitted = st.form_submit_button(
            "INGRESAR AL PANEL",
            use_container_width=True,
            type="primary"
        )

        # Footer dentro del form para que esté en la tarjeta
        st.markdown("""
        <div class="ia-footer">
          Desarrollado por Bilstein Group México<br>
          <strong>Bilstein Group</strong> © 2026
        </div>
        """, unsafe_allow_html=True)

        if submitted:
            if USUARIOS.get(usuario) == password:
                st.session_state.logged_in   = True
                st.session_state.usuario     = usuario
                st.session_state.login_error = False
                st.rerun()
            else:
                st.session_state.login_error = True
                st.rerun()

    return False