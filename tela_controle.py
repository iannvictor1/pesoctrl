import streamlit as st
import os
import html
from datetime import datetime
import pandas as pd
from streamlit_autorefresh import st_autorefresh

from config import CONFIG_IMPRESSORAS, montar_config
from monitor import (
    carregar_status,
    carregar_recebimento,
    iniciar_recebimento,
    encerrar_recebimento,
    iniciar_sessao,
    encerrar_sessao,
    novo_pallet,
    iniciar_monitor_processo,
    parar_monitor_processo,
    status_monitor,
    garantir_pastas,
    garantir_config_limpeza,
    carregar_config_limpeza,
    salvar_config_limpeza,
    executar_limpeza_automatica,
    existe_pallet_ativo_em_qualquer_impressora,
)

st.set_page_config(
    page_title="PesoCtrl — Controle de Pesagens",
    page_icon="⚙",
    layout="wide",
)
st_autorefresh(interval=5000, key="auto_refresh_pesagens_multi")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=Share+Tech+Mono&family=Barlow:wght@300;400;500;600&display=swap');

:root {
    --sidebar-w: 230px;
    --header-h: 52px;
    --content-pad-x: 0.75rem;
}

/* Base */
html, body, .stApp, [class*="css"] {
    font-family: 'Barlow', sans-serif !important;
    background-color: #0a0a0a !important;
    color: #f2eadb !important;
    margin: 0 !important;
    padding: 0 !important;
}

body::before {
    content: '';
    position: fixed;
    inset: 0;
    background: repeating-linear-gradient(
        0deg, transparent, transparent 2px,
        rgba(0,0,0,0.04) 2px, rgba(0,0,0,0.04) 4px
    );
    pointer-events: none;
    z-index: 9999;
}

#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"] {
    display: none !important;
    height: 0 !important;
    min-height: 0 !important;
    background: transparent !important;
}

/* Containers principais */
[data-testid="stAppViewContainer"] {
    background: #0a0a0a !important;
    margin: 0 !important;
    padding: 0 !important;
}

[data-testid="stMain"] {
    margin: 0 !important;
    padding: 0 !important;
    width: 100% !important;
}

main {
    margin: 0 !important;
    padding: 0 !important;
    background: #0a0a0a !important;
    width: 100% !important;
}

/* Conteúdo deslocado para não ficar atrás da sidebar */
.block-container {
    padding-top: 0 !important;
    padding-right: 1rem !important;
    padding-bottom: 0 !important;
    padding-left: calc(var(--sidebar-w) + 0.35rem) !important;
    margin: 0 !important;
    max-width: 100% !important;
}

[data-testid="stMainBlockContainer"] {
    padding-left: calc(var(--sidebar-w) + 0.35rem) !important;
    padding-right: 1rem !important;
    max-width: calc(100vw - var(--sidebar-w)) !important;
}

section.main {
    width: 100% !important;
}

section.main > div {
    width: 100% !important;
    max-width: 100% !important;
}

[data-testid="stVerticalBlock"] {
    gap: 0 !important;
}

div[data-testid="stHorizontalBlock"] {
    gap: 1.25rem !important;
}

/* Sidebar fixa */
.sidebar-impressoras {
    position: fixed;
    top: var(--header-h);
    left: 0;
    width: var(--sidebar-w);
    height: calc(100vh - var(--header-h));
    background: linear-gradient(180deg, #0d0d0d 0%, #111111 100%);
    border-right: 1px solid #1f1f1f;
    z-index: 9996;
    padding: 16px 10px;
    box-sizing: border-box;
    overflow-y: auto;
}

.sidebar-impressoras-titulo {
    font-family: 'Rajdhani', sans-serif;
    font-weight: 700;
    font-size: 16px;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #f2eadb;
    margin-bottom: 10px;
}

.sidebar-impressoras-subtitulo {
    font-family: 'Share Tech Mono', monospace;
    font-size: 10px;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #8f8f8f;
    margin-bottom: 14px;
}

.sidebar-impressora-card {
    background: linear-gradient(180deg, #121212 0%, #151515 100%);
    border: 1px solid #222222;
    border-radius: 12px;
    padding: 14px;
    margin-bottom: 14px;
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.02);
}

.sidebar-impressora-topo {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 10px;
}

.sidebar-impressora-icone {
    font-size: 26px;
    line-height: 1;
}

.sidebar-impressora-nome {
    font-family: 'Rajdhani', sans-serif;
    font-size: 18px;
    font-weight: 600;
    color: #f2eadb;
    line-height: 1.1;
}

.sidebar-status-linha {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-top: 2px;
}

.sidebar-status-bolinha {
    width: 12px;
    height: 12px;
    border-radius: 50%;
    flex-shrink: 0;
}

.sidebar-status-ativa {
    background: #10b981;
    box-shadow: 0 0 10px rgba(16,185,129,0.45);
}

.sidebar-status-inativa {
    background: #ef4444;
    box-shadow: 0 0 10px rgba(239,68,68,0.35);
}

.sidebar-status-texto {
    font-family: 'Share Tech Mono', monospace;
    font-size: 11px;
    letter-spacing: 2px;
    text-transform: uppercase;
}

.sidebar-status-texto.ativa {
    color: #6ee7b7;
}

.sidebar-status-texto.inativa {
    color: #fca5a5;
}

.sidebar-divisor {
    height: 1px;
    background: #222222;
    margin: 14px 0 20px 0;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: #0a0a0a !important;
    border-bottom: 1px solid #1a1a1a !important;
    gap: 0 !important;
    padding: 0 !important;
    margin: 0 !important;
}

.stTabs [data-baseweb="tab"] {
    font-family: 'Rajdhani', sans-serif !important;
    font-weight: 600 !important;
    font-size: 15px !important;
    letter-spacing: 2px !important;
    text-transform: uppercase !important;
    color: #8f8f8f !important;
    background: transparent !important;
    border: none !important;
    border-bottom: 2px solid transparent !important;
    padding: 14px 20px !important;
    transition: all 0.2s !important;
}

.stTabs [data-baseweb="tab"]:hover {
    color: #d7d7d7 !important;
    background: transparent !important;
}

.stTabs [aria-selected="true"] {
    color: #d97706 !important;
    border-bottom: 2px solid #d97706 !important;
    background: transparent !important;
}

.stTabs [data-baseweb="tab-highlight"],
.stTabs [data-baseweb="tab-border"] {
    display: none !important;
}

.stTabs [data-baseweb="tab-panel"] {
    padding: 1.5rem 0 2rem !important;
    background: #0a0a0a !important;
    margin: 0 !important;
}

/* Botões */
.stButton {
    margin-top: 0.25rem !important;
    margin-bottom: 0.75rem !important;
}

.stButton > button {
    font-family: 'Rajdhani', sans-serif !important;
    font-weight: 600 !important;
    font-size: 15px !important;
    letter-spacing: 2px !important;
    text-transform: uppercase !important;
    border-radius: 3px !important;
    min-height: 60px !important;
    padding: 14px 22px !important;
    transition: all 0.15s !important;
    width: 100% !important;
    background: linear-gradient(180deg, #111111 0%, #0c0c0c 100%) !important;
    border: 1px solid #2d2d2d !important;
    color: #c9c9c9 !important;
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.03);
}

.stButton > button:hover {
    background: rgba(217,119,6,0.08) !important;
    border-color: rgba(217,119,6,0.7) !important;
    color: #f0b35a !important;
}

.stButton > button:active {
    transform: scale(0.98) !important;
}

/* Inputs */
div[data-baseweb="input"] > div {
    background: #0f0f0f !important;
    border: 1px solid #2a2a2a !important;
}

div[data-baseweb="input"] input {
    background: #0f0f0f !important;
    color: #f2eadb !important;
    font-family: 'Share Tech Mono', monospace !important;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border: 1px solid #1a1a1a !important;
    border-radius: 4px !important;
    overflow: hidden !important;
}

[data-testid="stDataFrame"] thead th {
    font-family: 'Share Tech Mono', monospace !important;
    font-size: 10px !important;
    letter-spacing: 2px !important;
    background: #0a0a0a !important;
    color: #b1b1b1 !important;
    text-transform: uppercase !important;
    border-bottom: 1px solid #1a1a1a !important;
}

[data-testid="stDataFrame"] tbody td {
    font-family: 'Barlow', sans-serif !important;
    font-size: 14px !important;
    color: #d2d2d2 !important;
    background: #0f0f0f !important;
    border-bottom: 1px solid #141414 !important;
}

[data-testid="stDataFrame"] tbody tr:hover td {
    background: rgba(217,119,6,0.03) !important;
    color: #ffffff !important;
}

/* Outros */
.stAlert {
    font-family: 'Share Tech Mono', monospace !important;
    font-size: 11px !important;
    letter-spacing: 1px !important;
    border-radius: 3px !important;
    border-left-width: 3px !important;
    background: #0f0f0f !important;
}

.stCaption {
    font-family: 'Share Tech Mono', monospace !important;
    font-size: 11px !important;
    letter-spacing: 1px !important;
    color: #9b9b9b !important;
}

h2, h3, [data-testid="stHeading"] {
    font-family: 'Rajdhani', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: 2px !important;
    text-transform: uppercase !important;
    color: #f2eadb !important;
}

::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: #0a0a0a; }
::-webkit-scrollbar-thumb { background: #4a4a4a; border-radius: 2px; }
::-webkit-scrollbar-thumb:hover { background: #d97706; }
</style>
""", unsafe_allow_html=True)


def esc(texto):
    return html.escape(str(texto)) if texto is not None else ""


def abrir_caminho(caminho):
    try:
        os.startfile(str(caminho))
        return True, None
    except Exception as e:
        return False, str(e)


def section_label(text: str):
    text = esc(text)
    st.markdown(
        f"""
        <div style="
            font-family: 'Share Tech Mono', monospace;
            font-size: 11px;
            letter-spacing: 3px;
            color: #9a9a9a;
            text-transform: uppercase;
            margin: 1.6rem 0 1.2rem;
            display: flex;
            align-items: center;
            gap: 10px;
        ">
            {text}
            <div style="flex:1; height:1px; background:#2a2a2a;"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def status_banner(label: str, value: str, color: str, icon: str, subtitle: str = ""):
    label = esc(label)
    value = esc(value)
    icon = esc(icon)
    subtitle = esc(subtitle) if subtitle else ""

    subtitle_html = f"""<div style="
font-family:'Share Tech Mono', monospace;
font-size:11px;
color:#8f8f8f;
margin-top:10px;
letter-spacing:1px;
">{subtitle}</div>""" if subtitle else ""

    return f"""<div style="
background:linear-gradient(180deg, #111111 0%, #0d0d0d 100%);
border:1px solid #242424;
border-left:4px solid {color};
border-radius:6px;
padding:22px 26px;
min-height:118px;
box-sizing:border-box;
overflow:hidden;
box-shadow: inset 0 1px 0 rgba(255,255,255,0.03);
">
<div style="
display:flex;
align-items:center;
gap:18px;
">
<div style="
width:66px;
height:66px;
border-radius:8px;
background:rgba(255,255,255,0.035);
border:1px solid rgba(255,255,255,0.03);
display:flex;
align-items:center;
justify-content:center;
font-size:28px;
flex-shrink:0;
">{icon}</div>

<div style="min-width:0;">
<div style="
font-family:'Share Tech Mono', monospace;
font-size:10px;
letter-spacing:4px;
color:#a0a0a0;
text-transform:uppercase;
margin-bottom:8px;
">{label}</div>

<div style="
font-family:'Rajdhani', sans-serif;
font-weight:700;
font-size:22px;
color:{color};
line-height:1.05;
word-break:break-word;
">● {value}</div>

{subtitle_html}
</div>
</div>
</div>"""


def big_metric(label: str, value, unit: str = ""):
    label = esc(label)
    value = esc(value)
    unit = esc(unit)

    return f"""<div style="
background:#0f0f0f;
border:1px solid #1e1e1e;
border-radius:4px;
padding:22px;
text-align:center;
position:relative;
">
<div style="position:absolute;top:-1px;left:-1px;width:8px;height:8px;border-top:2px solid #d97706;border-left:2px solid #d97706;"></div>
<div style="position:absolute;top:-1px;right:-1px;width:8px;height:8px;border-top:2px solid #d97706;border-right:2px solid #d97706;"></div>
<div style="position:absolute;bottom:-1px;left:-1px;width:8px;height:8px;border-bottom:2px solid #d97706;border-left:2px solid #d97706;"></div>
<div style="position:absolute;bottom:-1px;right:-1px;width:8px;height:8px;border-bottom:2px solid #d97706;border-right:2px solid #d97706;"></div>

<div style="
font-family:'Share Tech Mono', monospace;
font-size:11px;
letter-spacing:3px;
color:#9a9a9a;
text-transform:uppercase;
margin-bottom:10px;
">{label}</div>

<div style="
font-family:'Rajdhani', sans-serif;
font-weight:700;
font-size:42px;
color:#f5f5f5;
line-height:1;
">{value}</div>

<div style="
font-family:'Share Tech Mono', monospace;
font-size:12px;
color:#b08b52;
margin-top:6px;
letter-spacing:2px;
">{unit}</div>
</div>"""


def card_config_info(titulo: str, valor: str, subtitulo: str = ""):
    titulo = esc(titulo)
    valor = esc(valor)
    subtitulo = esc(subtitulo)

    return f"""
    <div style="
        background:#0f0f0f;
        border:1px solid #1a1a1a;
        border-radius:4px;
        padding:16px;
        position:relative;
        min-height:95px;
    ">
        <div style="font-family:'Share Tech Mono',monospace;font-size:10px;letter-spacing:2px;color:#9a9a9a;text-transform:uppercase;margin-bottom:8px;">{titulo}</div>
        <div style="font-family:'Rajdhani',sans-serif;font-size:30px;font-weight:700;color:#f2eadb;line-height:1;">{valor}</div>
        <div style="font-family:'Share Tech Mono',monospace;font-size:10px;letter-spacing:1px;color:#b08b52;margin-top:8px;">{subtitulo}</div>
    </div>
    """


def pallet_info_banner(numero_pallet: int, descricao: str, recebimento_id: str):
    desc_exibe = descricao if descricao else "aguardando primeira etiqueta..."
    desc_exibe = esc(desc_exibe)
    recebimento_id = esc(recebimento_id)

    return f"""
    <div style="
        background: linear-gradient(135deg, #0f0f0f 0%, #141208 100%);
        border: 1px solid #3a2a0a;
        border-left: 4px solid #d97706;
        border-radius: 4px;
        padding: 16px 20px;
        display: flex;
        align-items: center;
        gap: 20px;
        margin-bottom: 4px;
    ">
        <div style="
            background: rgba(217,119,6,0.15);
            border: 1px solid rgba(217,119,6,0.4);
            border-radius: 4px;
            padding: 8px 16px;
            text-align: center;
            flex-shrink: 0;
        ">
            <div style="font-family:'Share Tech Mono',monospace;font-size:9px;letter-spacing:3px;color:#b08b52;text-transform:uppercase;">Pallet</div>
            <div style="font-family:'Rajdhani',sans-serif;font-weight:700;font-size:36px;color:#d97706;line-height:1;">{numero_pallet}</div>
        </div>
        <div style="flex:1; min-width:0;">
            <div style="font-family:'Share Tech Mono',monospace;font-size:9px;letter-spacing:3px;color:#8f8f8f;text-transform:uppercase;margin-bottom:6px;">Produto identificado</div>
            <div style="font-family:'Rajdhani',sans-serif;font-weight:600;font-size:20px;color:#f2eadb;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{desc_exibe}</div>
            <div style="font-family:'Share Tech Mono',monospace;font-size:10px;color:#6a6a6a;margin-top:4px;">Recebimento: {recebimento_id}</div>
        </div>
    </div>
    """


def ler_excel_resumo_e_pesagens(cfg: dict):
    df = pd.DataFrame()
    total_etiquetas = 0
    soma_pesos = 0

    if cfg["arquivo_excel_atual"].exists():
        try:
            df = pd.read_excel(cfg["arquivo_excel_atual"], sheet_name="Pesagens")
            df = df.reset_index(drop=True)
            df.index += 1
        except Exception:
            df = pd.DataFrame()

        try:
            resumo = pd.read_excel(cfg["arquivo_excel_atual"], sheet_name="Resumo", header=None)
            total_etiquetas = resumo.iloc[1, 1] if len(resumo) > 1 else 0
            soma_pesos = resumo.iloc[2, 1] if len(resumo) > 2 else 0
        except Exception:
            total_etiquetas = 0
            soma_pesos = 0

    return df, total_etiquetas, soma_pesos


def renderizar_config_limpeza(nome_impressora: str, cfg: dict):
    garantir_config_limpeza(cfg)
    conf = carregar_config_limpeza(cfg)

    section_label("Configuração da limpeza automática")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(card_config_info("Histórico de etiquetas", str(conf["dias_reter_etiquetas"]), "dias de retenção"), unsafe_allow_html=True)
    with c2:
        st.markdown(card_config_info("Histórico de pesagens", str(conf["dias_reter_pesagens"]), "dias de retenção"), unsafe_allow_html=True)
    with c3:
        st.markdown(card_config_info("Ciclo da limpeza", str(conf["intervalo_limpeza_automatica_segundos"]), "segundos"), unsafe_allow_html=True)

    st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

    f1, f2, f3 = st.columns(3)
    with f1:
        dias_etiquetas = st.number_input(
            "Dias para manter etiquetas",
            min_value=0,
            max_value=3650,
            value=int(conf["dias_reter_etiquetas"]),
            step=1,
            key=f"dias_etiquetas_{nome_impressora}"
        )
    with f2:
        dias_pesagens = st.number_input(
            "Dias para manter pesagens",
            min_value=0,
            max_value=3650,
            value=int(conf["dias_reter_pesagens"]),
            step=1,
            key=f"dias_pesagens_{nome_impressora}"
        )
    with f3:
        intervalo_limpeza = st.number_input(
            "Intervalo da limpeza (segundos)",
            min_value=10,
            max_value=86400,
            value=int(conf["intervalo_limpeza_automatica_segundos"]),
            step=10,
            key=f"intervalo_limpeza_{nome_impressora}"
        )

    b1, b2 = st.columns(2)
    with b1:
        if st.button("💾  Salvar configuração de limpeza", key=f"salvar_limpeza_{nome_impressora}"):
            novos_dados = {
                "dias_reter_etiquetas": int(dias_etiquetas),
                "dias_reter_pesagens": int(dias_pesagens),
                "intervalo_limpeza_automatica_segundos": int(intervalo_limpeza),
            }
            try:
                salvar_config_limpeza(cfg, novos_dados)
                st.success("Configuração de limpeza salva com sucesso.")
            except Exception as e:
                st.error(f"Erro ao salvar configuração: {e}")

    with b2:
        if st.button("🧹  Executar limpeza agora", key=f"executar_limpeza_agora_{nome_impressora}"):
            try:
                executar_limpeza_automatica(cfg)
                st.success("Limpeza automática executada com sucesso.")
            except Exception as e:
                st.error(f"Erro ao executar limpeza: {e}")

    st.caption("A limpeza automática afeta apenas os históricos. Ela não apaga o Excel atual, a fila ativa, a pasta de entrada nem o descarte da fila.")


def renderizar_sidebar_impressoras():
    cards_html = []

    for chave in CONFIG_IMPRESSORAS:
        cfg = montar_config(chave)
        garantir_pastas(cfg)

        status = carregar_status(cfg)
        sessao_ativa = status.get("sessao_ativa", False)

        nome_amigavel = esc(CONFIG_IMPRESSORAS[chave]["nome_amigavel"])
        status_classe = "ativa" if sessao_ativa else "inativa"
        status_texto = "ATIVA" if sessao_ativa else "INATIVA"
        bolinha_classe = "sidebar-status-ativa" if sessao_ativa else "sidebar-status-inativa"

        card_html = f"""<div class="sidebar-impressora-card">
<div class="sidebar-impressora-topo">
<div class="sidebar-impressora-icone">🖨️</div>
<div class="sidebar-impressora-nome">{nome_amigavel}</div>
</div>
<div class="sidebar-status-linha">
<div class="sidebar-status-bolinha {bolinha_classe}"></div>
<div class="sidebar-status-texto {status_classe}">{status_texto}</div>
</div>
</div>"""
        cards_html.append(card_html)

    sidebar_html = f"""<div class="sidebar-impressoras">
<div class="sidebar-impressoras-titulo">IMPRESSORAS</div>
<div class="sidebar-impressoras-subtitulo">STATUS EM TEMPO REAL</div>
<div class="sidebar-divisor"></div>
{''.join(cards_html)}
</div>"""

    st.markdown(sidebar_html, unsafe_allow_html=True)


def renderizar_bloco_recebimento_global():
    rec = carregar_recebimento()
    recebimento_ativo = rec.get("recebimento_ativo", False)
    recebimento_id = rec.get("recebimento_id") or "—"
    contador_pallets = rec.get("contador_pallets", 0)
    existe_pallet_ativo = existe_pallet_ativo_em_qualquer_impressora()

    section_label("Recebimento global — caminhão")

    cor = "#3b82f6" if recebimento_ativo else "#9ca3af"
    valor = "EM ANDAMENTO" if recebimento_ativo else "AGUARDANDO"
    icone = "🚚" if recebimento_ativo else "○"
    subtitulo = f"ID: {recebimento_id}  ·  Pallets totais: {contador_pallets}" if recebimento_ativo else ""

    st.markdown(
        f"""<div style="
background:linear-gradient(180deg, #111111 0%, #0d0d0d 100%);
border:1px solid #242424;
border-left:4px solid {cor};
border-radius:6px;
padding:22px 26px;
min-height:118px;
box-sizing:border-box;
overflow:hidden;
box-shadow: inset 0 1px 0 rgba(255,255,255,0.03);
margin-bottom:22px;
">
<div style="
display:flex;
align-items:center;
gap:18px;
">
<div style="
width:66px;
height:66px;
border-radius:8px;
background:rgba(255,255,255,0.035);
border:1px solid rgba(255,255,255,0.03);
display:flex;
align-items:center;
justify-content:center;
font-size:28px;
flex-shrink:0;
">{icone}</div>

<div style="min-width:0;">
<div style="
font-family:'Share Tech Mono', monospace;
font-size:10px;
letter-spacing:4px;
color:#a0a0a0;
text-transform:uppercase;
margin-bottom:8px;
">Recebimento global</div>

<div style="
font-family:'Rajdhani', sans-serif;
font-weight:700;
font-size:22px;
color:{cor};
line-height:1.05;
word-break:break-word;
">● {valor}</div>

<div style="
font-family:'Share Tech Mono', monospace;
font-size:11px;
color:#8f8f8f;
margin-top:10px;
letter-spacing:1px;
">{subtitulo}</div>
</div>
</div>
</div>""",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

    r1, r2 = st.columns(2)

    with r1:
        btn_iniciar_rec = st.button(
            "🚚  Iniciar recebimento",
            key="iniciar_recebimento_global",
            disabled=recebimento_ativo,
        )
        if btn_iniciar_rec:
            with st.spinner("Iniciando recebimento..."):
                ok, msg = iniciar_recebimento()
            if ok:
                st.success(msg)
            else:
                st.error(msg)
            st.rerun()

    with r2:
        btn_encerrar_rec = st.button(
            "🏁  Encerrar recebimento",
            key="encerrar_recebimento_global",
            disabled=not recebimento_ativo or existe_pallet_ativo,
        )
        if btn_encerrar_rec:
            with st.spinner("Encerrando recebimento..."):
                ok, msg = encerrar_recebimento()
            if ok:
                st.success(msg)
            else:
                st.error(msg)
            st.rerun()

    if recebimento_ativo and existe_pallet_ativo:
        st.caption("⚠ Encerre todos os pallets ativos antes de encerrar o recebimento.")
    elif not recebimento_ativo:
        st.caption("Inicie um recebimento global para liberar os controles de pallet.")


def renderizar_painel(nome_impressora: str):
    cfg = montar_config(nome_impressora)
    garantir_pastas(cfg)
    garantir_config_limpeza(cfg)

    status = carregar_status(cfg)
    rec = carregar_recebimento()
    monitor_rodando, pid = status_monitor(cfg)

    recebimento_ativo = rec.get("recebimento_ativo", False)
    sessao_ativa = status.get("sessao_ativa", False)
    numero_pallet = status.get("numero_pallet")
    descricao_produto = status.get("descricao_produto") or ""
    recebimento_id = status.get("recebimento_id") or rec.get("recebimento_id") or "—"

    section_label("Estado do sistema")

    col_rec, col_pal, col_mon = st.columns(3, gap="medium")

    with col_rec:
        st.markdown(
            status_banner(
                label="Recebimento global",
                value="EM ANDAMENTO" if recebimento_ativo else "AGUARDANDO",
                color="#3b82f6" if recebimento_ativo else "#9ca3af",
                icon="🚚" if recebimento_ativo else "○",
                subtitle=f"ID: {rec.get('recebimento_id')}" if recebimento_ativo else "",
            ),
            unsafe_allow_html=True,
        )

    with col_pal:
        st.markdown(
            status_banner(
                label="Pallet atual",
                value=f"PALLET {numero_pallet} — ATIVO" if sessao_ativa else "NENHUM ATIVO",
                color="#10b981" if sessao_ativa else "#ef4444",
                icon="📦" if sessao_ativa else "○",
                subtitle=descricao_produto if sessao_ativa and descricao_produto else "",
            ),
            unsafe_allow_html=True,
        )

    with col_mon:
        st.markdown(
            status_banner(
                label="Monitor",
                value="EM EXECUÇÃO" if monitor_rodando else "PARADO",
                color="#d97706" if monitor_rodando else "#9ca3af",
                icon="⚙" if monitor_rodando else "■",
                subtitle=f"PID: {pid}" if monitor_rodando and pid else "",
            ),
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)

    if sessao_ativa and numero_pallet:
        st.markdown(
            pallet_info_banner(numero_pallet, descricao_produto, recebimento_id),
            unsafe_allow_html=True,
        )

    section_label("Pallet")

    p1, p2 = st.columns(2)

    with p1:
        if not sessao_ativa:
            btn_iniciar_pal = st.button(
                "📦 Iniciar pallet",
                key=f"iniciar_{nome_impressora}",
                disabled=not recebimento_ativo,
            )
            if btn_iniciar_pal:
                with st.spinner("Iniciando pallet..."):
                    ok, msg = iniciar_sessao(cfg)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
                st.rerun()
        else:
            btn_novo_pal = st.button(
                "🔄  Novo pallet",
                key=f"novo_{nome_impressora}",
            )
            if btn_novo_pal:
                with st.spinner("Fechando pallet atual e iniciando o próximo..."):
                    ok, msg = novo_pallet(cfg)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
                st.rerun()

    with p2:
        btn_encerrar_pal = st.button(
            "✅  Encerrar pallet",
            key=f"encerrar_{nome_impressora}",
            disabled=not sessao_ativa,
        )
        if btn_encerrar_pal:
            with st.spinner("Encerrando pallet..."):
                ok, msg = encerrar_sessao(cfg)
            if ok:
                st.success(msg)
            else:
                st.error(msg)
            st.rerun()

    if not recebimento_ativo:
        st.caption("Inicie o recebimento global para liberar os controles de pallet.")

    section_label("Monitor de arquivos")

    m1, m2 = st.columns(2)

    with m1:
        if st.button("▶  Iniciar monitor", key=f"start_monitor_{nome_impressora}"):
            with st.spinner("Iniciando monitor..."):
                ok, msg = iniciar_monitor_processo(nome_impressora)
            if ok:
                st.success(msg)
            else:
                st.error(msg)
            st.rerun()

    with m2:
        if st.button("■  Parar monitor", key=f"stop_monitor_{nome_impressora}"):
            with st.spinner("Parando monitor..."):
                ok, msg = parar_monitor_processo(nome_impressora)
            if ok:
                st.success(msg)
            else:
                st.error(msg)
            st.rerun()

    section_label("Acesso rápido")

    a1, a2, a3, a4 = st.columns(4)
    with a1:
        if st.button("📄  Excel atual", key=f"excel_{nome_impressora}"):
            if cfg["arquivo_excel_atual"].exists():
                ok, erro = abrir_caminho(cfg["arquivo_excel_atual"])
                if not ok:
                    st.error(f"Erro: {erro}")
            else:
                st.warning("Excel atual não existe.")
    with a2:
        if st.button("🗂  Históricos", key=f"hist_{nome_impressora}"):
            if cfg["pasta_historico_pesagens"].exists():
                ok, erro = abrir_caminho(cfg["pasta_historico_pesagens"])
                if not ok:
                    st.error(f"Erro: {erro}")
            else:
                st.warning("Pasta de históricos não encontrada.")
    with a3:
        if st.button("📥  Entrada", key=f"entrada_{nome_impressora}"):
            if cfg["pasta_origem"].exists():
                ok, erro = abrir_caminho(cfg["pasta_origem"])
                if not ok:
                    st.error(f"Erro: {erro}")
            else:
                st.warning("Pasta de entrada não encontrada.")
    with a4:
        if st.button("📁  Pasta base", key=f"base_{nome_impressora}"):
            if cfg["base"].exists():
                ok, erro = abrir_caminho(cfg["base"])
                if not ok:
                    st.error(f"Erro: {erro}")
            else:
                st.warning("Pasta base não encontrada.")

    st.caption(
        f"Excel: {cfg['arquivo_excel_atual']}  ·  Históricos: {cfg['pasta_historico_pesagens']}  ·  Entrada: {cfg['pasta_origem']}  ·  Base: {cfg['base']}"
    )

    section_label("Métricas do pallet atual")

    df, total_etiquetas, soma_pesos = ler_excel_resumo_e_pesagens(cfg)

    try:
        val_etiquetas = int(total_etiquetas) if total_etiquetas == total_etiquetas else 0
    except (ValueError, TypeError):
        val_etiquetas = 0

    try:
        val_pesos = f"{float(soma_pesos):.3f}" if soma_pesos == soma_pesos else "0.000"
    except (ValueError, TypeError):
        val_pesos = "0.000"

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.markdown(big_metric("Total de etiquetas", val_etiquetas, "UNIDADES"), unsafe_allow_html=True)
    with col_m2:
        st.markdown(big_metric("Soma total dos pesos", val_pesos, "KG"), unsafe_allow_html=True)

    section_label("Pesagens do pallet atual")

    st.markdown(
        """
        <div style="
            background:#0f0f0f;
            border:1px solid #1a1a1a;
            border-radius:4px;
            padding:10px 16px 6px;
            margin-bottom:4px;
            display:flex; align-items:center; justify-content:space-between;
        ">
            <span style="font-family:'Share Tech Mono',monospace;font-size:11px;letter-spacing:3px;color:#a1a1a1;text-transform:uppercase;">Registros</span>
            <span style="font-family:'Share Tech Mono',monospace;font-size:10px;padding:2px 8px;background:rgba(217,119,6,0.1);border:1px solid rgba(217,119,6,0.3);color:#f0b35a;border-radius:2px;">LIVE</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if df.empty:
        st.markdown(
            """
            <div style="text-align:center;padding:40px;background:#0f0f0f;border:1px solid #1a1a1a;border-radius:4px;">
                <div style="font-size:28px;opacity:0.35;margin-bottom:10px;">⚖</div>
                <div style="font-family:'Share Tech Mono',monospace;font-size:11px;letter-spacing:3px;color:#9a9a9a;text-transform:uppercase;">
                    Nenhuma pesagem registrada
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.dataframe(df, width="stretch")

    section_label("Histórico de pallets")

    arquivos_historico = sorted(
        cfg["pasta_historico_pesagens"].glob("*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    if arquivos_historico:
        dados = [
            {
                "Arquivo": arq.name,
                "Modificado em": datetime.fromtimestamp(arq.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            }
            for arq in arquivos_historico[:20]
        ]
        st.dataframe(pd.DataFrame(dados), width="stretch", hide_index=True)
    else:
        st.markdown(
            """
            <div style="text-align:center;padding:30px;background:#0f0f0f;border:1px solid #1a1a1a;border-radius:4px;">
                <div style="font-family:'Share Tech Mono',monospace;font-size:11px;letter-spacing:3px;color:#9a9a9a;text-transform:uppercase;">
                    Nenhum histórico encontrado
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    renderizar_config_limpeza(nome_impressora, cfg)
    st.markdown("<div style='height:2rem;'></div>", unsafe_allow_html=True)


st.markdown("""<div style="
position:fixed;
top:0;
left:0;
right:0;
background:#0f0f0f;
border-bottom:1px solid #1e1e1e;
margin:0;
height:var(--header-h);
display:flex;
align-items:center;
z-index:10000;
box-sizing:border-box;
">
<div style="
width:var(--sidebar-w);
min-width:var(--sidebar-w);
height:100%;
padding:0 14px;
display:flex;
align-items:center;
gap:12px;
border-right:1px solid #1e1e1e;
box-sizing:border-box;
">
<div style="
width:28px;
height:28px;
background:#d97706;
clip-path:polygon(50% 0%,100% 25%,100% 75%,50% 100%,0% 75%,0% 25%);
flex-shrink:0;
"></div>
<span style="
font-family:'Rajdhani',sans-serif;
font-weight:700;
font-size:20px;
letter-spacing:3px;
text-transform:uppercase;
color:#f2eadb;
line-height:1;
">
PESO<span style="color:#d97706;">CTRL</span>
</span>
</div>

<div style="
flex:1;
height:100%;
display:flex;
align-items:center;
justify-content:space-between;
padding:0 1.5rem;
box-sizing:border-box;
">
<div style="font-family:'Share Tech Mono',monospace;font-size:11px;color:#9a9a9a;letter-spacing:2px;">
CONTROLE DE PESAGENS
</div>
<div style="display:flex;align-items:center;gap:8px;">
<div style="width:6px;height:6px;border-radius:50%;background:#d97706;animation:blink 1.5s ease-in-out infinite;"></div>
<span style="font-family:'Share Tech Mono',monospace;font-size:11px;color:#f0b35a;letter-spacing:2px;">
SISTEMA ATIVO
</span>
</div>
</div>
</div>
<div style="height:52px;"></div>
<style>
@keyframes blink { 0%,100%{opacity:1} 50%{opacity:0.2} }
</style>""", unsafe_allow_html=True)

if not CONFIG_IMPRESSORAS:
    st.error("Nenhuma impressora configurada em config.py.")
    st.stop()

renderizar_sidebar_impressoras()
renderizar_bloco_recebimento_global()

abas = st.tabs([CONFIG_IMPRESSORAS[k]["nome_amigavel"] for k in CONFIG_IMPRESSORAS])

for aba, nome_impressora in zip(abas, CONFIG_IMPRESSORAS):
    with aba:
        renderizar_painel(nome_impressora)