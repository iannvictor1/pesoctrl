import json
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
from config import BASE

ARQUIVO_SOMADOR = BASE / "somador_entrega.json"
PASTA_HISTORICO_SOMADOR = BASE / "historico_somador_entregas"

def aplicar_estilo_coletor():
    st.markdown(
        """
        <style>
        .somador-metric {
            background: #ffffff !important;
            border: 2px solid #111111 !important;
            border-radius: 6px !important;
            padding: 14px 16px !important;
            margin-bottom: 12px !important;
        }

        .somador-metric-label,
        .somador-metric-value {
            color: #000000 !important;
            font-weight: 800 !important;
            opacity: 1 !important;
        }

        .somador-metric-label {
            font-size: 15px !important;
            text-transform: uppercase !important;
        }

        .somador-metric-value {
            font-size: 34px !important;
            line-height: 1.1 !important;
            margin-top: 6px !important;
        }

        @media (max-width: 800px) {
            :root {
                --sidebar-w: 0px;
                --header-h: 0px;
            }

            .sidebar-impressoras,
            [style*="position:fixed"][style*="top:0"] {
                display: none !important;
            }

            .block-container,
            [data-testid="stMainBlockContainer"] {
                padding-left: 0.75rem !important;
                padding-right: 0.75rem !important;
                max-width: 100vw !important;
            }

            .somador-metric-value {
                font-size: 30px !important;
            }

            .stButton > button,
            .stDownloadButton > button,
            div[data-testid="stFormSubmitButton"] button {
                background: #ffffff !important;
                color: #000000 !important;
                border: 2px solid #000000 !important;
                font-weight: 800 !important;
                min-height: 48px !important;
            }

            .stTextInput label,
            .stTextInput input {
                color: #000000 !important;
                opacity: 1 !important;
                font-weight: 800 !important;
            }

            .stTextInput input {
                background: #ffffff !important;
                border: 2px solid #000000 !important;
                min-height: 48px !important;
                font-size: 20px !important;
            }

            [data-testid="stDataFrame"] {
                border: 2px solid #000000 !important;
            }

            [data-testid="stDataFrame"] * {
                color: #000000 !important;
                opacity: 1 !important;
                font-weight: 700 !important;
            }

            [data-testid="stDataFrame"] thead th,
            [data-testid="stDataFrame"] tbody td {
                background: #ffffff !important;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

def extrair_peso_etiqueta(codigo: str):
    codigo = codigo.strip().replace(",", ".")
    
    match = re.search(r"37([0-9]+(?:\.[0-9]+)?)\s*15", codigo)
    
    if not match:
        return None
    
    peso = match.group(1)
    
    if len(peso.split(".")[0]) > 2:
        peso = peso[-5:]
        
    try:
        return float(peso)
    except ValueError:
        return None
    
def carregar_leituras():
    if not ARQUIVO_SOMADOR.exists():
        return []
    
    try:
        with open(ARQUIVO_SOMADOR, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []
    
def salvar_leituras(leituras):
    with open(ARQUIVO_SOMADOR, "w", encoding="utf-8") as f:
        json.dump(leituras, f, ensure_ascii=False, indent=2)

def formatar_peso_csv(valor):
    try:
        return f"{float(valor):.3f}".replace(".", ",")
    except (TypeError, ValueError):
        return ""

def gerar_csv_somador(leituras):
    df_export = pd.DataFrame(leituras)
    df_export = df_export.rename(
        columns={
            "data_hora": "Data/Hora",
            "codigo": "Código",
            "peso": "Peso (KG)",
        }
    )

    if "Peso (KG)" in df_export.columns:
        df_export["Peso (KG)"] = df_export["Peso (KG)"].map(formatar_peso_csv)

    csv_texto = "sep=;\r\n" + df_export.to_csv(index=False, sep=";", lineterminator="\r\n")
    return csv_texto.encode("utf-8-sig")

def salvar_historico_entrega(leituras):
    if not leituras:
        return None

    PASTA_HISTORICO_SOMADOR.mkdir(parents=True, exist_ok=True)
    agora = datetime.now()
    nome_arquivo = f"entrega_{agora.strftime('%Y%m%d_%H%M%S_%f')}.csv"
    caminho = PASTA_HISTORICO_SOMADOR / nome_arquivo
    caminho.write_bytes(gerar_csv_somador(leituras))
    return caminho

def renderizar_somador_entrega():
    aplicar_estilo_coletor()
    st.subheader("Somador de Entrega")

    mensagem_historico = st.session_state.pop("mensagem_historico_somador", None)
    if mensagem_historico:
        st.success(mensagem_historico)
    
    leituras = carregar_leituras()
    
    with st.form("form_leitura_entrega", clear_on_submit=True):
        codigo = st.text_input(
            "Bipe a etiqueta",
            key="codigo_etiqueta_entrega",
            placeholder="Aguardando leitura..."
        )
        
        enviado = st.form_submit_button("Somar")
        
    if enviado and codigo:
        peso = extrair_peso_etiqueta(codigo)
        
        if peso is not None:
            leituras.append({
                "data_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "codigo": codigo,
                "peso": peso,
            })
            salvar_leituras(leituras)
            st.success(f"Peso adicionado: {peso:.2f} KG")
            st.rerun()
        else: 
            st.error("Peso não identificado nessa etiqueta.")
            
    total = sum(float(item.get("peso", 0)) for item in leituras)
    
    c1, c2 = st.columns(2)
    
    with c1:
        st.markdown(
            f"""
            <div class="somador-metric">
                <div class="somador-metric-label">Total acumulado</div>
                <div class="somador-metric-value">{total:.2f} KG</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        
    with c2:
        st.markdown(
            f"""
            <div class="somador-metric">
                <div class="somador-metric-label">Caixas lidas</div>
                <div class="somador-metric-value">{len(leituras)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        
    b1, b2, b3 = st.columns(3)
    
    with b1:
        if st.button("Desfazer última"):
            if leituras:
                leituras.pop()
                salvar_leituras(leituras)
                st.rerun()
                
    with b2:
        if st.button("Zerar entrega"):
            if leituras:
                try:
                    caminho_historico = salvar_historico_entrega(leituras)
                    salvar_leituras([])
                    st.session_state["mensagem_historico_somador"] = (
                        f"Entrega zerada e salva no histórico: {caminho_historico.name}"
                    )
                    st.rerun()
                except OSError as erro:
                    st.error(f"Não foi possível salvar o histórico. A entrega não foi zerada: {erro}")
    
    with b3:
        if leituras:
            csv = gerar_csv_somador(leituras)
            
            st.download_button(
                "Exportar CSV",
                data=csv,
                file_name="somador_entrega.csv",
                mime="text/csv; charset=utf-8"
            )
    
    if leituras:
        st.markdown("### Últimas leituras")
        df = pd.DataFrame(leituras)
        df = df.tail(100).iloc[::-1]
        st.dataframe(df, use_container_width=True)
    else:
        st.info("Nenhuma etiqueta lida ainda.")
    
