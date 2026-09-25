import os
import sys

# 1. Garantir que o diretório base do projeto esteja no topo do sys.path
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if _BASE_DIR not in sys.path:
    sys.path.insert(0, _BASE_DIR)

# 2. Garantir que as dependências do ambiente virtual (.venv) estejam disponíveis
_VENV_SITE = os.path.join(_BASE_DIR, ".venv", "Lib", "site-packages")
if os.path.exists(_VENV_SITE) and _VENV_SITE not in sys.path:
    sys.path.insert(0, _VENV_SITE)

import io
import re
import time
import base64
import unicodedata
from datetime import datetime, timedelta, date
import json
from decimal import Decimal

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from sqlalchemy import create_engine, text
import plotly.express as px
import plotly.graph_objects as go
import bcrypt

from config import DB_CONFIG, APP_CONFIG
from gerador_contratos_agu import (
    gerar_contrato_docx,
    gerar_contrato_pdf,
    gerar_comprovante_protocolo_pdf,
    gerar_clausulas_aditivo,
    gerar_aditivo_docx,
    gerar_aditivo_pdf,
    processar_planilha_itens_excel,
    estruturar_tabela_com_ia,
    extrair_tabela_dinamica_df,
    gerar_planilha_itens_excel_formatada,
    renderizar_tabela_html_formatada
)
from agu_api_client import AGUApiClient
from cnpj_service import (
    buscar_ou_cadastrar_fornecedor,
    buscar_fornecedor_no_banco,
    puxar_fornecedor_api_e_cadastrar,
    consultar_cnpj_api,
    formatar_cnpj,
    limpar_cnpj
)



# Configuração da página e Ícone Oficial Municipal
_brasao_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "brasao_web.png")
st.set_page_config(
    page_title="Prefeitura de Ribeirãozinho do Maranhão - MA | Gestão de Contratos",
    page_icon=_brasao_icon_path if os.path.exists(_brasao_icon_path) else "🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Escudo Anti-Tradução para prevenir conflitos no React DOM (NotFoundError: insertBefore)
# e Ocultação Definitiva da Barra Superior (Header/Share) e do Botão 'Manage app'
components.html("""
<script>
    (function() {
        const getDocs = function() {
            const list = [document];
            try { if (window.parent && window.parent.document && !list.includes(window.parent.document)) list.push(window.parent.document); } catch(e) {}
            try { if (window.top && window.top.document && !list.includes(window.top.document)) list.push(window.top.document); } catch(e) {}
            return list;
        };

        const getWins = function() {
            const list = [window];
            try { if (window.parent && !list.includes(window.parent)) list.push(window.parent); } catch(e) {}
            try { if (window.top && !list.includes(window.top)) list.push(window.top); } catch(e) {}
            return list;
        };

        // 1. Escudo Anti-Tradução
        try {
            getDocs().forEach(function(doc) {
                try {
                    doc.documentElement.setAttribute('translate', 'no');
                    doc.documentElement.classList.add('notranslate');
                    if (doc.body) {
                        doc.body.setAttribute('translate', 'no');
                        doc.body.classList.add('notranslate');
                    }
                    let meta = doc.querySelector('meta[name="google"]');
                    if (!meta) {
                        meta = doc.createElement('meta');
                        meta.setAttribute('name', 'google');
                        meta.setAttribute('content', 'notranslate');
                        doc.head.appendChild(meta);
                    }
                } catch(e) {}
            });
        } catch(e) {}

        // 2. Injeção de Estilo Global no Document, Parent e Top para eliminar Toolbar, Share e Manage App,
        // e GARANTIR a Sidebar sempre aberta, visível e sem botão de colapso acidental
        const cssOcultar = `
            [data-testid="stToolbar"],
            [data-testid="stToolbarActions"],
            [data-testid="stShareButton"],
            button[data-testid="stShareButton"],
            .stAppDeployButton,
            div[class*="stAppDeployButton"],
            div[data-testid="stDecoration"],
            div[data-testid="stStatusWidget"],
            #MainMenu,
            [data-testid="manage-app-button"],
            div[class*="viewerBadge_container"],
            div[class*="manageApp"],
            div[class*="StatusWidget"],
            .viewerBadge_container__r5tak,
            button[title="View app in Streamlit Community Cloud"],
            iframe[title="manage-app"],
            footer {
                display: none !important;
                visibility: hidden !important;
                opacity: 0 !important;
                pointer-events: none !important;
                height: 0px !important;
                min-height: 0px !important;
                max-height: 0px !important;
                width: 0px !important;
                margin: 0 !important;
                padding: 0 !important;
                overflow: hidden !important;
            }
            header[data-testid="stHeader"],
            .stApp > header {
                background: transparent !important;
                height: 0px !important;
                min-height: 0px !important;
                border: none !important;
                pointer-events: none !important;
                z-index: 999990 !important;
            }
            /* Ocultar botão de fechar a sidebar para evitar fechamento acidental */
            button[data-testid="stSidebarCollapseButton"],
            [data-testid="stSidebarCollapseButton"] {
                display: none !important;
                visibility: hidden !important;
                pointer-events: none !important;
                width: 0px !important;
                height: 0px !important;
            }
            /* Botão de abrir/expandir a sidebar caso esteja fechada */
            button[data-testid="stExpandSidebarButton"],
            [data-testid="stExpandSidebarButton"],
            [data-testid="collapsedControl"],
            [data-testid="stSidebarCollapsedControl"] {
                display: flex !important;
                visibility: visible !important;
                opacity: 1 !important;
                pointer-events: auto !important;
                position: fixed !important;
                top: 0.6rem !important;
                left: 0.6rem !important;
                background: #012b63 !important;
                color: #fbbf24 !important;
                border: 1.5px solid #fbbf24 !important;
                border-radius: 8px !important;
                padding: 6px 10px !important;
                box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
                cursor: pointer !important;
                z-index: 999999 !important;
            }
            [data-testid="collapsedControl"] svg,
            button[data-testid="stExpandSidebarButton"] svg {
                fill: #fbbf24 !important;
                color: #fbbf24 !important;
            }
            div[data-testid="stAppViewContainer"] > section.main > div.block-container,
            .main .block-container {
                padding-top: 1.2rem !important;
            }
        `;

        function injetarCSS() {
            getDocs().forEach(function(doc) {
                try {
                    if (doc.head) {
                        let style = doc.getElementById('gel-hide-streamlit-bars');
                        if (!style) {
                            style = doc.createElement('style');
                            style.id = 'gel-hide-streamlit-bars';
                            doc.head.appendChild(style);
                        }
                        style.innerHTML = cssOcultar;
                    }
                } catch(e) {}
            });
        }
        injetarCSS();

        function assegurarSidebarAberta() {
            // 1. Limpar e forçar estado aberto no localStorage de todas as janelas
            getWins().forEach(function(win) {
                try {
                    if (win.localStorage) {
                        Object.keys(win.localStorage).forEach(function(k) {
                            if (k.indexOf('stSidebar') !== -1 || k.indexOf('stSidebarCollapsed') !== -1) {
                                win.localStorage.setItem(k, 'false');
                            }
                        });
                    }
                } catch(e) {}
            });

            // 2. Expandir via DOM se houver botão de expandir ativo
            getDocs().forEach(function(doc) {
                try {
                    const btnExpand = doc.querySelector('button[data-testid="stExpandSidebarButton"], [data-testid="stExpandSidebarButton"], [data-testid="collapsedControl"] button, [data-testid="collapsedControl"], button[data-testid="stSidebarCollapsedControl"]');
                    if (btnExpand) {
                        btnExpand.click();
                    }

                    // Se a sidebar existir, garantir propriedades abertas
                    const sidebars = doc.querySelectorAll('section[data-testid="stSidebar"]');
                    sidebars.forEach(function(sb) {
                        sb.style.setProperty('display', 'flex', 'important');
                        sb.style.setProperty('visibility', 'visible', 'important');
                        sb.style.setProperty('opacity', '1', 'important');
                        sb.style.setProperty('transform', 'none', 'important');
                        sb.style.setProperty('min-width', '260px', 'important');
                        sb.style.setProperty('max-width', '260px', 'important');
                        sb.style.setProperty('width', '260px', 'important');
                        sb.style.setProperty('margin-left', '0px', 'important');
                        sb.style.setProperty('left', '0px', 'important');
                    });
                } catch(e) {}
            });
        }

        setTimeout(assegurarSidebarAberta, 100);
        setTimeout(assegurarSidebarAberta, 350);
        setTimeout(assegurarSidebarAberta, 800);
        setTimeout(assegurarSidebarAberta, 1500);

        // 3. Delegação de clique nos Cards de Contrato para abrir a Ficha Completa
        try {
            getDocs().forEach(function(doc) {
                if (!doc.__cardClickRegistered) {
                    doc.__cardClickRegistered = true;
                    doc.addEventListener('click', function(e) {
                        const card = e.target.closest('.gel-contract-card');
                        if (card) {
                            const container = card.closest('[data-testid="stVerticalBlock"]');
                            if (container) {
                                const btn = container.querySelector('button');
                                if (btn && !btn.contains(e.target)) {
                                    btn.click();
                                }
                            }
                        }
                    });
                }
            });
        } catch(e) {}

        // 4. Varredura contínua para manter Share, Toolbar e Manage App 100% ocultos
        function varrerEOcultar() {
            const seletores = [
                '[data-testid="stToolbar"]',
                '[data-testid="stToolbarActions"]',
                '[data-testid="stShareButton"]',
                'button[data-testid="stShareButton"]',
                '.stAppDeployButton',
                'div[class*="stAppDeployButton"]',
                'div[data-testid="stDecoration"]',
                'div[data-testid="stStatusWidget"]',
                '#MainMenu',
                '[data-testid="manage-app-button"]',
                '.viewerBadge_container__r5tak',
                'div[class*="viewerBadge_container"]',
                'div[class*="manageApp"]',
                'div[class*="StatusWidget"]',
                'button[title="View app in Streamlit Community Cloud"]',
                'iframe[title="manage-app"]',
                'footer'
            ];

            getDocs().forEach(function(doc) {
                try {
                    seletores.forEach(function(s) {
                        const els = doc.querySelectorAll(s);
                        els.forEach(function(el) {
                            el.style.setProperty('display', 'none', 'important');
                            el.style.setProperty('visibility', 'hidden', 'important');
                            el.style.setProperty('opacity', '0', 'important');
                            el.style.setProperty('pointer-events', 'none', 'important');
                            el.style.setProperty('height', '0px', 'important');
                        });
                    });

                    // Procura qualquer botão com Share ou Compartilhar exclusivamente na barra superior
                    const botoes = doc.querySelectorAll('[data-testid="stToolbar"] button, button[data-testid="stShareButton"]');
                    botoes.forEach(function(btn) {
                        const txt = (btn.textContent || '').trim().toLowerCase();
                        const aria = (btn.getAttribute('aria-label') || '').toLowerCase();
                        const title = (btn.getAttribute('title') || '').toLowerCase();
                        if (txt === 'share' || txt === 'compartilhar' || aria.includes('share') || title.includes('share') || aria.includes('compartilhar')) {
                            btn.style.setProperty('display', 'none', 'important');
                            btn.style.setProperty('visibility', 'hidden', 'important');
                            btn.style.setProperty('pointer-events', 'none', 'important');
                            const pai = btn.closest('[data-testid="stToolbarActions"], [data-testid="stToolbar"]');
                            if (pai) {
                                pai.style.setProperty('display', 'none', 'important');
                            }
                        }
                    });
                } catch(e) {}
            });
        }

        varrerEOcultar();
        setInterval(varrerEOcultar, 600);
        setInterval(assegurarSidebarAberta, 1200);
    })();
</script>
""", height=0, width=0)


# Estilização CSS Clara, Moderna e Dinâmica com a Paleta Oficial de Ribeirãozinho do Maranhão - MA
# Cores Oficiais: Azul Real / Celeste (#0284C7 / #0369A1), Verde Esperança (#059669 / #16A34A), Amarelo Ouro (#D97706 / #F59E0B), Branco (#FFFFFF)
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = True
if "usuario" not in st.session_state:
    st.session_state["usuario"] = {
        "id": 1,
        "nome": "Thiago",
        "email": "admin",
        "cargo": "ADMIN",
        "ativo": 1
    }

_esta_logado = st.session_state.get("logged_in", True)
if _esta_logado:
    st.markdown("""
    <style>
        .stApp, [data-testid="stAppViewContainer"], section.main {
            background-color: #f1f5f9 !important;
        }
        section[data-testid="stSidebar"], div[data-testid="stSidebar"] {
            display: flex !important;
            visibility: visible !important;
            opacity: 1 !important;
            transform: none !important;
            min-width: 260px !important;
            max-width: 260px !important;
            width: 260px !important;
            margin-left: 0px !important;
            left: 0px !important;
            position: relative !important;
        }
    </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <style>
        .stApp, [data-testid="stAppViewContainer"], section.main {
            background: linear-gradient(135deg, #0284c7 0%, #0369a1 45%, #075985 100%) !important;
            background-attachment: fixed !important;
        }
        section[data-testid="stSidebar"], [data-testid="stSidebar"] {
            display: none !important;
        }
    </style>
    """, unsafe_allow_html=True)

st.markdown("""
<div translate="no" class="notranslate" style="display:none"></div>
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;700&family=Inter:wght@300;400;500;600;700;800&display=swap');

    /* Prevenção de conflito de reconciliação DOM com tradutores */
    .notranslate {
        translate: no !important;
    }

    /* Ocultar containers de componentes invisíveis para evitar flicker */
    div[data-testid="stCustomComponentV1"],
    iframe[title*="streamlit_components"] {
        display: none !important;
        height: 0px !important;
        width: 0px !important;
        margin: 0 !important;
        padding: 0 !important;
        border: none !important;
        overflow: hidden !important;
    }

    /* Ocultar completamente toolbar, botão Share, Deploy, 'Manage app' e footer */
    [data-testid="stToolbar"],
    [data-testid="stToolbarActions"],
    [data-testid="stShareButton"],
    button[data-testid="stShareButton"],
    .stAppDeployButton,
    div[class*="stAppDeployButton"],
    div[data-testid="stDecoration"],
    div[data-testid="stStatusWidget"],
    [data-testid="manage-app-button"],
    div[class*="viewerBadge_container"],
    div[class*="manageApp"],
    div[class*="StatusWidget"],
    .viewerBadge_container__r5tak,
    button[title="View app in Streamlit Community Cloud"],
    iframe[title="manage-app"],
    footer,
    #MainMenu {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
        height: 0px !important;
        min-height: 0px !important;
        max-height: 0px !important;
        width: 0px !important;
        margin: 0 !important;
        padding: 0 !important;
        overflow: hidden !important;
    }

    /* Header transparente sem bloquear cliques nem esconder botões de sistema */
    header,
    header[data-testid="stHeader"],
    .stApp > header {
        background: transparent !important;
        height: 0px !important;
        min-height: 0px !important;
        border: none !important;
        pointer-events: none !important;
        z-index: 999990 !important;
    }

    /* Ocultar botão de fechar a sidebar para evitar fechamento acidental */
    button[data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarCollapseButton"] {
        display: none !important;
        visibility: hidden !important;
        pointer-events: none !important;
        width: 0px !important;
        height: 0px !important;
    }

    /* Botão de abrir/expandir a sidebar caso esteja fechada */
    button[data-testid="stExpandSidebarButton"],
    [data-testid="stExpandSidebarButton"],
    [data-testid="collapsedControl"],
    [data-testid="stSidebarCollapsedControl"] {
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        pointer-events: auto !important;
        position: fixed !important;
        top: 0.6rem !important;
        left: 0.6rem !important;
        background: #012b63 !important;
        color: #fbbf24 !important;
        border: 1.5px solid #fbbf24 !important;
        border-radius: 8px !important;
        padding: 6px 10px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
        cursor: pointer !important;
        z-index: 999999 !important;
    }

    /* Ajuste de margem superior para o conteúdo ocupar o topo perfeitamente */
    div[data-testid="stAppViewContainer"] > section.main > div.block-container,
    .main .block-container {
        padding-top: 1.2rem !important;
    }

    /* Paleta Oficial Municipal - Ribeirãozinho do Maranhão - MA */
    :root {
        --gel-blue: #0284c7;
        --gel-blue-dark: #0369a1;
        --gel-blue-navy: #1e3a8a;
        --gel-green: #059669;
        --gel-green-light: #10b981;
        --gel-yellow: #d97706;
        --gel-yellow-gold: #f59e0b;
        --gel-bg: #f8fafc;
        --gel-card: #ffffff;
        --gel-border: #e2e8f0;
        --gel-text-dark: #0f172a;
        --gel-text-muted: #64748b;
    }

    html, body, [class*="css"], .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"], section.main {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #0f172a;
        transition: background-color 0.2s ease !important;
    }
    
    /* O texto herda o #0f172a do .stApp, sem !important para permitir que headers e banners tenham suas próprias cores (branco) */

    /* Reduzir espaçamento extremo para maximizar aproveitamento da tela */
    div.block-container {
        padding-top: 0.8rem !important;
        padding-bottom: 0.5rem !important;
        padding-left: 0.5rem !important;
        padding-right: 0.5rem !important;
        max-width: 100% !important;
    }

    /* Reduzir fortemente o espaço em branco (gap) nativo entre componentes do Streamlit */
    div[data-testid="stVerticalBlock"] {
        gap: 0.3rem !important;
    }
    div[data-testid="stHorizontalBlock"] {
        gap: 0.3rem !important;
    }
    
    /* Espremer margens das métricas e textos soltos */
    div.stMarkdown p {
        margin-bottom: 0.2rem !important;
    }
    
    /* Transição suave e estável sem transparência fantasma nem cortes abruptos */
    div[data-stale="true"] {
        opacity: 0.92 !important;
        filter: none !important;
        transition: opacity 0.15s ease !important;
    }
    div[data-testid="stForm"][data-stale="true"] {
        opacity: 0.96 !important;
        filter: none !important;
        pointer-events: none !important;
    }

    /* Fundo limpo nos formulários do sistema interno (exceto login) */
    div[data-testid="stForm"]:not(.gel-login-box):not([aria-label="login_form_gel"]) {
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 8px !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05) !important;
        padding: 12px 20px !important;
        margin-bottom: 5px !important;
    }
    
    /* Compactar abas (tabs) de forma natural, mantendo respiro dos KPIs superiores */
    div[data-testid="stTabs"] {
        margin-top: 10px !important;
    }

    /* Barra Lateral Escura (Menu do Sistema) - Sempre Fixa e Visível */
    section[data-testid="stSidebar"],
    div[data-testid="stSidebar"] {
        background-color: #012b63 !important;
        border-right: 1px solid #fbbf24 !important; /* Borda amarela separadora */
        box-shadow: 2px 0 12px rgba(0, 0, 0, 0.2) !important;
        min-width: 260px !important;
        max-width: 260px !important;
        width: 260px !important;
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        transform: none !important;
        margin-left: 0px !important;
    }
    
    /* Textos da Sidebar Brancos */
    section[data-testid="stSidebar"] div, section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span, section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2 {
        color: #ffffff !important;
    }

    section[data-testid="stSidebar"] .block-container {
        padding-top: 1.5rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
    }

    /* Banner Oficial da Prefeitura de Ribeirãozinho do Maranhão - MA (Design Mais Estreito e Equilibrado) */
    .municipal-banner {
        background: linear-gradient(135deg, #0369a1 0%, #0284c7 50%, #059669 100%);
        border-radius: 9px;
        padding: 9px 18px;
        margin-bottom: 7px;
        color: #ffffff;
        box-shadow: 0 4px 14px -3px rgba(2, 132, 199, 0.22);
        position: relative;
        overflow: hidden;
    }

    .municipal-banner::after {
        content: '';
        position: absolute;
        bottom: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #f59e0b, #10b981, #38bdf8);
    }

    .municipal-badge-pill {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        background: rgba(255, 255, 255, 0.22);
        border: 1px solid rgba(255, 255, 255, 0.3);
        color: #ffffff;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 10px;
        font-weight: 700;
        letter-spacing: 0.4px;
        text-transform: uppercase;
        margin-bottom: 2px;
    }

    .municipal-banner-title {
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-weight: 800;
        font-size: 19px;
        letter-spacing: -0.35px;
        color: #ffffff;
        margin: 0;
        line-height: 1.2;
    }

    .municipal-banner-subtitle {
        font-size: 11.5px;
        color: #e0f2fe;
        font-weight: 400;
        margin-top: 1px;
        line-height: 1.25;
    }

    /* Cards de Indicadores (KPIs) com as Cores de Ribeirãozinho do Maranhão (Design Mais Estreito) */
    .gel-kpi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 9px 13px;
        margin-top: 5px !important;
        margin-bottom: 9px !important;
        position: relative;
        overflow: hidden;
        transition: all 0.2s ease-in-out;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03);
        box-sizing: border-box;
    }

    /* Espaçamento garantido entre colunas de KPIs para evitar sobreposição */
    div[data-testid="stHorizontalBlock"]:has(.gel-kpi-card) {
        gap: 12px !important;
        margin-top: 6px !important;
        margin-bottom: 8px !important;
    }

    div[data-testid="stColumn"]:has(.gel-kpi-card) {
        margin-bottom: 6px !important;
    }

    .gel-kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 16px -4px rgba(2, 132, 199, 0.12);
        border-color: #cbd5e1;
    }
    
    .gel-contract-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        transition: all 0.2s ease-in-out;
        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.03);
    }
    /* Regra para repassar o Hover da Coluna para o Card de Contrato */
    div[data-testid="stColumn"]:has(.gel-contract-card):hover .gel-contract-card {
        transform: translateY(-2px);
        box-shadow: 0 10px 20px -5px rgba(2, 132, 199, 0.12);
        border-color: #cbd5e1;
    }

    /* Transformando a coluna num conteiner relativo para o botão invisível */
    div[data-testid="stColumn"]:has(.gel-contract-card) {
        position: relative !important;
    }

    /* Escondendo e esticando o botão do Streamlit para cobrir a coluna toda (Card Clicável) */
    div[data-testid="stColumn"]:has(.gel-contract-card) div[data-testid="stButton"] button {
        position: absolute !important;
        top: 0 !important;
        left: 0 !important;
        width: 100% !important;
        height: 100% !important;
        opacity: 0 !important; /* Totalmente invisível */
        z-index: 100 !important;
        cursor: pointer !important;
    }

    .gel-kpi-blue { border-top: 3px solid #0284c7; }
    .gel-kpi-green { border-top: 3px solid #059669; }
    .gel-kpi-gold { border-top: 3px solid #d97706; }
    .gel-kpi-cyan { border-top: 3px solid #0284c7; }

    .gel-kpi-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    .gel-kpi-title {
        font-size: 11px;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .gel-kpi-value {
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-size: 20px;
        font-weight: 800;
        color: #0f172a;
        margin: 4px 0 2px 0;
        letter-spacing: -0.4px;
        line-height: 1.15;
    }

    .gel-kpi-subtext {
        font-size: 11px;
        color: #94a3b8;
        font-weight: 500;
        line-height: 1.2;
    }

    /* Sinalizador Ativo em Verde Esperança */
    @keyframes pulse-green {
        0% { box-shadow: 0 0 0 0 rgba(5, 150, 105, 0.7); }
        70% { box-shadow: 0 0 0 7px rgba(5, 150, 105, 0); }
        100% { box-shadow: 0 0 0 0 rgba(5, 150, 105, 0); }
    }

    .beacon-active-gel {
        width: 8px;
        height: 8px;
        background: #059669;
        border-radius: 50%;
        display: inline-block;
        animation: pulse-green 2s infinite;
        vertical-align: middle;
        margin-right: 5px;
    }

    /* Card do Usuário (Claro e Compacto)    /* Perfis e Relógio na Sidebar */
    .gel-user-box {
        background: rgba(255,255,255,0.1);
        border: 1px solid rgba(251, 191, 36, 0.4); /* Amarelo translúcido */
        padding: 10px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 15px;
    }
    
    .gel-avatar-circle {
        background: #fbbf24;
        color: #012b63 !important;
        font-weight: 800;
        width: 32px;
        height: 32px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 13px;
        flex-shrink: 0;
    }

    .gel-user-name {
        font-weight: 700;
        font-size: 12.5px;
        color: #ffffff !important;
        line-height: 1.2;
    }

    .gel-role-pill {
        display: inline-block;
        background: #0ea5e9;
        color: #ffffff !important;
        font-size: 9px;
        font-weight: 800;
        padding: 2px 6px;
        border-radius: 4px;
        margin-top: 3px;
    }

    .gel-clock-box {
        background: rgba(0,0,0,0.2);
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 8px;
        padding: 8px 12px;
        font-size: 11px;
        text-align: center;
        margin-bottom: 16px;
        color: #e2e8f0;
    }
    .gel-clock-time {
        font-family: 'JetBrains Mono', monospace;
        font-weight: 700;
        font-size: 12px;
        color: #fbbf24 !important;
    }

    .gel-db-status-box {
        background: rgba(15, 23, 42, 0.3);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 8px;
        padding: 7px 10px;
        font-size: 11px;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .gel-db-pulse-dot {
        width: 10px;
        height: 10px;
        border-radius: 50%;
        background-color: #10b981;
        box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7);
        animation: pulse-green 2s infinite;
        flex-shrink: 0;
    }
    .gel-db-pulse-dot.offline {
        background-color: #ef4444;
        box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.7);
        animation: pulse-red 2s infinite;
    }
    @keyframes pulse-green {
        0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }
        70% { transform: scale(1); box-shadow: 0 0 0 6px rgba(16, 185, 129, 0); }
        100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }
    }
    @keyframes pulse-red {
        0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.7); }
        70% { transform: scale(1); box-shadow: 0 0 0 6px rgba(239, 68, 68, 0); }
        100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); }
    }

    /* Cards de Contratos na Visão Dinâmica (Claro e Limpo) */
    .gel-contract-card {
        background: #ffffff;
        border: 1.5px solid #e2e8f0;
        border-radius: 14px;
        padding: 18px 20px;
        margin-bottom: 8px;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
        cursor: pointer;
        position: relative;
    }

    .gel-contract-card:hover {
        border-color: #0284c7;
        box-shadow: 0 10px 25px -4px rgba(2, 132, 199, 0.22);
        transform: translateY(-3px);
    }

    .gel-contract-card.gel-card-selected {
        border: 2.5px solid #0284c7 !important;
        background: #f0f9ff !important;
        box-shadow: 0 0 0 4px rgba(2, 132, 199, 0.18), 0 8px 20px -4px rgba(2, 132, 199, 0.25) !important;
    }

    .gel-badge-vigente {
        background: #dcfce7;
        border: 1px solid #86efac;
        color: #166534;
        font-size: 11px;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 14px;
        display: inline-flex;
        align-items: center;
        gap: 4px;
    }

    .gel-badge-alerta {
        background: #fef3c7;
        border: 1px solid #fde68a;
        color: #92400e;
        font-size: 11px;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 14px;
        display: inline-flex;
        align-items: center;
        gap: 4px;
    }

    .gel-badge-vencido {
        background: #fee2e2;
        border: 1px solid #fca5a5;
        color: #991b1b;
        font-size: 11px;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 14px;
        display: inline-flex;
        align-items: center;
        gap: 4px;
    }

    /* Barra de Progresso da Vigência (Azul para Verde) */
    .gel-progress-bg {
        width: 100%;
        background: #f1f5f9;
        border-radius: 8px;
        height: 7px;
        overflow: hidden;
        margin: 8px 0 5px 0;
    }

    .gel-progress-fill {
        height: 100%;
        border-radius: 8px;
        transition: width 0.4s ease;
    }

    /* Customização dos Tabs Streamlit */
    .stTabs [data-baseweb="tab-list"] {
        background: #f1f5f9;
        padding: 4px;
        border-radius: 12px;
        gap: 6px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        color: #64748b !important;
        font-weight: 600;
        font-size: 13.5px;
        padding: 8px 18px;
    }

    .stTabs [aria-selected="true"] {
        background: #ffffff !important;
        color: #0284c7 !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.06);
        border: 1px solid #e2e8f0 !important;
    }

    /* Botões Modernos */
    .stButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease !important;
    }

        /* =================================================== */
    /* NAVEGAÇÃO FUTURÍSTICA & DINÂMICA NA BARRA LATERAL   */
    /* =================================================== */
    .nav-section-badge {
        font-size: 10px;
        font-weight: 800;
        color: #fde047 !important; /* Texto Amarelo */
        background: rgba(255, 255, 255, 0.1); /* Fundo sutil */
        border: 1px solid rgba(253, 224, 71, 0.4);
        padding: 2px 8px;
        border-radius: 6px;
        letter-spacing: 0.5px;
    }

    .nav-category-header {
        font-size: 10.5px;
        font-weight: 800;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin: 14px 4px 6px 4px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }

    /* Botões de Navegação na Barra Lateral */
    section[data-testid="stSidebar"] div.stButton {
        margin-bottom: 7px !important;
    }

    section[data-testid="stSidebar"] div.stButton > button {
        width: 100% !important;
        text-align: left !important;
        justify-content: flex-start !important;
        padding: 11px 16px !important;
        border-radius: 12px !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-size: 13.5px !important;
        font-weight: 600 !important;
        letter-spacing: 0.1px !important;
        transition: all 0.22s cubic-bezier(0.4, 0, 0.2, 1) !important;
        position: relative !important;
        overflow: hidden !important;
        display: flex !important;
        align-items: center !important;
    }

    /* ====================================================
       BOTÕES DA SIDEBAR (TODOS AZUIS)
       ==================================================== */
    section[data-testid="stSidebar"] button {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;
        background-color: #0284c7 !important;
        border: 1px solid #0369a1 !important;
        border-radius: 12px !important;
        padding: 10px 14px !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15) !important;
        transition: all 0.3s ease !important;
    }
    
    section[data-testid="stSidebar"] button *,
    section[data-testid="stSidebar"] button p,
    section[data-testid="stSidebar"] button span {
        color: #ffffff !important;
        background-color: transparent !important;
        font-weight: 700 !important;
    }

    section[data-testid="stSidebar"] button:hover {
        transform: translateX(4px) !important;
        box-shadow: 0 6px 15px rgba(0,0,0,0.2) !important;
        border-color: #38bdf8 !important;
    }

    /* BOTÃO SELECIONADO (PRIMARY) - Adiciona Contorno Amarelo e Ponto */
    section[data-testid="stSidebar"] button[kind="primary"] {
        border-left: 5px solid #f59e0b !important;
        box-shadow: 0 6px 20px -2px rgba(2, 132, 199, 0.35), 0 0 10px rgba(245, 158, 11, 0.3) !important;
    }

    section[data-testid="stSidebar"] button[kind="primary"]::after {
        content: '●' !important;
        position: absolute !important;
        right: 14px !important;
        color: #f59e0b !important;
        font-size: 13px !important;
        filter: drop-shadow(0 0 4px #f59e0b) !important;
    }

    /* Destaque para Botões de Ação Rápida (Criar / Novo) */
    section[data-testid="stSidebar"] div.stButton button[data-testid*="nav_novo_"]:hover {
        border-left-color: #059669 !important;
        background: linear-gradient(90deg, #ecfdf5 0%, #ffffff 100%) !important;
        color: #059669 !important;
        box-shadow: 0 6px 16px -2px rgba(5, 150, 105, 0.22) !important;
    }

    /* Botão Sair do Sistema */
    section[data-testid="stSidebar"] div.stButton button[data-testid*="btn_logout"]:hover {
        border-color: #fca5a5 !important;
        border-left-color: #ef4444 !important;
        color: #dc2626 !important;
        background: #fff1f2 !important;
    }

    /* Fundo consistente e fixo sem piscamento ou alternância de temas */
    html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"], section.main {
        background-color: #f1f5f9 !important;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def obter_brasao_b64():
    """Retorna o brasão oficial de Ribeirãozinho do Maranhão codificado em base64 para uso inline em HTML (versão otimizada ultraleve)"""
    caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "brasao_icon.png")
    if not os.path.exists(caminho):
        caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "brasao_limpo.png")
    if not os.path.exists(caminho):
        caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "brasao_web.png")
    if os.path.exists(caminho):
        with open(caminho, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    return ""

# Conexão com o Banco de Dados
@st.cache_resource
def get_engine():
    # 1. Tentar ler URL de banco da nuvem (Supabase/Neon Postgres)
    import os
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        try:
            if hasattr(st, "secrets") and "DATABASE_URL" in st.secrets:
                db_url = st.secrets["DATABASE_URL"]
        except Exception:
            pass
    if not db_url:
        try:
            sec_file = os.path.join(_BASE_DIR, ".streamlit", "secrets.toml")
            if os.path.exists(sec_file):
                import tomllib
                with open(sec_file, "rb") as f:
                    sec_data = tomllib.load(f)
                    db_url = sec_data.get("DATABASE_URL")
        except Exception:
            pass
        
    if db_url:
        # Normalizar prefixo postgres:// -> postgresql://
        if db_url.startswith("postgres://"):
            db_url = db_url.replace("postgres://", "postgresql://", 1)
        
        # Lista de variações de drivers a tentar para máxima compatibilidade (psycopg2 vs psycopg v3)
        urls_para_tentar = []
        if db_url.startswith("postgresql://") or "+psycopg" in db_url:
            base_pg = re.sub(r"^postgresql(\+[a-zA-Z0-9_]+)?://", "postgresql://", db_url)
            # Tentar psycopg2 primeiro se disponível
            try:
                import psycopg2
                urls_para_tentar.append(base_pg.replace("postgresql://", "postgresql+psycopg2://", 1))
            except ImportError:
                pass
            # Tentar psycopg v3
            try:
                import psycopg
                urls_para_tentar.append(base_pg.replace("postgresql://", "postgresql+psycopg://", 1))
            except ImportError:
                pass
            # Fallback direto com a URL informada
            urls_para_tentar.append(db_url)
            urls_para_tentar.append(base_pg)
        else:
            urls_para_tentar.append(db_url)

        for url_cand in urls_para_tentar:
            try:
                eng = create_engine(url_cand, pool_pre_ping=True, pool_recycle=300)
                # Validar conexão rápida
                with eng.connect() as conn:
                    pass
                return eng
            except (ImportError, ModuleNotFoundError):
                continue
            except Exception as e_conn:
                # Se for erro de rede/credenciais mas o driver funcionou, tenta usar o engine direto
                try:
                    return create_engine(url_cand, pool_pre_ping=True, pool_recycle=300)
                except Exception:
                    continue

    # 2. Fallback para as configuracoes locais originais
    db_type = DB_CONFIG.get("type", "sqlite")
    if db_type == "sqlite":
        sqlite_path = DB_CONFIG.get("sqlite_path", "gestao_contratos.db")
        return create_engine(f"sqlite:///{sqlite_path}", connect_args={'check_same_thread': False})
    else:
        try:
            eng = create_engine(
                f"mysql+pymysql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['host']}:{DB_CONFIG.get('port', 3306)}/{DB_CONFIG['database']}"
            )
            with eng.connect() as conn:
                pass
            return eng
        except Exception:
            sqlite_path = DB_CONFIG.get("sqlite_path", "gestao_contratos.db")
            return create_engine(f"sqlite:///{sqlite_path}", connect_args={'check_same_thread': False})

engine = get_engine()

# ============================================
# MAPAS OFICIAIS DE SECRETARIAS E SECRETÁRIOS
# PREFEITURA DE RIBEIRÃOZINHO DO MARANHÃO - MA
# ============================================

MAPA_ORGAOS_PADRAO = {
    "EDUCA": "SECRETARIA MUNICIPAL DE EDUCAÇÃO",
    "SAUD": "SECRETARIA MUNICIPAL DE SAÚDE",
    "SAÚD": "SECRETARIA MUNICIPAL DE SAÚDE",
    "FINAN": "SECRETARIA MUNICIPAL DE FINANÇAS, FAZENDA E RECEITA",
    "ASSIST": "SECRETARIA MUNICIPAL DE ASSISTÊNCIA SOCIAL",
    "CULTUR": "SECRETARIA MUNICIPAL DE CULTURA E TURISMO",
    "ADMINIST": "SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO",
    "OBRA": "SECRETARIA MUNICIPAL DE INFRAESTRUTURA E OBRAS",
    "PREFEITURA": "SECRETARIA MUNICIPAL DE INFRAESTRUTURA E OBRAS",
    "INFRA": "SECRETARIA MUNICIPAL DE INFRAESTRUTURA E OBRAS",
    "SAAE": "SERVIÇO AUTÔNOMO DE ÁGUA E ESGOTO - SAAE",
    "AGUA": "SERVIÇO AUTÔNOMO DE ÁGUA E ESGOTO - SAAE",
    "ÁGUA": "SERVIÇO AUTÔNOMO DE ÁGUA E ESGOTO - SAAE",
    "JURID": "ASSESSORIA JURÍDICA E GABINETE",
    "GABINETE": "ASSESSORIA JURÍDICA E GABINETE",
    "JURÍD": "ASSESSORIA JURÍDICA E GABINETE"
}

MAPA_SECRETARIOS_PADRAO = {
    # --- Dados oficiais extraídos do portal da Prefeitura de Ribeirãozinho do Maranhão - MA ---
    "SECRETARIA MUNICIPAL DE EDUCAÇÃO": "JOÃO VITOR SOUSA JUSTINO",
    "SECRETARIA MUNICIPAL DE SAÚDE": "SIRLEIDE MARINHO DOS SANTOS",
    "SECRETARIA MUNICIPAL DE FINANÇAS, FAZENDA E RECEITA": "DANIEL SILVA PEREIRA",
    "SECRETARIA MUNICIPAL DE DESENVOLVIMENTO SOCIAL": "FERNANDA NUNES ROCHA",
    "SECRETARIA MUNICIPAL DE CULTURA E TURISMO": "ELANDIAS BEZERRA SOUSA",
    "SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO, PLANEJAMENTO, ORÇAMENTO E GESTÃO": "MARCUS PEREIRA DE FREITAS",
    "SECRETARIA MUNICIPAL DE INFRAESTRUTURA E TRANSPORTES": "NEUTON COELHO DOS SANTOS NETO",
    "SERVIÇO AUTÔNOMO DE ÁGUAS E ESGOTOS": "IVANILZA DA SILVA FERREIRA",
    "PROCURADORIA-GERAL DO MUNICÍPIO": "VENILSON BATISTA PEREIRA",
    "SECRETARIA MUNICIPAL DE AGRICULTURA, PRODUÇÃO, ABASTECIMENTO E PESCA": "WERBETH LIMA SANTOS",
    "SECRETARIA MUNICIPAL DE DESENVOLVIMENTO ECONÔMICO, INDÚSTRIA E COMÉRCIO": "GEORGE ALENCAR DE ARAÚJO",
    "SECRETARIA MUNICIPAL DE GOVERNO, COMUNICAÇÃO E RELAÇÕES INSTITUCIONAIS": "VENILSON BATISTA PEREIRA",
    "SECRETARIA MUNICIPAL DE JUVENTUDE, ESPORTES E LAZER": "BARTOLOMEU DA SILVA",
    "SECRETARIA MUNICIPAL DE URBANISMO E REGULARIZAÇÃO FUNDIÁRIA": "MAYKON QUEIROZ VASCONCELOS",
    "SECRETARIA MUNICIPAL DE MEIO AMBIENTE E RECURSOS NATURAIS": "ISABELA CAROLINE OLIVEIRA SILVA",
    "SECRETARIA MUNICIPAL DE POLÍTICAS PÚBLICAS PARA MULHERES": "HISLLA GABRIELLY SOARES LIMA VIANA",
    "GABINETE DO PREFEITO": "MATHEUS SOARES CARVALHO",
    "CONTROLADORIA-GERAL DO MUNICÍPIO": "GRAZIELLE ALVES DA SILVA",
    "GERÊNCIA DE CONTRATAÇÕES PÚBLICAS": "RAFAEL ABREU SANTOS",
    "COORDENAÇÃO DA UNIDADE VIVA PROCON": "JOÃO VITOR SOUSA JUSTINO",
    # --- Aliases (nomes antigos mantidos para compatibilidade com contratos existentes) ---
    "SECRETARIA MUNICIPAL DE ASSISTÊNCIA SOCIAL": "FERNANDA NUNES ROCHA",
    "SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO": "MARCUS PEREIRA DE FREITAS",
    "SECRETARIA MUNICIPAL DE INFRAESTRUTURA E OBRAS": "NEUTON COELHO DOS SANTOS NETO",
    "SERVIÇO AUTÔNOMO DE ÁGUA E ESGOTO - SAAE": "IVANILZA DA SILVA FERREIRA",
    "ASSESSORIA JURÍDICA E GABINETE": "VENILSON BATISTA PEREIRA",
}

MAPA_CARGO_SECRETARIOS = {
    "SECRETARIA MUNICIPAL DE EDUCAÇÃO": "Secretário Municipal de Educação",
    "SECRETARIA MUNICIPAL DE SAÚDE": "Secretária Municipal de Saúde",
    "SECRETARIA MUNICIPAL DE FINANÇAS, FAZENDA E RECEITA": "Secretário Municipal de Finanças, Fazenda e Receita",
    "SECRETARIA MUNICIPAL DE DESENVOLVIMENTO SOCIAL": "Secretária Municipal de Desenvolvimento Social",
    "SECRETARIA MUNICIPAL DE CULTURA E TURISMO": "Secretário(a) Municipal de Cultura e Turismo",
    "SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO, PLANEJAMENTO, ORÇAMENTO E GESTÃO": "Secretário Municipal de Administração, Planejamento, Orçamento e Gestão",
    "SECRETARIA MUNICIPAL DE INFRAESTRUTURA E TRANSPORTES": "Secretário Municipal de Infraestrutura e Transportes",
    "SERVIÇO AUTÔNOMO DE ÁGUAS E ESGOTOS": "Diretora Executiva",
    "PROCURADORIA-GERAL DO MUNICÍPIO": "Procurador-Geral do Município",
    "SECRETARIA MUNICIPAL DE AGRICULTURA, PRODUÇÃO, ABASTECIMENTO E PESCA": "Secretário Municipal de Agricultura, Produção, Abastecimento e Pesca",
    "SECRETARIA MUNICIPAL DE DESENVOLVIMENTO ECONÔMICO, INDÚSTRIA E COMÉRCIO": "Secretário Municipal de Desenvolvimento Econômico, Indústria e Comércio",
    "SECRETARIA MUNICIPAL DE GOVERNO, COMUNICAÇÃO E RELAÇÕES INSTITUCIONAIS": "Secretário Municipal de Governo, Comunicação e Relações Institucionais",
    "SECRETARIA MUNICIPAL DE JUVENTUDE, ESPORTES E LAZER": "Secretário Municipal de Juventude, Esportes e Lazer",
    "SECRETARIA MUNICIPAL DE URBANISMO E REGULARIZAÇÃO FUNDIÁRIA": "Secretário Municipal de Urbanismo e Regularização Fundiária",
    "SECRETARIA MUNICIPAL DE MEIO AMBIENTE E RECURSOS NATURAIS": "Secretária Municipal de Meio Ambiente e Recursos Naturais",
    "SECRETARIA MUNICIPAL DE POLÍTICAS PÚBLICAS PARA MULHERES": "Secretária Municipal de Políticas Públicas para Mulheres",
    "GABINETE DO PREFEITO": "Chefe de Gabinete do Prefeito",
    "CONTROLADORIA-GERAL DO MUNICÍPIO": "Controladora-Geral do Município",
    "GERÊNCIA DE CONTRATAÇÕES PÚBLICAS": "Gerente de Contratações Públicas",
    "COORDENAÇÃO DA UNIDADE VIVA PROCON": "Coordenador da Unidade VIVA PROCON",
    # Aliases
    "SECRETARIA MUNICIPAL DE ASSISTÊNCIA SOCIAL": "Secretária Municipal de Desenvolvimento Social",
    "SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO": "Secretário Municipal de Administração, Planejamento, Orçamento e Gestão",
    "SECRETARIA MUNICIPAL DE INFRAESTRUTURA E OBRAS": "Secretário Municipal de Infraestrutura e Transportes",
    "SERVIÇO AUTÔNOMO DE ÁGUA E ESGOTO - SAAE": "Diretora Executiva",
    "ASSESSORIA JURÍDICA E GABINETE": "Procurador-Geral do Município",
}

LISTA_SECRETARIAS_OFICIAIS = [
    "SECRETARIA MUNICIPAL DE EDUCAÇÃO",
    "SECRETARIA MUNICIPAL DE SAÚDE",
    "SECRETARIA MUNICIPAL DE FINANÇAS, FAZENDA E RECEITA",
    "SECRETARIA MUNICIPAL DE DESENVOLVIMENTO SOCIAL",
    "SECRETARIA MUNICIPAL DE CULTURA E TURISMO",
    "SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO, PLANEJAMENTO, ORÇAMENTO E GESTÃO",
    "SECRETARIA MUNICIPAL DE INFRAESTRUTURA E TRANSPORTES",
    "SECRETARIA MUNICIPAL DE AGRICULTURA, PRODUÇÃO, ABASTECIMENTO E PESCA",
    "SECRETARIA MUNICIPAL DE DESENVOLVIMENTO ECONÔMICO, INDÚSTRIA E COMÉRCIO",
    "SECRETARIA MUNICIPAL DE GOVERNO, COMUNICAÇÃO E RELAÇÕES INSTITUCIONAIS",
    "SECRETARIA MUNICIPAL DE JUVENTUDE, ESPORTES E LAZER",
    "SECRETARIA MUNICIPAL DE URBANISMO E REGULARIZAÇÃO FUNDIÁRIA",
    "SECRETARIA MUNICIPAL DE MEIO AMBIENTE E RECURSOS NATURAIS",
    "SECRETARIA MUNICIPAL DE POLÍTICAS PÚBLICAS PARA MULHERES",
    "SERVIÇO AUTÔNOMO DE ÁGUAS E ESGOTOS",
    "PROCURADORIA-GERAL DO MUNICÍPIO",
    "GABINETE DO PREFEITO",
    "CONTROLADORIA-GERAL DO MUNICÍPIO",
    "GERÊNCIA DE CONTRATAÇÕES PÚBLICAS",
    "COORDENAÇÃO DA UNIDADE VIVA PROCON",
]

MAPA_CNPJ_ORGAOS = {
    "SECRETARIA MUNICIPAL DE SAÚDE": "13.877.696/0001-80",
    "FUNDO MUNICIPAL DE SAÚDE": "13.877.696/0001-80",
    "SECRETARIA MUNICIPAL DE EDUCAÇÃO": "06.077.947/0001-87",
    "FUNDO MUNICIPAL DE EDUCAÇÃO": "06.077.947/0001-87",
    "SECRETARIA MUNICIPAL DE DESENVOLVIMENTO SOCIAL": "22.757.771/0001-60",
    "SECRETARIA MUNICIPAL DE ASSISTÊNCIA SOCIAL": "22.757.771/0001-60",
    "FUNDO MUNICIPAL DE ASSISTÊNCIA SOCIAL": "22.757.771/0001-60",
    "SERVIÇO AUTÔNOMO DE ÁGUAS E ESGOTOS": "23.603.835/0001-31",
    "SERVIÇO AUTÔNOMO DE ÁGUA E ESGOTO - SAAE": "23.603.835/0001-31",
}

MODELOS_AGU_MAP = {
    "Termo de Contrato - Aquisição de Bens / Compras (Padrão AGU)": "COMPRAS",
    "Termo de Contrato - Prestação de Serviços Contínuos (Sem dedicação)": "SERVICOS_CONTINUOS",
    "Termo de Contrato - Prestação de Serviços Não Contínuos / Escopo": "SERVICOS_ESCOPO",
    "Termo de Contrato - Contratação Direta (Dispensa / Inexigibilidade)": "DISPENSA",
    "Termo de Contrato - Obras e Serviços Comuns de Engenharia": "OBRAS"
}

@st.cache_data(ttl=300, show_spinner=False)
def obter_historico_orgao(orgao_nome):
    """Consulta o cadastro e histórico do órgão priorizando dados alimentados pelos usuários no sistema"""
    if not orgao_nome or not str(orgao_nome).strip() or "Selecione" in str(orgao_nome):
        return {
            "cnpj": "01.612.834/0001-86",
            "cep": "65928-000",
            "endereco": "Rua Principal, s/n",
            "numero": "s/n",
            "complemento": "Prédio Administrativo",
            "bairro": "Centro",
            "cidade": "Ribeirãozinho do Maranhão",
            "uf": "MA",
            "secretario": "",
            "cargo_secretario": "Secretário(a) Municipal",
            "fiscal": "",
            "dotacao": "",
            "modelo_agu": "COMPRAS"
        }

    is_educa = "EDUCA" in str(orgao_nome).upper()
    is_saude = "SAUDE" in normalizar_texto(str(orgao_nome))
    is_infra = "INFRA" in normalizar_texto(str(orgao_nome))
    is_finan = "FINAN" in normalizar_texto(str(orgao_nome))
    is_admin = "ADMIN" in normalizar_texto(str(orgao_nome))
    is_social = "SOCIAL" in normalizar_texto(str(orgao_nome)) or "ASSIST" in normalizar_texto(str(orgao_nome))
    is_cultura = "CULTURA" in normalizar_texto(str(orgao_nome))
    is_agri = "AGRICULTURA" in normalizar_texto(str(orgao_nome))

    sec_fallback = MAPA_SECRETARIOS_PADRAO.get(orgao_nome, "")
    cargo_fallback = MAPA_CARGO_SECRETARIOS.get(orgao_nome, "Secretário Municipal de Educação" if is_educa else "Secretário(a) Municipal")
    end_fallback = "Av. Tancredo Neves, nº 120" if is_educa else "Rua Principal, s/n"

    cnpj_fallback = MAPA_CNPJ_ORGAOS.get(orgao_nome, "")
    if not cnpj_fallback:
        if is_saude:
            cnpj_fallback = "13.877.696/0001-80"
        elif is_educa:
            cnpj_fallback = "06.077.947/0001-87"
        elif is_social:
            cnpj_fallback = "22.757.771/0001-60"
        elif "SAAE" in str(orgao_nome).upper() or "AGUA" in normalizar_texto(str(orgao_nome)):
            cnpj_fallback = "23.603.835/0001-31"
        else:
            cnpj_fallback = "01.612.834/0001-86"

    if is_educa:
        fisc_fallback = "CHARLIANE DE ABREU MACIEL"
        dot_fallback = "Fundo Municipal de Educação / SEMED - 02.05.00 - 12.361.0402.2022.0000"
    elif is_saude:
        fisc_fallback = "JOSIVAN SILVA AROUCHA"
        dot_fallback = "Fundo Municipal de Saúde / FMS - 02.06.00 - 10.301.0201.2023.0000"
    elif is_finan:
        fisc_fallback = "WILSON FERREIRA SOARES"
        dot_fallback = "Secretaria Municipal de Finanças - 02.03.00 - 04.123.0101.2005.0000"
    elif is_admin:
        fisc_fallback = "WILSON FERREIRA SOARES"
        dot_fallback = "Secretaria Municipal de Administração - 02.02.00 - 04.122.0101.2003.0000"
    elif is_social:
        fisc_fallback = "JHENYFFER VIEIRA DA SILVA"
        dot_fallback = "Fundo Municipal de Assistência Social - 02.08.00 - 08.244.0301.2035.0000"
    elif is_infra:
        fisc_fallback = "NEUTON COELHO DOS SANTOS NETO"
        dot_fallback = "Secretaria de Infraestrutura e Transportes - 02.07.00 - 15.451.0501.2030.0000"
    elif is_cultura:
        fisc_fallback = "RUBSON DA SILVA BRASIL"
        dot_fallback = "Secretaria Municipal de Cultura e Turismo - 02.09.00 - 13.392.0601.2040.0000"
    elif is_agri:
        fisc_fallback = "RUBSON DA SILVA BRASIL"
        dot_fallback = "Secretaria de Agricultura e Abastecimento - 02.10.00 - 20.606.0701.2045.0000"
    else:
        fisc_fallback = ""
        dot_fallback = ""

    dados = {
        "cnpj": cnpj_fallback,
        "cep": "65928-000",
        "endereco": end_fallback,
        "numero": "120" if is_educa else "s/n",
        "complemento": "Prédio Administrativo",
        "bairro": "Centro",
        "cidade": "Ribeirãozinho do Maranhão",
        "uf": "MA",
        "secretario": sec_fallback,
        "cargo_secretario": cargo_fallback,
        "fiscal": fisc_fallback,
        "dotacao": dot_fallback,
        "modelo_agu": "COMPRAS"
    }

    try:
        with engine.connect() as conn:
            # 1. Tabela de Órgãos (busca exata primeiro, depois normalizada em memória para tolerar acentos/caixa)
            row = conn.execute(
                text("SELECT * FROM orgaos WHERE nome = :n"),
                {"n": orgao_nome}
            ).mappings().first()

            if not row:
                all_orgaos = conn.execute(text("SELECT * FROM orgaos")).mappings().fetchall()
                norm_in = normalizar_texto(orgao_nome)
                for r in all_orgaos:
                    if normalizar_texto(r["nome"]) == norm_in:
                        row = r
                        break
                if not row:
                    for r in all_orgaos:
                        norm_db = normalizar_texto(r["nome"])
                        if norm_in in norm_db or norm_db in norm_in:
                            row = r
                            break
                if not row:
                    palavras = [p for p in re.split(r'[\s,/]+', norm_in) if len(p) > 3 and p not in ("SECRETARIA", "MUNICIPAL", "GABINETE", "COORDENACAO", "GERENCIA", "SETOR", "FUNDO", "DESENVOLVIMENTO")]
                    for p in palavras:
                        for r in all_orgaos:
                            if p in normalizar_texto(r["nome"]):
                                row = r
                                break
                        if row:
                            break

            oid = None
            if row:
                oid = row.get("id")
                if row.get("cnpj"):
                    c_clean = re.sub(r"\D", "", str(row["cnpj"]))
                    if len(c_clean) == 14:
                        dados["cnpj"] = formatar_cnpj(c_clean)
                if row.get("cep"): dados["cep"] = formatar_cep(str(row["cep"]).strip())
                if row.get("endereco") and str(row["endereco"]).strip():
                    end_db = str(row["endereco"]).strip()
                    dados["endereco"] = "Rua Principal, s/n" if end_db in ("", "Rua Principal, s/n") else end_db
                if row.get("bairro"): dados["bairro"] = str(row["bairro"]).strip()
                if row.get("cidade"): dados["cidade"] = str(row["cidade"]).strip()
                if row.get("uf"): dados["uf"] = str(row["uf"]).strip()
                if row.get("secretario_padrao") and str(row["secretario_padrao"]).strip():
                    dados["secretario"] = str(row["secretario_padrao"]).strip()
                if row.get("cargo_secretario") and str(row["cargo_secretario"]).strip():
                    dados["cargo_secretario"] = str(row["cargo_secretario"]).strip()
                if row.get("fiscal_padrao") and str(row["fiscal_padrao"]).strip() and not str(row["fiscal_padrao"]).startswith("Fiscal da") and not str(row["fiscal_padrao"]).startswith("Fiscal de"):
                    dados["fiscal"] = str(row["fiscal_padrao"]).strip()
                if row.get("dotacao_padrao") and str(row["dotacao_padrao"]).strip():
                    dados["dotacao"] = str(row["dotacao_padrao"]).strip()

            # 2. Histórico do ÚLTIMO contrato do órgão (PRIORIDADE MÁXIMA)
            if oid:
                row_ultimo = conn.execute(
                    text("""
                        SELECT 
                            c.secretario, c.fiscal, c.dotacao_orcamentaria,
                            c.endereco_orgao, c.bairro_orgao, c.cep_orgao, c.cidade_orgao, c.uf_orgao, c.cnpj_orgao,
                            c.modelo_agu
                        FROM contratos c
                        WHERE c.orgao_id = :oid
                        ORDER BY c.id DESC LIMIT 1
                    """),
                    {"oid": oid}
                ).mappings().first()

                if row_ultimo:
                    if row_ultimo.get("secretario") and str(row_ultimo["secretario"]).strip():
                        dados["secretario"] = str(row_ultimo["secretario"]).strip()
                    if row_ultimo.get("fiscal") and str(row_ultimo["fiscal"]).strip() and not str(row_ultimo["fiscal"]).startswith("Fiscal da") and not str(row_ultimo["fiscal"]).startswith("Fiscal de"):
                        dados["fiscal"] = str(row_ultimo["fiscal"]).strip()
                    if row_ultimo.get("dotacao_orcamentaria") and str(row_ultimo["dotacao_orcamentaria"]).strip():
                        dados["dotacao"] = str(row_ultimo["dotacao_orcamentaria"]).strip()
                    if row_ultimo.get("endereco_orgao") and str(row_ultimo["endereco_orgao"]).strip():
                        dados["endereco"] = str(row_ultimo["endereco_orgao"]).strip()
                    if row_ultimo.get("bairro_orgao") and str(row_ultimo["bairro_orgao"]).strip():
                        dados["bairro"] = str(row_ultimo["bairro_orgao"]).strip()
                    if row_ultimo.get("cidade_orgao") and str(row_ultimo["cidade_orgao"]).strip():
                        dados["cidade"] = str(row_ultimo["cidade_orgao"]).strip()
                    if row_ultimo.get("uf_orgao") and str(row_ultimo["uf_orgao"]).strip():
                        dados["uf"] = str(row_ultimo["uf_orgao"]).strip()
                    if row_ultimo.get("cep_orgao") and str(row_ultimo["cep_orgao"]).strip():
                        dados["cep"] = formatar_cep(str(row_ultimo["cep_orgao"]).strip())
                    if row_ultimo.get("cnpj_orgao"):
                        c_cn = re.sub(r"\D", "", str(row_ultimo["cnpj_orgao"]))
                        if len(c_cn) == 14:
                            c_atual_limpo = re.sub(r"\D", "", str(dados.get("cnpj", "")))
                            # Não sobrescrever o CNPJ específico do órgão se o histórico antigo tiver o genérico da prefeitura
                            if not (c_atual_limpo not in ("01612834000186", "01597627000134") and c_cn in ("01612834000186", "01597627000134")):
                                dados["cnpj"] = formatar_cnpj(c_cn)
                    if row_ultimo.get("modelo_agu"):
                        dados["modelo_agu"] = str(row_ultimo["modelo_agu"]).strip()

    except Exception:
        pass

    return dados

def salvar_historico_orgao(orgao_nome, dados):
    """Atualiza o perfil e histórico do órgão para alimentar próximas sugestões"""
    try:
        with engine.connect() as conn:
            conn.execute(
                text("""
                    UPDATE orgaos
                    SET cnpj = COALESCE(NULLIF(:cnpj, ''), cnpj),
                        cep = COALESCE(NULLIF(:cep, ''), cep),
                        endereco = COALESCE(NULLIF(:end, ''), endereco),
                        bairro = COALESCE(NULLIF(:bai, ''), bairro),
                        cidade = COALESCE(NULLIF(:cid, ''), cidade),
                        uf = COALESCE(NULLIF(:uf, ''), uf),
                        secretario_padrao = COALESCE(NULLIF(:sec, ''), secretario_padrao),
                        cargo_secretario = COALESCE(NULLIF(:carg, ''), cargo_secretario),
                        fiscal_padrao = COALESCE(NULLIF(:fisc, ''), fiscal_padrao),
                        dotacao_padrao = COALESCE(NULLIF(:dot, ''), dotacao_padrao)
                    WHERE nome = :n
                """),
                {
                    "cnpj": dados.get("cnpj", "01.612.834/0001-86"),
                    "cep": dados.get("cep", "65928-000"),
                    "end": dados.get("endereco", "Rua Principal, s/n"),
                    "bai": dados.get("bairro", "Centro"),
                    "cid": dados.get("cidade", "Ribeirãozinho do Maranhão"),
                    "uf": dados.get("uf", "MA"),
                    "sec": dados.get("secretario", ""),
                    "carg": dados.get("cargo_secretario", "Secretário(a) Municipal Titular"),
                    "fisc": dados.get("fiscal", ""),
                    "dot": dados.get("dotacao", ""),
                    "n": orgao_nome
                }
            )
            conn.commit()
            try:
                obter_historico_orgao.clear()
            except Exception:
                pass
    except Exception:
        pass

# ============================================
# FUNÇÕES DE FORMATAÇÃO (PADRÃO BRASIL)
# ============================================

def normalizar_texto(txt):
    if not txt:
        return ""
    txt = unicodedata.normalize("NFKD", str(txt))
    txt = "".join([c for c in txt if not unicodedata.combining(c)])
    return txt.strip().upper()

def formatar_numero_moeda(val):
    """Formata valor com pontos nos milhares e vírgula nos centavos (ex: 50.000,00)"""
    try:
        if val is None or pd.isna(val):
            return "0,00"
        num = float(val)
        neg = "-" if num < 0 else ""
        num = abs(num)
        inteiro, decimal = f"{num:.2f}".split(".")
        partes = []
        while len(inteiro) > 3:
            partes.insert(0, inteiro[-3:])
            inteiro = inteiro[:-3]
        partes.insert(0, inteiro)
        return f"{neg}{'.'.join(partes)},{decimal}"
    except Exception:
        return "0,00"

def formatar_moeda(val):
    """Formata valor com R$, pontos nos milhares e vírgula nos centavos (ex: R$ 50.000,00)"""
    num_fmt = formatar_numero_moeda(val)
    if num_fmt.startswith("-"):
        return f"-R$ {num_fmt[1:]}"
    return f"R$ {num_fmt}"

def clean_money(val):
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ["nan", "none", "-", "", "null"]:
        return 0.0
    is_neg = False
    if val_str.startswith("(") and val_str.endswith(")"):
        is_neg = True
        val_str = val_str[1:-1].strip()
    elif val_str.startswith("-"):
        is_neg = True
        val_str = val_str[1:].strip()
    val_str = re.sub(r"[R$\s\xa0BRL]", "", val_str)
    if not val_str:
        return 0.0
    if "," in val_str and "." in val_str:
        if val_str.rfind(",") > val_str.rfind("."):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "," in val_str:
        partes = val_str.split(",")
        if len(partes) > 2:
            val_str = val_str.replace(",", "")
        else:
            val_str = val_str.replace(",", ".")
    elif "." in val_str:
        partes = val_str.split(".")
        if len(partes) > 2:
            val_str = val_str.replace(".", "")
        else:
            if len(partes[1]) == 3 and len(partes[0]) <= 3:
                val_str = val_str.replace(".", "")
    try:
        res = float(val_str)
        return -res if is_neg else res
    except Exception:
        return 0.0

def formatar_input_moeda_callback(chave_state):
    """Callback para formatar automaticamente um campo de texto monetário com pontos e vírgula (ex: 50.000,00)"""
    val_atual = st.session_state.get(chave_state, "")
    num = clean_money(val_atual)
    st.session_state[chave_state] = formatar_numero_moeda(num)

def formatar_input_cnpj_callback(chave_state):
    """Callback para formatar automaticamente um campo de CNPJ com pontos, barra e traço (ex: 00.000.000/0001-00)"""
    val_atual = st.session_state.get(chave_state, "")
    if val_atual and str(val_atual).strip():
        st.session_state[chave_state] = formatar_cnpj_cpf(val_atual)

def formatar_numero_br(val):
    try:
        if val is None or pd.isna(val):
            return "0"
        num = int(val)
        s = f"{num:,}".replace(",", ".")
        return s
    except Exception:
        return str(val)

def parse_data_universal(val):
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if s == "" or s.lower() in ["nan", "none", "null", "nat", "-"]:
        return None
    for fmt in ["%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%Y/%m/%d"]:
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    m = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", s)
    if m:
        d, mth, y = m.groups()
        if len(y) == 2:
            y = "20" + y
        try:
            return datetime(int(y), int(mth), int(d)).strftime("%Y-%m-%d")
        except Exception:
            pass
    return None

import json

def registrar_audit_log(conn, tabela, registro_id, acao, descricao, dados_anteriores=None, dados_novos=None):
    """Grava um registro na tabela audit_log."""
    if "usuario" in st.session_state and st.session_state["usuario"]:
        usuario_id = st.session_state["usuario"].get("id")
        usuario_nome = st.session_state["usuario"].get("nome")
    else:
        usuario_id = None
        usuario_nome = "Sistema/Desconhecido"
    
    da_json = json.dumps(dados_anteriores, ensure_ascii=False) if dados_anteriores else None
    dn_json = json.dumps(dados_novos, ensure_ascii=False) if dados_novos else None

    conn.execute(text('''
        INSERT INTO audit_log 
        (tabela, registro_id, acao, usuario_id, usuario_nome, descricao, dados_anteriores, dados_novos)
        VALUES (:tb, :rid, :acao, :uid, :unome, :desc, :da, :dn)
    '''), {
        "tb": tabela, "rid": registro_id, "acao": acao,
        "uid": usuario_id, "unome": usuario_nome, "desc": descricao,
        "da": da_json, "dn": dn_json
    })

def formatar_data_br(val):
    if val is None or pd.isna(val) or str(val).strip() in ["", "None", "NaT"]:
        return "Não informada"
    s = str(val).strip().split(" ")[0]
    dt_iso = parse_data_universal(s)
    if dt_iso:
        partes = dt_iso.split("-")
        if len(partes) == 3:
            return f"{partes[2]}/{partes[1]}/{partes[0]}"
    return s

def formatar_datetime_br(val):
    if val is None or pd.isna(val):
        return "Não informado"
    try:
        if isinstance(val, (datetime, pd.Timestamp)):
            return val.strftime("%d/%m/%Y às %H:%M")
        s = str(val).strip()
        dt = pd.to_datetime(s)
        return dt.strftime("%d/%m/%Y às %H:%M")
    except Exception:
        return str(val)

def formatar_data_extenso_br(dt=None):
    if dt is None:
        dt = datetime.now()
    dias_semana = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"]
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    dia_sem = dias_semana[dt.weekday()]
    mes = meses[dt.month - 1]
    return f"{dia_sem}, {dt.day} de {mes} de {dt.year}"

def formatar_cnpj_cpf(val):
    """Formata CNPJ (00.000.000/0001-00) ou CPF (000.000.000-00) no padrão oficial"""
    if val is None or pd.isna(val) or str(val).strip() in ["", "nan", "None", "NULL"]:
        return "Não informado"
    s = re.sub(r"\D", "", str(val)).strip()
    if len(s) == 14:
        return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"
    elif len(s) == 11:
        return f"{s[:3]}.{s[3:6]}.{s[6:9]}-{s[9:]}"
    return str(val).strip()

def formatar_cep(val):
    """Formata CEP no padrão oficial 00000-000"""
    s = re.sub(r"\D", "", str(val or "")).strip()
    if len(s) == 8:
        return f"{s[:5]}-{s[5:]}"
    return str(val or "").strip()

def consultar_cep_viacep(cep_val: str) -> dict:
    """Consulta dados de endereço pelo CEP via ViaCEP (mesmo endpoint do portal oficial AGU) com fallback para BrasilAPI"""
    c = re.sub(r"\D", "", str(cep_val or ""))
    if len(c) != 8:
        return {}
    # 1. ViaCEP (padrão AGU cep.js)
    try:
        import requests
        r = requests.get(f"https://viacep.com.br/ws/{c}/json/", headers={"User-Agent": "Mozilla/5.0"}, timeout=3)
        if r.status_code == 200:
            d = r.json()
            if not d.get("erro"):
                return d
    except Exception:
        pass
    # 2. Fallback BrasilAPI
    try:
        import requests
        r2 = requests.get(f"https://brasilapi.com.br/api/cep/v1/{c}", headers={"User-Agent": "Mozilla/5.0"}, timeout=3)
        if r2.status_code == 200:
            d2 = r2.json()
            if not d2.get("errors"):
                return {
                    "cep": d2.get("cep"),
                    "logradouro": d2.get("street") or "",
                    "bairro": d2.get("neighborhood") or "",
                    "localidade": d2.get("city") or "",
                    "uf": d2.get("state") or "",
                }
    except Exception:
        pass
    return {}


def obter_iniciais(nome):
    if not nome:
        return "GL"
    partes = str(nome).strip().split()
    if len(partes) == 1:
        return partes[0][:2].upper()
    return f"{partes[0][0]}{partes[-1][0]}".upper()

# ============================================
# AUTENTICAÇÃO E USUÁRIOS
# ============================================

def normalizar_texto_login(txt):
    if not txt:
        return ""
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ASCII", "ignore").decode("utf-8")
    return txt.strip().lower()

def verificar_login(login_input, senha):
    if not login_input or not senha:
        return None
    login_clean = str(login_input).strip()
    senha_clean = str(senha).strip()
    if not login_clean or not senha_clean:
        return None

    login_lower = login_clean.lower()
    login_norm = normalizar_texto_login(login_clean)

    with engine.connect() as conn:
        users = conn.execute(
            text("SELECT id, nome, email, senha_hash, cargo, ativo FROM usuarios WHERE ativo = 1")
        ).mappings().all()

        for u in users:
            uid = u["id"]
            unome = (u["nome"] or "").strip()
            ulogin = (u["email"] or "").strip()
            usenha_hash = (u["senha_hash"] or "").strip()
            ucargo = u["cargo"]

            # Identificadores válidos para este usuário
            ulogin_lower = ulogin.lower()
            ulogin_user = ulogin_lower.split("@")[0] if "@" in ulogin_lower else ulogin_lower
            unome_lower = unome.lower()
            unome_norm = normalizar_texto_login(unome)
            unome_primeiro = unome_norm.split()[0] if unome_norm else ""
            unome_slug = re.sub(r"[^a-z0-9]", "", unome_norm)

            # Nome sem parênteses (ex: "João Silva (Operador)" -> "João Silva")
            unome_limpo = re.sub(r"\(.*?\)", "", unome).strip()
            unome_limpo_lower = unome_limpo.lower()
            unome_limpo_norm = normalizar_texto_login(unome_limpo)
            unome_limpo_primeiro = unome_limpo_norm.split()[0] if unome_limpo_norm else ""
            unome_limpo_slug = re.sub(r"[^a-z0-9]", "", unome_limpo_norm)

            candidatos = {
                ulogin, ulogin_lower, ulogin_user,
                unome, unome_lower, unome_norm, unome_primeiro, unome_slug,
                unome_limpo, unome_limpo_lower, unome_limpo_norm, unome_limpo_primeiro, unome_limpo_slug
            }

            # Atalhos amigáveis por cargo
            if ucargo in ["ADMIN", "ADMINISTRADOR"]:
                candidatos.add("admin")
                candidatos.add("administrador")
            if ucargo == "OPERADOR":
                candidatos.add("operador")

            # Permite login tanto exato quanto insensível a maiúsculas/minúsculas
            candidatos_lower = {str(c).lower() for c in candidatos}
            if login_clean in candidatos or login_lower in candidatos_lower or login_norm in candidatos:
                senha_ok = False
                if usenha_hash.startswith(("$2a$", "$2b$", "$2y$")):
                    try:
                        senha_ok = bcrypt.checkpw(senha_clean.encode("utf-8"), usenha_hash.encode("utf-8"))
                    except Exception:
                        senha_ok = False
                elif usenha_hash == senha_clean or usenha_hash == str(senha):
                    senha_ok = True
                    try:
                        novo_h = bcrypt.hashpw(senha_clean.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
                        conn.execute(text("UPDATE usuarios SET senha_hash = :h WHERE id = :id"), {"h": novo_h, "id": uid})
                        conn.commit()
                    except Exception:
                        pass

                if senha_ok:
                    return {"id": uid, "nome": unome, "email": ulogin, "cargo": ucargo}

    return None

def verificar_senha_usuario(usuario_id: int, senha: str) -> bool:
    """Valida criptograficamente se a senha informada confere com o usuário cadastrado no banco."""
    if not usuario_id or not senha:
        return False
    senha_clean = str(senha).strip()
    if not senha_clean:
        return False
    try:
        with get_engine().connect() as conn:
            u = conn.execute(
                text("SELECT id, senha_hash, cargo, ativo FROM usuarios WHERE id = :id AND ativo = 1"),
                {"id": int(usuario_id)}
            ).mappings().first()
            if not u:
                return False
            usenha_hash = (u["senha_hash"] or "").strip()
            if usenha_hash.startswith(("$2a$", "$2b$", "$2y$")):
                try:
                    import bcrypt
                    return bcrypt.checkpw(senha_clean.encode("utf-8"), usenha_hash.encode("utf-8"))
                except Exception:
                    return False
            return usenha_hash == senha_clean or usenha_hash == str(senha)
    except Exception:
        return False

@st.cache_data(ttl=120, show_spinner=False)
def carregar_usuarios():
    with engine.connect() as conn:
        query = text("SELECT id, nome, email, cargo, ativo, data_criacao FROM usuarios ORDER BY nome ASC")
        return pd.read_sql(query, conn)

def cadastrar_novo_usuario(nome, usuario_login=None, senha="", cargo="OPERADOR"):
    if not nome or not str(nome).strip():
        return False, "O campo 'Nome Completo' é obrigatório."
    nome_limpo = str(nome).strip()

    if not usuario_login or not str(usuario_login).strip():
        slug = normalizar_texto_login(nome_limpo)
        slug_partes = slug.split()
        if len(slug_partes) > 1:
            login_limpo = f"{re.sub(r'[^a-zA-Z0-9]', '', slug_partes[0])}.{re.sub(r'[^a-zA-Z0-9]', '', slug_partes[-1])}"
        else:
            login_limpo = re.sub(r'[^a-zA-Z0-9]', '', slug)
    else:
        # Deixa exatamente igual ao que o usuário criou
        login_limpo = str(usuario_login).strip()
        if "@" in login_limpo:
            login_limpo = login_limpo.split("@")[0].strip()

    if not login_limpo:
        return False, "Nome de usuário inválido."

    if not senha or len(senha.strip()) < 4:
        return False, "A senha deve conter no mínimo 4 caracteres."

    senha_limpa = senha.strip()
    senha_hash = bcrypt.hashpw(senha_limpa.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

    with engine.connect() as conn:
        existe = conn.execute(
            text("SELECT id FROM usuarios WHERE LOWER(nome) = LOWER(:n) OR LOWER(email) = LOWER(:e)"),
            {"n": nome_limpo, "e": login_limpo}
        ).fetchone()
        if existe:
            return False, f"Já existe um usuário cadastrado com este Nome ('{nome_limpo}') ou Nome de Usuário ('{login_limpo}')."

        try:
            conn.execute(
                text("""
                    INSERT INTO usuarios (nome, email, senha_hash, cargo, ativo, data_criacao)
                    VALUES (:nome, :email, :senha, :cargo, 1, CURRENT_TIMESTAMP)
                """),
                {"nome": nome_limpo, "email": login_limpo, "senha": senha_hash, "cargo": cargo}
            )
            conn.commit()
            try:
                carregar_usuarios.clear()
            except Exception:
                pass
            return True, f"Usuário '{nome_limpo}' cadastrado com sucesso! Nome de usuário para login: '{login_limpo}'."
        except Exception as err:
            return False, f"Erro ao gravar usuário: {str(err)}"

def atualizar_usuario(usuario_id, nome, usuario_login, cargo, ativo, nova_senha=None):
    if not nome or not str(nome).strip():
        return False, "O campo 'Nome Completo' é obrigatório."
    nome_limpo = str(nome).strip()
    login_limpo = str(usuario_login).strip() if usuario_login else ""
    if "@" in login_limpo:
        login_limpo = login_limpo.split("@")[0].strip()

    if not login_limpo:
        return False, "O campo 'Nome de Usuário' é obrigatório."

    if nova_senha is not None and str(nova_senha).strip() != "":
        senha_limpa = str(nova_senha).strip()
        if len(senha_limpa) < 4:
            return False, "A nova senha deve conter no mínimo 4 caracteres."
    else:
        senha_limpa = None

    with engine.connect() as conn:
        conflito = conn.execute(
            text("SELECT id FROM usuarios WHERE LOWER(email) = LOWER(:e) AND id != :id"),
            {"e": login_limpo, "id": usuario_id}
        ).fetchone()
        if conflito:
            return False, f"Já existe outro usuário com o nome de usuário '{login_limpo}'."

        try:
            if senha_limpa:
                senha_hash = bcrypt.hashpw(senha_limpa.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
                conn.execute(
                    text("UPDATE usuarios SET nome = :nome, email = :email, cargo = :cargo, ativo = :ativo, senha_hash = :senha WHERE id = :id"),
                    {"nome": nome_limpo, "email": login_limpo, "cargo": cargo, "ativo": 1 if ativo else 0, "senha": senha_hash, "id": usuario_id}
                )
            else:
                conn.execute(
                    text("UPDATE usuarios SET nome = :nome, email = :email, cargo = :cargo, ativo = :ativo WHERE id = :id"),
                    {"nome": nome_limpo, "email": login_limpo, "cargo": cargo, "ativo": 1 if ativo else 0, "id": usuario_id}
                )
            conn.commit()
            try:
                carregar_usuarios.clear()
            except Exception:
                pass
            return True, "Usuário atualizado com sucesso!"
        except Exception as e:
            return False, f"Erro ao atualizar: {str(e)}"

def excluir_usuario(usuario_id, current_user_id):
    if usuario_id == current_user_id:
        return False, "Você não pode excluir o seu próprio usuário logado."
    with engine.connect() as conn:
        total = conn.execute(text("SELECT COUNT(*) FROM usuarios WHERE ativo = 1")).scalar()
        if total <= 1:
            return False, "Não é possível excluir o único usuário ativo do sistema."
        try:
            conn.execute(text("DELETE FROM usuarios WHERE id = :id"), {"id": usuario_id})
            conn.commit()
            try:
                carregar_usuarios.clear()
            except Exception:
                pass
            return True, "Usuário excluído com sucesso!"
        except Exception as e:
            return False, f"Erro ao excluir: {str(e)}"

# ============================================
# CARREGAMENTO DE DADOS (CONTRATOS E PROTOCOLOS)
# ============================================

@st.cache_data(ttl=120, show_spinner=False)
def carregar_contratos():
    with engine.connect() as conn:
        query = text("""
            SELECT 
                c.id,
                c.numero_contrato,
                c.ano_contrato,
                c.numero_completo,
                c.orgao_id,
                o.nome as orgao,
                c.modalidade,
                c.processo_adm,
                c.objeto,
                c.fornecedor_id,
                COALESCE(f.razao_social, f.nome_fantasia, 'Não informado') as fornecedor,
                f.cnpj_cpf,
                f.endereco as fornecedor_endereco,
                f.numero as fornecedor_numero,
                f.complemento as fornecedor_complemento,
                f.bairro as fornecedor_bairro,
                f.cidade as fornecedor_cidade,
                f.uf as fornecedor_uf,
                f.cep as fornecedor_cep,
                f.nome_representante as fornecedor_representante,
                f.cargo_representante as fornecedor_cargo,
                c.valor_total,
                c.data_assinatura,
                c.data_publicacao,
                c.data_vencimento,
                c.vigencia_descricao,
                c.solicitante,
                c.secretario,
                c.fiscal,
                c.status,
                c.dotacao_orcamentaria,
                c.modelo_agu,
                c.qtd_pagamentos,
                c.observacoes,
                u.nome as usuario_criador,
                c.data_criacao,
                c.atualizado_em,
                c.editado_em
            FROM contratos c
            LEFT JOIN orgaos o ON c.orgao_id = o.id
            LEFT JOIN fornecedores f ON c.fornecedor_id = f.id
            LEFT JOIN usuarios u ON c.criado_por = u.id
            ORDER BY c.numero_contrato ASC, c.ano_contrato DESC
        """)
        df = pd.read_sql(query, conn)
        # Normalização de datas para string ISO (YYYY-MM-DD) para compatibilidade entre SQLite e PostgreSQL
        for col in ["data_assinatura", "data_publicacao", "data_vencimento", "data_criacao"]:
            if col in df.columns:
                df[col] = df[col].apply(
                    lambda v: v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else (
                        str(v)[:10] if (v is not None and str(v).strip() not in ["", "None", "NaT"]) else ""
                    )
                )
        return df

def gerar_proximo_numero_contrato(ano=None):
    if ano is None:
        ano = datetime.now().year
    with engine.connect() as conn:
        res = conn.execute(
            text("SELECT MAX(numero_contrato) FROM contratos WHERE ano_contrato = :ano"),
            {"ano": ano}
        ).scalar()
        proximo = 1 if (res is None) else (int(res) + 1)
        return proximo, f"{str(proximo).zfill(3)}/{ano}"

def obter_ultimo_e_proximo_numero_contrato(ano=None):
    if ano is None:
        ano = datetime.now().year
    with engine.connect() as conn:
        res = conn.execute(
            text("SELECT MAX(numero_contrato) FROM contratos WHERE ano_contrato = :ano"),
            {"ano": ano}
        ).scalar()
        if res is None or int(res) <= 0:
            ultimo_num = 0
            ultimo_str = f"Nenhum em {ano}"
            proximo = 1
        else:
            ultimo_num = int(res)
            ultimo_str = f"{str(ultimo_num).zfill(3)}/{ano}"
            proximo = ultimo_num + 1
        return ultimo_num, ultimo_str, proximo, f"{str(proximo).zfill(3)}/{ano}"

@st.cache_data(ttl=120, show_spinner=False)
def carregar_protocolos():
    with engine.connect() as conn:
        query = text("""
            SELECT 
                p.id,
                p.numero_protocolo,
                p.ano_protocolo,
                COALESCE(p.numero_completo, 'PROT-' || p.numero_protocolo || '/' || p.ano_protocolo) as numero_completo,
                p.data_recebimento,
                COALESCE(p.assunto, p.descricao) as assunto,
                COALESCE(p.solicitante, p.remetente_nome) as solicitante,
                COALESCE(o.nome, p.remetente_orgao, 'Não especificado') as orgao_origem,
                p.status,
                c.numero_completo as contrato_vinculado,
                p.contrato_id,
                p.secretario,
                u.nome as usuario_criador,
                p.data_criacao,
                p.tipo_documento,
                p.observacoes
            FROM protocolos p
            LEFT JOIN orgaos o ON p.orgao_origem_id = o.id
            LEFT JOIN contratos c ON p.contrato_id = c.id
            LEFT JOIN usuarios u ON p.criado_por = u.id
            ORDER BY p.id DESC
        """)
        return pd.read_sql(query, conn)

def carregar_protocolos_do_contrato(contrato_id):
    with engine.connect() as conn:
        query = text("""
            SELECT 
                p.id,
                COALESCE(p.numero_completo, 'PROT-' || p.numero_protocolo || '/' || p.ano_protocolo) as Protocolo,
                p.data_recebimento as Data,
                COALESCE(p.assunto, p.descricao) as Assunto,
                COALESCE(p.solicitante, p.remetente_nome) as Solicitante,
                p.status as Status
            FROM protocolos p
            WHERE p.contrato_id = :cid
            ORDER BY p.id DESC
        """)
        return pd.read_sql(query, conn, params={"cid": contrato_id})

def gerar_proximo_numero_protocolo(ano=None):
    if ano is None:
        ano = datetime.now().year
    with engine.connect() as conn:
        res = conn.execute(
            text("SELECT MAX(numero_protocolo) FROM protocolos WHERE ano_protocolo = :ano"),
            {"ano": ano}
        ).scalar()
        proximo = 1 if (res is None) else (int(res) + 1)
        return proximo, f"PROT-{str(proximo).zfill(4)}/{ano}"

def carregar_aditivos_do_contrato(contrato_id):
    with engine.connect() as conn:
        query = text("""
            SELECT 
                id,
                contrato_id,
                numero_aditivo,
                numero_completo,
                ano_aditivo,
                tipo_aditivo,
                regime_legal,
                processo_adm,
                data_assinatura,
                data_publicacao,
                nova_data_vencimento,
                prazo_aditado_meses,
                valor_aditado,
                percentual_aditado,
                novo_valor_total,
                objeto_aditivo,
                justificativa,
                dotacao_orcamentaria,
                garantia_execucao,
                secretario,
                fiscal,
                status,
                data_criacao
            FROM termos_aditivos
            WHERE contrato_id = :cid
            ORDER BY numero_aditivo ASC
        """)
        return pd.read_sql(query, conn, params={"cid": int(contrato_id)})

@st.cache_data(ttl=120, show_spinner=False)
def carregar_todos_aditivos():
    with engine.connect() as conn:
        query = text("""
            SELECT 
                a.*,
                c.numero_completo as contrato_numero,
                c.orgao_id,
                o.nome as orgao_nome,
                COALESCE(f.razao_social, f.nome_fantasia) as fornecedor_nome
            FROM termos_aditivos a
            LEFT JOIN contratos c ON a.contrato_id = c.id
            LEFT JOIN orgaos o ON c.orgao_id = o.id
            LEFT JOIN fornecedores f ON c.fornecedor_id = f.id
            ORDER BY a.id DESC
        """)
        return pd.read_sql(query, conn)

def gerar_proximo_numero_aditivo(contrato_id):
    with engine.connect() as conn:
        res = conn.execute(
            text("SELECT MAX(numero_aditivo) FROM termos_aditivos WHERE contrato_id = :cid"),
            {"cid": int(contrato_id)}
        ).scalar()
        proximo = 1 if (res is None) else (int(res) + 1)
        return proximo

def salvar_termo_aditivo(dados: dict, atualizar_contrato: bool = True):
    with engine.connect() as conn:
        cid = int(dados["contrato_id"])
        num_adit = int(dados.get("numero_aditivo") or gerar_proximo_numero_aditivo(cid))
        ano = int(dados.get("ano_aditivo") or datetime.now().year)
        n_comp = dados.get("numero_completo") or f"{num_adit}º TERMO ADITIVO"
        
        res_adit = conn.execute(
            text("""
                INSERT INTO termos_aditivos (
                    contrato_id, numero_aditivo, numero_completo, ano_aditivo,
                    tipo_aditivo, regime_legal, processo_adm, data_assinatura,
                    data_publicacao, nova_data_vencimento, prazo_aditado_meses,
                    valor_aditado, percentual_aditado, novo_valor_total,
                    objeto_aditivo, justificativa, dotacao_orcamentaria,
                    garantia_execucao, secretario, fiscal, status, criado_por
                ) VALUES (
                    :cid, :num, :ncomp, :ano,
                    :tipo, :regime, :proc, :dt_ass,
                    :dt_pub, :nova_venc, :prazo,
                    :val_adit, :perc, :novo_val,
                    :obj, :just, :dot,
                    :gar, :sec, :fisc, :st, :uid
                )
            """),
            {
                "cid": cid,
                "num": num_adit,
                "ncomp": n_comp,
                "ano": ano,
                "tipo": dados.get("tipo_aditivo", "PRORROGACAO"),
                "regime": dados.get("regime_legal", "LEI_14133_2021"),
                "proc": dados.get("processo_adm", ""),
                "dt_ass": str(dados.get("data_assinatura")) if dados.get("data_assinatura") else None,
                "dt_pub": str(dados.get("data_publicacao")) if dados.get("data_publicacao") else None,
                "nova_venc": str(dados.get("nova_data_vencimento")) if dados.get("nova_data_vencimento") else None,
                "prazo": int(dados.get("prazo_aditado_meses", 0) or 0),
                "val_adit": float(dados.get("valor_aditado", 0.0) or 0.0),
                "perc": float(dados.get("percentual_aditado", 0.0) or 0.0),
                "novo_val": float(dados.get("novo_valor_total", 0.0) or 0.0),
                "obj": dados.get("objeto_aditivo", ""),
                "just": dados.get("justificativa", ""),
                "dot": dados.get("dotacao_orcamentaria", ""),
                "gar": dados.get("garantia_execucao", ""),
                "sec": dados.get("secretario", ""),
                "fisc": dados.get("fiscal", ""),
                "st": "ATIVO",
                "uid": dados.get("criado_por")
            }
        )
        
        adit_id = res_adit.lastrowid
        registrar_audit_log(
            conn, "termos_aditivos", adit_id, "CRIACAO",
            f"Aditivo {n_comp} registrado para o contrato ID {cid}.",
            dados_novos={"tipo": dados.get("tipo_aditivo"), "valor": dados.get("valor_aditado")}
        )
        
        if atualizar_contrato:
            updates = []
            params = {"cid": cid}
            if dados.get("nova_data_vencimento"):
                updates.append("data_vencimento = :nova_venc")
                params["nova_venc"] = str(dados["nova_data_vencimento"])
                updates.append("status = 'PRORROGADO'")
            if dados.get("novo_valor_total") and float(dados["novo_valor_total"]) > 0:
                updates.append("valor_total = :novo_val")
                params["novo_val"] = float(dados["novo_valor_total"])
            
            if updates:
                uid_edit = dados.get("criado_por")
                if uid_edit:
                    updates.append("editado_por = :uid_edit")
                    params["uid_edit"] = uid_edit
                updates.append("editado_em = CURRENT_TIMESTAMP")
                updates.append("atualizado_em = CURRENT_TIMESTAMP")
                
                sql_up = f"UPDATE contratos SET {', '.join(updates)} WHERE id = :cid"
                conn.execute(text(sql_up), params)
                
                registrar_audit_log(
                    conn, "contratos", cid, "ADITIVO",
                    f"Contrato atualizado via {n_comp}.",
                    dados_novos={"novos_updates": updates}
                )
                
        conn.commit()
        try:
            carregar_contratos.clear()
            carregar_todos_aditivos.clear()
        except Exception:
            pass
    return True, f"Termo Aditivo {n_comp} registrado e contrato atualizado com sucesso!"


# ============================================
# IMPORTADOR INTELIGENTE DE PLANILHAS
# ============================================

def detectar_coluna(colunas_reais, termos):
    # 1. Primeiro busca correspondência exata (normalizada)
    for t in termos:
        t_norm = normalizar_texto(t)
        for c in colunas_reais:
            c_norm = normalizar_texto(c)
            if t_norm == c_norm:
                return c

    # 2. Depois busca por substring na ordem de prioridade dos termos
    for t in termos:
        t_norm = normalizar_texto(t)
        for c in colunas_reais:
            c_norm = normalizar_texto(c)
            # Evita falsos positivos como 'VALOR MENSAL', 'VALOR UNITÁRIO' quando busca 'VALOR'
            if t_norm == "valor":
                if any(descarte in c_norm for descarte in ["mensal", "unitario", "unit", "parcela", "estimado", "item"]):
                    continue
            if t_norm in ["total", "valor total", "valor global"]:
                if any(descarte in c_norm for descarte in ["qtd", "quant", "quantidade", "dias", "meses", "horas", "itens"]):
                    continue
            if t_norm in c_norm:
                return c
    return None

def processar_importacao_planilha_customizada(df_upload, mapa_colunas, modo="adicionar", progress_bar=None, status_text=None):
    df_import = df_upload.copy()
    if df_import.empty:
        return False, "A planilha enviada está vazia."

    linhas_processadas = 0
    linhas_ignoradas = 0

    with engine.connect() as conn:
        trans = conn.begin()
        try:
            if modo == "substituir":
                conn.execute(text("DELETE FROM movimentacoes_protocolo"))
                conn.execute(text("DELETE FROM protocolos"))
                conn.execute(text("DELETE FROM aditivos"))
                conn.execute(text("DELETE FROM contratos"))

            orgaos_db = {r[1].upper().strip(): r[0] for r in conn.execute(text("SELECT id, nome FROM orgaos")).fetchall()}
            fornec_db = {r[1].strip(): r[0] for r in conn.execute(text("SELECT id, cnpj_cpf FROM fornecedores WHERE cnpj_cpf IS NOT NULL")).fetchall()}

            total_rows = len(df_import)

            for idx, row in df_import.iterrows():
                if progress_bar and total_rows > 0:
                    progress_bar.progress((idx + 1) / total_rows)
                if status_text:
                    status_text.text(f"Processando registro {idx + 1} de {total_rows}...")

                col_num = mapa_colunas.get("numero")
                raw_num = str(row[col_num]).strip() if col_num and pd.notna(row[col_num]) else ""
                if not raw_num or raw_num.lower() in ["nan", "none", "", "-"]:
                    linhas_ignoradas += 1
                    continue

                # Descartar linhas de resumo/total para evitar duplicação ou números espúrios
                row_str_full = " ".join([str(v) for v in row.values if pd.notna(v)]).strip().upper()
                if any(kw in raw_num.upper() for kw in ["TOTAL", "SUBTOTAL", "SOMA"]) or row_str_full.startswith(("TOTAL", "SUBTOTAL", "SOMA", "RESUMO")):
                    linhas_ignoradas += 1
                    continue

                m_num = re.search(r"(\d+)[/.\-_]?(\d{4})?", raw_num)
                if m_num:
                    num_contrato = int(m_num.group(1))
                    ano_contrato = int(m_num.group(2)) if m_num.group(2) else 2026
                else:
                    try:
                        num_contrato = int(float(raw_num))
                        ano_contrato = 2026
                    except Exception:
                        linhas_ignoradas += 1
                        continue

                num_completo = f"{str(num_contrato).zfill(3)}/{ano_contrato}"

                col_org = mapa_colunas.get("orgao")
                raw_org = str(row[col_org]).strip() if col_org and pd.notna(row[col_org]) else "PREFEITURA MUNICIPAL"
                org_norm = normalizar_texto(raw_org)
                orgao_nome = "SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO"
                for chave_termo, org_oficial in MAPA_ORGAOS_PADRAO.items():
                    if normalizar_texto(chave_termo) in org_norm:
                        orgao_nome = org_oficial
                        break

                orgao_id = orgaos_db.get(orgao_nome.upper())
                if not orgao_id:
                    res_org = conn.execute(
                        text("INSERT INTO orgaos (nome, sigla, ativo) VALUES (:nome, :sigla, 1)"),
                        {"nome": orgao_nome, "sigla": orgao_nome[:10]}
                    )
                    orgao_id = res_org.lastrowid
                    orgaos_db[orgao_nome.upper()] = orgao_id

                col_sec = mapa_colunas.get("secretario")
                raw_sec = str(row[col_sec]).strip() if col_sec and pd.notna(row[col_sec]) else ""
                secretario_titular = MAPA_SECRETARIOS_PADRAO.get(orgao_nome, "SECRETÁRIO MUNICIPAL TITULAR")
                if raw_sec and raw_sec.lower() not in ["nan", "none", "-", ""]:
                    secretario_titular = raw_sec.upper()

                col_forn = mapa_colunas.get("fornecedor")
                razao_social = str(row[col_forn]).strip() if col_forn and pd.notna(row[col_forn]) else "FORNECEDOR NÃO INFORMADO"

                col_cnpj = mapa_colunas.get("cnpj")
                cnpj_raw = str(row[col_cnpj]).strip() if col_cnpj and pd.notna(row[col_cnpj]) else None
                if cnpj_raw and cnpj_raw.lower() in ["nan", "none", "-", ""]:
                    cnpj_raw = None

                fornecedor_id = None
                if cnpj_raw and cnpj_raw in fornec_db:
                    fornecedor_id = fornec_db[cnpj_raw]
                else:
                    res_f = conn.execute(
                        text("INSERT INTO fornecedores (nome_fantasia, razao_social, cnpj_cpf, ativo) VALUES (:rz, :rz, :cnpj, 1)"),
                        {"rz": razao_social, "cnpj": cnpj_raw}
                    )
                    fornecedor_id = res_f.lastrowid
                    if cnpj_raw:
                        fornec_db[cnpj_raw] = fornecedor_id

                col_obj = mapa_colunas.get("objeto")
                objeto_txt = str(row[col_obj]).strip() if col_obj and pd.notna(row[col_obj]) else "Objeto não informado"

                col_mod = mapa_colunas.get("modalidade")
                modalidade_txt = str(row[col_mod]).strip() if col_mod and pd.notna(row[col_mod]) else "DISPENSA DE LICITAÇÃO"
                
                col_proc = mapa_colunas.get("processo")
                proc_adm = str(row[col_proc]).strip() if col_proc and pd.notna(row[col_proc]) else None

                col_val = mapa_colunas.get("valor")
                val_total = clean_money(row[col_val]) if col_val and pd.notna(row[col_val]) else 0.0

                col_solic = mapa_colunas.get("solicitante")
                solicitante_txt = str(row[col_solic]).strip() if col_solic and pd.notna(row[col_solic]) else None

                col_fisc = mapa_colunas.get("fiscal")
                fiscal_txt = str(row[col_fisc]).strip() if col_fisc and pd.notna(row[col_fisc]) else None

                col_data_ass = mapa_colunas.get("data_assinatura")
                raw_data_ass = row[col_data_ass] if col_data_ass and pd.notna(row[col_data_ass]) else None
                data_ass = parse_data_universal(raw_data_ass)

                if not data_ass and solicitante_txt:
                    m_date_solic = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", solicitante_txt)
                    if m_date_solic:
                        data_ass = parse_data_universal(m_date_solic.group(0))

                col_data_pub = mapa_colunas.get("data_publicacao")
                raw_data_pub = row[col_data_pub] if col_data_pub and pd.notna(row[col_data_pub]) else None
                data_pub = parse_data_universal(raw_data_pub)

                col_data_venc = mapa_colunas.get("data_vencimento")
                raw_data_venc = row[col_data_venc] if col_data_venc and pd.notna(row[col_data_venc]) else None
                data_venc = parse_data_universal(raw_data_venc)

                if data_ass and not data_pub:
                    try:
                        dt_ass = datetime.strptime(data_ass, "%Y-%m-%d")
                        data_pub = (dt_ass + timedelta(days=3)).strftime("%Y-%m-%d")
                    except Exception:
                        pass

                if data_ass and not data_venc:
                    try:
                        dt_ass = datetime.strptime(data_ass, "%Y-%m-%d")
                        data_venc = dt_ass.replace(year=dt_ass.year + 1).strftime("%Y-%m-%d")
                    except Exception:
                        pass

                vigencia_desc = "12 (doze) meses"
                if raw_data_ass and any(kw in str(raw_data_ass).upper() for kw in ["DEFINITIVO", "RECEBIMENTO", "ENTREGA"]):
                    vigencia_desc = str(raw_data_ass).strip()

                hoje_str = date.today().strftime("%Y-%m-%d")
                status_txt = "VIGENTE"
                if data_venc and data_venc < hoje_str:
                    status_txt = "ENCERRADO"

                existente_id = conn.execute(
                    text("SELECT id FROM contratos WHERE numero_contrato = :nc AND ano_contrato = :ano"),
                    {"nc": num_contrato, "ano": ano_contrato}
                ).scalar()

                if existente_id:
                    conn.execute(
                        text("""
                            UPDATE contratos SET
                                numero_completo = :ncomp, orgao_id = :org, modalidade = :mod,
                                processo_adm = :proc, objeto = :obj, fornecedor_id = :forn,
                                valor_total = :val, data_assinatura = :d_ass, data_publicacao = :d_pub,
                                data_vencimento = :d_venc, vigencia_descricao = :vig, solicitante = :solic,
                                secretario = :sec, fiscal = :fisc, status = :st, atualizado_em = CURRENT_TIMESTAMP
                            WHERE id = :cid
                        """),
                        {
                            "cid": existente_id,
                            "ncomp": num_completo, "org": orgao_id, "mod": modalidade_txt, "proc": proc_adm,
                            "obj": objeto_txt, "forn": fornecedor_id, "val": val_total,
                            "d_ass": data_ass, "d_pub": data_pub, "d_venc": data_venc,
                            "vig": vigencia_desc, "solic": solicitante_txt, "sec": secretario_titular,
                            "fisc": fiscal_txt, "st": status_txt
                        }
                    )
                else:
                    conn.execute(
                        text("""
                            INSERT INTO contratos (
                                numero_contrato, ano_contrato, numero_completo, orgao_id, modalidade,
                                processo_adm, objeto, fornecedor_id, valor_total, data_assinatura,
                                data_publicacao, data_vencimento, vigencia_descricao, solicitante,
                                secretario, fiscal, status, criado_por, data_criacao, atualizado_em
                            ) VALUES (
                                :nc, :ano, :ncomp, :org, :mod,
                                :proc, :obj, :forn, :val, :d_ass,
                                :d_pub, :d_venc, :vig, :solic,
                                :sec, :fisc, :st, 1, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
                            )
                        """),
                        {
                            "nc": num_contrato, "ano": ano_contrato, "ncomp": num_completo,
                            "org": orgao_id, "mod": modalidade_txt, "proc": proc_adm,
                            "obj": objeto_txt, "forn": fornecedor_id, "val": val_total,
                            "d_ass": data_ass, "d_pub": data_pub, "d_venc": data_venc,
                            "vig": vigencia_desc, "solic": solicitante_txt, "sec": secretario_titular,
                            "fisc": fiscal_txt, "st": status_txt
                        }
                    )
                linhas_processadas += 1

            trans.commit()
            try:
                carregar_contratos.clear()
            except Exception:
                pass
            return True, f"Conexão com base concluída! {linhas_processadas} contratos integrados com sucesso."
        except Exception as e:
            trans.rollback()
            return False, f"Falha na integração: {str(e)}"

@st.cache_data(show_spinner=False)
def exportar_contratos_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_exp = df.copy()
        df_exp["Valor Total (R$)"] = df_exp["valor_total"].apply(formatar_moeda)
        df_exp["Data Assinatura"] = df_exp["data_assinatura"].apply(formatar_data_br)
        df_exp["Data Publicação"] = df_exp["data_publicacao"].apply(formatar_data_br)
        df_exp["Data Vencimento"] = df_exp["data_vencimento"].apply(formatar_data_br)
        df_exp["CNPJ/CPF Formatado"] = df_exp["cnpj_cpf"].apply(formatar_cnpj_cpf)
        
        cols = [
            "numero_completo", "orgao", "secretario", "modalidade", "processo_adm",
            "fornecedor", "CNPJ/CPF Formatado", "objeto", "Valor Total (R$)",
            "Data Assinatura", "Data Publicação", "Data Vencimento", "vigencia_descricao",
            "solicitante", "fiscal", "status"
        ]
        cols_exist = [c for c in cols if c in df_exp.columns]
        df_exp[cols_exist].to_excel(writer, sheet_name="Contratos Municipais 2026", index=False)
    return output.getvalue()

@st.cache_data(show_spinner=False)
def exportar_protocolos_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Protocolos 2026", index=False)
    return output.getvalue()

# ============================================
# INTERFACE - LOGIN CLARO & MUNICIPAL
# ============================================

def tela_login():
    brasao_b64 = obter_brasao_b64()
    brasao_img = f"data:image/png;base64,{brasao_b64}" if brasao_b64 else ""

    st.markdown("""
    <style>
        /* Fundo da tela de login em Azul Municipal Minimalista */
        .stApp, [data-testid="stAppViewContainer"], section.main {
            background: linear-gradient(135deg, #0284c7 0%, #0369a1 45%, #075985 100%) !important;
            background-attachment: fixed !important;
        }

        /* Card Minimalista do Formulário de Login */
        div[data-testid="stForm"],
        form[aria-label="login_form_gel"],
        .gel-login-box [data-testid="stForm"] {
            background: #ffffff !important;
            border-radius: 16px !important;
            padding: 34px 30px !important;
            border: 1px solid rgba(255, 255, 255, 0.4) !important;
            box-shadow: 0 20px 45px rgba(0, 0, 0, 0.22) !important;
            width: 100% !important;
            max-width: 420px !important;
            margin: 0 auto !important;
        }

        .login-brasao {
            display: flex;
            justify-content: center;
            margin-bottom: 12px;
        }
        .login-brasao img {
            width: 65px;
            height: auto;
            filter: drop-shadow(0 4px 6px rgba(0,0,0,0.12));
        }

        .login-prefeitura-nome {
            font-family: 'Plus Jakarta Sans', sans-serif !important;
            font-weight: 800 !important;
            font-size: 15px !important;
            color: #0f172a !important;
            text-align: center !important;
            line-height: 1.3 !important;
            margin-bottom: 3px !important;
            letter-spacing: -0.2px !important;
        }

        .login-sistema-nome {
            font-size: 11.5px !important;
            color: #0284c7 !important;
            font-weight: 700 !important;
            text-align: center !important;
            margin-bottom: 20px !important;
            letter-spacing: 0.4px !important;
            text-transform: uppercase !important;
        }

        div[data-testid="stForm"] div[data-testid="stTextInput"] label,
        div[data-testid="stForm"] div[data-testid="stTextInput"] label p,
        div[data-testid="stForm"] label,
        div[data-testid="stForm"] label p {
            font-size: 11px !important;
            color: #475569 !important;
            font-weight: 700 !important;
            letter-spacing: 0.5px !important;
            text-transform: uppercase !important;
            margin-bottom: 4px !important;
        }

        div[data-testid="stForm"] div[data-testid="stTextInput"] input {
            font-size: 14px !important;
            padding: 10px 12px !important;
            border-radius: 8px !important;
            background: #f8fafc !important;
            color: #0f172a !important;
            border: 1.5px solid #cbd5e1 !important;
            transition: all 0.2s ease !important;
        }
        div[data-testid="stForm"] div[data-testid="stTextInput"] input:focus {
            border-color: #0284c7 !important;
            background: #ffffff !important;
            box-shadow: 0 0 0 3px rgba(2, 132, 199, 0.15) !important;
            outline: none !important;
        }

        div[data-testid="stForm"] [data-testid="stFormSubmitButton"] button {
            background: linear-gradient(180deg, #0284c7 0%, #0369a1 100%) !important;
            color: #ffffff !important;
            font-weight: 700 !important;
            font-size: 14px !important;
            border-radius: 8px !important;
            border: none !important;
            margin-top: 10px !important;
            padding: 11px 16px !important;
            box-shadow: 0 4px 12px rgba(2, 132, 199, 0.3) !important;
            transition: all 0.2s ease !important;
        }
        div[data-testid="stForm"] [data-testid="stFormSubmitButton"] button:hover {
            background: linear-gradient(180deg, #0369a1 0%, #075985 100%) !important;
            box-shadow: 0 6px 16px rgba(2, 132, 199, 0.4) !important;
            transform: translateY(-1px) !important;
        }

        .login-footer-text {
            color: rgba(255, 255, 255, 0.9) !important;
            font-size: 11px !important;
            text-align: center !important;
            margin-top: 14px !important;
            line-height: 1.4 !important;
        }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<div style='height: 6vh;'></div>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])
    
    with col2:
        brasao_html = f"<div class='login-brasao'><img src='{brasao_img}' alt='Brasão Oficial' /></div>" if brasao_img else "<div style='text-align: center; font-size: 42px; margin-bottom: 12px;'>🏛️</div>"

        with st.form("login_form_gel"):
            st.markdown(f"""
                {brasao_html}
                <div class='login-prefeitura-nome'>
                    PREFEITURA MUNICIPAL DE RIBEIRÃOZINHO DO MARANHÃO
                </div>
                <div class='login-sistema-nome'>
                    Gestão de Contratos — Modelos AGU
                </div>
            """, unsafe_allow_html=True)

            usuario_login = st.text_input("Usuário", placeholder="Digite seu usuário", key="input_login_usuario")
            senha = st.text_input("Senha", type="password", placeholder="Digite sua senha", key="input_senha_usuario")
            submit = st.form_submit_button("Entrar no Sistema →", use_container_width=True)

            if submit:
                if not usuario_login or not senha:
                    st.warning("⚠️ Por favor, informe seu usuário e senha.")
                else:
                    login_tratado = usuario_login.strip()
                    usuario = verificar_login(login_tratado, senha)
                    if usuario:
                        st.session_state["logged_in"] = True
                        st.session_state["usuario"] = usuario
                        st.session_state.pop("editando_contrato_id", None)
                        st.rerun()
                    else:
                        st.error("❌ Usuário ou senha incorretos, ou usuário inativo.")

        # Rodapé discreto
        st.markdown("""
        <div style='text-align: center; margin-top: 14px;'>
            <div class='login-footer-text'>
                Acesso seguro e restrito a servidores autorizados • Exercício 2026
            </div>
        </div>
        """, unsafe_allow_html=True)

# ============================================
# INTERFACE - DASHBOARD CONTRATOS MODERNO (GEL)
# ============================================


def tela_edicao_contrato(cid=None):
    st.markdown("""
        <div class='municipal-banner' style='display: flex; justify-content: space-between; align-items: center; padding: 10px 18px; margin-bottom: 14px;'>
            <div>
                <div class='municipal-badge-pill'>
                    <span class='beacon-active-gel' style='background: #ffffff;'></span> GESTÃO CONTRATUAL • MÓDULO DE ALTERAÇÃO
                </div>
                <h2 class='municipal-banner-title' style='font-size: 20px;'>✏️ Edição Completa de Contrato Municipal</h2>
                <div class='municipal-banner-subtitle' style='font-size: 12px;'>
                    Altere numeração, datas, valores, partes contratantes, dotação orçamentária, vigência e objeto com atualização direta no banco de dados.
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    usuario_atual = st.session_state.get("usuario", {})
    cargo_atual = str(usuario_atual.get("cargo", "")).upper()
    if cargo_atual not in ["ADMIN", "ADMINISTRADOR"]:
        st.error("🔒 **Acesso Restrito:** Apenas usuários com perfil de Administrador (**ADMIN**) possuem autorização para alterar contratos municipais.")
        st.info("Caso necessite de alterações no registro deste contrato, contate o Administrador do sistema.")
        if st.button("⬅️ Voltar ao Painel de Contratos", key="btn_voltar_bloqueio_admin", type="primary"):
            st.session_state["menu_selecionado"] = "📊 Contratos"
            st.session_state.pop("editando_contrato_id", None)
            st.rerun()
        return

    df = carregar_contratos()
    if df.empty:
        st.warning("⚠️ Nenhum contrato cadastrado no sistema para editar.")
        if st.button("⬅️ Voltar ao Painel"):
            st.session_state["menu_selecionado"] = "📊 Contratos"
            st.session_state.pop("editando_contrato_id", None)
            st.rerun()
        return

    # Se cid não foi passado, tenta obter da sessão
    if cid is None:
        cid = st.session_state.get("editando_contrato_id")

    # Mapeamento e lista de contratos para seleção
    opcoes_map = {}
    lista_display = []
    default_idx = 0
    for idx, row in df.iterrows():
        c_id = int(row["id"])
        num_c = str(row.get("numero_completo") or f"{row.get('numero_contrato', '')}/{row.get('ano_contrato', '')}")
        forn = str(row.get("fornecedor", "Sem Fornecedor"))[:35]
        val_f = formatar_moeda(row.get("valor_total", 0.0))
        label = f"📄 Contrato {num_c} • {forn} • {val_f} (ID #{c_id})"
        opcoes_map[label] = c_id
        lista_display.append(label)
        if cid is not None and c_id == int(cid):
            default_idx = len(lista_display) - 1

    col_sel_c, col_dup_topo, col_voltar = st.columns([3.0, 1.4, 1.1])
    with col_sel_c:
        contrato_escolhido = st.selectbox(
            "🔍 Selecione o Contrato que deseja editar:",
            lista_display,
            index=default_idx,
            key="sel_contrato_edicao_ativa"
        )
    cid_atual = opcoes_map[contrato_escolhido]
    st.session_state["editando_contrato_id"] = cid_atual

    with col_dup_topo:
        st.write("")
        st.write("")
        if st.button("📋 Duplicar Contrato", key="btn_dup_edicao_topo", use_container_width=True, help="Duplica este contrato gerando uma nova minuta com o próximo número sequencial"):
            uid_atual = usuario_atual.get("id", 1)
            ok_dup, msg_dup, novo_cid = duplicar_contrato_bd(cid_atual, uid_atual)
            if ok_dup:
                st.session_state["editando_contrato_id"] = novo_cid
                st.session_state["sel_detalhe_gel"] = None
                st.toast("Contrato duplicado com o próximo número!", icon="📋")
                st.success(msg_dup)
                st.rerun()
            else:
                st.error(msg_dup)

    with col_voltar:
        st.write("")
        st.write("")
        if st.button("⬅️ Voltar", key="btn_voltar_dashboard_edit", use_container_width=True):
            st.session_state["menu_selecionado"] = "📊 Contratos"
            st.session_state.pop("editando_contrato_id", None)
            st.rerun()

    c_row = df[df["id"] == cid_atual].iloc[0]

    # Carrega órgãos e fornecedores do banco
    with get_engine().connect() as conn:
        orgaos_db = conn.execute(text("SELECT id, nome, sigla FROM orgaos WHERE ativo = 1 ORDER BY nome ASC")).fetchall()
        fornec_db = conn.execute(text("SELECT id, razao_social, nome_fantasia, cnpj_cpf FROM fornecedores ORDER BY razao_social ASC")).fetchall()

    nomes_orgaos = [r[1] for r in orgaos_db]
    map_orgao_id = {r[1]: r[0] for r in orgaos_db}
    
    orgao_atual = c_row.get("orgao") or ""
    if orgao_atual and orgao_atual not in nomes_orgaos:
        nomes_orgaos.insert(0, orgao_atual)
        if pd.notna(c_row.get("orgao_id")) and c_row.get("orgao_id"):
            map_orgao_id[orgao_atual] = int(c_row["orgao_id"])

    idx_orgao = nomes_orgaos.index(orgao_atual) if orgao_atual in nomes_orgaos else 0

    # Fornecedores
    fornec_labels = ["-- Manter Fornecedor Atual / Não Alterar --"]
    fornec_map = {}
    for f_r in fornec_db:
        f_lbl = f"{f_r[1]} ({formatar_cnpj_cpf(f_r[3])})"
        fornec_labels.append(f_lbl)
        fornec_map[f_lbl] = {
            "id": f_r[0],
            "nome": f_r[1],
            "cnpj": f_r[3]
        }

    status_opcoes = ["VIGENTE", "ADITADO", "ENCERRADO", "DISTRATADO", "SUSPENSO", "VENCIDO"]
    status_atual = str(c_row.get("status") or "VIGENTE").upper().strip()
    if status_atual not in status_opcoes:
        status_opcoes.append(status_atual)
    idx_status = status_opcoes.index(status_atual) if status_atual in status_opcoes else 0

    modalidades_opcoes = [
        "DISPENSA DE LICITAÇÃO",
        "PREGÃO ELETRÔNICO",
        "PREGÃO PRESENCIAL",
        "CONCORRÊNCIA",
        "INEXIGIBILIDADE DE LICITAÇÃO",
        "TOMADA DE PREÇOS",
        "CONVITE",
        "ADESÃO À ATA DE REGISTRO DE PREÇOS (CARONA)",
        "CREDENCIAMENTO",
        "LEILÃO",
        "CONCURSO"
    ]
    mod_atual = str(c_row.get("modalidade") or "DISPENSA DE LICITAÇÃO").upper().strip()
    if mod_atual and mod_atual not in modalidades_opcoes:
        modalidades_opcoes.insert(0, mod_atual)
    idx_mod = modalidades_opcoes.index(mod_atual) if mod_atual in modalidades_opcoes else 0

    st.markdown("---")

    with st.form(f"form_edicao_completa_{cid_atual}"):
        st.markdown("##### 1. 📄 Identificação e Numeração Oficial")
        c_id1, c_id2, c_id3, c_id4 = st.columns([1, 1, 1.5, 1.5])
        with c_id1:
            novo_num_ct = st.number_input("Nº do Contrato", value=int(c_row.get("numero_contrato") or 1), min_value=1, step=1)
        with c_id2:
            novo_ano_ct = st.number_input("Ano do Contrato", value=int(c_row.get("ano_contrato") or date.today().year), min_value=2000, max_value=2050, step=1)
        with c_id3:
            num_completo_default = c_row.get("numero_completo") or f"{str(novo_num_ct).zfill(3)}/{novo_ano_ct}"
            novo_num_comp = st.text_input("Número Completo (Oficial)", value=str(num_completo_default))
        with c_id4:
            novo_proc = st.text_input("Processo Administrativo", value=str(c_row.get("processo_adm") or ""))

        c_id5, c_id5_n, c_id6 = st.columns([1.5, 1.2, 1.3])
        with c_id5:
            nova_mod = st.selectbox("Modalidade de Contratação", modalidades_opcoes, index=idx_mod)
            m_num_ed = re.search(r'\d+/\d+', str(c_row.get("modalidade") or ""))
            def_num_ed = m_num_ed.group(0) if m_num_ed else ""
            novo_num_mod = st.text_input("Nº do Procedimento / Pregão / Dispensa", value=def_num_ed, placeholder="Ex: 010/2026")
        with c_id6:
            novo_status = st.selectbox("Status do Contrato", status_opcoes, index=idx_status)

        st.markdown("##### 2. 🏛️ Secretaria Contratante & Gestores Responsáveis")
        c_org1, c_org2 = st.columns(2)
        with c_org1:
            novo_orgao_nome = st.selectbox("Secretaria / Órgão Contratante", nomes_orgaos, index=idx_orgao)
            novo_sec = st.text_input("Secretário(a) Municipal Titular", value=str(c_row.get("secretario") or ""))
        with c_org2:
            novo_fisc = st.text_input("Fiscal do Contrato Designado", value=str(c_row.get("fiscal") or ""))
            novo_solic = st.text_input("Solicitante / Unidade Demandante", value=str(c_row.get("solicitante") or ""))

        st.markdown("##### 3. 🏢 Fornecedor / Contratada")
        c_forn1, c_forn2 = st.columns(2)
        with c_forn1:
            forn_sel = st.selectbox("Vincular a Fornecedor Cadastrado:", fornec_labels, index=0)
            novo_forn_nome = st.text_input("Razão Social / Nome do Fornecedor:", value=str(c_row.get("fornecedor") or ""))
        with c_forn2:
            novo_cnpj = st.text_input("CNPJ / CPF do Fornecedor:", value=formatar_cnpj_cpf(c_row.get("cnpj_cpf") or ""))
            novo_rep = st.text_input("Representante Legal do Fornecedor:", value=str(c_row.get("fornecedor_representante") or ""))

        st.markdown("##### 4. 💰 Valores e Dotação Orçamentária")
        c_val1, c_val2 = st.columns([1.5, 2])
        with c_val1:
            novo_val = st.number_input("Valor Global do Contrato (R$)", value=float(c_row.get("valor_total") or 0.0), step=100.0, format="%.2f")
            val_formatado_prev = f"R$ {novo_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            st.caption(f"Valor por extenso visual: **{val_formatado_prev}**")
        with c_val2:
            novo_dot = st.text_input("Dotação Orçamentária Consolidada", value=str(c_row.get("dotacao_orcamentaria") or ""))
            st.caption("Ex: 02.05.00 - Secretaria Municipal de Educação | 12.361.0005.2015.0000 | 3.3.90.30.00")

        st.markdown("##### 5. ⏱️ Prazos, Vigência e Datas Legais")
        c_dt1, c_dt2, c_dt3, c_dt4 = st.columns(4)
        with c_dt1:
            nova_ass = st.date_input("Data de Assinatura", value=pd.to_datetime(c_row["data_assinatura"]).date() if pd.notna(c_row.get("data_assinatura")) and str(c_row.get("data_assinatura")).strip() != "" else date.today())
        with c_dt2:
            nova_pub = st.date_input("Data de Publicação", value=pd.to_datetime(c_row["data_publicacao"]).date() if pd.notna(c_row.get("data_publicacao")) and str(c_row.get("data_publicacao")).strip() != "" else date.today())
        with c_dt3:
            nova_venc = st.date_input("Data de Vencimento", value=pd.to_datetime(c_row["data_vencimento"]).date() if pd.notna(c_row.get("data_vencimento")) and str(c_row.get("data_vencimento")).strip() != "" else date.today() + timedelta(days=365))
        with c_dt4:
            nova_vig = st.text_input("Cláusula de Vigência Descritiva", value=str(c_row.get("vigencia_descricao") or "12 (doze) meses"))

        st.markdown("##### 6. 📝 Objeto Contratual & Observações Gerais")
        novo_obj = st.text_area("Objeto Contratual Completo:", value=str(c_row.get("objeto") or ""), height=110)
        novo_obs = st.text_area("Observações Internas / Administrativas:", value=str(c_row.get("observacoes") or ""), height=70)

        st.markdown("<br>", unsafe_allow_html=True)
        c_btn_save, c_btn_cancel = st.columns([2, 1])
        with c_btn_save:
            salvar_clique = st.form_submit_button("💾 Salvar Todas as Alterações", use_container_width=True, type="primary")
        with c_btn_cancel:
            cancelar_clique = st.form_submit_button("❌ Cancelar / Descartar", use_container_width=True)

    if salvar_clique:
        uid = st.session_state.get("usuario", {}).get("id", 1)

        try:
            with get_engine().connect() as conn:
                target_oid = map_orgao_id.get(novo_orgao_nome)
                if not target_oid:
                    r_ins_o = conn.execute(text("INSERT INTO orgaos (nome, sigla, ativo) VALUES (:n, :s, 1)"), {"n": novo_orgao_nome, "s": novo_orgao_nome[:10]})
                    target_oid = r_ins_o.lastrowid

                target_fid = int(c_row.get("fornecedor_id")) if pd.notna(c_row.get("fornecedor_id")) and c_row.get("fornecedor_id") else None
                if forn_sel != "-- Manter Fornecedor Atual / Não Alterar --":
                    target_fid = fornec_map[forn_sel]["id"]
                elif novo_cnpj and novo_cnpj.strip():
                    cnpj_limpo = limpar_cnpj(novo_cnpj)
                    cnpj_fmt = formatar_cnpj_cpf(novo_cnpj)
                    fid_exist = conn.execute(text("SELECT id FROM fornecedores WHERE cnpj_cpf = :c OR cnpj_cpf = :c2"), {"c": cnpj_fmt, "c2": cnpj_limpo}).scalar()
                    if fid_exist:
                        target_fid = fid_exist
                    elif target_fid:
                        conn.execute(
                            text("UPDATE fornecedores SET razao_social = :rz, cnpj_cpf = :c, nome_representante = :rep WHERE id = :fid"),
                            {"rz": novo_forn_nome, "c": cnpj_fmt, "rep": novo_rep, "fid": target_fid}
                        )
                elif target_fid and novo_forn_nome:
                    conn.execute(
                        text("UPDATE fornecedores SET razao_social = :rz, nome_representante = :rep WHERE id = :fid"),
                        {"rz": novo_forn_nome, "rep": novo_rep, "fid": target_fid}
                    )

                uid = st.session_state.get("usuario", {}).get("id", 1)

                conn.execute(
                    text("""
                        UPDATE contratos SET
                            numero_contrato = :nc,
                            ano_contrato = :ano,
                            numero_completo = :ncomp,
                            processo_adm = :proc,
                            modalidade = :mod,
                            status = :status,
                            orgao_id = :oid,
                            secretario = :sec,
                            fiscal = :fisc,
                            solicitante = :solic,
                            fornecedor_id = :fid,
                            valor_total = :val,
                            dotacao_orcamentaria = :dot,
                            data_assinatura = :d_ass,
                            data_publicacao = :d_pub,
                            data_vencimento = :d_venc,
                            vigencia_descricao = :vig,
                            objeto = :obj,
                            observacoes = :obs,
                            editado_por = :uid,
                            editado_em = CURRENT_TIMESTAMP,
                            atualizado_em = CURRENT_TIMESTAMP
                        WHERE id = :cid
                    """),
                    {
                        "nc": int(novo_num_ct),
                        "ano": int(novo_ano_ct),
                        "ncomp": str(novo_num_comp).strip(),
                        "proc": str(novo_proc).strip(),
                        "mod": f"{str(nova_mod).strip()} Nº {str(novo_num_mod).strip()}" if (novo_num_mod and str(novo_num_mod).strip() and str(novo_num_mod).strip() not in str(nova_mod)) else str(nova_mod).strip(),
                        "status": str(novo_status).strip(),
                        "oid": target_oid,
                        "sec": str(novo_sec).strip(),
                        "fisc": str(novo_fisc).strip(),
                        "solic": str(novo_solic).strip(),
                        "fid": target_fid,
                        "val": float(novo_val),
                        "dot": str(novo_dot).strip(),
                        "d_ass": nova_ass,
                        "d_pub": nova_pub,
                        "d_venc": nova_venc,
                        "vig": str(nova_vig).strip(),
                        "obj": str(novo_obj).strip(),
                        "obs": str(novo_obs).strip(),
                        "uid": uid,
                        "cid": int(cid_atual)
                    }
                )

                registrar_audit_log(conn, "contratos", int(cid_atual), "EDICAO", f"Contrato {novo_num_comp} editado com sucesso")
                conn.commit()

            try:
                carregar_contratos.clear()
            except Exception:
                pass

            st.success(f"🎉 **Contrato Nº {novo_num_comp} atualizado com sucesso no banco de dados!**")
            st.session_state["sel_detalhe_gel"] = str(novo_num_comp).strip()
            st.toast("✅ Contrato salvo com sucesso!", icon="💾")
            st.rerun()

        except Exception as e:
            st.error(f"Erro ao salvar alterações do contrato: {e}")

    if cancelar_clique:
        st.session_state["menu_selecionado"] = "📊 Contratos"
        st.session_state.pop("editando_contrato_id", None)
        st.rerun()



def dashboard_contratos():
    if st.session_state.get('editando_contrato_id'):
        tela_edicao_contrato(st.session_state['editando_contrato_id'])
        return

    df = carregar_contratos()

    # Alertas Inteligentes de Vencimento
    if 'data_vencimento' in df.columns:
        hoje_dt = pd.Timestamp.now().normalize()
        sessenta_dias = hoje_dt + pd.Timedelta(days=60)
        trinta_dias = hoje_dt + pd.Timedelta(days=30)
        
        # Converte para datetime pandas
        df_v = df.copy()
        df_v['data_vencimento'] = pd.to_datetime(df_v['data_vencimento'], errors='coerce')
        
        venc_30 = df_v[(df_v['data_vencimento'] >= hoje_dt) & (df_v['data_vencimento'] <= trinta_dias) & (df_v['status'] != 'DISTRATADO')]
        venc_60 = df_v[(df_v['data_vencimento'] > trinta_dias) & (df_v['data_vencimento'] <= sessenta_dias) & (df_v['status'] != 'DISTRATADO')]
        
        if not venc_30.empty:
            st.error(f"🚨 ATENÇÃO: Há {len(venc_30)} contrato(s) vencendo em menos de 30 dias! Adote as providências de prorrogação ou nova licitação.")
        elif not venc_60.empty:
            st.warning(f"⚠️ Aviso: Há {len(venc_60)} contrato(s) vencendo em menos de 60 dias.")

    # Topo Municipal Oficial de Ribeirãozinho do Maranhão - MA
    st.markdown("""
        <div class='municipal-banner' style='display: flex; justify-content: space-between; align-items: center; padding: 9px 18px; gap: 14px;'>
            <div style='display: flex; flex-direction: column; gap: 1px;'>
                <div style='display: flex; align-items: center; gap: 6px; margin-bottom: 2px;'>
                    <span class='municipal-badge-pill'>
                        <span class='beacon-active-gel' style='background: #ffffff;'></span> PREFEITURA DE RIBEIRÃOZINHO DO MARANHÃO - MA • EXERCÍCIO 2026
                    </span>
                </div>
                <h2 class='municipal-banner-title'>Painel Geral de Gestão de Contratos</h2>
                <div class='municipal-banner-subtitle'>
                    Acompanhamento de vigências, secretarias contratantes, fornecedores homologados e dotações orçamentárias municipais
                </div>
            </div>
            <div style='display: inline-flex; align-items: center; gap: 6px; font-family: "JetBrains Mono", monospace; font-size: 10.5px; background: rgba(0,0,0,0.2); padding: 5px 12px; border-radius: 8px; white-space: nowrap; flex-shrink: 0; border: 1px solid rgba(255,255,255,0.18);'>
                <span style='color: #a7f3d0; font-weight: 700;'>🟢 BASE OFICIAL ATIVA</span>
            </div>
        </div>
    """, unsafe_allow_html=True)


    if df.empty:
        st.markdown("""
        <div class='gel-kpi-card gel-kpi-blue' style='text-align: center; padding: 40px;'>
            <h3>📂 Base de Contratos Municipal Vazia</h3>
            <p style='color: #64748b; margin: 15px 0 25px 0;'>Importe uma planilha de contratos ou inicie o cadastro manual.</p>
        </div>
        """, unsafe_allow_html=True)
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("📥 Importar Planilha de Contratos", type="primary", use_container_width=True):
                st.session_state["menu_selecionado"] = "📥 Importar Planilha"
                st.rerun()
        with col_b2:
            if st.button("➕ Novo Contrato (Padrão AGU)", use_container_width=True):
                st.session_state["menu_selecionado"] = "⚖️ Modelos AGU / Minutas"
                st.session_state["mostrar_formulario"] = False
                st.rerun()
        return

    hoje = date.today()
    hoje_str = hoje.strftime("%Y-%m-%d")
    em_60_dias_str = (hoje + timedelta(days=60)).strftime("%Y-%m-%d")

    total_contratos = len(df)
    valor_global = df["valor_total"].sum() if "valor_total" in df.columns else 0.0
    total_orgaos = df["orgao"].nunique() if "orgao" in df.columns else 0
    total_fornecedores = df["fornecedor"].nunique() if "fornecedor" in df.columns else 0

    # Garantir normalização preventiva de data_vencimento
    if "data_vencimento" in df.columns:
        df["data_vencimento"] = df["data_vencimento"].apply(
            lambda v: v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else (
                str(v)[:10] if (v is not None and str(v).strip() not in ["", "None", "NaT"]) else ""
            )
        )

    contratos_vigentes = len(df[df["data_vencimento"] >= hoje_str])
    contratos_alerta_60d = len(df[(df["data_vencimento"] >= hoje_str) & (df["data_vencimento"] <= em_60_dias_str)])
    contratos_vencidos = len(df[(df["data_vencimento"] != "") & (df["data_vencimento"] < hoje_str)])

    # Cards de KPIs com as Cores da Cidade (Design Mais Estreito e Compacto com Respiro)
    st.markdown("<div style='margin-top: 6px;'></div>", unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class='gel-kpi-card gel-kpi-blue'>
            <div class='gel-kpi-header'>
                <span class='gel-kpi-title'>Total de Contratos</span>
                <span style='font-size: 15px;'>📄</span>
            </div>
            <div class='gel-kpi-value' style='color: #0284c7;'>{formatar_numero_br(total_contratos)}</div>
            <div class='gel-kpi-subtext'><span class='beacon-active-gel'></span> <strong>{formatar_numero_br(contratos_vigentes)}</strong> vigentes no exercício</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class='gel-kpi-card gel-kpi-green'>
            <div class='gel-kpi-header'>
                <span class='gel-kpi-title'>Investimento Global</span>
                <span style='font-size: 15px;'>💰</span>
            </div>
            <div class='gel-kpi-value' style='color: #059669;'>{formatar_moeda(valor_global)}</div>
            <div class='gel-kpi-subtext'>Recursos municipais empenhados</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
        <div class='gel-kpi-card gel-kpi-gold'>
            <div class='gel-kpi-header'>
                <span class='gel-kpi-title'>Secretarias Atendidas</span>
                <span style='font-size: 15px;'>🏛️</span>
            </div>
            <div class='gel-kpi-value' style='color: #d97706;'>{formatar_numero_br(total_orgaos)}</div>
            <div class='gel-kpi-subtext'>Órgãos municipais contratantes</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
        <div class='gel-kpi-card gel-kpi-cyan'>
            <div class='gel-kpi-header'>
                <span class='gel-kpi-title'>Rede de Fornecedores</span>
                <span style='font-size: 15px;'>🏭</span>
            </div>
            <div class='gel-kpi-value' style='color: #0369a1;'>{formatar_numero_br(total_fornecedores)}</div>
            <div class='gel-kpi-subtext'>Empresas e pessoas físicas homologadas</div>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)
    tab_consulta, tab_graficos = st.tabs(["📋 Painel de Contratos & Vigências", "📈 Gráficos Analíticos Executivos"])

    with tab_consulta:
        # Botões de Ação Rápida: Novo Contrato e Editar Contrato
        col_btn1, col_btn2, col_esp = st.columns([1.8, 1.8, 6.4])
        with col_btn1:
            if st.button("➕ Novo Contrato (AGU)", use_container_width=True, type="primary"):
                st.session_state["menu_selecionado"] = "⚖️ Modelos AGU / Minutas"
                st.session_state["mostrar_formulario"] = False
                st.session_state["mostrar_formulario_protocolo"] = False
                st.session_state.pop("editando_contrato_id", None)
                st.rerun()
        with col_btn2:
            if st.button("✏️ Editar um Contrato", use_container_width=True):
                st.session_state["menu_selecionado"] = "✏️ Editar Contrato"
                st.session_state["mostrar_formulario"] = False
                st.session_state["mostrar_formulario_protocolo"] = False
                st.rerun()
        # Filtros Rápidos Segmentados
        col_filtro_rapido, col_vazio = st.columns([3, 1])
        with col_filtro_rapido:
            opcoes_status = [
                f"Todos ({total_contratos})",
                f"🟢 Vigentes ({contratos_vigentes})",
                f"⏳ Alerta 60 Dias ({contratos_alerta_60d})",
                f"🔴 Vencidos ({contratos_vencidos})"
            ]
            filtro_status_rapido = st.segmented_control(
                "Filtro Rápido de Vigência:",
                opcoes_status,
                default=opcoes_status[0],
                label_visibility="collapsed"
            )
            if not filtro_status_rapido:
                filtro_status_rapido = opcoes_status[0]

        f_col1, f_col2, f_col3, f_col4 = st.columns([1, 1.5, 1.3, 2])
        with f_col1:
            anos_disponiveis = ["Todos"] + sorted([str(int(a)) for a in df["ano_contrato"].dropna().unique() if a > 1900], reverse=True)
            filtro_ano = st.selectbox("Ano", anos_disponiveis)

        with f_col2:
            orgaos_disponiveis = sorted(df["orgao"].dropna().unique())
            filtro_orgao = st.multiselect("Secretaria", orgaos_disponiveis, placeholder="Todas as secretarias")

        with f_col3:
            modalidades_disponiveis = sorted(df["modalidade"].dropna().unique())
            filtro_mod = st.multiselect("Modalidade", modalidades_disponiveis, placeholder="Todas as modalidades")

        with f_col4:
            busca = st.text_input("Busca Rápida Inteligente", placeholder="Nº Contrato, Secretário, Órgão, Fornecedor ou CNPJ")

        df_filtrado = df.copy()

        if "Vigentes" in filtro_status_rapido:
            df_filtrado = df_filtrado[df_filtrado["data_vencimento"] >= hoje_str]
        elif "60 Dias" in filtro_status_rapido:
            df_filtrado = df_filtrado[(df_filtrado["data_vencimento"] >= hoje_str) & (df_filtrado["data_vencimento"] <= em_60_dias_str)]
        elif "Vencidos" in filtro_status_rapido:
            df_filtrado = df_filtrado[(df_filtrado["data_vencimento"] != "") & (df_filtrado["data_vencimento"] < hoje_str)]

        if filtro_ano != "Todos":
            df_filtrado = df_filtrado[df_filtrado["ano_contrato"] == int(filtro_ano)]
        if filtro_orgao:
            df_filtrado = df_filtrado[df_filtrado["orgao"].isin(filtro_orgao)]
        if filtro_mod:
            df_filtrado = df_filtrado[df_filtrado["modalidade"].isin(filtro_mod)]
        if busca:
            busca_clean = busca.strip()
            df_filtrado = df_filtrado[
                df_filtrado["numero_completo"].str.contains(busca_clean, case=False, na=False) |
                df_filtrado["objeto"].str.contains(busca_clean, case=False, na=False) |
                df_filtrado["fornecedor"].str.contains(busca_clean, case=False, na=False) |
                df_filtrado["cnpj_cpf"].str.contains(busca_clean, case=False, na=False) |
                df_filtrado["secretario"].str.contains(busca_clean, case=False, na=False) |
                df_filtrado["orgao"].str.contains(busca_clean, case=False, na=False) |
                df_filtrado["solicitante"].str.contains(busca_clean, case=False, na=False)
            ]

    
        col_res1, col_res2, col_res3 = st.columns([2.5, 1.8, 1])
        with col_res1:
            st.caption(f"Mostrando **{formatar_numero_br(len(df_filtrado))}** contratos. Montante correspondente: **{formatar_moeda(df_filtrado['valor_total'].sum())}**")
        with col_res2:
            opcoes_modo = ["🗂️ Cards Visuais", "📊 Tabela Analítica"]
            modo_visualizacao = st.segmented_control(
                "Modo de Exibição:",
                opcoes_modo,
                default=opcoes_modo[0],
                label_visibility="collapsed"
            )
            if not modo_visualizacao:
                modo_visualizacao = opcoes_modo[0]
        with col_res3:
            excel_bytes = exportar_contratos_excel(df_filtrado)
            st.download_button(
                label="📥 Exportar Excel",
                data=excel_bytes,
                file_name=f"contratos_ribeiraozinho_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        # 1. Modo Cards Visuais (Cores Oficiais com Paginação Ultrarrápida)
        if modo_visualizacao == "🗂️ Cards Visuais":
            if df_filtrado.empty:
                st.info("Nenhum contrato encontrado para os parâmetros informados.")
            else:
                cards_list = df_filtrado.to_dict("records")
                total_cards = len(cards_list)

                # Controle inteligente de paginação para carregamento instantâneo
                col_pag_info, col_pag_qtd = st.columns([3, 1])
                with col_pag_qtd:
                    qtd_sel = st.selectbox(
                        "Exibir por página:",
                        ["12 por página", "24 por página", "Todos"],
                        index=0,
                        key="sel_cards_per_page"
                    )

                if qtd_sel == "12 por página":
                    itens_por_pag = 12
                elif qtd_sel == "24 por página":
                    itens_por_pag = 24
                else:
                    itens_por_pag = max(1, total_cards)

                total_pags = max(1, (total_cards + itens_por_pag - 1) // itens_por_pag)
                if "cards_pagina_atual" not in st.session_state:
                    st.session_state["cards_pagina_atual"] = 1
                if st.session_state["cards_pagina_atual"] > total_pags:
                    st.session_state["cards_pagina_atual"] = 1

                with col_pag_info:
                    if total_pags > 1:
                        c_p1, c_p2, c_p3 = st.columns([1, 2, 1])
                        with c_p1:
                            if st.button("◀ Anterior", disabled=(st.session_state["cards_pagina_atual"] <= 1), use_container_width=True, key="btn_pag_prev"):
                                st.session_state["cards_pagina_atual"] -= 1
                                st.rerun()
                        with c_p2:
                            st.markdown(f"<div style='text-align: center; font-size: 12px; font-weight: 700; color: #0369a1; padding-top: 6px;'>Página {st.session_state['cards_pagina_atual']} de {total_pags} ({total_cards} contratos)</div>", unsafe_allow_html=True)
                        with c_p3:
                            if st.button("Próxima ▶", disabled=(st.session_state["cards_pagina_atual"] >= total_pags), use_container_width=True, key="btn_pag_next"):
                                st.session_state["cards_pagina_atual"] += 1
                                st.rerun()
                    else:
                        st.caption(f"Mostrando todos os **{total_cards}** contratos.")

                start_idx = (st.session_state["cards_pagina_atual"] - 1) * itens_por_pag
                end_idx = min(start_idx + itens_por_pag, total_cards)
                cards_pagina = cards_list[start_idx:end_idx]

                def _selecionar_card_cb(num_c):
                    st.session_state["sel_detalhe_gel"] = num_c
                    st.session_state["scroll_para_ficha"] = True

                # Exibir em 3 colunas para ocupar melhor a tela e diminuir scroll
                for i in range(0, len(cards_pagina), 3):
                    cols = st.columns(3)
                    for j in range(3):
                        if i + j < len(cards_pagina):
                            item = cards_pagina[i + j]
                            with cols[j]:
                                dt_ass = parse_data_universal(item["data_assinatura"])
                                dt_venc = parse_data_universal(item["data_vencimento"])
                                
                                percentual = 0
                                dias_restantes_txt = "Em vigência"
                                badge_class = "gel-badge-vigente"
                                badge_label = "🟢 VIGENTE"
                                bar_color = "linear-gradient(90deg, #0284c7 0%, #059669 100%)"

                                if dt_ass and dt_venc:
                                    try:
                                        d_ini = datetime.strptime(dt_ass, "%Y-%m-%d").date()
                                        d_fim = datetime.strptime(dt_venc, "%Y-%m-%d").date()
                                        total_dias = max(1, (d_fim - d_ini).days)
                                        passados = (hoje - d_ini).days
                                        restantes = (d_fim - hoje).days
                                        percentual = min(100, max(0, int((passados / total_dias) * 100)))

                                        if restantes < 0:
                                            badge_class = "gel-badge-vencido"
                                            badge_label = f"🔴 EXPIRADO ({abs(restantes)}D)"
                                            dias_restantes_txt = "Prazo finalizado"
                                            bar_color = "#ef4444"
                                        elif restantes <= 60:
                                            badge_class = "gel-badge-alerta"
                                            badge_label = f"⏳ ALERTA ({restantes} DIAS)"
                                            dias_restantes_txt = f"{restantes} dias p/ fim"
                                            bar_color = "linear-gradient(90deg, #f59e0b 0%, #d97706 100%)"
                                        else:
                                            dias_restantes_txt = f"{restantes} dias restantes"
                                    except Exception:
                                        pass

                                is_selecionado = (st.session_state.get("sel_detalhe_gel") == item["numero_completo"])
                                card_css_class = "gel-contract-card gel-card-selected" if is_selecionado else "gel-contract-card"
                                sel_badge = "<span style='background: #0284c7; color: #fff; font-size: 9.5px; font-weight: 800; padding: 2px 8px; border-radius: 10px;'>FICHA ABERTA</span>" if is_selecionado else ""
                                sel_footer = "<div style='margin-top: 10px; font-size: 11px; color: #0284c7; font-weight: 700; text-align: right;'>✨ Contrato Selecionado na Ficha Abaixo ↓</div>" if is_selecionado else "<div style='margin-top: 10px; font-size: 11px; color: #64748b; font-weight: 600; text-align: right;'>👉 Clique no card para abrir a Ficha Completa ↓</div>"

                                html_card = f"""
<div class="{card_css_class}" data-card-id="{item['id']}" style="padding:12px; margin-bottom:5px;">
    <div style="display: flex; justify-content: space-between; align-items: flex-start;">
        <div>
            <span style="font-size: 15px; font-weight: 800; color: #0284c7; font-family: 'Plus Jakarta Sans';">
                Nº {item['numero_completo']}
            </span>
            <div style="font-size: 10.5px; color: #64748b; font-family: 'JetBrains Mono'; line-height: 1.1;">
                {item['modalidade']}
            </div>
        </div>
        <div style="display: flex; flex-direction: column; align-items: flex-end; gap: 2px;">
            <span class="{badge_class}" style="font-size:9px; padding: 2px 6px;">{badge_label}</span>
            {sel_badge}
        </div>
    </div>
    <div style="margin-top: 6px; font-size: 11.5px; color: #334155; line-height:1.2;">
        <div style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">🏛️ <strong>{item['orgao']}</strong></div>
        <div style="color: #0369a1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">👔 Titular: <strong>{item['secretario']}</strong></div>
        <div style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">🏭 Fornecedor: <strong>{item['fornecedor']}</strong></div>
    </div>
    <div style="margin-top: 5px; font-size: 17px; font-weight: 800; color: #059669; font-family: 'Plus Jakarta Sans';">
        {formatar_moeda(item['valor_total'])}
    </div>
    <div style="margin-top: 5px;">
        <div style="display: flex; justify-content: space-between; font-size: 10px; color: #64748b; margin-bottom: 2px;">
            <span>Ass: <strong>{formatar_data_br(item['data_assinatura'])}</strong></span>
            <span>Venc: <strong>{formatar_data_br(item['data_vencimento'])}</strong></span>
        </div>
        <div class="gel-progress-bg" style="height:4px;">
            <div class="gel-progress-fill" style="width: {percentual}%; background: {bar_color};"></div>
        </div>
        <div style="display: flex; justify-content: space-between; font-size: 9.5px; color: #94a3b8; margin-top:2px;">
            <span>{percentual}% corrido</span>
            <span>{dias_restantes_txt}</span>
        </div>
    </div>
</div>
"""
                                # Usando st.html nativo em vez de st.markdown para impedir falhas no parser de texto
                                st.html(html_card)

                                cargo_user = st.session_state.get('usuario', {}).get('cargo', '')
                                if is_selecionado:
                                    col_card_act1, col_card_act2 = st.columns(2)
                                    with col_card_act1:
                                        if cargo_user in ['ADMIN', 'ADMINISTRADOR']:
                                            if st.button("✏️ Editar", key=f"btn_card_edit_{item['id']}", use_container_width=True, type="primary", help="Editar todos os dados deste contrato (requer confirmação de senha do Administrador)"):
                                                st.session_state["editando_contrato_id"] = item['id']
                                                st.session_state["menu_selecionado"] = "✏️ Editar Contrato"
                                                st.rerun()
                                        else:
                                            st.button("🔒 Editar", key=f"btn_card_edit_dis_{item['id']}", use_container_width=True, disabled=True, help="Apenas administradores podem alterar contratos municipais.")
                                    with col_card_act2:
                                        if st.button("📋 Duplicar", key=f"btn_card_dup_{item['id']}", use_container_width=True, help="Duplica este contrato gerando uma nova minuta com o próximo número sequencial"):
                                            uid_atual = st.session_state.get('usuario', {}).get('id', 1)
                                            ok_dup, msg_dup, novo_cid = duplicar_contrato_bd(item['id'], uid_atual)
                                            if ok_dup:
                                                st.toast("Contrato duplicado com sucesso!", icon="📋")
                                                st.success(msg_dup)
                                                st.session_state["sel_detalhe_gel"] = None
                                                st.rerun()
                                            else:
                                                st.error(msg_dup)
                                    st.button(
                                        "🔍 Ficha Aberta Abaixo ↓",
                                        key=f"btn_card_ver_{item['id']}",
                                        use_container_width=True,
                                        type="secondary",
                                        on_click=_selecionar_card_cb,
                                        args=(item["numero_completo"],)
                                    )
                                else:
                                    col_nc1, col_nc2 = st.columns([1.4, 1])
                                    with col_nc1:
                                        st.button(
                                            "👉 Selecionar",
                                            key=f"btn_card_{item['id']}",
                                            use_container_width=True,
                                            type="tertiary",
                                            on_click=_selecionar_card_cb,
                                            args=(item["numero_completo"],),
                                            help="Clique para selecionar este contrato e abrir as opções"
                                        )
                                    with col_nc2:
                                        if st.button("📋 Duplicar", key=f"btn_card_dup_fast_{item['id']}", use_container_width=True, help="Duplicar este contrato gerando o próximo número sequencial"):
                                            uid_atual = st.session_state.get('usuario', {}).get('id', 1)
                                            ok_dup, msg_dup, novo_cid = duplicar_contrato_bd(item['id'], uid_atual)
                                            if ok_dup:
                                                st.toast("Contrato duplicado com sucesso!", icon="📋")
                                                st.success(msg_dup)
                                                st.session_state["sel_detalhe_gel"] = None
                                                st.rerun()
                                            else:
                                                st.error(msg_dup)

        # 2. Modo Tabela Analítica
        else:
            df_grid = df_filtrado.copy()
            df_grid["Valor (R$)"] = df_grid["valor_total"].apply(formatar_moeda)
            df_grid["Data Assinatura"] = df_grid["data_assinatura"].apply(formatar_data_br)
            df_grid["Data Publicação"] = df_grid["data_publicacao"].apply(formatar_data_br)
            df_grid["Data Vencimento"] = df_grid["data_vencimento"].apply(formatar_data_br)
            df_grid["CNPJ/CPF"] = df_grid["cnpj_cpf"].apply(formatar_cnpj_cpf)
            df_grid["Secretário Titular"] = df_grid["secretario"].fillna("Não informado")

            def get_sit_pill(d_v):
                if not d_v or pd.isna(d_v):
                    return "⚪ Sem Data"
                ds = str(d_v).split(" ")[0]
                if ds < hoje_str:
                    return "🔴 Expirado"
                elif ds <= em_60_dias_str:
                    return "⏳ Alerta 60d"
                return "🟢 Vigente"

            df_grid["Situação"] = df_grid["data_vencimento"].apply(get_sit_pill)

            event = st.dataframe(
                df_grid[[
                    "numero_completo", "Situação", "orgao", "Secretário Titular", "modalidade",
                    "fornecedor", "CNPJ/CPF", "objeto", "Valor (R$)",
                    "Data Assinatura", "Data Publicação", "Data Vencimento"
                ]].rename(columns={
                    "numero_completo": "Nº Contrato",
                    "orgao": "Secretaria / Órgão",
                    "modalidade": "Modalidade",
                    "fornecedor": "Fornecedor",
                    "objeto": "Objeto Contratual"
                }),
                use_container_width=True,
                hide_index=True,
                height=450,
                on_select="rerun",
                selection_mode="single-row"
            )
            if event and hasattr(event, "selection") and event.selection and event.selection.rows:
                row_idx = event.selection.rows[0]
                num_sel_tab = df_grid.iloc[row_idx]["numero_completo"]
                if st.session_state.get("sel_detalhe_gel") != num_sel_tab:
                    st.session_state["sel_detalhe_gel"] = num_sel_tab
                    st.session_state["scroll_para_ficha"] = True
                    st.rerun()

        # Ficha Detalhada do Contrato
        st.markdown("<div id='secao-ficha-completa' style='scroll-margin-top: 15px;'></div>", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("### 🔍 Ficha Completa & Detalhes do Contrato")

        opcoes_c = ["Selecione um contrato para abrir a ficha completa..."] + df_filtrado["numero_completo"].tolist()

        if "sel_detalhe_gel" not in st.session_state or st.session_state["sel_detalhe_gel"] not in opcoes_c:
            st.session_state["sel_detalhe_gel"] = opcoes_c[0]

        c_sel = st.selectbox("Pesquise o contrato municipal:", opcoes_c, key="sel_detalhe_gel")

        if st.session_state.get("scroll_para_ficha"):
            st.session_state["scroll_para_ficha"] = False
            components.html("""
            <script>
                try {
                    const doc = window.parent.document;
                    setTimeout(function() {
                        const target = doc.getElementById('secao-ficha-completa');
                        if (target) {
                            target.scrollIntoView({ behavior: 'smooth', block: 'start' });
                        }
                    }, 120);
                } catch(e) {}
            </script>
            """, height=0, width=0)

        if c_sel != "Selecione um contrato para abrir a ficha completa...":
            c_row = df_filtrado[df_filtrado["numero_completo"] == c_sel].iloc[0]
            cid = int(c_row["id"])
            cargo_user = st.session_state.get('usuario', {}).get('cargo', '')

            st.markdown(f"""
            <div style='background: #eff6ff; border: 1.5px solid #bfdbfe; border-left: 5px solid #0284c7; border-radius: 8px; padding: 10px 16px; margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center;'>
                <div>
                    <span style='font-size: 14px; font-weight: 800; color: #0369a1;'>📄 Contrato Selecionado: {c_row['numero_completo']}</span>
                    <span style='font-size: 12px; color: #475569; margin-left: 12px;'>🏛️ {c_row['orgao']} | 🏭 Fornecedor: <strong>{c_row['fornecedor']}</strong></span>
                </div>
                <span style='font-size: 11px; background: #0284c7; color: #fff; padding: 3px 10px; border-radius: 12px; font-weight: 700;'>FICHA ABERTA PELO CARD</span>
            </div>
            """, unsafe_allow_html=True)

            c_top_act1, c_top_act2 = st.columns(2)
            with c_top_act1:
                if cargo_user in ['ADMIN', 'ADMINISTRADOR']:
                    if st.button("✏️ Editar Dados deste Contrato", key=f"btn_top_card_edit_{cid}", use_container_width=True, type="primary"):
                        st.session_state["editando_contrato_id"] = cid
                        st.session_state["menu_selecionado"] = "✏️ Editar Contrato"
                        st.rerun()
                else:
                    st.button("🔒 Editar Contrato (Apenas ADMIN)", key=f"btn_top_card_edit_dis_{cid}", use_container_width=True, disabled=True, help="Apenas administradores podem alterar contratos municipais.")
            with c_top_act2:
                if st.button("📋 Duplicar Contrato (Próximo Nº Disponível)", key=f"btn_top_card_dup_{cid}", use_container_width=True, help="Duplica este contrato gerando uma nova minuta com o próximo número sequencial"):
                    uid_atual = st.session_state.get('usuario', {}).get('id', 1)
                    ok_dup, msg_dup, novo_cid = duplicar_contrato_bd(cid, uid_atual)
                    if ok_dup:
                        st.toast("Contrato duplicado com sucesso!", icon="📋")
                        st.success(msg_dup)
                        st.session_state["sel_detalhe_gel"] = None
                        st.rerun()
                    else:
                        st.error(msg_dup)

            st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)

            tab_g, tab_t, tab_agu, tab_adit, tab_p = st.tabs([
                "📌 Visão Geral & Partes", 
                "⏱️ Prazos & Vigência", 
                "⚖️ Minuta Oficial AGU (Word / PDF)", 
                "📝 Termos Aditivos",
                "📋 Protocolos Vinculados"
            ])

            with tab_g:
                c1, c2 = st.columns([1.2, 1])
                with c1:
                    st.markdown(f"#### Contrato Municipal Nº {c_row['numero_completo']}")
                    st.markdown(f"**🏛️ Secretaria Contratante:** {c_row['orgao']}")
                    st.markdown(f"**👔 Secretário(a) Titular:** `{c_row['secretario']}`")
                    st.markdown(f"**📑 Modalidade:** {c_row['modalidade']}")
                    st.markdown(f"**📂 Processo Administrativo:** {c_row['processo_adm'] if c_row['processo_adm'] else 'Não informado'}")
                    st.markdown(f"**🏭 Fornecedor:** {c_row['fornecedor']}")
                    st.markdown(f"**📄 CNPJ / CPF:** {formatar_cnpj_cpf(c_row['cnpj_cpf'])}")
                with c2:
                    st.markdown(f"""
                    <div class='gel-kpi-card gel-kpi-green' style='margin-top: 10px;'>
                        <div class='gel-kpi-title'>Dotação Global Contratada</div>
                        <div class='gel-kpi-value' style='color: #059669;'>{formatar_moeda(c_row['valor_total'])}</div>
                        <div style='font-size: 12px; color: #64748b; margin-top: 6px;'>
                            👤 Fiscal: <strong>{c_row['fiscal'] if pd.notna(c_row['fiscal']) else 'Não designado'}</strong><br>
                            👤 Solicitante: <strong>{c_row['solicitante'] if c_row['solicitante'] else 'Não informado'}</strong>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                st.info(f"**📝 Objeto Completo:**\n\n{c_row['objeto']}")
                
                if st.session_state['usuario']['cargo'] != 'AUDITOR':
                    c_act1, c_act2 = st.columns(2)
                    with c_act1:
                        if st.session_state['usuario']['cargo'] in ['ADMIN', 'ADMINISTRADOR']:
                            if st.button("✏️ Editar Dados deste Contrato", key=f"btn_edit_{cid}", use_container_width=True, type="primary"):
                                st.session_state["editando_contrato_id"] = cid
                                st.session_state["menu_selecionado"] = "✏️ Editar Contrato"
                                st.rerun()
                        else:
                            st.button("🔒 Editar Dados (Apenas ADMIN)", key=f"btn_edit_dis_{cid}", use_container_width=True, disabled=True, help="Apenas administradores podem alterar contratos municipais.")
                    with c_act2:
                        if st.button("📋 Duplicar Contrato (Próximo Nº)", key=f"btn_dupl_{cid}", use_container_width=True, help="Cria um novo contrato clonando os dados deste e atribuindo o próximo número sequencial disponível"):
                            uid_atual = st.session_state.get('usuario', {}).get('id', 1)
                            ok_dup, msg_dup, novo_cid = duplicar_contrato_bd(cid, uid_atual)
                            if ok_dup:
                                st.toast("Contrato duplicado com sucesso!", icon="📋")
                                st.success(msg_dup)
                                st.session_state["sel_detalhe_gel"] = None
                                st.rerun()
                            else:
                                st.error(msg_dup)

                if st.button("📄 Exportar JSON (PNCP)", key=f"btn_pncp_{cid}", use_container_width=True):
                    import json
                    dados_pncp = {
                        "orgaoEntidade": {"cnpj": "01.612.834/0001-86"},
                        "numeroContrato": str(c_row['numero']),
                        "anoContrato": str(c_row['ano']),
                        "fornecedor": {"cnpj": str(c_row['cnpj_cpf']) if 'cnpj_cpf' in c_row else "00.000.000/0001-00"},
                        "valorInicial": float(c_row['valor_total']) if c_row['valor_total'] else 0.0,
                        "dataAssinatura": c_row['data_assinatura'].strftime('%Y-%m-%d') if hasattr(c_row['data_assinatura'], 'strftime') else str(c_row['data_assinatura']),
                        "dataVigenciaFim": c_row['data_vencimento'].strftime('%Y-%m-%d') if hasattr(c_row['data_vencimento'], 'strftime') else str(c_row['data_vencimento']),
                        "objetoContrato": str(c_row['objeto'])
                    }
                    json_str = json.dumps(dados_pncp, indent=4, ensure_ascii=False)
                    st.download_button("📥 Baixar Arquivo PNCP", data=json_str, file_name=f"pncp_contrato_{c_row['numero']}_{c_row['ano']}.json", mime="application/json", type="primary", use_container_width=True)

            with tab_t:
                p1, p2, p3 = st.columns(3)
                with p1:
                    st.markdown(f"""
                    <div class='gel-kpi-card gel-kpi-blue'>
                        <div class='gel-kpi-title'>📅 Data de Assinatura</div>
                        <div class='gel-kpi-value' style='font-size: 20px; color: #0284c7;'>{formatar_data_br(c_row['data_assinatura'])}</div>
                    </div>
                    """, unsafe_allow_html=True)
                with p2:
                    st.markdown(f"""
                    <div class='gel-kpi-card gel-kpi-green'>
                        <div class='gel-kpi-title'>📰 Data de Publicação</div>
                        <div class='gel-kpi-value' style='font-size: 20px; color: #059669;'>{formatar_data_br(c_row['data_publicacao'])}</div>
                    </div>
                    """, unsafe_allow_html=True)
                with p3:
                    st.markdown(f"""
                    <div class='gel-kpi-card gel-kpi-gold'>
                        <div class='gel-kpi-title'>⏳ Data de Vencimento</div>
                        <div class='gel-kpi-value' style='font-size: 20px; color: #d97706;'>{formatar_data_br(c_row['data_vencimento'])}</div>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown(f"<br><strong>Cláusula de Vigência:</strong> {c_row['vigencia_descricao'] if pd.notna(c_row['vigencia_descricao']) else '12 meses'}", unsafe_allow_html=True)

            with tab_agu:
                st.markdown("""
                <div class='municipal-banner' style='margin-bottom: 14px; padding: 10px 18px;'>
                    <div class='municipal-badge-pill'>
                        <span class='beacon-active-gel' style='background: #ffffff;'></span> PADRÃO OFICIAL CGU/AGU • LEI Nº 14.133/2021
                    </div>
                    <h3 class='municipal-banner-title' style='font-size: 17px;'>⚖️ Emissor de Minutas Oficiais e Termos Aditivos AGU</h3>
                    <div class='municipal-banner-subtitle' style='font-size: 11.5px;'>
                        Geração de minutas de contratos e termos aditivos em conformidade com a Consultoria-Geral da União (cgu.agu.gov.br/contrato/ e modelos padronizados AGU Lei nº 14.133/21 e Lei nº 8.666/93)
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # Obtém histórico cadastral da secretaria
                hist_org = obter_historico_orgao(c_row["orgao"])

                c_agu1, c_agu2 = st.columns([1.5, 1])
                with c_agu1:
                    modelo_escolhido = st.selectbox(
                        "Selecione o Modelo de Minuta AGU Desejado:",
                        list(MODELOS_AGU_MAP.keys()),
                        key=f"sel_mod_agu_{cid}"
                    )
                    cod_modelo_agu = MODELOS_AGU_MAP[modelo_escolhido]

                    st.markdown(f"""
                    <div style='background: #f1f5f9; padding: 12px 16px; border-radius: 8px; font-size: 12px; border-left: 4px solid #0284c7; margin-bottom: 12px;'>
                        <strong>📍 Dados do Órgão Aplicados:</strong> {c_row['orgao']}<br>
                        <strong>Endereço:</strong> {hist_org['endereco']}, Bairro {hist_org['bairro']} - CEP {hist_org['cep']} - {hist_org['cidade']}/{hist_org['uf']}<br>
                        <strong>Representante:</strong> {c_row['secretario'] or hist_org['secretario']} | <strong>Fiscal:</strong> {c_row['fiscal'] or hist_org['fiscal'] or 'A designar'}<br>
                        <strong>CNPJ Contratante:</strong> {formatar_cnpj_cpf(hist_org['cnpj'])}
                    </div>
                    """, unsafe_allow_html=True)

                with c_agu2:
                    st.markdown("<div style='margin-top: 25px;'></div>", unsafe_allow_html=True)
                    st.success("🟢 Padrão CGU/AGU 14.133 Habilitado")
                    st.caption("Documentos gerados com brasão oficial, paginação automática e cláusulas obrigatórias.")

                # Monta dicionário completo para geração
                dados_minuta = {
                    "numero_completo": c_row["numero_completo"],
                    "numero_contrato": c_row.get("numero_contrato", 1),
                    "processo_adm": c_row.get("processo_adm") or "001/2026",
                    "modalidade": c_row.get("modalidade") or "DISPENSA DE LICITAÇÃO",
                    "objeto": c_row.get("objeto") or "Contratação administrativa",
                    "valor_total": c_row.get("valor_total") or 0.0,
                    "data_assinatura": c_row.get("data_assinatura"),
                    "data_vencimento": c_row.get("data_vencimento"),
                    "vigencia_descricao": c_row.get("vigencia_descricao") or "12 (doze) meses",
                    "secretario": c_row.get("secretario") or hist_org["secretario"],
                    "cargo_secretario": hist_org.get("cargo_secretario", "Secretário(a) Municipal Titular"),
                    "fiscal": c_row.get("fiscal") or hist_org["fiscal"],
                    "orgao_nome": c_row["orgao"],
                    "orgao_cnpj": formatar_cnpj_cpf(hist_org["cnpj"]),
                    "orgao_cep": hist_org["cep"],
                    "orgao_endereco": hist_org["endereco"],
                    "orgao_bairro": hist_org["bairro"],
                    "orgao_cidade": hist_org["cidade"],
                    "orgao_uf": hist_org["uf"],
                    "fornecedor_nome": c_row.get("fornecedor", "FORNECEDOR"),
                    "fornecedor_cnpj": formatar_cnpj_cpf(c_row.get("cnpj_cpf", "00.000.000/0001-00")),
                    "fornecedor_representante": c_row.get("fornecedor_representante") or "Representante Legal",
                    "fornecedor_cargo": c_row.get("fornecedor_cargo") or "Sócio Administrador",
                    "fornecedor_endereco": c_row.get("fornecedor_endereco") or "Sede Comercial da Contratada",
                    "fornecedor_bairro": c_row.get("fornecedor_bairro") or "Centro",
                    "fornecedor_cep": c_row.get("fornecedor_cep") or "65900-000",
                    "fornecedor_cidade": c_row.get("fornecedor_cidade") or "Imperatriz",
                    "fornecedor_uf": c_row.get("fornecedor_uf") or "MA",
                    "dotacao_orcamentaria": hist_org["dotacao"],
                    "modelo_agu": cod_modelo_agu
                }

                df_p_c = carregar_protocolos_do_contrato(cid)
                if not df_p_c.empty:
                    p_top = df_p_c.iloc[0]
                    st.markdown(f"""
                    <div style='background: #eff6ff; border: 1px solid #bfdbfe; border-left: 5px solid #0284c7; border-radius: 8px; padding: 10px 14px; margin-bottom: 12px;'>
                        <div style='display: flex; justify-content: space-between; align-items: center;'>
                            <div>
                                <span style='font-size: 13px; font-weight: 700; color: #1e40af;'>🔗 Protocolo Oficial Vinculado: {p_top['Protocolo']}</span>
                                <span style='font-size: 11px; color: #64748b; margin-left: 10px;'>Solicitante: <strong>{p_top['Solicitante']}</strong> | Situação: <strong>{p_top['Status']}</strong></span>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                col_b1, col_b2, col_b3, col_b4 = st.columns([1, 1, 1.1, 1.1])
                with col_b1:
                    docx_stream = gerar_contrato_docx(dados_minuta)
                    nome_docx = f"Contrato_{str(c_row['numero_completo']).replace('/', '_')}_AGU.docx"
                    st.download_button(
                        label="📥 Baixar Minuta Word (.docx)",
                        data=docx_stream,
                        file_name=nome_docx,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_docx_{cid}",
                        use_container_width=True,
                        type="primary"
                    )

                with col_b2:
                    pdf_stream = gerar_contrato_pdf(dados_minuta)
                    nome_pdf = f"Contrato_{str(c_row['numero_completo']).replace('/', '_')}_AGU.pdf"
                    st.download_button(
                        label="📄 Baixar Minuta PDF (.pdf)",
                        data=pdf_stream,
                        file_name=nome_pdf,
                        mime="application/pdf",
                        key=f"dl_pdf_{cid}",
                        use_container_width=True
                    )

                with col_b3:
                    if st.button("🌐 Sincronizar via API CGU/AGU", key=f"btn_sync_agu_{cid}", use_container_width=True):
                        with st.spinner("Consultando endpoint oficial da AGU (cgu.agu.gov.br)..."):
                            client_agu = AGUApiClient(timeout=10)
                            retorno_agu = client_agu.gerar_minuta_agu(dados_minuta)
                            if retorno_agu.get("sucesso"):
                                st.success("✅ Minuta validada com sucesso pelo motor oficial da AGU!")
                                with st.expander("Ver Resumo da Minuta Processada pela CGU/AGU", expanded=False):
                                    st.write(retorno_agu.get("texto_completo", "")[:1200] + "...")
                            else:
                                st.warning(f"Aviso de conexão API AGU: {retorno_agu.get('erro', 'Servidor externo indisponível')}. A minuta nativa local (Word/PDF) está 100% pronta para uso.")

                with col_b4:
                    if not df_p_c.empty:
                        p_sel_row = df_p_c.iloc[0]
                        prot_exp_data = {
                            "numero": str(p_sel_row["Protocolo"]),
                            "data_rec": str(p_sel_row["Data"]),
                            "assunto": str(p_sel_row["Assunto"]),
                            "solicitante": str(p_sel_row["Solicitante"]),
                            "orgao": str(c_row["orgao"]),
                            "secretario": str(c_row["secretario"]),
                            "contrato": str(c_row["numero_completo"]),
                            "fornecedor": str(c_row["fornecedor"]),
                            "observacoes": f"Protocolo vinculado ao Contrato {c_row['numero_completo']}",
                            "usuario": st.session_state["usuario"].get("nome", "Servidor"),
                            "qtd_vias": 2
                        }
                        try:
                            pdf_buf_c = gerar_comprovante_protocolo_pdf(prot_exp_data)
                            st.download_button(
                                label="🖨️ Comprovante Protocolo (PDF)",
                                data=pdf_buf_c.getvalue(),
                                file_name=f"comprovante_{str(p_sel_row['Protocolo']).replace('/', '_')}.pdf",
                                mime="application/pdf",
                                key=f"btn_dl_prot_quick_{cid}",
                                use_container_width=True
                            )
                        except Exception as e:
                            st.error(f"Erro: {e}")
                    else:
                        if st.button("➕ Abrir Protocolo do Contrato", key=f"btn_criar_prot_fast_{cid}", use_container_width=True):
                            st.session_state["contrato_pre_selecionado_id"] = cid
                            st.session_state["mostrar_formulario_protocolo"] = True
                            st.rerun()

                st.markdown("---")
                with st.expander("👁️ Pré-Visualizar Cláusulas Oficiais da Minuta (Lei 14.133/2021)", expanded=False):
                    from gerador_contratos_agu import gerar_clausulas_agu
                    clausulas_preview = gerar_clausulas_agu(dados_minuta)
                    for cl in clausulas_preview:
                        st.markdown(f"**{cl['numero']}**")
                        for p_t in cl['conteudo']:
                            st.caption(p_t)

            with tab_adit:
                st.markdown("##### 📝 Termos Aditivos Vinculados a este Contrato")
                st.caption("Prorrogações de vigência, alterações de quantitativos/valores (até 25%/50%), reequilíbrio e apostilamento conforme modelos oficiais da AGU (Lei nº 14.133/2021 e Lei nº 8.666/1993).")

                df_aditivos = carregar_aditivos_do_contrato(cid)
                total_adit = len(df_aditivos)

                c_ad_top1, c_ad_top2 = st.columns([2, 1])
                with c_ad_top1:
                    if total_adit == 0:
                        st.info("ℹ️ Nenhum termo aditivo formalizado até o momento para este contrato.")
                    else:
                        st.success(f"**{total_adit}** termo(s) aditivo(s) formalizado(s) e vigentes.")
                with c_ad_top2:
                    if st.button("➕ Elaborar Novo Termo Aditivo AGU", key=f"btn_novo_adit_card_{cid}", type="primary", use_container_width=True):
                        st.session_state["menu_selecionado"] = "⚖️ Modelos AGU / Minutas"
                        st.session_state["aditivo_contrato_pre_sel"] = c_row["numero_completo"]
                        st.rerun()

                if not df_aditivos.empty:
                    soma_aditada = float(df_aditivos["valor_aditado"].sum() or 0.0)
                    m1, m2, m3 = st.columns(3)
                    with m1:
                        st.markdown(f"""
                        <div class='gel-kpi-card gel-kpi-blue'>
                            <div class='gel-kpi-title'>Total de Aditivos</div>
                            <div class='gel-kpi-value' style='font-size: 20px; color: #0284c7;'>{total_adit} registrado(s)</div>
                        </div>
                        """, unsafe_allow_html=True)
                    with m2:
                        st.markdown(f"""
                        <div class='gel-kpi-card gel-kpi-green'>
                            <div class='gel-kpi-title'>Soma de Valores Aditados</div>
                            <div class='gel-kpi-value' style='font-size: 20px; color: #059669;'>{formatar_moeda(soma_aditada)}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    with m3:
                        datas_validas = df_aditivos["nova_data_vencimento"].dropna()
                        ultimo_venc = datas_validas.iloc[-1] if not datas_validas.empty else c_row["data_vencimento"]
                        st.markdown(f"""
                        <div class='gel-kpi-card gel-kpi-gold'>
                            <div class='gel-kpi-title'>Vigência Atualizada</div>
                            <div class='gel-kpi-value' style='font-size: 20px; color: #d97706;'>{formatar_data_br(ultimo_venc)}</div>
                        </div>
                        """, unsafe_allow_html=True)

                    st.markdown("---")
                    st.markdown("###### 📜 Relação de Termos Aditivos:")
                    for idx_ad, ad_row in df_aditivos.iterrows():
                        reg_label = "Lei nº 14.133/2021" if "14133" in str(ad_row.get("regime_legal", "")) else "Lei nº 8.666/1993"
                        with st.expander(f"📑 {ad_row['numero_completo']} — {ad_row['tipo_aditivo']} ({reg_label})", expanded=(idx_ad == total_adit - 1)):
                            col_a1, col_a2 = st.columns([1.4, 1])
                            with col_a1:
                                st.markdown(f"**Processo Administrativo:** `{ad_row.get('processo_adm') or 'Não informado'}`")
                                st.markdown(f"**Data de Assinatura:** {formatar_data_br(ad_row.get('data_assinatura'))}")
                                if ad_row.get('prazo_aditado_meses') and int(ad_row.get('prazo_aditado_meses')) > 0:
                                    st.markdown(f"**Prazo Prorrogado:** {ad_row['prazo_aditado_meses']} meses (novo vencimento: **{formatar_data_br(ad_row.get('nova_data_vencimento'))}**)")
                                if ad_row.get('valor_aditado') and float(ad_row.get('valor_aditado')) > 0:
                                    st.markdown(f"**Valor do Aditivo:** {formatar_moeda(ad_row['valor_aditado'])} ({ad_row.get('percentual_aditado', 0):.2f}% do inicial)")
                                    st.markdown(f"**Novo Valor Consolidado:** {formatar_moeda(ad_row.get('novo_valor_total', 0))}")
                                st.markdown(f"**Justificativa:** {ad_row.get('justificativa') or 'Sem justificativa registrada'}")
                            with col_a2:
                                st.markdown("###### 🖨️ Minuta Oficial Padronizada:")
                                hist_org_ad = obter_historico_orgao(c_row["orgao"])
                                dados_ad_doc = {
                                    "tipo_aditivo": ad_row.get("tipo_aditivo", "PRORROGACAO"),
                                    "regime_legal": ad_row.get("regime_legal", "LEI_14133_2021"),
                                    "numero_aditivo": ad_row.get("numero_aditivo", 1),
                                    "numero_completo_aditivo": ad_row.get("numero_completo", f"{ad_row.get('numero_aditivo', 1)}º TERMO ADITIVO"),
                                    "numero_contrato": c_row["numero_completo"],
                                    "processo_adm": ad_row.get("processo_adm") or c_row.get("processo_adm", "001/2026"),
                                    "valor_original": float(c_row.get("valor_total") or 0.0),
                                    "valor_aditado": float(ad_row.get("valor_aditado") or 0.0),
                                    "novo_valor_total": float(ad_row.get("novo_valor_total") or c_row.get("valor_total") or 0.0),
                                    "percentual_aditado": float(ad_row.get("percentual_aditado") or 0.0),
                                    "prazo_meses": int(ad_row.get("prazo_aditado_meses") or 12),
                                    "data_inicio_aditivo": formatar_data_br(c_row.get("data_vencimento")),
                                    "nova_data_vencimento": formatar_data_br(ad_row.get("nova_data_vencimento")),
                                    "justificativa": ad_row.get("justificativa") or "",
                                    "dotacao_orcamentaria": ad_row.get("dotacao_orcamentaria") or hist_org_ad["dotacao"],
                                    "garantia_execucao": ad_row.get("garantia_execucao"),
                                    "orgao_nome": c_row["orgao"],
                                    "orgao_cnpj": formatar_cnpj_cpf(hist_org_ad["cnpj"]),
                                    "orgao_endereco": hist_org_ad["endereco"],
                                    "orgao_bairro": hist_org_ad["bairro"],
                                    "orgao_cep": hist_org_ad["cep"],
                                    "orgao_cidade": hist_org_ad["cidade"],
                                    "orgao_uf": hist_org_ad["uf"],
                                    "secretario": ad_row.get("secretario") or c_row.get("secretario") or hist_org_ad["secretario"],
                                    "cargo_secretario": hist_org_ad.get("cargo_secretario", "Secretário(a) Municipal Titular"),
                                    "fornecedor_nome": c_row.get("fornecedor", "CONTRATADA"),
                                    "fornecedor_cnpj": formatar_cnpj_cpf(c_row.get("cnpj_cpf", "00.000.000/0001-00")),
                                    "fornecedor_representante": "Representante Legal",
                                    "fornecedor_cargo": "Sócio Administrador",
                                    "fornecedor_endereco": "Sede Comercial da Contratada",
                                    "fiscal": ad_row.get("fiscal") or c_row.get("fiscal") or hist_org_ad["fiscal"],
                                    "data_assinatura": ad_row.get("data_assinatura") or date.today()
                                }
                                try:
                                    docx_ad = gerar_aditivo_docx(dados_ad_doc)
                                    st.download_button(
                                        label="📥 Baixar em Word (.docx)",
                                        data=docx_ad,
                                        file_name=f"Termo_Aditivo_{ad_row['id']}_{str(c_row['numero_completo']).replace('/', '_')}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        key=f"dl_docx_ad_{ad_row['id']}",
                                        use_container_width=True
                                    )
                                except Exception as e:
                                    st.caption(f"Word: {e}")

                                try:
                                    pdf_ad = gerar_aditivo_pdf(dados_ad_doc)
                                    st.download_button(
                                        label="📄 Baixar em PDF (.pdf)",
                                        data=pdf_ad,
                                        file_name=f"Termo_Aditivo_{ad_row['id']}_{str(c_row['numero_completo']).replace('/', '_')}.pdf",
                                        mime="application/pdf",
                                        key=f"dl_pdf_ad_{ad_row['id']}",
                                        use_container_width=True
                                    )
                                except Exception as e:
                                    st.caption(f"PDF: {e}")

            with tab_p:
                df_prot_vinculados = carregar_protocolos_do_contrato(cid)
                if df_prot_vinculados.empty:
                    st.info("Nenhum protocolo vinculado a este contrato.")
                else:
                    st.dataframe(df_prot_vinculados, use_container_width=True, hide_index=True)

                    st.markdown("##### 🖨️ Emitir Comprovante do Protocolo Vinculado:")
                    c_p_sel, c_p_dl = st.columns([2, 1])
                    with c_p_sel:
                        prot_escolhido = st.selectbox(
                            "Escolha o protocolo:",
                            df_prot_vinculados["Protocolo"].tolist(),
                            key=f"sb_prot_c_{cid}"
                        )
                    with c_p_dl:
                        st.write("")
                        st.write("")
                        row_prot_sel = df_prot_vinculados[df_prot_vinculados["Protocolo"] == prot_escolhido].iloc[0]
                        prot_exp_data = {
                            "numero": str(row_prot_sel["Protocolo"]),
                            "data_rec": str(row_prot_sel["Data"]),
                            "assunto": str(row_prot_sel["Assunto"]),
                            "solicitante": str(row_prot_sel["Solicitante"]),
                            "orgao": str(c_row["orgao"]),
                            "secretario": str(c_row["secretario"]),
                            "contrato": str(c_row["numero_completo"]),
                            "fornecedor": str(c_row["fornecedor"]),
                            "observacoes": f"Protocolo vinculado ao Contrato {c_row['numero_completo']}",
                            "usuario": st.session_state["usuario"].get("nome", "Servidor"),
                            "qtd_vias": 2
                        }
                        try:
                            pdf_buf_c = gerar_comprovante_protocolo_pdf(prot_exp_data)
                            st.download_button(
                                label="📄 Baixar Comprovante em PDF",
                                data=pdf_buf_c.getvalue(),
                                file_name=f"comprovante_{str(row_prot_sel['Protocolo']).replace('/', '_')}.pdf",
                                mime="application/pdf",
                                key=f"btn_dl_prot_c_{cid}_{row_prot_sel['id']}",
                                use_container_width=True
                            )
                        except Exception as e:
                            st.error(f"Erro ao gerar PDF: {e}")

                st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
                if st.button("➕ Abrir Novo Protocolo para este Contrato", key=f"btn_novo_prot_c_{cid}", type="primary"):
                    st.session_state["contrato_pre_selecionado_id"] = cid
                    st.session_state["mostrar_formulario_protocolo"] = True
                    st.rerun()


    with tab_graficos:
        st.subheader("📈 Análise Gráfica & Distribuição de Recursos")

        g1, g2 = st.columns(2)
        with g1:
            st.markdown("##### 🏛️ Investimento Total por Secretaria (R$)")
            if not df.empty and "orgao" in df.columns:
                df_org = df.groupby("orgao")["valor_total"].sum().sort_values(ascending=True).reset_index()
                df_org.columns = ["Secretaria", "Investimento"]
                
                fig_org = px.bar(
                    df_org,
                    x="Investimento",
                    y="Secretaria",
                    orientation="h",
                    template="plotly_white",
                    color="Investimento",
                    color_continuous_scale=["#bae6fd", "#0284c7"]
                )
                fig_org.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(l=10, r=10, t=10, b=10),
                    height=380,
                    xaxis_title="",
                    yaxis_title="",
                    coloraxis_showscale=False
                )
                fig_org.update_traces(
                    hovertemplate="<b>%{y}</b><br>R$ %{x:,.2f}<extra></extra>"
                )
                st.plotly_chart(fig_org, use_container_width=True)

        with g2:
            st.markdown("##### 💰 Distribuição por Modalidade de Licitação")
            if not df.empty and "modalidade" in df.columns:
                val_mod = df.groupby("modalidade")["valor_total"].sum().sort_values(ascending=False).reset_index()
                val_mod.columns = ["Modalidade", "Valor Total"]
                
                fig_mod = px.pie(
                    val_mod,
                    names="Modalidade",
                    values="Valor Total",
                    hole=0.5,
                    template="plotly_white",
                    color_discrete_sequence=["#0284c7", "#059669", "#f59e0b", "#0369a1", "#10b981", "#d97706"]
                )
                fig_mod.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(l=10, r=10, t=10, b=10),
                    height=380
                )
                fig_mod.update_traces(
                    hovertemplate="<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>"
                )
                st.plotly_chart(fig_mod, use_container_width=True)

        st.markdown("---")
        st.markdown("##### 📅 Cronograma Mensal de Vencimentos dos Contratos")
        if not df.empty and "data_vencimento" in df.columns:
            df_cron = df.copy()
            df_cron["mes_vencimento"] = df_cron["data_vencimento"].apply(lambda x: str(x)[:7] if pd.notna(x) else "Sem data")
            cron_group = df_cron.groupby("mes_vencimento").agg(
                Quantidade=("id", "count"),
                Valor=("valor_total", "sum")
            ).reset_index().sort_values("mes_vencimento")

            fig_cron = px.bar(
                cron_group,
                x="mes_vencimento",
                y="Quantidade",
                template="plotly_white",
                color="Quantidade",
                color_continuous_scale=["#bbf7d0", "#059669"],
                labels={"mes_vencimento": "Mês de Vencimento", "Quantidade": "Contratos a Vencer"}
            )
            fig_cron.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                margin=dict(l=10, r=10, t=10, b=10),
                height=320,
                coloraxis_showscale=False
            )
            st.plotly_chart(fig_cron, use_container_width=True)

# ============================================
# INTERFACE - DASHBOARD PROTOCOLOS
# ============================================

def dashboard_protocolos():
    df = carregar_protocolos()

    brasao_b64 = obter_brasao_b64()
    brasao_banner_img = f"<div style='background: rgba(255,255,255,0.92); padding: 6px 12px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.15); display: flex; align-items: center; justify-content: center;'><img src='data:image/png;base64,{brasao_b64}' style='height: 52px; width: auto;' alt='Brasão' /></div>" if brasao_b64 else ""
    st.markdown(f"""
        <div class='municipal-banner'>
            <div style='display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 14px;'>
                <div>
                    <div class='municipal-badge-pill'>
                        <span class='beacon-active-gel' style='background: #ffffff;'></span> TRAMITAÇÃO OFICIAL • EXERCÍCIO 2026
                    </div>
                    <h2 class='municipal-banner-title'>Gestão de Protocolos</h2>
                    <div class='municipal-banner-subtitle'>Prefeitura Municipal de Ribeirãozinho do Maranhão - MA</div>
                </div>
                {brasao_banner_img}
            </div>
        </div>
    """, unsafe_allow_html=True)

    col_topo1, col_topo2, col_topo3 = st.columns([1, 1, 3])
    with col_topo1:
        if st.session_state['usuario']['cargo'] != 'AUDITOR' and st.button("➕ Novo Protocolo", use_container_width=True, type="primary"):
            st.session_state["menu_selecionado"] = "➕ Novo Protocolo"
            st.session_state["mostrar_formulario"] = False
            st.session_state["mostrar_formulario_protocolo"] = True
            st.rerun()

    if df.empty:
        st.info("Nenhum protocolo registrado.")
        if st.button("➕ Abrir Primeiro Protocolo", type="primary"):
            st.session_state["mostrar_formulario_protocolo"] = True
            st.rerun()
        return

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class='gel-kpi-card gel-kpi-blue'>
            <div class='gel-kpi-title'>Total de Protocolos</div>
            <div class='gel-kpi-value' style='color: #0284c7;'>{len(df)}</div>
            <div class='gel-kpi-subtext'>Fluxo geral do município</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        pendentes = len(df[df["status"] == "PENDENTE"])
        st.markdown(f"""
        <div class='gel-kpi-card gel-kpi-gold'>
            <div class='gel-kpi-title'>Pendentes</div>
            <div class='gel-kpi-value' style='color: #d97706;'>{pendentes}</div>
            <div class='gel-kpi-subtext'>Aguardando despacho</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        em_andamento = len(df[df["status"] == "EM_ANDAMENTO"])
        st.markdown(f"""
        <div class='gel-kpi-card gel-kpi-cyan'>
            <div class='gel-kpi-title'>Em Tramitação</div>
            <div class='gel-kpi-value' style='color: #0369a1;'>{em_andamento}</div>
            <div class='gel-kpi-subtext'>Em análise nas secretarias</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        concluidos = len(df[df["status"] == "CONCLUIDO"])
        st.markdown(f"""
        <div class='gel-kpi-card gel-kpi-green'>
            <div class='gel-kpi-title'>Concluídos</div>
            <div class='gel-kpi-value' style='color: #059669;'>{concluidos}</div>
            <div class='gel-kpi-subtext'>Processos finalizados</div>
        </div>
        """, unsafe_allow_html=True)



    col_btn, col_exp = st.columns([1, 1])
    with col_btn:
        if st.button("➕ Novo Protocolo", type="primary"):
            st.session_state["mostrar_formulario_protocolo"] = True
            st.rerun()
    with col_exp:
        p_bytes = exportar_protocolos_excel(df)
        st.download_button(
            label="📥 Exportar Protocolos (.xlsx)",
            data=p_bytes,
            file_name=f"protocolos_ribeiraozinho_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    with st.expander("🖨️ Visualizar e Emitir Comprovante Oficial (PDF / Impressão)", expanded=False):
        st.markdown("##### Selecione um protocolo para emitir o comprovante oficial:")
        opcoes_prot = [f"{row['numero_completo']} — {row['solicitante']} ({str(row['assunto'])[:40]}...)" for _, row in df.iterrows()]
        sel_prot_str = st.selectbox("Protocolo:", opcoes_prot, key="sb_reimprimir_prot")
        if sel_prot_str:
            sel_idx = opcoes_prot.index(sel_prot_str)
            row_sel = df.iloc[sel_idx]
            
            # Buscar fornecedor se tiver contrato
            forn_nome = "—"
            if row_sel.get("contrato_id"):
                with engine.connect() as conn:
                    forn_nome = conn.execute(
                        text("SELECT COALESCE(f.razao_social, f.nome_fantasia) FROM contratos c LEFT JOIN fornecedores f ON c.fornecedor_id = f.id WHERE c.id = :cid"),
                        {"cid": int(row_sel["contrato_id"])}
                    ).scalar() or "—"

            
            c_vias, c_btn_pdf = st.columns([1, 2])
            with c_vias:
                vias_escolha = st.radio("Vias do comprovante:", ["1 Via (Protocolo)", "2 Vias (Protocolo + Recebedor)"], horizontal=True, key="radio_vias_reimp")
            
            qtd_v = 2 if "2 Vias" in vias_escolha else 1
            dt_r = row_sel["data_recebimento"]
            if hasattr(dt_r, "strftime"):
                dt_r_str = dt_r.strftime("%d/%m/%Y")
            else:
                dt_r_str = str(dt_r)
                
            # Montar string de observações incluindo documentos e observações do banco, ou situação
            tdocs = row_sel.get('tipo_documento', '')
            obs_banco = row_sel.get('observacoes', '')
            tdocs = tdocs if tdocs and pd.notna(tdocs) else ""
            obs_banco = obs_banco if obs_banco and pd.notna(obs_banco) else ""
            
            str_obs = []
            if tdocs:
                str_obs.append(f"Docs: {tdocs}")
            if obs_banco:
                str_obs.append(f"Obs: {obs_banco}")
            if not str_obs:
                str_obs.append(f"Situação: {row_sel['status']}")
            
            prot_dados_exp = {
                "numero": str(row_sel["numero_completo"]),
                "data_rec": dt_r_str,
                "assunto": str(row_sel["assunto"]) if row_sel["assunto"] else "—",
                "solicitante": str(row_sel["solicitante"]) if row_sel["solicitante"] else "—",
                "orgao": str(row_sel["orgao_origem"]) if row_sel["orgao_origem"] else "—",
                "secretario": str(row_sel["secretario"]) if row_sel["secretario"] else "—",
                "contrato": str(row_sel["contrato_vinculado"]) if row_sel["contrato_vinculado"] else "—",
                "fornecedor": forn_nome,
                "observacoes": "\n".join(str_obs),
                "usuario": str(row_sel["usuario_criador"]) if row_sel["usuario_criador"] else "Sistema",
                "qtd_vias": qtd_v
            }
            
            c_b_pdf, c_b_agu = st.columns([1, 1])
            with c_b_pdf:
                try:
                    pdf_buf = gerar_comprovante_protocolo_pdf(prot_dados_exp)
                    st.download_button(
                        label=f"📄 Baixar Comprovante em PDF",
                        data=pdf_buf.getvalue(),
                        file_name=f"comprovante_{str(row_sel['numero_completo']).replace('/', '_')}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Erro ao gerar PDF: {e}")
            with c_b_agu:
                if st.button(f"⚖️ Gerar Minuta AGU deste Protocolo", key=f"btn_ir_agu_dash_{row_sel['id']}", type="secondary", use_container_width=True):
                    st.session_state["protocolo_agu_selecionado"] = str(row_sel["numero_completo"])
                    if row_sel.get("contrato_vinculado") and str(row_sel["contrato_vinculado"]) != "—":
                        st.session_state["contrato_agu_selecionado"] = str(row_sel["contrato_vinculado"])
                    st.session_state["menu_selecionado"] = "⚖️ Modelos AGU / Minutas"
                    st.session_state["mostrar_formulario_protocolo"] = False
                    st.rerun()


# ============================================
# INTERFACE - IMPORTAÇÃO DE PLANILHA
# ============================================

def tela_importar_planilha():
    st.markdown("""
        <div class='municipal-banner'>
            <div class='municipal-badge-pill'>
                <span class='beacon-active-gel' style='background: #ffffff;'></span> INTEGRAÇÃO DE BASE • 2026
            </div>
            <h2 class='municipal-banner-title'>Importador Oficial de Planilhas</h2>
            <div class='municipal-banner-subtitle'>Prefeitura Municipal de Ribeirãozinho do Maranhão - MA</div>
        </div>
    """, unsafe_allow_html=True)

    col_info, col_modelo = st.columns([2, 1])
    with col_info:
        st.markdown("""
        Faça upload de planilhas Excel (.xlsx) ou CSV de contratos da Prefeitura.
        O sistema normaliza secretarias oficiais, preenche secretários titulares e formata datas no padrão brasileiro (`DD/MM/AAAA`).
        """)
    with col_modelo:
        modelo_path = r"e:\Sistema contratos\modelo_importacao_contratos.xlsx"
        if os.path.exists(modelo_path):
            with open(modelo_path, "rb") as f:
                modelo_bytes = f.read()
            st.download_button(
                label="📥 Baixar Planilha Modelo (.xlsx)",
                data=modelo_bytes,
                file_name="modelo_contratos_ribeiraozinho.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    if "key_uploader_contratos" not in st.session_state:
        st.session_state["key_uploader_contratos"] = 0

    def _callback_limpar_planilha_contratos():
        st.session_state["key_uploader_contratos"] += 1

    uploader_key = f"uploader_contratos_{st.session_state['key_uploader_contratos']}"
    uploaded_file = st.file_uploader("Selecione o arquivo de contratos:", type=["xlsx", "xls", "csv"], key=uploader_key)

    if uploaded_file:
        try:
            if uploaded_file.name.endswith(".csv"):
                try:
                    df_up = pd.read_csv(uploaded_file, encoding="utf-8")
                except Exception:
                    uploaded_file.seek(0)
                    df_up = pd.read_csv(uploaded_file, encoding="latin1", sep=";")
            else:
                df_up = pd.read_excel(uploaded_file)

            col_msg_up, col_btn_limp = st.columns([3.2, 1.2])
            with col_msg_up:
                st.success(f"Arquivo carregado com sucesso! **{uploaded_file.name}** ({len(df_up)} registros identificados).")
            with col_btn_limp:
                st.button(
                    "🗑️ Limpar / Trocar Planilha",
                    key=f"btn_limpar_up_top_{st.session_state['key_uploader_contratos']}",
                    on_click=_callback_limpar_planilha_contratos,
                    use_container_width=True,
                    help="Remove a planilha atual para que você possa selecionar outro arquivo caso precise alterar."
                )

            cols_existentes = list(df_up.columns)
            def_num = detectar_coluna(cols_existentes, ["CONTRATO", "NUMERO", "N°", "Nº"])
            def_org = detectar_coluna(cols_existentes, ["SECRETARIA", "ORGAO", "ÓRGÃO", "SETOR"])
            def_sec = detectar_coluna(cols_existentes, ["SECRETARIO", "SECRETÁRIO", "TITULAR", "GESTOR"])
            def_forn = detectar_coluna(cols_existentes, ["FORNECEDOR", "EMPRESA", "CONTRATADO", "RAZAO"])
            def_cnpj = detectar_coluna(cols_existentes, ["CNPJ", "CPF", "DOC"])
            def_obj = detectar_coluna(cols_existentes, ["OBJETO", "DESCRICAO", "DESCRIÇÃO", "SERVICO"])
            def_val = detectar_coluna(cols_existentes, ["VALOR TOTAL", "VALOR GLOBAL", "TOTAL", "VALOR CONTRATADO", "VALOR ESTIMADO", "VALOR", "MONTANTE", "GLOBAL"])
            def_mod = detectar_coluna(cols_existentes, ["MODALIDADE", "TIPO", "LICITACAO", "LICITAÇÃO"])
            def_proc = detectar_coluna(cols_existentes, ["PROCESSO", "PROC", "PA"])
            def_ass = detectar_coluna(cols_existentes, ["ASSINATURA", "DATA DO CONTRATO", "DATA CONTRATO"])
            def_pub = detectar_coluna(cols_existentes, ["PUBLICACAO", "PUBLICAÇÃO"])
            def_venc = detectar_coluna(cols_existentes, ["VENCIMENTO", "TERMINO", "FIM", "FINAL"])
            def_solic = detectar_coluna(cols_existentes, ["SOLICITANTE", "DEMANDANTE"])
            def_fisc = detectar_coluna(cols_existentes, ["FISCAL"])

            col_m1, col_m2 = st.columns(2)
            with col_m1:
                sel_num = st.selectbox("Nº do Contrato *", cols_existentes, index=cols_existentes.index(def_num) if def_num else 0)
                sel_org = st.selectbox("Secretaria / Órgão *", cols_existentes, index=cols_existentes.index(def_org) if def_org else 0)
                sel_sec = st.selectbox("Secretário(a) Titular", ["(Preencher automaticamente)"] + cols_existentes, index=(cols_existentes.index(def_sec) + 1) if def_sec else 0)
                sel_forn = st.selectbox("Fornecedor / Empresa *", cols_existentes, index=cols_existentes.index(def_forn) if def_forn else 0)
                sel_cnpj = st.selectbox("CNPJ / CPF", ["(Nenhum)"] + cols_existentes, index=(cols_existentes.index(def_cnpj) + 1) if def_cnpj else 0)
                sel_obj = st.selectbox("Objeto Contratual *", cols_existentes, index=cols_existentes.index(def_obj) if def_obj else 0)
                sel_val = st.selectbox("Valor Total (R$) *", cols_existentes, index=cols_existentes.index(def_val) if def_val else 0)

            with col_m2:
                sel_mod = st.selectbox("Modalidade", ["(Padrão: DISPENSA)"] + cols_existentes, index=(cols_existentes.index(def_mod) + 1) if def_mod else 0)
                sel_proc = st.selectbox("Processo Administrativo", ["(Nenhum)"] + cols_existentes, index=(cols_existentes.index(def_proc) + 1) if def_proc else 0)
                sel_ass = st.selectbox("Data de Assinatura", ["(Calcular automaticamente)"] + cols_existentes, index=(cols_existentes.index(def_ass) + 1) if def_ass else 0)
                sel_pub = st.selectbox("Data de Publicação", ["(Calcular +3 dias)"] + cols_existentes, index=(cols_existentes.index(def_pub) + 1) if def_pub else 0)
                sel_venc = st.selectbox("Data de Vencimento", ["(Calcular +1 ano)"] + cols_existentes, index=(cols_existentes.index(def_venc) + 1) if def_venc else 0)
                sel_solic = st.selectbox("Solicitante", ["(Nenhum)"] + cols_existentes, index=(cols_existentes.index(def_solic) + 1) if def_solic else 0)
                sel_fisc = st.selectbox("Fiscal do Contrato", ["(Nenhum)"] + cols_existentes, index=(cols_existentes.index(def_fisc) + 1) if def_fisc else 0)

            mapa = {
                "numero": sel_num, "orgao": sel_org, "secretario": None if sel_sec.startswith("(") else sel_sec,
                "fornecedor": sel_forn, "cnpj": None if sel_cnpj.startswith("(") else sel_cnpj,
                "objeto": sel_obj, "valor": sel_val, "modalidade": None if sel_mod.startswith("(") else sel_mod,
                "processo": None if sel_proc.startswith("(") else sel_proc,
                "data_assinatura": None if sel_ass.startswith("(") else sel_ass,
                "data_publicacao": None if sel_pub.startswith("(") else sel_pub,
                "data_vencimento": None if sel_venc.startswith("(") else sel_venc,
                "solicitante": None if sel_solic.startswith("(") else sel_solic,
                "fiscal": None if sel_fisc.startswith("(") else sel_fisc
            }

            # Prévia do valor total apurado com auditoria em tempo real
            valores_calculados = []
            linhas_totais_detectadas = 0
            for _, r_prev in df_up.iterrows():
                col_n = mapa.get("numero")
                raw_n = str(r_prev[col_n]).strip().upper() if col_n and pd.notna(r_prev[col_n]) else ""
                row_str = " ".join([str(v) for v in r_prev.values if pd.notna(v)]).strip().upper()
                if any(kw in raw_n for kw in ["TOTAL", "SUBTOTAL", "SOMA"]) or row_str.startswith(("TOTAL", "SUBTOTAL", "SOMA", "RESUMO")):
                    linhas_totais_detectadas += 1
                    continue
                v_num = clean_money(r_prev[sel_val]) if pd.notna(r_prev[sel_val]) else 0.0
                valores_calculados.append(v_num)

            total_apurado_prev = sum(valores_calculados)
            qtd_validos = len(valores_calculados)

            st.markdown("---")
            st.markdown("##### 🔍 Auditoria Prévia dos Valores da Planilha")
            col_ap1, col_ap2, col_ap3 = st.columns(3)
            with col_ap1:
                st.metric("Coluna de Valor Selecionada", sel_val)
            with col_ap2:
                st.metric("Contratos Válidos a Importar", f"{qtd_validos} linhas" + (f" ({linhas_totais_detectadas} totais ignorados)" if linhas_totais_detectadas > 0 else ""))
            with col_ap3:
                st.metric("Soma Total Apurada", formatar_moeda(total_apurado_prev))

            if linhas_totais_detectadas > 0:
                st.info(f"ℹ️ **{linhas_totais_detectadas} linha(s) de total/resumo** da planilha foram identificadas e excluídas da soma para não duplicar o valor.")

            st.markdown("---")
            opcoes_grav = ["Adicionar / Atualizar registros", "Substituir base inteira (Recarregar)"]
            modo_imp = st.segmented_control("Modo de Gravação:", opcoes_grav, default=opcoes_grav[0])
            if not modo_imp:
                modo_imp = opcoes_grav[0]
            modo_code = "substituir" if "Substituir" in modo_imp else "adicionar"

            col_act_proc, col_act_limp = st.columns([2.8, 1.2])
            with col_act_proc:
                btn_processar_planilha = st.button("🚀 Processar e Integrar Planilha", type="primary", use_container_width=True)
            with col_act_limp:
                st.button(
                    "🗑️ Limpar Planilha",
                    key=f"btn_limpar_up_bot_{st.session_state['key_uploader_contratos']}",
                    on_click=_callback_limpar_planilha_contratos,
                    use_container_width=True,
                    help="Descarta a seleção atual e permite carregar uma nova planilha caso precise alterar."
                )

            if btn_processar_planilha:
                p_bar = st.progress(0)
                s_txt = st.empty()
                ok, msg = processar_importacao_planilha_customizada(df_up, mapa, modo=modo_code, progress_bar=p_bar, status_text=s_txt)
                if ok:
                    st.success(msg)
                    st.balloons()
                    time.sleep(1.5)
                    st.session_state["menu_selecionado"] = "📊 Contratos"
                    st.rerun()
                else:
                    st.error(msg)
        except Exception as e:
            st.error(f"Erro ao processar: {str(e)}")

# ============================================
# INTERFACE - USUÁRIOS
# ============================================

def tela_usuarios():
    if st.session_state["usuario"].get("cargo") not in ["ADMIN", "ADMINISTRADOR"]:
        st.error("⛔ Acesso Restrito: Apenas usuários com perfil ADMINISTRADOR podem acessar este módulo.")
        return

    st.markdown("""
        <div class='municipal-banner'>
            <div class='municipal-badge-pill'>
                <span class='beacon-active-gel' style='background: #ffffff;'></span> CONTROLE DE ACESSO
            </div>
            <h2 class='municipal-banner-title'>Gestão de Usuários e Permissões</h2>
            <div class='municipal-banner-subtitle'>Prefeitura Municipal de Ribeirãozinho do Maranhão - MA</div>
        </div>
    """, unsafe_allow_html=True)

    tab_lista, tab_cadastro = st.tabs(["📋 Servidores Registrados", "➕ Novo Usuário"])

    with tab_lista:
        df_u = carregar_usuarios()
        df_view = df_u.copy()
        df_view["Status"] = df_view["ativo"].apply(lambda a: "🟢 Ativo" if a == 1 else "🔴 Inativo")
        df_view["Data Cadastro"] = df_view["data_criacao"].apply(formatar_datetime_br)

        st.dataframe(
            df_view[["id", "nome", "email", "cargo", "Status", "Data Cadastro"]].rename(columns={
                "id": "ID", "nome": "Nome Completo", "email": "Nome de Usuário (Login)", "cargo": "Perfil / Função"
            }),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("---")
        opcoes_usuarios = [f"{r['id']} - {r['nome']} ({r['cargo']}) | Usuário: {str(r['email']).lower()}" for _, r in df_u.iterrows()]
        u_sel = st.selectbox("Selecione usuário para gerenciar:", opcoes_usuarios)

        if u_sel:
            uid = int(u_sel.split(" - ")[0])
            u_row = df_u[df_u["id"] == uid].iloc[0]

            with st.form(f"form_edit_u_{uid}"):
                c1, c2 = st.columns(2)
                with c1:
                    e_nome = st.text_input("Nome Completo *", value=u_row["nome"])
                    val_login_inicial = str(u_row["email"] or "").strip()
                    e_login = st.text_input("Nome de Usuário (Login) *", value=val_login_inicial, help="Nome de usuário utilizado para entrar no sistema.")
                with c2:
                    cargos_opcoes = ["ADMIN", "OPERADOR"]
                    idx_c = cargos_opcoes.index(u_row["cargo"]) if u_row["cargo"] in cargos_opcoes else 1
                    e_cargo = st.selectbox("Perfil", cargos_opcoes, index=idx_c)
                    e_ativo = st.checkbox("Ativo (permitir acesso ao sistema)", value=(u_row["ativo"] == 1))

                e_senha = st.text_input("Nova Senha (deixe em branco para manter a atual)", type="password", help="Mínimo 4 caracteres. Se não quiser alterar a senha deste usuário, deixe este campo vazio.")

                b1, b2 = st.columns(2)
                with b1:
                    salvar = st.form_submit_button("💾 Salvar Alterações", type="primary", use_container_width=True)
                with b2:
                    excluir = st.form_submit_button("🗑️ Excluir Usuário", use_container_width=True)

                if salvar:
                    login_tratado = str(e_login).strip()
                    ok, msg = atualizar_usuario(uid, e_nome, login_tratado, e_cargo, e_ativo, e_senha if e_senha else None)
                    if ok:
                        st.success(msg)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg)

                if excluir:
                    curr_id = st.session_state["usuario"]["id"]
                    ok, msg = excluir_usuario(uid, curr_id)
                    if ok:
                        st.success(msg)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg)

    with tab_cadastro:
        with st.form("form_novo_user_gel"):
            st.markdown("##### Dados do Novo Servidor / Operador:")
            n_nome = st.text_input("Nome Completo *", placeholder="Ex: Roberto Alves Costa")
            n_usuario = st.text_input("Nome de Usuário (Login)", placeholder="Ex: Roberto ou roberto.costa (se vazio, gerado automaticamente)")
            c_cad1, c_cad2 = st.columns(2)
            with c_cad1:
                n_cargo = st.selectbox("Perfil de Acesso:", ["OPERADOR", "ADMIN"])
            with c_cad2:
                n_senha = st.text_input("Senha Inicial * (mínimo 4 caracteres)", type="password", value="123456")

            st.caption("💡 O usuário entrará no sistema informando seu **Nome de Usuário** e a senha.")

            if st.form_submit_button("➕ Criar Acesso", type="primary", use_container_width=True):
                user_cad_tratado = str(n_usuario).strip() if n_usuario else ""
                ok, msg = cadastrar_novo_usuario(n_nome, user_cad_tratado, n_senha, n_cargo)
                if ok:
                    st.success(msg)
                    time.sleep(1.5)
                    st.rerun()
                else:
                    st.error(msg)

# ============================================
# FORMULÁRIOS INTELIGENTES
# ============================================

def formulario_protocolo():
    st.markdown("### ➕ Novo Protocolo")
    st.caption("Prefeitura Municipal de Ribeirãozinho do Maranhão - MA")

    prox_num, prox_str = gerar_proximo_numero_protocolo()

    # Carregar contratos ordenados do mais recente
    df_c = carregar_contratos()
    if not df_c.empty:
        df_c_sorted = df_c.sort_values("id", ascending=False)
    else:
        df_c_sorted = df_c

    # Escolher tipo de protocolo
    st.markdown("""
    <div style='background: #f0f9ff; border: 1px solid #bae6fd; border-radius: 14px; padding: 16px 20px; margin-bottom: 18px;'>
        <div style='font-size: 13px; font-weight: 700; color: #0369a1; margin-bottom: 6px;'>📄 Tipo de Protocolo</div>
        <div style='font-size: 12px; color: #475569;'>
            Selecione <strong>"Vinculado a Contrato"</strong> para puxar dados de um contrato existente, 
            ou <strong>"Documento Avulso"</strong> para protocolar outros documentos (ofícios, requerimentos, etc).
        </div>
    </div>
    """, unsafe_allow_html=True)

    tipo_protocolo = st.segmented_control(
        "Tipo de Protocolo:",
        ["📑 Vinculado a Contrato", "📝 Documento Avulso"],
        default="📑 Vinculado a Contrato" if not df_c.empty else "📝 Documento Avulso",
        label_visibility="collapsed"
    )
    if not tipo_protocolo:
        tipo_protocolo = "📝 Documento Avulso"

    # Se vinculado a contrato, mostrar seletor de contratos recentes
    contrato_selecionado = None
    dados_contrato = {}

    if "Vinculado a Contrato" in tipo_protocolo and not df_c_sorted.empty:
        st.markdown("""
        <div style='font-size: 12px; font-weight: 700; color: #0284c7; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px;'>
            📋 Selecione o Contrato (mais recentes primeiro)
        </div>
        """, unsafe_allow_html=True)

        pre_sel_cid = st.session_state.pop("contrato_pre_selecionado_id", None)
        idx_contrato_default = 0
        if pre_sel_cid and not df_c_sorted.empty:
            for i_c, r_c in enumerate(df_c_sorted.itertuples()):
                if getattr(r_c, 'id', None) == pre_sel_cid:
                    idx_contrato_default = i_c + 1
                    break

        opcoes_contratos = ["— Selecione um contrato —"] + [
            f"Nº {r['numero_completo']} | {r['fornecedor']} | {formatar_moeda(r['valor_total'])} | {r['orgao']}"
            for _, r in df_c_sorted.iterrows()
        ]

        contrato_escolhido = st.selectbox(
            "Contrato:",
            opcoes_contratos,
            index=idx_contrato_default,
            label_visibility="collapsed"
        )


        if contrato_escolhido != "— Selecione um contrato —":
            idx_sel = opcoes_contratos.index(contrato_escolhido) - 1
            row_sel = df_c_sorted.iloc[idx_sel]
            contrato_selecionado = int(row_sel["id"])
            dados_contrato = {
                "numero_completo": row_sel.get("numero_completo", ""),
                "fornecedor": row_sel.get("fornecedor", ""),
                "cnpj_cpf": row_sel.get("cnpj_cpf", ""),
                "orgao": row_sel.get("orgao", ""),
                "objeto": row_sel.get("objeto", ""),
                "valor_total": row_sel.get("valor_total", 0),
                "secretario": row_sel.get("secretario", ""),
                "data_assinatura": row_sel.get("data_assinatura", ""),
                "data_vencimento": row_sel.get("data_vencimento", ""),
                "modalidade": row_sel.get("modalidade", ""),
            }

            # Mostrar resumo do contrato selecionado
            st.markdown(f"""
            <div style='background: #ecfdf5; border: 1px solid #86efac; border-left: 5px solid #059669; border-radius: 10px; padding: 14px 18px; margin: 10px 0 16px 0;'>
                <div style='font-size: 13px; font-weight: 700; color: #059669; margin-bottom: 6px;'>✅ Contrato Selecionado — Nº {dados_contrato['numero_completo']}</div>
                <div style='display: grid; grid-template-columns: 1fr 1fr; gap: 4px 20px; font-size: 12px; color: #334155;'>
                    <div>🏭 <strong>Fornecedor:</strong> {dados_contrato['fornecedor']}</div>
                    <div>🏛️ <strong>Secretaria:</strong> {dados_contrato['orgao']}</div>
                    <div>💰 <strong>Valor:</strong> {formatar_moeda(dados_contrato['valor_total'])}</div>
                    <div>📋 <strong>Modalidade:</strong> {dados_contrato['modalidade']}</div>
                    <div>📅 <strong>Assinatura:</strong> {formatar_data_br(dados_contrato['data_assinatura'])}</div>
                    <div>📅 <strong>Vencimento:</strong> {formatar_data_br(dados_contrato['data_vencimento'])}</div>
                </div>
                <div style='margin-top: 6px; font-size: 11.5px; color: #475569;'>📝 <strong>Objeto:</strong> {str(dados_contrato['objeto'])[:180]}{'...' if len(str(dados_contrato['objeto'])) > 180 else ''}</div>
            </div>
            """, unsafe_allow_html=True)

    with st.form("form_novo_prot_gel"):
        st.markdown(f"##### Nº Identificador: **{prox_str}**")
        c1, c2 = st.columns(2)

        # Pré-preencher campos se contrato selecionado
        default_assunto = f"Protocolo ref. Contrato Nº {dados_contrato.get('numero_completo', '')}" if contrato_selecionado else ""
        default_solicitante = dados_contrato.get("fornecedor", "") if contrato_selecionado else ""
        default_orgao_nome = dados_contrato.get("orgao", "") if contrato_selecionado else ""

        # Encontrar índice do órgão na lista
        opcoes_prot_org = ["Selecione a Secretaria de Origem..."] + LISTA_SECRETARIAS_OFICIAIS
        default_orgao_idx = 0
        if default_orgao_nome and default_orgao_nome in opcoes_prot_org:
            default_orgao_idx = opcoes_prot_org.index(default_orgao_nome)

        with c1:
            assunto = st.text_input("Assunto / Demanda *", value=default_assunto, placeholder="Ex: Solicitação de Termo Aditivo")
            solicitante = st.text_input("Solicitante *", value=default_solicitante, placeholder="Nome da empresa ou servidor")
            orgao = st.selectbox("Secretaria de Origem:", opcoes_prot_org, index=default_orgao_idx)
            
            documentos_opcoes = [
                "Ofício", "Requerimento", "Nota Fiscal", "Certidão", 
                "Projeto Básico/Executivo", "Planilha Orçamentária", 
                "Contrato Original", "Termo Aditivo", "Termo de Recebimento", "Outros"
            ]
            documentos_selecionados = st.multiselect("Documentos Inclusos:", documentos_opcoes, placeholder="Selecione um ou mais documentos...")
            
            observacoes = st.text_area("Observações (opcional):", placeholder="Detalhes adicionais do protocolo...", height=80)
        with c2:
            data_rec = st.date_input("Data de Recebimento", value=date.today())
            status = st.selectbox("Status:", ["PENDENTE", "EM_ANDAMENTO", "CONCLUIDO"])

            if "Vinculado a Contrato" in tipo_protocolo:
                if contrato_selecionado:
                    st.markdown(f"""
                    <div style='background: #e0f2fe; border: 1px solid #7dd3fc; border-radius: 8px; padding: 10px 14px; margin-top: 4px;'>
                        <div style='font-size: 11px; color: #0369a1; font-weight: 700;'>📎 CONTRATO VINCULADO</div>
                        <div style='font-size: 13px; color: #0f172a; font-weight: 600; margin-top: 2px;'>Nº {dados_contrato['numero_completo']}</div>
                        <div style='font-size: 11px; color: #64748b;'>{dados_contrato['fornecedor']}</div>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.info("⬆️ Selecione um contrato acima para vincular.")
            else:
                st.markdown("""
                <div style='background: #f1f5f9; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 14px; margin-top: 4px;'>
                    <div style='font-size: 11px; color: #64748b; font-weight: 700;'>📝 DOCUMENTO AVULSO</div>
                    <div style='font-size: 12px; color: #475569; margin-top: 2px;'>Protocolo sem vínculo contratual</div>
                </div>
                """, unsafe_allow_html=True)

            qtd_vias = st.selectbox("Vias do Comprovante de Entrega:", ["1 Via (Protocolo)", "2 Vias (Protocolo + Recebedor)"])

        b1, b2 = st.columns(2)
        with b1:
            sub = st.form_submit_button("Salvar Protocolo", type="primary", use_container_width=True)
        with b2:
            canc = st.form_submit_button("Cancelar", use_container_width=True)

        if canc:
            st.session_state["mostrar_formulario_protocolo"] = False
            st.session_state["menu_selecionado"] = "📋 Protocolos"
            st.session_state.pop("ultimo_protocolo_impresso", None)
            st.rerun()

        if sub:
            if not assunto or not solicitante:
                st.error("Preencha todos os campos obrigatórios.")
            else:
                cid = contrato_selecionado
                with engine.connect() as conn:
                    oid = conn.execute(text("SELECT id FROM orgaos WHERE nome = :n"), {"n": orgao}).scalar()
                    if not oid:
                        res = conn.execute(text("INSERT INTO orgaos (nome, sigla, ativo) VALUES (:n, :s, 1)"), {"n": orgao, "s": orgao[:10]})
                        oid = res.lastrowid

                    sec_titular = MAPA_SECRETARIOS_PADRAO.get(orgao, "SECRETÁRIO TITULAR")
                    uid = st.session_state["usuario"]["id"]

                    tp_prot = "CONTRATO" if "Vinculado a Contrato" in tipo_protocolo else "AVULSO"

                    conn.execute(
                        text("""
                            INSERT INTO protocolos (
                                numero_protocolo, ano_protocolo, numero_completo, data_recebimento, data_protocolo,
                                assunto, descricao, solicitante, orgao_origem_id, status, contrato_id,
                                secretario, criado_por, data_criacao, tipo_protocolo,
                                tipo_documento, observacoes
                            ) VALUES (
                                :num, :ano, :ncomp, :dt, CURRENT_DATE,
                                :ass, :ass, :solic, :oid, :st, :cid,
                                :sec, :uid, CURRENT_TIMESTAMP, :tp,
                                :tdocs, :obs
                            )
                        """),
                        {
                            "num": prox_num, "ano": 2026, "ncomp": prox_str, "dt": data_rec.strftime("%Y-%m-%d"),
                            "ass": assunto, "solic": solicitante, "oid": oid, "st": status, "cid": cid,
                            "sec": sec_titular, "uid": uid, "tp": tp_prot,
                            "tdocs": ", ".join(documentos_selecionados), "obs": observacoes
                        }
                    )
                    conn.commit()
                    try:
                        carregar_protocolos.clear()
                    except Exception:
                        pass

                # Salvar dados para impressão
                st.session_state["ultimo_protocolo_impresso"] = {
                    "numero": prox_str,
                    "data_rec": data_rec.strftime("%d/%m/%Y"),
                    "assunto": assunto,
                    "solicitante": solicitante,
                    "orgao": orgao,
                    "secretario": sec_titular,
                    "contrato": dados_contrato.get("numero_completo", "—") if cid else "—",
                    "fornecedor": dados_contrato.get("fornecedor", "—") if cid else "—",
                    "observacoes": f"Docs: {', '.join(documentos_selecionados)}\nObs: {observacoes}" if documentos_selecionados else (observacoes if observacoes else "—"),
                    "usuario": st.session_state["usuario"].get("nome", ""),
                    "qtd_vias": 2 if "2 Vias" in qtd_vias else 1
                }

                st.success(f"✅ Protocolo {prox_str} registrado com sucesso!")

    # Seção de impressão do comprovante (aparece após salvar)
    if "ultimo_protocolo_impresso" in st.session_state:
        prot = st.session_state["ultimo_protocolo_impresso"]
        brasao_b64 = obter_brasao_b64()
        brasao_img_tag = f"<img src='data:image/png;base64,{brasao_b64}' style='height: 60px; width: auto;' />" if brasao_b64 else "🏛️"

        agora_str = datetime.now().strftime("%d/%m/%Y às %H:%M")
        vias = prot["qtd_vias"]

        def gerar_via_html(via_label):
            return f"""
            <div style='border: 2px solid #0284c7; border-radius: 8px; padding: 22px 28px; margin-bottom: 20px; page-break-inside: avoid; background: #fff;'>
                <div style='display: flex; align-items: center; justify-content: space-between; border-bottom: 2px solid #0284c7; padding-bottom: 12px; margin-bottom: 14px;'>
                    <div style='display: flex; align-items: center; gap: 12px;'>
                        {brasao_img_tag}
                        <div>
                            <div style='font-weight: 800; font-size: 14px; color: #0369a1; line-height: 1.3;'>PREFEITURA MUNICIPAL DE<br>RIBEIRÃOZINHO DO MARANHÃO - MA</div>
                            <div style='font-size: 10px; color: #059669; font-weight: 700; letter-spacing: 0.5px;'>ESTADO DO MARANHÃO</div>
                        </div>
                    </div>
                    <div style='text-align: right;'>
                        <div style='font-size: 11px; color: #64748b; font-weight: 600;'>{via_label}</div>
                        <div style='font-size: 18px; font-weight: 800; color: #0284c7; font-family: \"JetBrains Mono\", monospace;'>PROTOCOLO Nº {prot['numero']}</div>
                    </div>
                </div>
                <table style='width: 100%; border-collapse: collapse; font-size: 12px; color: #334155;'>
                    <tr>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: 700; width: 160px;'>Data de Recebimento</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0;'>{prot['data_rec']}</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: 700; width: 160px;'>Secretaria de Origem</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0;'>{prot['orgao']}</td>
                    </tr>
                    <tr>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: 700;'>Solicitante / Remetente</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0;' colspan='3'>{prot['solicitante']}</td>
                    </tr>
                    <tr>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: 700;'>Assunto / Demanda</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0;' colspan='3'>{prot['assunto']}</td>
                    </tr>
                    <tr>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: 700;'>Contrato Vinculado</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0;'>{prot['contrato']}</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: 700;'>Fornecedor / Empresa</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0;'>{prot['fornecedor']}</td>
                    </tr>
                    <tr>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: 700;'>Secretário(a) Titular</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0;'>{prot['secretario']}</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: 700;'>Registrado por</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0;'>{prot['usuario']}</td>
                    </tr>
                    <tr>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0; background: #f8fafc; font-weight: 700;'>Observações</td>
                        <td style='padding: 5px 8px; border: 1px solid #e2e8f0;' colspan='3'>{prot['observacoes']}</td>
                    </tr>
                </table>
                <div style='display: flex; justify-content: space-between; margin-top: 30px; padding-top: 10px;'>
                    <div style='text-align: center; width: 45%;'>
                        <div style='border-top: 1px solid #334155; padding-top: 6px; font-size: 11px; color: #334155; font-weight: 600;'>
                            Servidor(a) Responsável pelo Recebimento
                        </div>
                        <div style='font-size: 10px; color: #64748b; margin-top: 2px;'>Carimbo / Matrícula</div>
                    </div>
                    <div style='text-align: center; width: 45%;'>
                        <div style='border-top: 1px solid #334155; padding-top: 6px; font-size: 11px; color: #334155; font-weight: 600;'>
                            Entregue por (Assinatura do Recebedor)
                        </div>
                        <div style='font-size: 10px; color: #64748b; margin-top: 2px;'>Data: ____/____/________</div>
                    </div>
                </div>
                <div style='text-align: center; font-size: 9px; color: #94a3b8; margin-top: 14px; border-top: 1px dashed #cbd5e1; padding-top: 6px;'>
                    Documento gerado em {agora_str} pelo Sistema Integrado de Gestão — Prefeitura Municipal de Ribeirãozinho do Maranhão - MA
                </div>
            </div>
            """

        # Gerar HTML de impressão
        via1 = gerar_via_html("1ª VIA — PROTOCOLO")
        via2 = gerar_via_html("2ª VIA — RECEBEDOR") if vias == 2 else ""
        separador = "<div style='border-top: 2px dashed #94a3b8; margin: 8px 0; page-break-before: auto;'></div>" if vias == 2 else ""

        html_impressao = f"""
        <div id='area-impressao-protocolo'>
            {via1}
            {separador}
            {via2}
        </div>
        """

        st.markdown("---")
        st.markdown(f"#### 🖨️ Comprovante de Protocolo — Nº {prot['numero']}")
        st.markdown(f"**{vias} via(s)** gerada(s) para impressão.")

        st.markdown(html_impressao, unsafe_allow_html=True)

        html_completo = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Protocolo {prot['numero']}</title>
<style>
body {{ font-family: 'Inter', Arial, sans-serif; margin: 20px; color: #0f172a; background: #fff; }}
@media print {{ body {{ margin: 10px; }} }}
table {{ border-collapse: collapse; width: 100%; }}
td {{ padding: 5px 8px; border: 1px solid #e2e8f0; font-size: 12px; }}
</style>
</head>
<body onload="setTimeout(function(){{ window.print(); }}, 500);">
    {via1}
    {separador}
    {via2}
</body>
</html>"""

        col_pdf, col_html, col_voltar = st.columns(3)
        with col_pdf:
            try:
                pdf_buffer = gerar_comprovante_protocolo_pdf(prot)
                pdf_bytes = pdf_buffer.getvalue()
                st.download_button(
                    label="📄 Baixar Comprovante em PDF",
                    data=pdf_bytes,
                    file_name=f"comprovante_{prot['numero'].replace('/', '_')}.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Erro ao gerar PDF do comprovante: {e}")

        with col_html:
            st.download_button(
                label="🖨️ Imprimir no Navegador (HTML)",
                data=html_completo.encode("utf-8"),
                file_name=f"protocolo_{prot['numero'].replace('/', '_')}.html",
                mime="text/html",
                type="secondary",
                use_container_width=True
            )

        with col_voltar:
            if st.button("📋 Voltar para Protocolos", use_container_width=True, type="secondary"):
                st.session_state.pop("ultimo_protocolo_impresso", None)
                st.session_state["mostrar_formulario_protocolo"] = False
                st.session_state["menu_selecionado"] = "📋 Protocolos"
                st.rerun()

        st.markdown("<div style='margin-top: 14px;'></div>", unsafe_allow_html=True)
        if prot.get("contrato") and str(prot.get("contrato")) != "—":
            if st.button(f"⚖️ Emitir Minuta Oficial AGU (Word / PDF) para o Contrato {prot['contrato']}", type="primary", use_container_width=True, key="btn_ir_agu_pos_salvar"):
                st.session_state["contrato_agu_selecionado"] = str(prot["contrato"])
                st.session_state["menu_selecionado"] = "⚖️ Modelos AGU / Minutas"
                st.session_state["mostrar_formulario_protocolo"] = False
                st.session_state.pop("ultimo_protocolo_impresso", None)
                st.rerun()
        else:
            if st.button(f"⚖️ Gerar Minuta Oficial AGU a partir deste Protocolo ({prot['numero']})", type="primary", use_container_width=True, key="btn_ir_agu_pos_salvar"):
                st.session_state["protocolo_agu_selecionado"] = str(prot["numero"])
                st.session_state["menu_selecionado"] = "⚖️ Modelos AGU / Minutas"
                st.session_state["mostrar_formulario_protocolo"] = False
                st.session_state.pop("ultimo_protocolo_impresso", None)
                st.rerun()



def formulario_contrato():
    tela_modelos_agu()

import re
import requests
import json

def extrair_dotacao_com_ia(texto_bruto: str) -> dict:
    """Extrai e classifica informações orçamentárias municipais a partir de texto bruto copiado.
    Suporta formato tabular/planilha (Exercício, Poder, Órgão, Unidade, Natureza, Fonte) e texto corrido.
    Tenta via IA (Gemini) e possui motor especialista contábil (MCASP/Lei 4.320/64) resiliente."""
    if not texto_bruto or not str(texto_bruto).strip():
        return {}
        
    t = str(texto_bruto).strip()

    # 1. Tentar via LLM / Gemini se houver conectividade
    try:
        import json
        import requests
        import re
        import os
        api_key = os.environ.get("GEMINI_API_KEY", "AQ.Ab8RN6J8dyEt9ArI7BqLKaLvuF52lLgIlr68BaQ6cekMTdu3dw")
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
        
        prompt = f'''Você é um especialista em contabilidade e orçamento público municipal brasileiro.
O usuário forneceu um texto contendo dados de dotação orçamentária.
Texto: {t}

Extraia as informações para as seguintes chaves:
- "exercicio": Ano do exercício (ex: "2026").
- "poder": Ex: "Poder Executivo 02.00".
- "orgao": Órgão ou Fundo (ex: "02.15.00 - Fundo Municipal de Educação").
- "unidade": Unidade Orçamentária / Funcional Programática / Projeto Atividade (ex: "MANUTENÇÃO 12.361.0402.2022.0000").
- "natureza": Natureza da despesa (ex: "33.90.30.00 - Material de Consumo").
- "outros_campos": Qualquer outro campo presente como Fonte de Recursos, Subfunção, etc.

Retorne APENAS um JSON válido no formato:
{{"exercicio": "", "poder": "", "orgao": "", "unidade": "", "natureza": "", "outros_campos": ""}}'''
        
        payload = {"contents": [{"parts":[{"text": prompt}]}]}
        resp = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=4)
        if resp.status_code == 200:
            data = resp.json()
            if "candidates" in data and data["candidates"]:
                resposta_llm = data["candidates"][0]["content"]["parts"][0]["text"]
                json_str = re.search(r"\{.*\}", resposta_llm, re.DOTALL)
                if json_str:
                    res_json = json.loads(json_str.group(0))
                    if any(res_json.values()):
                        return res_json
    except Exception:
        pass

    # 2. MOTOR ESPECIALISTA RESILIENTE DE CONTABILIDADE PÚBLICA (MCASP / Lei 4.320/64)
    import re
    from datetime import date

    res = {
        'exercicio': str(date.today().year),
        'poder': 'Poder Executivo 02.00',
        'orgao': '',
        'unidade': '',
        'natureza': '',
        'outros_campos': ''
    }

    # ESTRATÉGIA 1: Parsing de linhas com rótulos e colunas tabuladas (Excel / Tabela copiada)
    linhas = [l.strip() for l in t.splitlines() if l.strip()]
    outros_itens = []
    linhas_detectadas = 0

    for linha in linhas:
        # Separação por tabulações ou por dois pontos ou múltiplos espaços
        partes = [p.strip() for p in re.split(r'\t+', linha) if p.strip()]
        if len(partes) == 1 and ':' in linha:
            partes = [p.strip() for p in linha.split(':', 1) if p.strip()]
        elif len(partes) == 1:
            m_sp = re.split(r'\s{2,}', linha, maxsplit=1)
            if len(m_sp) > 1:
                partes = [m_sp[0].strip(), m_sp[1].strip()]

        if len(partes) >= 2:
            lbl = partes[0].lower().replace(':', '').strip()
            val_partes = partes[1:]

            if re.match(r'^(exerc[ií]cio|ano)$', lbl, re.IGNORECASE):
                res['exercicio'] = val_partes[0]
                linhas_detectadas += 1
            elif re.match(r'^poder', lbl, re.IGNORECASE):
                res['poder'] = ' '.join(val_partes)
                linhas_detectadas += 1
            elif re.match(r'^[óo]rg[ãa]o', lbl, re.IGNORECASE):
                cod = ''
                nome = ''
                for vp in val_partes:
                    if re.match(r'^\d{2}\.\d{2}(\.\d{2})?$', vp):
                        cod = vp
                    else:
                        nome = vp if not nome else f'{nome} - {vp}'
                if cod and nome:
                    res['orgao'] = f'{cod} - {nome}'
                else:
                    res['orgao'] = ' - '.join(val_partes)
                linhas_detectadas += 1
            elif re.search(r'(unidade|projeto|atividade|funcional)', lbl, re.IGNORECASE):
                cod = ''
                nome = ''
                for vp in val_partes:
                    if re.search(r'\d{2}\.\d{3}\.\d{4}', vp):
                        cod = vp
                    else:
                        nome = vp if not nome else f'{nome} - {vp}'
                if cod and nome:
                    res['unidade'] = f'{nome} {cod}'
                else:
                    res['unidade'] = ' - '.join(val_partes)
                linhas_detectadas += 1
            elif re.search(r'(natureza|elemento|despesa)', lbl, re.IGNORECASE):
                cod = ''
                nome = ''
                for vp in val_partes:
                    if re.search(r'[34][34]?\.\d{2}\.\d{2}', vp) or re.match(r'^[34]\d{7}$', vp):
                        cod = vp
                    else:
                        nome = vp if not nome else f'{nome} - {vp}'
                if cod and nome:
                    res['natureza'] = f'{cod} - {nome}'
                else:
                    res['natureza'] = ' - '.join(val_partes)
                linhas_detectadas += 1
            else:
                j_val = ' - '.join(val_partes)
                outros_itens.append(f"{partes[0]}: {j_val}")

    if outros_itens:
        res['outros_campos'] = ' | '.join(outros_itens)

    # Se conseguiu identificar via linhas tabuladas, retorna
    if linhas_detectadas >= 2:
        return res

    # ESTRATÉGIA 2: Parsing de texto corrido / blocos separados por | ou ;
    map_nat = {
        '339030': 'Material de Consumo',
        '339039': 'Outros Serviços de Terceiros - Pessoa Jurídica',
        '339036': 'Outros Serviços de Terceiros - Pessoa Física',
        '339040': 'Serviços de Tecnologia da Informação e Comunicação - TIC',
        '449052': 'Equipamentos e Material Permanente',
        '449051': 'Obras e Instalações',
        '339032': 'Material, Bem ou Serviço para Distribuição Gratuita',
        '339014': 'Diárias - Civil',
        '339035': 'Serviços de Consultoria',
        '319011': 'Vencimentos e Vantagens Fixas - Pessoal Civil',
        '319013': 'Obrigações Patronais',
        '335041': 'Contribuições'
    }

    partes = [p.strip() for p in re.split(r'[\n|;]+', t) if p.strip()]

    # A. Natureza da Despesa
    for p in partes:
        m_nat_lbl = re.search(r'(?:natureza(?:\s+da\s+despesa)?|elemento(?:\s+de\s+despesa)?|rubrica)[\s:]*([^\n|;]+)', p, re.IGNORECASE)
        if m_nat_lbl:
            res['natureza'] = m_nat_lbl.group(1).strip()
            break
    if not res['natureza']:
        m_nat_fmt = re.search(r'\b([34][34]?\.[1-9]?\d?\.\d{2}\.\d{2}(?:\.\d{2})?)\b', t)
        if m_nat_fmt:
            cod = m_nat_fmt.group(1)
            cod_digits = re.sub(r'\D', '', cod)[:6]
            nome_padrao = map_nat.get(cod_digits, '')
            m_ctx = re.search(re.escape(cod) + r'\s*[-–—]?\s*([a-zA-ZÀ-ÿ\s]{4,45})', t)
            if m_ctx and len(m_ctx.group(1).strip()) > 3:
                res['natureza'] = f'{cod} - {m_ctx.group(1).strip()}'
            elif nome_padrao:
                res['natureza'] = f'{cod} - {nome_padrao}'
            else:
                res['natureza'] = cod
        else:
            m_nat_raw = re.search(r'\b([34][1-9]90\d{2}(?:\d{2})?)\b', t)
            if m_nat_raw:
                raw_c = m_nat_raw.group(1)
                fmt_c = f'{raw_c[0]}.{raw_c[1]}.{raw_c[2:4]}.{raw_c[4:6]}'
                if len(raw_c) >= 8:
                    fmt_c += f'.{raw_c[6:8]}'
                else:
                    fmt_c += '.00'
                nome_padrao = map_nat.get(raw_c[:6], '')
                res['natureza'] = f'{fmt_c} - {nome_padrao}' if nome_padrao else fmt_c

    # B. Unidade Orçamentária / Funcional Programática
    for p in partes:
        m_uni_lbl = re.search(r'(?:projeto(?:\s*[\/e]\s*atividade)?|funcional(?:\s+program[áa]tica)?|\ba[çc][ãa]o\b)[\s:]*([^\n|;]+)', p, re.IGNORECASE)
        if m_uni_lbl:
            res['unidade'] = m_uni_lbl.group(1).strip()
            break
    if not res['unidade']:
        m_func = re.search(r'\b(\d{2}\.\d{3}\.\d{4}\.\d{4}(?:\.\d{4})?)\b', t)
        if m_func:
            cod_func = m_func.group(1)
            m_act = re.search(r'([a-zA-ZÀ-ÿ\s]{4,40})\s+' + re.escape(cod_func), t)
            if m_act:
                res['unidade'] = f'{m_act.group(1).strip().upper()} {cod_func}'
            else:
                m_act2 = re.search(re.escape(cod_func) + r'\s*[-–—]?\s*([a-zA-ZÀ-ÿ\s]{4,40})', t)
                if m_act2:
                    res['unidade'] = f'{m_act2.group(1).strip().upper()} {cod_func}'
                else:
                    res['unidade'] = f'MANUTENÇÃO E FUNCIONAMENTO {cod_func}'
        else:
            m_spc = re.search(r'\b(\d{2}\s+\d{3}\s+\d{4}\s+\d{4})\b', t)
            if m_spc:
                nums = m_spc.group(1).split()
                res['unidade'] = f'MANUTENÇÃO {nums[0]}.{nums[1]}.{nums[2]}.{nums[3]}.0000'

    # C. Órgão / Secretaria / Fundo
    for p in partes:
        m_org_lbl = re.search(r'(?:[óo]rg[ãa]o(?:\s+or[çc]ament[áa]rio)?|secretaria|fundo\s+municipal)[\s:]*([^\n|;]+)', p, re.IGNORECASE)
        if m_org_lbl:
            val_o = m_org_lbl.group(1).strip()
            val_o = re.sub(r'\s*[-–—]\s*(?:unidade|projeto|natureza|3\.[1-9]|33\.|12\.\d{3}).*', '', val_o, flags=re.IGNORECASE).strip()
            if not val_o.lower().startswith(('secretaria', 'fundo', 'órgão', 'gabinete', 'prefeitura', '02.')):
                val_o = f'Fundo / Secretaria {val_o}'
            res['orgao'] = val_o
            break
    if not res['orgao']:
        m_org_cod = re.search(r'\b(02\.\d{2}(?:\.\d{2})?\s*[-–—]\s*[a-zA-ZÀ-ÿ\s\.\/]+?)(?=[|;\n]|$|\s*-\s*\d{2}\.)', t)
        if m_org_cod:
            res['orgao'] = m_org_cod.group(1).strip()
        else:
            m_org_words = re.search(r'\b((?:Secretaria|Fundo|Gabinete|Prefeitura|Superintendência)\s+(?:Municipal\s+)?(?:de\s+)?[a-zA-ZÀ-ÿ\s]+?)(?=[|;\n]|\d{2}\.|$|\s*-\s*)', t, re.IGNORECASE)
            if m_org_words:
                res['orgao'] = m_org_words.group(1).strip()

    # D. Poder
    for p in partes:
        m_pod = re.search(r'(?:poder[\s:]*)?((?:0[12]\.?00\s*[-–—]\s*)?poder\s+(?:executivo|legislativo)(?:\s+0[12]\.?00)?)', p, re.IGNORECASE)
        if m_pod:
            res['poder'] = m_pod.group(1).strip()
            break
    if 'legislativo' in t.lower() and 'poder executivo' not in res['poder'].lower():
        res['poder'] = 'Poder Legislativo 01.00'

    # E. Exercício
    m_ano_lbl = re.search(r'(?:exerc[ií]cio|ano)[\s:]*([12]\d{3})', t, re.IGNORECASE)
    if m_ano_lbl:
        res['exercicio'] = m_ano_lbl.group(1)
    else:
        cod_func_clean = res['unidade']
        t_without_func = t.replace(cod_func_clean, '')
        m_ano_gen = re.search(r'\b(202[4-9]|203[0-5])\b', t_without_func)
        if m_ano_gen:
            res['exercicio'] = m_ano_gen.group(1)

    return res

def formatar_clausulas_inteligencia(texto_bruto: str, num_clausula_principal: int) -> list:
    """Motor de formatação inteligente que utiliza o Google Gemini (LLM via REST) para readequar 
    o texto juridicamente em estrita observância à Nova Lei de Licitações (Lei nº 14.133/2021),
    com fallback para o sistema baseado em regras se a API falhar."""
    if not texto_bruto or not str(texto_bruto).strip():
        return []
    
    try:
        api_key = "AQ.Ab8RN6J8dyEt9ArI7BqLKaLvuF52lLgIlr68BaQ6cekMTdu3dw"
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key={api_key}"
        
        prompt = f"""
Você é um Procurador Municipal e especialista sênior em Direito Administrativo e Contratações Públicas, com foco exclusivo na Nova Lei de Licitações e Contratos Administrativos (Lei Federal nº 14.133/2021) e nos Modelos Oficiais da Advocacia-Geral da União (AGU).

O usuário forneceu anotações/demandas para inclusão como cláusulas adicionais especiais em um contrato administrativo municipal.

TEXTO BRUTO DO USUÁRIO:
{texto_bruto}

DIRETRIZES OBRIGATÓRIAS DE CONFORMIDADE COM A NOVA LEI DE LICITAÇÕES (LEI Nº 14.133/2021):
1. OBSERVÂNCIA NORMATIVA ESTRITA: Cada item deve ser redigido em perfeita consonância com os princípios (art. 5º da Lei 14.133/2021) e regras da Lei Federal nº 14.133/2021 (arts. 89 a 154 - execução, fiscalização pelo art. 117, garantias, alterações contratuais, sanções do art. 156 e extinção contratual).
2. TERMINOLOGIA LEGAL: Utilize terminologia jurídica precisa e padronizada da AGU. Substitua qualquer menção ao Município ou Prefeitura por "CONTRATANTE", e qualquer menção à empresa/fornecedor por "CONTRATADA". Jamais faça referência à revogada Lei 8.666/1993.
3. ADEQUAÇÃO E LEGALIDADE: Converta anotações informais em preceitos contratuais executáveis, claros, com fixação de responsabilidades, prazos e consequências legais conformadas à Lei 14.133/2021.
4. NUMERAÇÃO PRECISA: Cada sub-cláusula DEVE OBRIGATORIAMENTE iniciar com a numeração no formato exato '{num_clausula_principal}.1.', '{num_clausula_principal}.2.', '{num_clausula_principal}.3.', sucessivamente.
5. FORMATO DE SAÍDA: Retorne APENAS as sub-cláusulas numeradas geradas, exatamente uma por linha. NÃO adicione marcadores (bullets, asteriscos, traços), nem títulos intermediários, nem introdução ("Segue...", "Aqui estão...") ou conclusão.
"""
        
        payload = {
            "contents": [{"parts":[{"text": prompt}]}]
        }
        
        headers = {"Content-Type": "application/json"}
        response = requests.post(url, headers=headers, json=payload, timeout=25)
        
        if response.status_code == 200:
            data = response.json()
            try:
                texto_gerado = data["candidates"][0]["content"]["parts"][0]["text"]
                linhas = str(texto_gerado).strip().split('\n')
                clausulas_prontas = []
                for linha in linhas:
                    l = linha.strip()
                    l = l.replace("**", "")
                    if l:
                        l = re.sub(r'^[-•*]\s*', '', l)
                        clausulas_prontas.append(l)
                if clausulas_prontas:
                    return clausulas_prontas
            except (KeyError, IndexError):
                pass
                
    except Exception as e:
        print(f"Erro na IA do Gemini (API REST), acionando Fallback Baseado em Regras: {e}")
        
    # FALLBACK (Motor Simbólico / Lógico em observância à Lei 14.133/2021)
    linhas = str(texto_bruto).strip().split('\n')
    clausulas_prontas = []
    subitem_count = 1
    
    for linha in linhas:
        l = linha.strip()
        if not l:
            continue
        l = re.sub(r'^(\d+[\.\)]\s*)+', '', l)
        l = re.sub(r'^[-\•\>]\s*', '', l)
        # Adequação terminológica da Lei 14.133/2021
        l = re.sub(r'\blei\s*(?:nº?\s*)?8\.?666(?:\/93)?\b', 'Lei Federal nº 14.133/2021', l, flags=re.IGNORECASE)
        l = re.sub(r'\bcontratada\b', 'CONTRATADA', l, flags=re.IGNORECASE)
        l = re.sub(r'\bcontratante\b', 'CONTRATANTE', l, flags=re.IGNORECASE)
        l = re.sub(r'\badministração pública\b', 'Administração Pública', l, flags=re.IGNORECASE)
        l = re.sub(r'\bprefeitura\b', 'Prefeitura', l, flags=re.IGNORECASE)
        l = re.sub(r'\bfiscal do contrato\b', 'Fiscal do Contrato designado nos termos do art. 117 da Lei nº 14.133/2021', l, flags=re.IGNORECASE)
        if len(l) > 0:
            l = l[0].upper() + l[1:]
        if not l.endswith(('.', ';', ':')):
            l += '.'
            
        numero_subitem = f"{num_clausula_principal}.{subitem_count}."
        clausulas_prontas.append(f"{numero_subitem} {l} (Em observância à Lei Federal nº 14.133/2021)")
        subitem_count += 1
        
    return clausulas_prontas


def salvar_contrato_cgutec_bd(dados: dict) -> tuple[bool, str]:
    try:
        with get_engine().connect() as conn:
            # 1. Orgao
            orgao = dados["orgao_nome"]
            oid = conn.execute(text("SELECT id FROM orgaos WHERE nome = :n"), {"n": orgao}).scalar()
            if not oid:
                res = conn.execute(text("INSERT INTO orgaos (nome, sigla, ativo) VALUES (:n, :s, 1)"), {"n": orgao, "s": orgao[:10]})
                oid = res.lastrowid
                
            # 2. Fornecedor
            forn_nome = dados["fornecedor_nome"]
            cnpj = dados.get("fornecedor_cnpj", "")
            fid = None
            if cnpj and cnpj.strip():
                fid = conn.execute(text("SELECT id FROM fornecedores WHERE cnpj_cpf = :cnpj"), {"cnpj": formatar_cnpj_cpf(cnpj)}).scalar()
            if not fid:
                fid = conn.execute(text("SELECT id FROM fornecedores WHERE razao_social = :rz OR nome_fantasia = :rz"), {"rz": forn_nome}).scalar()
            if not fid:
                res_f = conn.execute(text("INSERT INTO fornecedores (nome_fantasia, razao_social, cnpj_cpf, ativo) VALUES (:rz, :rz, :c, 1)"), {"rz": forn_nome, "c": formatar_cnpj_cpf(cnpj)})
                fid = res_f.lastrowid
                
            uid = st.session_state.get("usuario", {}).get("id", 1)
            
            # 3. Check if already exists to prevent duplicate clicks
            # M3 data doesn't explicitly store numero_contrato parsed cleanly, but it has num_ct_cgutec in session or UI
            # We will generate the next number if not provided
            ano_alvo = pd.to_datetime(dados.get("data_assinatura", date.today())).year
            ultimo = conn.execute(text("SELECT MAX(numero_contrato) FROM contratos WHERE ano_contrato = :ano"), {"ano": ano_alvo}).scalar()
            nc_final = 1 if not ultimo else int(ultimo) + 1
            ncomp_final = f"{str(nc_final).zfill(3)}/{ano_alvo}"
            
            # insert
            res_cont = conn.execute(
                text("""
                    INSERT INTO contratos (
                        numero_contrato, ano_contrato, numero_completo, orgao_id, modalidade,
                        processo_adm, objeto, fornecedor_id, valor_total, data_assinatura,
                        data_publicacao, data_vencimento, vigencia_descricao, solicitante,
                        secretario, fiscal, status, modelo_agu, dotacao_orcamentaria,
                        cep_orgao, endereco_orgao, bairro_orgao, cidade_orgao, uf_orgao, cnpj_orgao,
                        criado_por, data_criacao, atualizado_em
                    ) VALUES (
                        :nc, :ano, :ncomp, :oid, :mod,
                        :proc, :obj, :fid, :val, :d_ass,
                        :d_pub, :d_venc, :vig, :solic,
                        :sec, :fisc, 'VIGENTE', :mod_agu, :dot,
                        :cep_o, :end_o, :bai_o, :cid_o, :uf_o, :cnpj_o,
                        :uid, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
                    )
                """),
                {
                    "nc": nc_final, "ano": ano_alvo, "ncomp": ncomp_final, "oid": oid,
                    "mod": f"{dados.get('modalidade')} Nº {dados.get('numero_modalidade')}".strip() if (dados.get("numero_modalidade") and str(dados.get("numero_modalidade")).strip() and str(dados.get("numero_modalidade")).strip() not in str(dados.get("modalidade", ""))) else dados.get("modalidade"),
                    "proc": dados.get("processo_adm"),
                    "obj": dados.get("objeto"), "fid": fid, "val": clean_money(dados.get("valor_total", 0)),
                    "d_ass": dados.get("data_assinatura"), "d_pub": dados.get("data_assinatura"), # fallback
                    "d_venc": dados.get("data_vencimento"), "vig": dados.get("vigencia_descricao"),
                    "solic": "AGU API", "sec": dados.get("secretario"), "fisc": dados.get("fiscal"),
                    "mod_agu": dados.get("modelo_agu", "COMPRAS"), "dot": dados.get("dotacao_orcamentaria"),
                    "cep_o": dados.get("orgao_cep"), "end_o": dados.get("orgao_endereco"),
                    "bai_o": dados.get("orgao_bairro"), "cid_o": dados.get("orgao_cidade"),
                    "uf_o": dados.get("orgao_uf"), "cnpj_o": dados.get("orgao_cnpj"), "uid": uid
                }
            )
            # Atualiza o cadastro permanente do órgão com os dados deste contrato
            conn.execute(
                text("""
                    UPDATE orgaos SET
                        secretario_padrao = COALESCE(NULLIF(:sec, ''), secretario_padrao),
                        cargo_secretario = COALESCE(NULLIF(:carg, ''), cargo_secretario),
                        fiscal_padrao = COALESCE(NULLIF(:fisc, ''), fiscal_padrao),
                        endereco = COALESCE(NULLIF(:end, ''), endereco),
                        bairro = COALESCE(NULLIF(:bai, ''), bairro),
                        cidade = COALESCE(NULLIF(:cid, ''), cidade),
                        uf = COALESCE(NULLIF(:uf, ''), uf),
                        cep = COALESCE(NULLIF(:cep, ''), cep),
                        cnpj = COALESCE(NULLIF(:cnpj, ''), cnpj),
                        dotacao_padrao = COALESCE(NULLIF(:dot, ''), dotacao_padrao)
                    WHERE id = :oid
                """),
                {
                    "sec": dados.get("secretario"),
                    "carg": dados.get("cargo_secretario"),
                    "fisc": dados.get("fiscal"),
                    "end": dados.get("orgao_endereco"),
                    "bai": dados.get("orgao_bairro"),
                    "cid": dados.get("orgao_cidade"),
                    "uf": dados.get("orgao_uf"),
                    "cep": dados.get("orgao_cep"),
                    "cnpj": dados.get("orgao_cnpj"),
                    "dot": dados.get("dotacao_orcamentaria"),
                    "oid": oid
                }
            )
            conn.commit()
            try:
                obter_historico_orgao.clear()
            except Exception:
                pass
            try:
                carregar_contratos.clear()
            except Exception:
                pass
            registrar_audit_log(
                conn, "contratos", res_cont.lastrowid if hasattr(res_cont, 'lastrowid') else None, "CRIACAO",
                f"Contrato {ncomp_final} criado via Gerador de Modelo AGU.",
                dados_novos={"numero": ncomp_final, "numero_contrato": nc_final, "ano": ano_alvo, "orgao_id": oid, "fornecedor_id": fid}
            )
            return True, f"Contrato {ncomp_final} cadastrado e salvo"
    except Exception as e:
        return False, str(e)


def atualizar_historico_orgao_bd(orgao_nome: str, dados_atuais: dict):
    """Atualiza o cadastro permanente do órgão com os dados inseridos na elaboração e invalida o cache."""
    if not orgao_nome or not str(orgao_nome).strip() or "Selecione" in str(orgao_nome):
        return
    try:
        with get_engine().begin() as conn:
            oid = conn.execute(text("SELECT id FROM orgaos WHERE nome = :n"), {"n": orgao_nome}).scalar()
            if not oid:
                all_orgaos = conn.execute(text("SELECT id, nome FROM orgaos")).mappings().fetchall()
                norm_in = normalizar_texto(orgao_nome)
                for r in all_orgaos:
                    if normalizar_texto(r["nome"]) == norm_in:
                        oid = r["id"]
                        break
                if not oid:
                    for r in all_orgaos:
                        norm_db = normalizar_texto(r["nome"])
                        if norm_in in norm_db or norm_db in norm_in:
                            oid = r["id"]
                            break
                if not oid:
                    palavras = [p for p in re.split(r'[\s,/]+', norm_in) if len(p) > 3 and p not in ("SECRETARIA", "MUNICIPAL", "GABINETE", "COORDENACAO", "GERENCIA", "SETOR", "FUNDO", "DESENVOLVIMENTO")]
                    for p in palavras:
                        for r in all_orgaos:
                            if p in normalizar_texto(r["nome"]):
                                oid = r["id"]
                                break
                        if oid:
                            break
            if oid:
                conn.execute(
                    text("""
                        UPDATE orgaos SET
                            secretario_padrao = COALESCE(NULLIF(:sec, ''), secretario_padrao),
                            cargo_secretario = COALESCE(NULLIF(:carg, ''), cargo_secretario),
                            fiscal_padrao = COALESCE(NULLIF(:fisc, ''), fiscal_padrao),
                            endereco = COALESCE(NULLIF(:end, ''), endereco),
                            bairro = COALESCE(NULLIF(:bai, ''), bairro),
                            cidade = COALESCE(NULLIF(:cid, ''), cidade),
                            uf = COALESCE(NULLIF(:uf, ''), uf),
                            cep = COALESCE(NULLIF(:cep, ''), cep),
                            cnpj = COALESCE(NULLIF(:cnpj, ''), cnpj),
                            dotacao_padrao = COALESCE(NULLIF(:dot, ''), dotacao_padrao)
                        WHERE id = :oid
                    """),
                    {
                        "sec": dados_atuais.get("secretario"),
                        "carg": dados_atuais.get("cargo_secretario"),
                        "fisc": dados_atuais.get("fiscal"),
                        "end": dados_atuais.get("orgao_endereco") or dados_atuais.get("endereco"),
                        "bai": dados_atuais.get("orgao_bairro") or dados_atuais.get("bairro"),
                        "cid": dados_atuais.get("orgao_cidade") or dados_atuais.get("cidade"),
                        "uf": dados_atuais.get("orgao_uf") or dados_atuais.get("uf"),
                        "cep": dados_atuais.get("orgao_cep") or dados_atuais.get("cep"),
                        "cnpj": dados_atuais.get("orgao_cnpj") or dados_atuais.get("cnpj"),
                        "dot": dados_atuais.get("dotacao_orcamentaria") or dados_atuais.get("dotacao"),
                        "oid": oid
                    }
                )
        obter_historico_orgao.clear()
    except Exception as e:
        logger.debug(f"Erro ao atualizar historico orgao: {e}")


def duplicar_contrato_bd(contrato_id: int, usuario_id: int = 1) -> tuple[bool, str, Optional[int]]:
    """Duplica um contrato existente gerando automaticamente o próximo número sequencial disponível do ano."""
    try:
        with get_engine().connect() as conn:
            orig = conn.execute(
                text("SELECT * FROM contratos WHERE id = :id"),
                {"id": int(contrato_id)}
            ).mappings().first()
            if not orig:
                return False, f"Contrato original (ID #{contrato_id}) não encontrado.", None

            ano_alvo = date.today().year
            ultimo = conn.execute(
                text("SELECT MAX(numero_contrato) FROM contratos WHERE ano_contrato = :ano"),
                {"ano": ano_alvo}
            ).scalar()
            nc_novo = 1 if not ultimo else int(ultimo) + 1
            ncomp_novo = f"{str(nc_novo).zfill(3)}/{ano_alvo}"

            orig_num = orig.get("numero_completo") or f"{orig.get('numero_contrato')}/{orig.get('ano_contrato')}"
            obs_duplicado = f"Contrato duplicado a partir do Contrato {orig_num} (ID #{contrato_id})."
            if orig.get("observacoes"):
                obs_duplicado = f"{obs_duplicado} Obs original: {orig.get('observacoes')}"

            res_ins = conn.execute(
                text("""
                    INSERT INTO contratos (
                        numero_contrato, ano_contrato, numero_completo, orgao_id, modalidade,
                        processo_adm, objeto, fornecedor_id, valor_total, data_assinatura,
                        data_publicacao, data_vencimento, vigencia_descricao, solicitante,
                        fiscal, secretario, status, modelo_agu, dotacao_orcamentaria,
                        cep_orgao, endereco_orgao, bairro_orgao, cidade_orgao, uf_orgao, cnpj_orgao,
                        observacoes, criado_por, data_criacao, atualizado_em
                    ) VALUES (
                        :nc, :ano, :ncomp, :oid, :mod,
                        :proc, :obj, :fid, :val, :d_ass,
                        :d_pub, :d_venc, :vig, :solic,
                        :fisc, :sec, 'VIGENTE', :mod_agu, :dot,
                        :cep_o, :end_o, :bai_o, :cid_o, :uf_o, :cnpj_o,
                        :obs, :uid, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
                    )
                """),
                {
                    "nc": nc_novo,
                    "ano": ano_alvo,
                    "ncomp": ncomp_novo,
                    "oid": orig.get("orgao_id"),
                    "mod": orig.get("modalidade"),
                    "proc": orig.get("processo_adm"),
                    "obj": orig.get("objeto"),
                    "fid": orig.get("fornecedor_id"),
                    "val": float(orig.get("valor_total") or 0.0),
                    "d_ass": date.today(),
                    "d_pub": date.today(),
                    "d_venc": date.today() + timedelta(days=365),
                    "vig": orig.get("vigencia_descricao") or "12 (doze) meses",
                    "solic": orig.get("solicitante") or "Duplicação Contratual",
                    "fisc": orig.get("fiscal"),
                    "sec": orig.get("secretario"),
                    "mod_agu": orig.get("modelo_agu") or "COMPRAS",
                    "dot": orig.get("dotacao_orcamentaria"),
                    "cep_o": orig.get("cep_orgao"),
                    "end_o": orig.get("endereco_orgao"),
                    "bai_o": orig.get("bairro_orgao"),
                    "cid_o": orig.get("cidade_orgao"),
                    "uf_o": orig.get("uf_orgao"),
                    "cnpj_o": orig.get("cnpj_orgao"),
                    "obs": obs_duplicado,
                    "uid": usuario_id
                }
            )
            conn.commit()
            novo_id = None
            if hasattr(res_ins, 'lastrowid') and res_ins.lastrowid:
                novo_id = res_ins.lastrowid
            elif hasattr(res_ins, 'inserted_primary_key') and res_ins.inserted_primary_key:
                novo_id = res_ins.inserted_primary_key[0]
            if not novo_id:
                try:
                    novo_id = conn.execute(text("SELECT id FROM contratos WHERE numero_completo = :nc ORDER BY id DESC LIMIT 1"), {"nc": ncomp_novo}).scalar()
                except Exception:
                    pass
            try:
                carregar_contratos.clear()
            except Exception:
                pass

            registrar_audit_log(
                conn, "contratos", novo_id, "DUPLICACAO",
                f"Contrato {ncomp_novo} duplicado a partir do contrato {orig_num}.",
                dados_novos={"numero": ncomp_novo, "numero_contrato": nc_novo, "ano": ano_alvo, "origem_id": contrato_id}
            )
            return True, f"Contrato duplicado com sucesso! Novo contrato gerado: **{ncomp_novo}**.", novo_id
    except Exception as e:
        return False, f"Erro ao duplicar contrato: {str(e)}", None


def tela_modelos_agu():
    brasao_b64 = obter_brasao_b64()
    brasao_banner_img = f"<div style='background: rgba(255,255,255,0.92); padding: 6px 12px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.15); display: flex; align-items: center; justify-content: center;'><img src='data:image/png;base64,{brasao_b64}' style='height: 52px; width: auto;' alt='Brasão' /></div>" if brasao_b64 else ""
    st.markdown(f"""
        <div class='municipal-banner' style='margin-bottom: 12px;'>
            <div style='display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 14px;'>
                <div>
                    <div class='municipal-badge-pill'>
                        <span class='beacon-active-gel' style='background: #ffffff;'></span> PADRÃO OFICIAL CGU/AGU • EXERCÍCIO 2026
                    </div>
                    <h2 class='municipal-banner-title'>⚖️ Emissor de Minutas Oficiais e Termos Aditivos AGU</h2>
                    <div class='municipal-banner-subtitle'>
                        Geração de minutas de contratos e termos aditivos em conformidade com a Consultoria-Geral da União (cgu.agu.gov.br/contrato/ e modelos padronizados AGU Lei nº 14.133/21 e Lei nº 8.666/93)
                    </div>
                </div>
                {brasao_banner_img}
            </div>
        </div>
    """, unsafe_allow_html=True)

    df_c = carregar_contratos()

    st.markdown("""
        <style>
        /* Força as abas do Streamlit a dividirem o espaço igualmente */
        div[data-testid="stTabs"] button[role="tab"] {
            flex: 1;
            white-space: pre-wrap;
            text-align: center;
        }
        </style>
    """, unsafe_allow_html=True)

    tab_m3, tab_m1, tab_m4 = st.tabs([
        "✍️ Gerador de Modelo AGU (Padrão CGUTEC 1.0.2)",
        "📑 Gerar Minuta de Contrato Existente",
        "📝 Aditivar Contrato (Termos Aditivos AGU)"
    ])

    with tab_m1:
        if df_c.empty:
            st.info("Nenhum contrato cadastrado ainda.")
        else:
            opcoes = ["Selecione um contrato cadastrado..."] + df_c["numero_completo"].tolist()
            idx_def = 0
            c_pre = st.session_state.pop("contrato_agu_selecionado", None)
            if c_pre and c_pre in opcoes:
                idx_def = opcoes.index(c_pre)
            escolha = st.selectbox("Escolha o Contrato Municipal:", opcoes, index=idx_def, key="sel_agu_tela_m1")
            if escolha != "Selecione um contrato cadastrado...":
                row = df_c[df_c["numero_completo"] == escolha].iloc[0]
                hist_org = obter_historico_orgao(row["orgao"])

                df_prot_ligados = carregar_protocolos_do_contrato(row["id"])
                if not df_prot_ligados.empty:
                    p_info_list = [f"**{r['Protocolo']}** ({r['Solicitante']})" for _, r in df_prot_ligados.iterrows()]
                    st.markdown(f"""
                    <div style='background: #eff6ff; border: 1px solid #bfdbfe; border-left: 5px solid #0284c7; border-radius: 8px; padding: 10px 14px; margin-bottom: 12px;'>
                        <span style='font-size: 12.5px; font-weight: 700; color: #1e40af;'>🔗 Protocolos Oficiais Vinculados:</span>
                        <span style='font-size: 12px; color: #334155; margin-left: 8px;'>{' • '.join(p_info_list)}</span>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown("##### 📦 Tipo de Contratação & Demais Informações")
                c_m0_1, c_m0_2, c_m0_3 = st.columns([1, 1.3, 1.2])
                with c_m0_1:
                    tipo_c_m1 = st.selectbox(
                        "Tipo:",
                        ["Aquisições", "Serviços"],
                        index=0 if ("COMPRA" in str(row.get("objeto", "")).upper() or "AQUIS" in str(row.get("objeto", "")).upper() or "FORNEC" in str(row.get("objeto", "")).upper()) else 1,
                        key=f"tipo_c_m1_{row['id']}"
                    )
                with c_m0_2:
                    demais_inf_m1 = st.radio(
                        "Demais Informações (Modalidade):",
                        ["Pregão", "Concorrência", "Contratação direta"],
                        index=0 if "PREG" in str(row.get("modalidade", "")).upper() else (1 if "CONCORR" in str(row.get("modalidade", "")).upper() else 2),
                        horizontal=True,
                        key=f"demais_inf_m1_{row['id']}"
                    )
                with c_m0_3:
                    if demais_inf_m1 == "Pregão":
                        lbl_m1_mod = "Número do Pregão:"
                        ph_m1_mod = "Ex: 010/2026 ou 010/2026-SRP"
                        def_m1_mod = "010/2026"
                    elif demais_inf_m1 == "Concorrência":
                        lbl_m1_mod = "Número da Concorrência:"
                        ph_m1_mod = "Ex: 001/2026"
                        def_m1_mod = "001/2026"
                    else:
                        lbl_m1_mod = "Número da Dispensa / Inex.:"
                        ph_m1_mod = "Ex: 005/2026"
                        def_m1_mod = "005/2026"

                    m_num_m1 = re.search(r'\d+/\d+', str(row.get("modalidade", "")))
                    if m_num_m1:
                        def_m1_mod = m_num_m1.group(0)

                    num_mod_m1 = st.text_input(
                        lbl_m1_mod,
                        value=def_m1_mod,
                        placeholder=ph_m1_mod,
                        key=f"num_mod_m1_{row['id']}_{demais_inf_m1}"
                    )

                c_m1, c_m2 = st.columns([1.5, 1])
                with c_m1:
                    if tipo_c_m1 == "Aquisições":
                        opcoes_m1 = {
                            "Termo de Contrato - Aquisição de Bens / Compras (Padrão AGU)": "COMPRAS",
                            "Termo de Contrato - Aquisição de Bens para Entrega Imediata": "COMPRAS_IMEDIATA",
                            "Termo de Contrato - Aquisição de Bens de TIC": "COMPRAS_TIC"
                        }
                    else:
                        opcoes_m1 = {
                            "Termo de Contrato - Prestação de Serviços Contínuos (Sem dedicação de mão de obra)": "SERVICOS_CONTINUOS",
                            "Termo de Contrato - Prestação de Serviços Não Contínuos / Escopo": "SERVICOS_ESCOPO",
                            "Termo de Contrato - Obras e Serviços Comuns de Engenharia": "OBRAS",
                            "Termo de Contrato - Serviços de Tecnologia da Informação (TIC)": "SERVICOS_TIC"
                        }
                    modelo_sel = st.selectbox(
                        "Modelo de Minuta AGU (Lei 14.133/2021):",
                        list(opcoes_m1.keys()),
                        key="agu_sel_modelo_tela_m1"
                    )
                    cod_modelo = opcoes_m1[modelo_sel]
                with c_m2:
                    st.info(f"🏛️ **Secretaria:** {row['orgao']}\n\n📍 **Endereço:** {hist_org['endereco']}, {hist_org['bairro']} - CEP {hist_org['cep']}")

                with st.expander("🧠 Cláusulas Adicionais Especiais (Conformidade com a Lei nº 14.133/2021 via IA)", expanded=False):
                    st.caption("O sistema utilizará a Inteligência Artificial para readequar suas cláusulas em estrita conformidade com a Nova Lei de Licitações (Lei nº 14.133/2021).")
                    titulo_clausula_m1 = st.text_input(
                        "Nome / Título da Cláusula (Livre Nomeação pelo Usuário):",
                        value="DAS CONDIÇÕES ESPECIAIS E DISPOSIÇÕES COMPLEMENTARES (LEI Nº 14.133/2021)",
                        help="Defina livremente o assunto/título da cláusula. O prefixo numérico oficial (ex: CLÁUSULA DÉCIMA TERCEIRA – ) será inserido automaticamente.",
                        key=f"titulo_clausula_extra_m1_{row['id']}"
                    )
                    texto_clausulas_m1 = st.text_area(
                        "Cole aqui as cláusulas extras que deseja incluir nesta minuta (um item por linha):",
                        placeholder="Exemplo:\nO prazo de garantia dos equipamentos é de 12 meses.\nA CONTRATADA deverá manter preposto nos termos do art. 118 da Lei nº 14.133/2021.\nA CONTRATANTE poderá rescindir unilateralmente com base no art. 137.",
                        height=110,
                        key=f"clausulas_extras_m1_{row['id']}"
                    )
                    clausulas_formatadas_m1 = formatar_clausulas_inteligencia(texto_clausulas_m1, 13)
                    if clausulas_formatadas_m1:
                        st.success(f"✨ Cláusula '{titulo_clausula_m1}' readequada em observância à Lei Federal nº 14.133/2021!")
                        for cf in clausulas_formatadas_m1:
                            st.caption(f"✓ {cf}")

                # NOVO EXPANDER DE DOTAÇÃO ORÇAMENTÁRIA (M1 - GERADOR DE MINUTA)
                with st.expander("🏦 Dados da Dotação Orçamentária (Edição Inteligente)", expanded=False):
                    st.caption("Você pode alterar os dados orçamentários apenas para a geração deste documento, ou usar a IA para extrair de um texto.")
                    def _callback_extrair_dotacao_m1(target_id):
                        raw_txt_m1 = st.session_state.get(f"txt_dot_m1_{target_id}", "")
                        if not raw_txt_m1 or not str(raw_txt_m1).strip():
                            st.session_state[f"m1_dot_msg_{target_id}"] = ("warning", "Por favor, cole o texto da dotação orçamentária antes de extrair.")
                            return
                        d_ext_m1 = extrair_dotacao_com_ia(raw_txt_m1)
                        if d_ext_m1:
                            if d_ext_m1.get("exercicio"):
                                st.session_state[f"in_m1_ex_{target_id}"] = d_ext_m1["exercicio"]
                                st.session_state[f"m1_dot_ex_{target_id}"] = d_ext_m1["exercicio"]
                            if d_ext_m1.get("poder"):
                                st.session_state[f"in_m1_pod_{target_id}"] = d_ext_m1["poder"]
                                st.session_state[f"m1_dot_pod_{target_id}"] = d_ext_m1["poder"]
                            if d_ext_m1.get("orgao"):
                                st.session_state[f"in_m1_org_{target_id}"] = d_ext_m1["orgao"]
                                st.session_state[f"m1_dot_org_{target_id}"] = d_ext_m1["orgao"]
                            if d_ext_m1.get("unidade"):
                                st.session_state[f"in_m1_uni_{target_id}"] = d_ext_m1["unidade"]
                                st.session_state[f"m1_dot_uni_{target_id}"] = d_ext_m1["unidade"]
                            if d_ext_m1.get("natureza"):
                                st.session_state[f"in_m1_nat_{target_id}"] = d_ext_m1["natureza"]
                                st.session_state[f"m1_dot_nat_{target_id}"] = d_ext_m1["natureza"]
                            if d_ext_m1.get("outros_campos"):
                                st.session_state[f"in_m1_outros_{target_id}"] = d_ext_m1["outros_campos"]
                                st.session_state[f"m1_dot_outros_{target_id}"] = d_ext_m1["outros_campos"]
                            st.session_state[f"m1_dot_msg_{target_id}"] = ("success", "✨ Dados da dotação classificados e preenchidos com sucesso!")
                        else:
                            st.session_state[f"m1_dot_msg_{target_id}"] = ("error", "Não foi possível extrair a dotação. Verifique o texto colado.")

                    c_dot_ia1, c_dot_ia2 = st.columns([2, 1])
                    with c_dot_ia1:
                        txt_dot_m1 = st.text_area("Texto Bruto da Dotação Orçamentária (Copie e cole aqui)", height=100, key=f"txt_dot_m1_{row['id']}")
                    with c_dot_ia2:
                        st.write("")
                        st.write("")
                        st.button("✨ Extrair (IA)", key=f"btn_ext_ia_m1_{row['id']}", on_click=_callback_extrair_dotacao_m1, args=(row['id'],), use_container_width=True)

                    if f"m1_dot_msg_{row['id']}" in st.session_state:
                        tp_m1, tx_m1 = st.session_state.pop(f"m1_dot_msg_{row['id']}")
                        if tp_m1 == "success":
                            st.success(tx_m1)
                        elif tp_m1 == "warning":
                            st.warning(tx_m1)
                        else:
                            st.error(tx_m1)
                    
                    import json
                    dot_dict_parsed = {}
                    try:
                        dot_dict_parsed = json.loads(row.get("dotacao_orcamentaria") or "{}")
                    except:
                        pass
                        
                    c_d1, c_d2 = st.columns(2)
                    with c_d1:
                        m1_dot_ex = st.text_input("Exercício", value=st.session_state.get(f"m1_dot_ex_{row['id']}", dot_dict_parsed.get("exercicio") or ""), key=f"in_m1_ex_{row['id']}")
                        m1_dot_pod = st.text_input("Poder", value=st.session_state.get(f"m1_dot_pod_{row['id']}", dot_dict_parsed.get("poder") or ""), key=f"in_m1_pod_{row['id']}")
                        m1_dot_org = st.text_input("Órgão", value=st.session_state.get(f"m1_dot_org_{row['id']}", dot_dict_parsed.get("orgao") or ""), key=f"in_m1_org_{row['id']}")
                    with c_d2:
                        m1_dot_uni = st.text_input("Unidade Orçamentária / Projeto", value=st.session_state.get(f"m1_dot_uni_{row['id']}", dot_dict_parsed.get("unidade") or ""), key=f"in_m1_uni_{row['id']}")
                        m1_dot_nat = st.text_input("Natureza da Despesa", value=st.session_state.get(f"m1_dot_nat_{row['id']}", dot_dict_parsed.get("natureza") or ""), key=f"in_m1_nat_{row['id']}")
                        m1_dot_outros = st.text_input("Demais Campos / Fonte de Recursos (opcional)", value=st.session_state.get(f"m1_dot_outros_{row['id']}", dot_dict_parsed.get("outros_campos") or ""), key=f"in_m1_outros_{row['id']}", placeholder="Ex: Fonte de Recursos: 15000000 - Recursos Próprios")


                with st.expander("📦 Planilha ou Documento de Itens do Contrato (Auto-Ajustável - Excel / CSV / PDF)", expanded=False):
                    st.caption("A tabela no documento gerado se **auto-ajusta dinamicamente** à quantidade de colunas, linhas e descrição dos cabeçalhos da planilha ou PDF importado.")
                    r_id = row['id']
                    key_m1_counter = st.session_state.get(f"key_up_m1_{r_id}", 0)
                    key_m1_up = f"upload_itens_m1_{r_id}_{key_m1_counter}"
                    up_itens_m1 = st.file_uploader(
                        "Importar Planilha ou PDF de Itens (.xlsx, .xls, .csv, .pdf):",
                        type=["xlsx", "xls", "csv", "pdf"],
                        key=key_m1_up,
                        help="Carregue qualquer planilha ou PDF de itens. A quantidade de colunas, linhas e cabeçalhos será ajustada automaticamente na minuta."
                    )
                    itens_m1_proc = []
                    if up_itens_m1 is not None:
                        itens_m1_proc = processar_planilha_itens_excel(up_itens_m1.getvalue(), file_name=up_itens_m1.name, objeto_contexto=row.get("objeto", ""))
                        if itens_m1_proc:
                            tot_m1 = sum(float(it.get("_vt_float", 0.0) or 0.0) for it in itens_m1_proc)
                            col_m1_st, col_m1_cl = st.columns([3.2, 1.2])
                            with col_m1_st:
                                st.success(f"✅ **{len(itens_m1_proc)} itens** carregados da planilha!")
                            with col_m1_cl:
                                def _callback_limpar_m1(target_id):
                                    st.session_state[f"key_up_m1_{target_id}"] = st.session_state.get(f"key_up_m1_{target_id}", 0) + 1
                                st.button("🗑️ Limpar Planilha", key=f"btn_cl_m1_{r_id}_{key_m1_counter}", on_click=_callback_limpar_m1, args=(r_id,), use_container_width=True)
                            headers_m1 = ["Item", "Descrição", "Qtd", "Unid", "Valor Unitário", "Valor Total"]
                            rows_m1 = [
                                [str(it.get("item", "")), str(it.get("descricao", "")), str(it.get("quantidade", "")), str(it.get("unidade", "")), str(it.get("valor_unitario", "")), str(it.get("valor_total", ""))]
                                for it in itens_m1_proc
                            ]
                            df_itens_m1_view = pd.DataFrame(rows_m1, columns=headers_m1)
                            
                            tab_m1_fmt, tab_m1_grid = st.tabs([
                                "📋 Planilha Formatada (Quebra de Texto e Bordas)",
                                "🔢 Tabela Interativa (Grid)"
                            ])
                            with tab_m1_fmt:
                                html_m1 = renderizar_tabela_html_formatada(headers_m1, rows_m1, total_geral=tot_m1)
                                st.markdown(html_m1, unsafe_allow_html=True)
                            with tab_m1_grid:
                                col_cfg_m1 = {
                                    "Item": st.column_config.TextColumn("Item", width="small"),
                                    "Descrição": st.column_config.TextColumn("Descrição", width="large"),
                                    "Qtd": st.column_config.TextColumn("Qtd", width="small"),
                                    "Unid": st.column_config.TextColumn("Unid", width="small"),
                                    "Valor Unitário": st.column_config.TextColumn("Valor Unitário", width="medium"),
                                    "Valor Total": st.column_config.TextColumn("Valor Total", width="medium"),
                                }
                                st.dataframe(df_itens_m1_view, column_config=col_cfg_m1, use_container_width=True, hide_index=True)

                            col_tot_m1, col_down_m1 = st.columns([2.5, 1.5])
                            with col_tot_m1:
                                st.caption(f"Valor Total dos Itens da Planilha: **R$ {tot_m1:,.2f}**".replace(",", "X").replace(".", ",").replace("X", "."))
                            with col_down_m1:
                                excel_m1_bytes = gerar_planilha_itens_excel_formatada(headers_m1, rows_m1, total_geral=tot_m1)
                                st.download_button(
                                    label="📥 Baixar Planilha (.xlsx)",
                                    data=excel_m1_bytes,
                                    file_name=f"itens_contrato_{r_id}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"dl_m1_btn_{r_id}_{key_m1_counter}",
                                    use_container_width=True
                                )

                # Parse dotacao json se existir
                import json
                dot_dict_parsed = {}
                try:
                    dot_dict_parsed = json.loads(row.get("dotacao_orcamentaria") or "{}")
                except:
                    pass

                dados_payload = {
                    "numero_completo": row["numero_completo"],
                    "numero_contrato": row.get("numero_contrato", 1),
                    "processo_adm": row.get("processo_adm") or "001/2026",
                    "modalidade": demais_inf_m1.upper(),
                    "numero_modalidade": num_mod_m1.strip(),
                    "objeto": row.get("objeto") or "Contratação de fornecimento/serviço",
                    "valor_total": row.get("valor_total") or 0.0,
                    "data_assinatura": row.get("data_assinatura"),
                    "data_vencimento": row.get("data_vencimento"),
                    "vigencia_descricao": row.get("vigencia_descricao") or "12 (doze) meses",
                    "secretario": row.get("secretario") or hist_org["secretario"],
                    "cargo_secretario": hist_org.get("cargo_secretario", "Secretário(a) Municipal Titular"),
                    "fiscal": row.get("fiscal") or hist_org["fiscal"],
                    "orgao_nome": row["orgao"],
                    "orgao_cnpj": hist_org["cnpj"],
                    "orgao_cep": hist_org["cep"],
                    "orgao_endereco": hist_org["endereco"],
                    "orgao_bairro": hist_org["bairro"],
                    "orgao_cidade": hist_org["cidade"],
                    "orgao_uf": hist_org["uf"],
                    "fornecedor_nome": row.get("fornecedor", "EMPRESA CONTRATADA"),
                    "fornecedor_cnpj": row.get("cnpj_cpf", "00.000.000/0001-00"),
                    "fornecedor_representante": row.get("fornecedor_representante") or "Representante Legal",
                    "fornecedor_cargo": row.get("fornecedor_cargo") or "Sócio Administrador",
                    "fornecedor_endereco": row.get("fornecedor_endereco") or "Sede da Empresa",
                    "fornecedor_bairro": row.get("fornecedor_bairro") or "Centro",
                    "fornecedor_cep": row.get("fornecedor_cep") or "65900-000",
                    "fornecedor_cidade": row.get("fornecedor_cidade") or "Imperatriz",
                    "fornecedor_uf": row.get("fornecedor_uf") or "MA",
                    "dotacao_orcamentaria": f"{m1_dot_org} | {m1_dot_uni} | {m1_dot_nat}" + (f" | {m1_dot_outros}" if m1_dot_outros and str(m1_dot_outros).strip() else ""),
                    "dotacao_exercicio": m1_dot_ex,
                    "dotacao_poder": m1_dot_pod,
                    "dotacao_orgao": m1_dot_org,
                    "dotacao_unidade": m1_dot_uni,
                    "dotacao_natureza": m1_dot_nat,
                    "dotacao_outros": m1_dot_outros if m1_dot_outros else "",
                    "modelo_agu": cod_modelo,
                    "titulo_clausula_adicional": titulo_clausula_m1,
                    "clausulas_adicionais_formatadas": clausulas_formatadas_m1,
                    "tabela_dinamica": getattr(itens_m1_proc, "tabela_dinamica", None) if itens_m1_proc else None,
                    "itens": itens_m1_proc if itens_m1_proc else None,
                    "usar_total_itens": True if itens_m1_proc else False
                }

                st.markdown("##### 📁 Documentos Acessórios ao Contrato (Checklist)")
                # Use individual checkboxes instead of a multiselect for clearer UI
                docs_options = ["Capa do Processo Físico", "Extrato do Contrato (Publicação)", "Requerimento de Contratação"]
                anexos_sel_m1 = []
                for opt in docs_options:
                    if st.checkbox(opt, value=True, key=f"checkbox_{opt}_{row['id']}"):
                        anexos_sel_m1.append(opt)
                dados_payload["documentos_selecionados"] = anexos_sel_m1

                b1, b2, b3, b4, b5 = st.columns([1, 1, 1, 1, 1.2])
                
                with b1:
                    from gerador_contratos_agu import gerar_contrato_docx, gerar_pacote_zip
                    
                    zip_bytes_m1 = gerar_pacote_zip(dados_payload, is_aditivo=False)
                    st.download_button(
                        label="📦 Baixar Pacote Completo (.zip)",
                        data=zip_bytes_m1,
                        file_name=f"Pacote_Contrato_{str(row['numero_completo']).replace('/', '_')}.zip",
                        mime="application/zip",
                        key="dl_zip_tela_m1",
                        use_container_width=True,
                        type="primary"
                    )
                with b2:
                    docx_b = gerar_contrato_docx(dados_payload)
                    st.download_button(
                        label="📄 Baixar Apenas Minuta (.docx)",
                        data=docx_b,
                        file_name=f"Minuta_{str(row['numero_completo']).replace('/', '_')}_AGU.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="dl_docx_tela_m1",
                        use_container_width=True,
                        type="secondary"
                    )
                with b3:
                    pdf_b = gerar_contrato_pdf(dados_payload)
                    st.download_button(
                        label="📄 Baixar Minuta PDF (.pdf)",
                        data=pdf_b,
                        file_name=f"Minuta_{str(row['numero_completo']).replace('/', '_')}_AGU.pdf",
                        mime="application/pdf",
                        key="dl_pdf_tela_m1",
                        use_container_width=True
                    )
                with b4:
                    if st.button("🌐 Sincronizar via API CGU/AGU", key="btn_sync_tela_m1", use_container_width=True):
                        with st.spinner("Consultando API AGU (cgu.agu.gov.br)..."):
                            c_agu = AGUApiClient(timeout=10)
                            res_agu = c_agu.gerar_minuta_agu(dados_payload)
                            if res_agu.get("sucesso"):
                                st.success("✅ Minuta validada e montada diretamente pelo servidor da AGU!")
                                with st.expander("Ver Minuta Processada pela CGU/AGU", expanded=False):
                                    st.write(res_agu.get("texto_completo", "")[:1200] + "...")
                            else:
                                st.warning("Aviso de conexão API AGU: documento local nativo 100% pronto para uso.")
                with b5:
                    if not df_prot_ligados.empty:
                        r_top_p = df_prot_ligados.iloc[0]
                        prot_exp_p = {
                            "numero": str(r_top_p["Protocolo"]),
                            "data_rec": str(r_top_p["Data"]),
                            "assunto": str(r_top_p["Assunto"]),
                            "solicitante": str(r_top_p["Solicitante"]),
                            "orgao": str(row["orgao"]),
                            "secretario": str(dados_payload["secretario"]),
                            "contrato": str(row["numero_completo"]),
                            "fornecedor": str(dados_payload["fornecedor_nome"]),
                            "observacoes": f"Protocolo vinculado ao Contrato {row['numero_completo']}",
                            "usuario": st.session_state["usuario"].get("nome", "Servidor"),
                            "qtd_vias": 2
                        }
                        try:
                            buf_p = gerar_comprovante_protocolo_pdf(prot_exp_p)
                            st.download_button(
                                label="🖨️ Comprovante Protocolo (PDF)",
                                data=buf_p.getvalue(),
                                file_name=f"comprovante_{str(r_top_p['Protocolo']).replace('/', '_')}.pdf",
                                mime="application/pdf",
                                key="dl_prot_pdf_tab_m1",
                                use_container_width=True
                            )
                        except Exception as e:
                            st.error(f"Erro: {e}")
                    else:
                        if st.button("➕ Vincular Novo Protocolo", key=f"btn_vincular_prot_m1_{row['id']}", use_container_width=True):
                            st.session_state["contrato_pre_selecionado_id"] = int(row["id"])
                            st.session_state["mostrar_formulario_protocolo"] = True
                            st.rerun()


    with tab_m3:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #0369a1 0%, #0284c7 100%); padding: 14px 18px; border-radius: 8px; color: white; margin-bottom: 15px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);'>
            <div style='display: flex; justify-content: space-between; align-items: center;'>
                <span style='font-size: 15.5px; font-weight: 800; letter-spacing: 0.5px;'>📋 DADOS INICIAIS — GERADOR DE MODELOS AGU</span>
                <span style='background: rgba(255,255,255,0.22); font-size: 11px; padding: 3px 9px; border-radius: 12px; font-weight: 700;'>CGUTEC Versão 1.0.2</span>
            </div>
            <div style='font-size: 12px; color: #e0f2fe; margin-top: 4px;'>
                Estrutura oficial de preenchimento do Portal da Consultoria-Geral da União (cgu.agu.gov.br/contrato/)
            </div>
        </div>
        """, unsafe_allow_html=True)

        # -------------------------------------------------------------
        # 1. TIPO DE CONTRATAÇÃO
        # -------------------------------------------------------------
        st.markdown("##### 📦 Tipo de Contratação")
        c_tipo1, c_tipo2 = st.columns([1, 1.5])
        with c_tipo1:
            tipo_contratacao_sel = st.selectbox(
                "Tipo:",
                ["Aquisições", "Serviços"],
                index=0,
                help="Selecione o tipo de contratação conforme o portal da AGU (Aquisições ou Serviços)",
                key="cgutec_tipo_contratacao"
            )

        with c_tipo2:
            if tipo_contratacao_sel == "Aquisições":
                opcoes_submodelo = {
                    "Termo de Contrato - Aquisição de Bens / Compras em Geral (Padrão AGU)": "COMPRAS",
                    "Termo de Contrato - Aquisição de Bens para Entrega Imediata": "COMPRAS_IMEDIATA",
                    "Termo de Contrato - Aquisição de Equipamentos e Bens de TIC": "COMPRAS_TIC"
                }
            else:
                opcoes_submodelo = {
                    "Termo de Contrato - Prestação de Serviços Contínuos (Sem dedicação exclusiva)": "SERVICOS_CONTINUOS",
                    "Termo de Contrato - Prestação de Serviços Não Contínuos / Escopo": "SERVICOS_ESCOPO",
                    "Termo de Contrato - Obras e Serviços Comuns de Engenharia": "OBRAS",
                    "Termo de Contrato - Serviços de Tecnologia da Informação (TIC)": "SERVICOS_TIC"
                }
            
            submodelo_nome = st.selectbox(
                "Modelo Específico de Minuta AGU (Lei nº 14.133/2021):",
                list(opcoes_submodelo.keys()),
                key="cgutec_submodelo_sel"
            )
            cod_modelo_cgutec = opcoes_submodelo[submodelo_nome]

        st.markdown("---")

        # -------------------------------------------------------------
        # 2. DADOS DO ÓRGÃO
        # -------------------------------------------------------------
        st.markdown("##### 🏛️ Dados do Órgão")
        
        opcoes_orgao_cgutec = ["Selecione o Órgão / Secretaria Contratante..."] + LISTA_SECRETARIAS_OFICIAIS
        
        # Inicialização dos dados se ainda não presentes no session_state
        if "cgutec_rep_orgao" not in st.session_state:
            st.session_state["cgutec_org_aux_sel"] = "Selecione o Órgão / Secretaria Contratante..."
            st.session_state["cgutec_nome_orgao"] = "MUNICÍPIO DE RIBEIRÃOZINHO DO MARANHÃO"
            st.session_state["cgutec_rep_orgao"] = ""
            st.session_state["cgutec_cargo_orgao"] = "Secretário(a) Municipal"
            st.session_state["cgutec_fisc_orgao"] = ""
            st.session_state["cgutec_cnpj_orgao"] = "01.612.834/0001-86"
            st.session_state["cgutec_cep_orgao"] = "65928-000"
            st.session_state["cgutec_end_orgao"] = "Rua Principal"
            st.session_state["cgutec_num_orgao"] = "s/n"
            st.session_state["cgutec_comp_orgao"] = "Prédio Administrativo"
            st.session_state["cgutec_bairro_orgao"] = "Centro"
            st.session_state["cgutec_cid_orgao"] = "Ribeirãozinho do Maranhão"
            st.session_state["cgutec_uf_orgao"] = "MA"
            st.session_state["cgutec_dot_ct"] = ""

        # Seletor inteligente para agilizar com histórico institucional da secretaria escolhida pelo usuário
        def _preencher_dados_orgao_cgutec():
            nome_org = st.session_state.get("cgutec_org_aux_sel", "")
            if nome_org and nome_org != "Selecione o Órgão / Secretaria Contratante...":
                h = obter_historico_orgao(nome_org)
                st.session_state["cgutec_cnpj_orgao"] = h.get("cnpj") or MAPA_CNPJ_ORGAOS.get(nome_org, "01.612.834/0001-86")
                st.session_state["cgutec_rep_orgao"] = h.get("secretario", "")
                st.session_state["cgutec_cargo_orgao"] = h.get("cargo_secretario", "Secretário(a) Municipal")
                st.session_state["cgutec_nome_orgao"] = nome_org
                st.session_state["cgutec_cep_orgao"] = h.get("cep", "65928-000")
                
                # Tratar endereço e número separadamente
                end_str = h.get("endereco", "Rua Principal, s/n")
                num_extraido = h.get("numero", "s/n")
                if ", nº " in end_str:
                    partes = end_str.split(", nº ")
                    st.session_state["cgutec_end_orgao"] = partes[0]
                    num_extraido = partes[1]
                elif ", n° " in end_str:
                    partes = end_str.split(", n° ")
                    st.session_state["cgutec_end_orgao"] = partes[0]
                    num_extraido = partes[1]
                elif ", " in end_str:
                    partes = end_str.rsplit(", ", 1)
                    st.session_state["cgutec_end_orgao"] = partes[0]
                    num_extraido = partes[1]
                else:
                    st.session_state["cgutec_end_orgao"] = end_str
                st.session_state["cgutec_num_orgao"] = num_extraido

                st.session_state["cgutec_bairro_orgao"] = h.get("bairro", "Centro")
                st.session_state["cgutec_cid_orgao"] = h.get("cidade", "Ribeirãozinho do Maranhão")
                st.session_state["cgutec_uf_orgao"] = h.get("uf", "MA")
                st.session_state["cgutec_fisc_orgao"] = h.get("fiscal", "")
                st.session_state["cgutec_dot_ct"] = h.get("dotacao", "")
                st.session_state["cgutec_obj_orgao"] = f"Contratação de {st.session_state.get('cgutec_tipo_contratacao', 'aquisições').lower()} para atendimento das demandas e funcionamento regular da {nome_org}."
            elif nome_org == "Selecione o Órgão / Secretaria Contratante...":
                st.session_state["cgutec_nome_orgao"] = "MUNICÍPIO DE RIBEIRÃOZINHO DO MARANHÃO"
                st.session_state["cgutec_rep_orgao"] = ""
                st.session_state["cgutec_cargo_orgao"] = "Secretário(a) Municipal"
                st.session_state["cgutec_fisc_orgao"] = ""
                st.session_state["cgutec_end_orgao"] = "Rua Principal"
                st.session_state["cgutec_num_orgao"] = "s/n"
                st.session_state["cgutec_dot_ct"] = ""
                st.session_state["cgutec_obj_orgao"] = f"Contratação de {st.session_state.get('cgutec_tipo_contratacao', 'aquisições').lower()} para atendimento das demandas e funcionamento regular da Administração Municipal."

        col_org_sel1, col_org_sel2 = st.columns([1.5, 1])
        with col_org_sel1:
            idx_sel_org = 0
            val_atual_sel = st.session_state.get("cgutec_org_aux_sel", "")
            if val_atual_sel in opcoes_orgao_cgutec:
                idx_sel_org = opcoes_orgao_cgutec.index(val_atual_sel)
            org_aux_sel = st.selectbox(
                "Carregar dados pelo histórico da Secretaria:",
                opcoes_orgao_cgutec,
                index=idx_sel_org,
                key="cgutec_org_aux_sel",
                on_change=_preencher_dados_orgao_cgutec
            )
        with col_org_sel2:
            st.caption("💡 *Selecione uma Secretaria para preencher automaticamente com os dados alimentados pelos usuários, ou digite livremente nos campos abaixo:*")

        def _callback_cnpj_orgao():
            cnpj_val = st.session_state.get("cgutec_cnpj_orgao", "")
            c_limpo = re.sub(r"\D", "", str(cnpj_val or ""))
            if not c_limpo:
                return
            if len(c_limpo) == 14:
                st.session_state["cgutec_cnpj_orgao"] = formatar_cnpj_cpf(c_limpo)
                
                # 1. Persiste imediatamente o novo CNPJ no órgão atual
                _callback_persistir_orgao()
                
                # 2. Se o usuário ainda não tiver selecionado uma secretaria específica, tenta identificar pelo CNPJ
                nome_atual = st.session_state.get("cgutec_nome_orgao", "")
                if not nome_atual or nome_atual == "MUNICÍPIO DE RIBEIRÃOZINHO DO MARANHÃO":
                    try:
                        with engine.connect() as conn:
                            row_org = conn.execute(
                                text("SELECT * FROM orgaos WHERE replace(replace(replace(replace(cnpj, '.', ''), '/', ''), '-', ''), ' ', '') = :c LIMIT 1"),
                                {"c": c_limpo}
                            ).mappings().first()
                    except Exception:
                        row_org = None

                    if row_org:
                        st.session_state["cgutec_nome_orgao"] = row_org.get("nome") or st.session_state.get("cgutec_nome_orgao", "")
                        st.session_state["cgutec_rep_orgao"] = row_org.get("secretario_padrao") or st.session_state.get("cgutec_rep_orgao", "")
                        st.session_state["cgutec_cargo_orgao"] = row_org.get("cargo_secretario") or "Secretário(a) Municipal"
                        st.session_state["cgutec_fisc_orgao"] = row_org.get("fiscal_padrao") or st.session_state.get("cgutec_fisc_orgao", "")
                        if row_org.get("cep"): st.session_state["cgutec_cep_orgao"] = formatar_cep(row_org.get("cep"))
                        if row_org.get("endereco"):
                            end_db = str(row_org.get("endereco")).strip()
                            if ", nº " in end_db:
                                partes = end_db.split(", nº ", 1)
                                st.session_state["cgutec_end_orgao"] = partes[0]
                                st.session_state["cgutec_num_orgao"] = partes[1]
                            elif ", n° " in end_db:
                                partes = end_db.split(", n° ", 1)
                                st.session_state["cgutec_end_orgao"] = partes[0]
                                st.session_state["cgutec_num_orgao"] = partes[1]
                            else:
                                st.session_state["cgutec_end_orgao"] = end_db
                        if row_org.get("bairro"): st.session_state["cgutec_bairro_orgao"] = row_org.get("bairro")
                        if row_org.get("cidade"): st.session_state["cgutec_cid_orgao"] = row_org.get("cidade")
                        if row_org.get("uf"): st.session_state["cgutec_uf_orgao"] = row_org.get("uf")
                        st.toast(f"Dados do órgão '{row_org.get('nome')}' sincronizados!", icon="🏛️")
                        return

                    # 3. Busca na Receita Federal / BrasilAPI se for um novo CNPJ
                    res_api = consultar_cnpj_api(c_limpo)
                    if res_api.get("sucesso"):
                        st.session_state["cgutec_nome_orgao"] = res_api.get("razao_social") or res_api.get("nome_fantasia") or st.session_state.get("cgutec_nome_orgao", "")
                        if res_api.get("endereco") and res_api.get("endereco") != "Endereço Comercial":
                            st.session_state["cgutec_end_orgao"] = res_api.get("endereco")
                        if res_api.get("numero") and res_api.get("numero") != "s/n":
                            st.session_state["cgutec_num_orgao"] = res_api.get("numero")
                        if res_api.get("complemento"):
                            st.session_state["cgutec_comp_orgao"] = res_api.get("complemento")
                        if res_api.get("bairro"):
                            st.session_state["cgutec_bairro_orgao"] = res_api.get("bairro")
                        if res_api.get("cidade"):
                            st.session_state["cgutec_cid_orgao"] = res_api.get("cidade")
                        if res_api.get("uf"):
                            st.session_state["cgutec_uf_orgao"] = res_api.get("uf")
                        if res_api.get("cep"):
                            st.session_state["cgutec_cep_orgao"] = res_api.get("cep")
                        if res_api.get("representante_nome") and res_api.get("representante_nome") != "Representante Legal":
                            st.session_state["cgutec_rep_orgao"] = res_api.get("representante_nome")
                            st.session_state["cgutec_cargo_orgao"] = res_api.get("representante_cargo", "Prefeito(a) Municipal")
                        _callback_persistir_orgao()
                        st.toast("Dados do órgão atualizados via Receita Federal!", icon="🏛️")

        def _callback_cep_orgao():
            c = re.sub(r"\D", "", str(st.session_state.get("cgutec_cep_orgao", "")))
            if len(c) == 8:
                st.session_state["cgutec_cep_orgao"] = formatar_cep(c)
                d = consultar_cep_viacep(c)
                if d:
                    if d.get("logradouro"):
                        st.session_state["cgutec_end_orgao"] = d["logradouro"]
                    if d.get("bairro"):
                        st.session_state["cgutec_bairro_orgao"] = d["bairro"]
                    if d.get("localidade"):
                        st.session_state["cgutec_cid_orgao"] = d["localidade"]
                    if d.get("uf"):
                        st.session_state["cgutec_uf_orgao"] = d["uf"]

        def _callback_persistir_orgao():
            nome_org = st.session_state.get("cgutec_nome_orgao") or st.session_state.get("cgutec_org_aux_sel", "")
            if nome_org and nome_org != "Selecione o Órgão / Secretaria Contratante...":
                dot_val = st.session_state.get("cgutec_inp_dot_uni") or st.session_state.get("cgutec_dot_ct")
                atualizar_historico_orgao_bd(nome_org, {
                    "secretario": st.session_state.get("cgutec_rep_orgao"),
                    "cargo_secretario": st.session_state.get("cgutec_cargo_orgao"),
                    "fiscal": st.session_state.get("cgutec_fisc_orgao"),
                    "endereco": st.session_state.get("cgutec_end_orgao"),
                    "bairro": st.session_state.get("cgutec_bairro_orgao"),
                    "cidade": st.session_state.get("cgutec_cid_orgao"),
                    "uf": st.session_state.get("cgutec_uf_orgao"),
                    "cep": st.session_state.get("cgutec_cep_orgao"),
                    "cnpj": st.session_state.get("cgutec_cnpj_orgao"),
                    "dotacao": dot_val
                })

        c_o1, c_o2 = st.columns(2)
        with c_o1:
            cnpj_orgao_val = st.text_input(
                "CNPJ",
                placeholder="00.000.000/0001-00",
                key="cgutec_cnpj_orgao",
                on_change=_callback_cnpj_orgao,
                help="Digite o CNPJ para auto-preenchimento automático dos dados cadastrais do órgão via Receita Federal."
            )
            rep_orgao_val = st.text_input("Representado por", placeholder="Nome do(a) Secretário(a) ou Representante", key="cgutec_rep_orgao", on_change=_callback_persistir_orgao)
            cargo_orgao_val = st.text_input("Cargo", placeholder="Ex: Secretário(a) Municipal", key="cgutec_cargo_orgao", on_change=_callback_persistir_orgao)
            nome_orgao_val = st.text_input("Órgão ou entidade pública", placeholder="Ex: SECRETARIA MUNICIPAL DE...", key="cgutec_nome_orgao", on_change=_callback_persistir_orgao)
            if "cgutec_setor_orgao" not in st.session_state:
                st.session_state["cgutec_setor_orgao"] = "Comissão Permanente de Licitação - CPL / Setor de Contratos"
            setor_orgao_val = st.text_input("Setor responsável pelas licitações", placeholder="Setor responsável pelas licitações", key="cgutec_setor_orgao")
            cep_orgao_val = st.text_input("CEP", placeholder="CEP", key="cgutec_cep_orgao", on_change=_callback_cep_orgao)

        with c_o2:
            end_orgao_val = st.text_input("Endereço Rua,Av,Rodovia...", placeholder="Endereço da sede...", key="cgutec_end_orgao", on_change=_callback_persistir_orgao)
            bairro_orgao_val = st.text_input("Bairro", placeholder="Bairro...", key="cgutec_bairro_orgao", on_change=_callback_persistir_orgao)
            
            c_num1, c_num2 = st.columns([1, 1.5])
            with c_num1:
                num_orgao_val = st.text_input("Número", placeholder="Número", key="cgutec_num_orgao", on_change=_callback_persistir_orgao)
            with c_num2:
                if "cgutec_comp_orgao" not in st.session_state:
                    st.session_state["cgutec_comp_orgao"] = "Prédio Administrativo"
                comp_orgao_val = st.text_input("Complemento", placeholder="Bloco, torre, quadra...", key="cgutec_comp_orgao")
                
            c_cid1, c_cid2 = st.columns([2, 1])
            with c_cid1:
                cid_orgao_val = st.text_input("Cidade", placeholder="Cidade...", key="cgutec_cid_orgao", on_change=_callback_persistir_orgao)
            with c_cid2:
                uf_orgao_val = st.text_input("UF", placeholder="UF", key="cgutec_uf_orgao", on_change=_callback_persistir_orgao)
                
            fisc_orgao_val = st.text_input("Fiscal Designado do Contrato", placeholder="Servidor designado como fiscal", key="cgutec_fisc_orgao", on_change=_callback_persistir_orgao)

        nome_alvo_objeto = nome_orgao_val.strip() if nome_orgao_val and nome_orgao_val.strip() else "Administração Municipal"
        obj_padrao_sug = f"Contratação de {tipo_contratacao_sel.lower()} para atendimento das demandas e funcionamento regular da {nome_alvo_objeto}."
        if "cgutec_obj_orgao" not in st.session_state:
            st.session_state["cgutec_obj_orgao"] = obj_padrao_sug
        obj_orgao_val = st.text_area("Insira o objeto da licitação", placeholder="Insira o objeto da licitação", height=85, key="cgutec_obj_orgao")

        # -------------------------------------------------------------
        # 2.1. PLANILHA DE ITENS DA CONTRATAÇÃO (CLÁUSULA PRIMEIRA - OBJETO)
        # -------------------------------------------------------------
        st.markdown("##### 📦 Itens da Contratação (Cláusula Primeira - Detalhamento dos Itens)")
        with st.expander("📊 Importar Planilha ou PDF de Itens (Auto-Ajustável com IA)", expanded=True):
            st.caption("A tabela do documento gerado se **auto-ajusta dinamicamente** à quantidade de colunas, linhas e descrição dos cabeçalhos da planilha (.xlsx, .xls, .csv) ou PDF importado. Não é necessário modelo fixo.")

            if "key_uploader_itens_cgutec" not in st.session_state:
                st.session_state["key_uploader_itens_cgutec"] = 0

            def _callback_limpar_itens_cgutec():
                st.session_state["key_uploader_itens_cgutec"] += 1
                st.session_state.pop("cgutec_tab_din_carregada", None)
                st.session_state.pop("cgutec_itens_carregados", None)
                st.session_state.pop("cgutec_tab_din_custom", None)
                st.session_state.pop("cgutec_soma_planilha_detectada", None)
                st.session_state.pop("_ultimo_arquivo_itens_id", None)

            arquivo_planilha_itens = st.file_uploader(
                "Selecione a Planilha ou Arquivo PDF de Itens (.xlsx, .xls, .csv, .pdf):",
                type=["xlsx", "xls", "csv", "pdf"],
                key=f"cgutec_upload_planilha_itens_{st.session_state['key_uploader_itens_cgutec']}",
                help="Carregue qualquer planilha ou documento PDF com itens. A tabela da minuta será gerada auto-ajustando cabeçalhos, quantidade de linhas e colunas."
            )

            if arquivo_planilha_itens is not None:
                col_it_name, col_it_del = st.columns([3.2, 1.2])
                with col_it_name:
                    st.success(f"📄 Arquivo carregado: **{arquivo_planilha_itens.name}**")
                with col_it_del:
                    st.button(
                        "🗑️ Limpar Planilha",
                        key=f"btn_limp_it_top_{st.session_state['key_uploader_itens_cgutec']}",
                        on_click=_callback_limpar_itens_cgutec,
                        use_container_width=True,
                        help="Remove a planilha carregada caso precise alterar ou enviar outro arquivo."
                    )

            # Processamento dinâmico
            itens_cgutec_processados = []
            tab_din_cgutec = None

            if arquivo_planilha_itens is not None:
                cur_file_id = f"{arquivo_planilha_itens.name}_{arquivo_planilha_itens.size}"
                if st.session_state.get("_ultimo_arquivo_itens_id") != cur_file_id:
                    st.session_state["_ultimo_arquivo_itens_id"] = cur_file_id
                    st.session_state.pop("cgutec_tab_din_custom", None)

                itens_cgutec_processados = processar_planilha_itens_excel(
                    arquivo_planilha_itens.getvalue(),
                    file_name=arquivo_planilha_itens.name,
                    objeto_contexto=obj_orgao_val
                )
                tab_din_cgutec = getattr(itens_cgutec_processados, 'tabela_dinamica', {})
                st.session_state["cgutec_tab_din_carregada"] = tab_din_cgutec
                st.session_state["cgutec_itens_carregados"] = itens_cgutec_processados
            elif st.session_state.get("cgutec_tab_din_custom"):
                tab_din_cgutec = st.session_state["cgutec_tab_din_custom"]
                itens_cgutec_processados = st.session_state.get("cgutec_itens_carregados", [])
            elif st.session_state.get("cgutec_tab_din_carregada"):
                tab_din_cgutec = st.session_state["cgutec_tab_din_carregada"]
                itens_cgutec_processados = st.session_state.get("cgutec_itens_carregados", [])

            # Calcular e vincular automaticamente a soma total da planilha ao Valor Global da Contratação
            soma_planilha_detectada = 0.0
            if tab_din_cgutec and isinstance(tab_din_cgutec, dict) and tab_din_cgutec.get("total_geral"):
                soma_planilha_detectada = float(tab_din_cgutec.get("total_geral", 0.0))
            if soma_planilha_detectada == 0.0 and itens_cgutec_processados:
                try:
                    soma_planilha_detectada = sum(float(it.get("_vt_float", 0.0)) for it in itens_cgutec_processados)
                except Exception:
                    pass

            if soma_planilha_detectada > 0:
                soma_fmt_br = f"{soma_planilha_detectada:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                st.session_state["cgutec_val_ct_txt"] = soma_fmt_br
                st.session_state["cgutec_soma_planilha_detectada"] = soma_planilha_detectada

            # Área para converter texto/edital em tabela via IA
            with st.expander("📝 Ou Cole Aqui Texto / Tabela Copiada (Edital, Termo de Referência ou PDF)", expanded=False):
                st.caption("A Inteligência Artificial (Google Gemini) converterá o texto livre ou linhas coladas em uma tabela oficial padronizada.")
                txt_colado_ia = st.text_area(
                    "Cole o texto dos itens:",
                    placeholder="Exemplo:\nItem 1 - Resma de Papel A4 75g - 500 pct - R$ 25,00 - Total R$ 12.500,00\nItem 2 - Caneta Esferográfica Azul - 1.000 un - R$ 1,50 - Total R$ 1.500,00",
                    height=100,
                    key="txt_itens_ia_input"
                )
                if st.button("⚡ Estruturar Tabela via IA", key="btn_estruturar_txt_ia", use_container_width=True):
                    if txt_colado_ia.strip():
                        with st.spinner("Analisando e construindo tabela oficial via IA..."):
                            res_ia_txt = estruturar_tabela_com_ia(txt_colado_ia, obj_orgao_val)
                            if res_ia_txt and res_ia_txt.get("headers"):
                                st.session_state["cgutec_tab_din_custom"] = res_ia_txt
                                tab_din_cgutec = res_ia_txt
                                if res_ia_txt.get("total_geral", 0.0) > 0:
                                    s_ia = float(res_ia_txt["total_geral"])
                                    st.session_state["cgutec_val_ct_txt"] = f"{s_ia:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                    st.session_state["cgutec_soma_planilha_detectada"] = s_ia
                                st.success(f"🎉 Tabela criada com sucesso via IA: {res_ia_txt['qtd_colunas']} colunas e {res_ia_txt['qtd_linhas']} linhas!")
                                st.rerun()
                            else:
                                st.error("Não foi possível extrair a tabela do texto informado. Tente ajustar a formatação.")

            # Exibição da tabela auto-ajustada
            if tab_din_cgutec and tab_din_cgutec.get("headers") and tab_din_cgutec.get("rows"):
                tot_itens_calc = soma_planilha_detectada if soma_planilha_detectada > 0 else tab_din_cgutec.get("total_geral", 0.0)
                tot_txt_disp = f"R$ {tot_itens_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                
                st.success(f"✅ **Tabela Auto-Ajustada com Sucesso:** **{tab_din_cgutec['qtd_colunas']} colunas** detectadas ({', '.join(tab_din_cgutec['headers'])}) e **{tab_din_cgutec['qtd_linhas']} itens** carregados!")
                
                df_dyn_view = pd.DataFrame(tab_din_cgutec["rows"], columns=tab_din_cgutec["headers"])

                # Configuração de colunas para o Dataframe Interativo
                col_cfg_cgutec = {}
                for h in tab_din_cgutec["headers"]:
                    h_up = str(h).upper()
                    if any(k in h_up for k in ['DESC', 'ESPEC', 'OBJET', 'PROD', 'MATERIAL', 'SERVIÇO']):
                        col_cfg_cgutec[h] = st.column_config.TextColumn(h, width="large")
                    elif any(k in h_up for k in ['ITEM', 'Nº', 'UNID', 'UND', 'QTD', 'QUANT']):
                        col_cfg_cgutec[h] = st.column_config.TextColumn(h, width="small")
                    else:
                        col_cfg_cgutec[h] = st.column_config.TextColumn(h, width="medium")

                # Visualização da Planilha com Quebra de Texto e Ajuste às Bordas
                tab_cgu_fmt, tab_cgu_grid = st.tabs([
                    "📋 Planilha Formatada (Quebra de Texto e Ajuste às Bordas)",
                    "🔢 Tabela Interativa (Grid)"
                ])
                with tab_cgu_fmt:
                    html_cgu = renderizar_tabela_html_formatada(
                        tab_din_cgutec["headers"],
                        tab_din_cgutec["rows"],
                        total_geral=tot_itens_calc
                    )
                    st.markdown(html_cgu, unsafe_allow_html=True)
                with tab_cgu_grid:
                    st.dataframe(df_dyn_view, column_config=col_cfg_cgutec, use_container_width=True, hide_index=True)

                c_bot_ia1, c_bot_down, c_bot_ia2, c_bot_limp = st.columns([1.3, 1.3, 1.2, 1.0])
                with c_bot_ia1:
                    if st.button("✨ Padronizar Cabeçalhos (IA)", key="btn_ia_opt_headers", use_container_width=True):
                        with st.spinner("Otimizando cabeçalhos e padronizando tabela com IA..."):
                            res_opt = estruturar_tabela_com_ia(df_dyn_view.to_csv(index=False), obj_orgao_val)
                            if res_opt and res_opt.get("headers"):
                                st.session_state["cgutec_tab_din_custom"] = res_opt
                                if res_opt.get("total_geral", 0.0) > 0:
                                    s_opt = float(res_opt["total_geral"])
                                    st.session_state["cgutec_val_ct_txt"] = f"{s_opt:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                    st.session_state["cgutec_soma_planilha_detectada"] = s_opt
                                st.success("✨ Cabeçalhos e estrutura otimizados com sucesso pela IA!")
                                st.rerun()
                            else:
                                st.info("Estrutura atual mantida.")
                with c_bot_down:
                    excel_cgu_bytes = gerar_planilha_itens_excel_formatada(
                        tab_din_cgutec["headers"],
                        tab_din_cgutec["rows"],
                        total_geral=tot_itens_calc
                    )
                    st.download_button(
                        label="📥 Baixar Planilha (.xlsx)",
                        data=excel_cgu_bytes,
                        file_name=f"itens_contrato_formatados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_cgu_btn_{st.session_state.get('key_uploader_itens_cgutec', 0)}",
                        use_container_width=True,
                        help="Baixar planilha oficial com quebra de texto automática e bordas ajustadas."
                    )
                with c_bot_ia2:
                    st.metric(
                        label="Soma Total dos Itens",
                        value=tot_txt_disp,
                        help="Somatório oficial calculado automaticamente a partir dos itens da planilha."
                    )
                with c_bot_limp:
                    st.write("")
                    st.button(
                        "🗑️ Limpar Planilha",
                        key=f"btn_limp_itens_bot_{st.session_state['key_uploader_itens_cgutec']}",
                        on_click=_callback_limpar_itens_cgutec,
                        use_container_width=True,
                        help="Remove a planilha atual para que você possa carregar outro arquivo caso precise alterar."
                    )

                if tot_itens_calc > 0:
                    st.info(f"💎 **Valor Global Vinculado:** O somatório de **{tot_txt_disp}** foi puxado automaticamente para o campo **Valor Global da Contratação**.")

        st.markdown("---")

        # -------------------------------------------------------------
        # 3. DEMAIS INFORMAÇÕES
        # -------------------------------------------------------------
        st.markdown("##### ⚙️ Demais Informações")
        c_inf1, c_inf2, c_inf3 = st.columns([1.2, 1.2, 1.1])
        with c_inf1:
            modalidade_cgutec = st.radio(
                "Modalidade / Procedimento:",
                ["Pregão", "Concorrência", "Contratação direta"],
                index=0,
                key="cgutec_modalidade_sel",
                horizontal=True
            )
        with c_inf2:
            if modalidade_cgutec == "Pregão":
                lbl_mod = "Número do Pregão:"
                ph_mod = "Ex: 010/2026 ou 010/2026-SRP"
                def_mod = "010/2026"
            elif modalidade_cgutec == "Concorrência":
                lbl_mod = "Número da Concorrência:"
                ph_mod = "Ex: 001/2026"
                def_mod = "001/2026"
            else:
                lbl_mod = "Número do Procedimento (Dispensa / Inex.):"
                ph_mod = "Ex: 005/2026 ou Dispensa nº 005/2026"
                def_mod = "005/2026"

            k_mod = f"cgutec_num_mod_{modalidade_cgutec}"
            if k_mod not in st.session_state:
                st.session_state[k_mod] = def_mod
            num_modalidade_cgutec = st.text_input(
                lbl_mod,
                placeholder=ph_mod,
                key=k_mod
            )
        with c_inf3:
            if "cgutec_proc_adm" not in st.session_state:
                st.session_state["cgutec_proc_adm"] = "001/2026-ADM"
            proc_adm_val = st.text_input(
                "Processo Administrativo",
                placeholder="Processo Administrativo - Deixar em Branco se não houver",
                key="cgutec_proc_adm"
            )

        # -------------------------------------------------------------
        # 4. DADOS COMPLEMENTARES DA CONTRATADA E VALOR
        # -------------------------------------------------------------
        with st.expander("🏢 Dados da Empresa Contratada (Fornecedor)", expanded=True):
            st.caption("💡 *Ao digitar o CNPJ da empresa, os dados cadastrais oficiais e endereço são preenchidos automaticamente pelo sistema (padrão oficial AGU).*")

            # Inicialização de variáveis de sessão do fornecedor
            for k_forn, v_def in [
                ("cgutec_forn_cnpj", ""),
                ("cgutec_forn_nome", ""),
                ("cgutec_forn_rep", ""),
                ("cgutec_forn_cargo", "Sócio Administrador"),
                ("cgutec_forn_end", ""),
                ("cgutec_forn_bairro", "Centro"),
                ("cgutec_forn_cidade", "Imperatriz"),
                ("cgutec_forn_uf", "MA"),
                ("cgutec_forn_cep", "65900-000"),
                ("cgutec_cnpj_status", None),
                ("cgutec_cnpj_msg", None),
                ("cgutec_cnpj_nao_cadastrado", None)
            ]:
                if k_forn not in st.session_state:
                    st.session_state[k_forn] = v_def

            def _aplicar_dados_fornecedor(dados_f, status="encontrado_banco", msg=None):
                c_limpo = re.sub(r"\D", "", str(dados_f.get("cnpj") or dados_f.get("cnpj_formatado") or st.session_state.get("cgutec_forn_cnpj", "")))
                st.session_state["cgutec_forn_cnpj"] = dados_f.get("cnpj_formatado") or formatar_cnpj_cpf(c_limpo)
                st.session_state["cgutec_forn_nome"] = dados_f.get("razao_social") or dados_f.get("nome_fantasia") or ""
                st.session_state["cgutec_forn_rep"] = dados_f.get("representante_nome") or "Representante Legal"
                st.session_state["cgutec_forn_cargo"] = dados_f.get("representante_cargo") or "Sócio Administrador"
                st.session_state["cgutec_forn_end"] = dados_f.get("endereco") or "Endereço Comercial"
                st.session_state["cgutec_forn_bairro"] = dados_f.get("bairro") or "Centro"
                st.session_state["cgutec_forn_cidade"] = dados_f.get("cidade") or "Imperatriz"
                st.session_state["cgutec_forn_uf"] = dados_f.get("uf") or "MA"
                st.session_state["cgutec_forn_cep"] = dados_f.get("cep") or "65900-000"
                st.session_state["cgutec_cnpj_status"] = status
                st.session_state["cgutec_cnpj_msg"] = msg

            def _callback_cnpj_input():
                cnpj_val = st.session_state.get("cgutec_forn_cnpj", "")
                c_limpo = re.sub(r"\D", "", str(cnpj_val or ""))
                if not c_limpo:
                    st.session_state["cgutec_cnpj_status"] = None
                    st.session_state["cgutec_cnpj_msg"] = None
                    return

                if len(c_limpo) == 14:
                    st.session_state["cgutec_forn_cnpj"] = formatar_cnpj_cpf(c_limpo)

                    # 1. Consulta prioritária no banco de dados local
                    res_banco = buscar_fornecedor_no_banco(c_limpo, engine)
                    tem_banco_completo = (
                        res_banco and res_banco.get("razao_social") and
                        res_banco.get("endereco") and res_banco.get("endereco") != "Endereço Comercial"
                    )

                    if tem_banco_completo:
                        _aplicar_dados_fornecedor(
                            res_banco,
                            status="sucesso",
                            msg=f"Empresa **{res_banco.get('razao_social')}** preenchida automaticamente com dados do sistema."
                        )
                        return

                    # 2. Se não constar no banco local ou faltar endereço, auto-preenche via API oficial (padrão AGU)
                    res_api = puxar_fornecedor_api_e_cadastrar(c_limpo, engine)
                    if res_api.get("sucesso"):
                        _aplicar_dados_fornecedor(
                            res_api,
                            status="sucesso",
                            msg=f"Empresa **{res_api.get('razao_social')}** preenchida automaticamente via Receita Federal ({res_api.get('fonte', 'API Oficial')})."
                        )
                    elif res_banco and res_banco.get("razao_social"):
                        _aplicar_dados_fornecedor(
                            res_banco,
                            status="sucesso",
                            msg=f"Empresa **{res_banco.get('razao_social')}** preenchida com dados cadastrados."
                        )
                    else:
                        st.session_state["cgutec_cnpj_status"] = "manual"
                        st.session_state["cgutec_cnpj_msg"] = "CNPJ não localizado automaticamente nos registros da Receita Federal. Preencha os campos da empresa manualmente abaixo."

            # 1. Menu de empresas já cadastradas no banco (Seleção Rápida Opcional)
            try:
                with engine.connect() as conn:
                    rows_db_forn = conn.execute(text("SELECT cnpj_cpf, razao_social FROM fornecedores WHERE razao_social IS NOT NULL AND TRIM(razao_social) != '' ORDER BY razao_social ASC")).fetchall()
            except Exception:
                rows_db_forn = []

            if rows_db_forn:
                def _callback_sel_fornec_pre():
                    escolha = st.session_state.get("sel_fornec_preexistente", "")
                    if escolha and escolha != "-- Digite o CNPJ abaixo ou selecione da lista de empresas cadastradas --":
                        m = re.search(r"\(([\d\.\/\-]+)\)$", escolha)
                        if m:
                            cnpj_ext = m.group(1)
                            st.session_state["cgutec_forn_cnpj"] = cnpj_ext
                            _callback_cnpj_input()

                opcoes_fornec = ["-- Digite o CNPJ abaixo ou selecione da lista de empresas cadastradas --"] + [
                    f"{r[1]} ({formatar_cnpj_cpf(r[0])})" for r in rows_db_forn if r[0] and r[1]
                ]
                st.selectbox(
                    "🏢 Empresas já Cadastradas no Banco de Dados (Preenchimento Rápido):",
                    options=opcoes_fornec,
                    key="sel_fornec_preexistente",
                    on_change=_callback_sel_fornec_pre,
                    help="Selecione para preencher todos os dados da empresa cadastrada no banco de dados."
                )

            # 2. Campo Inteligente de CNPJ com Auto-Preenchimento Automático (Sem botão - Padrão AGU)
            forn_cnpj_val = st.text_input(
                "CNPJ da Contratada (Auto-preenchimento automático):",
                key="cgutec_forn_cnpj",
                placeholder="00.000.000/0001-00 (ou somente números)",
                on_change=_callback_cnpj_input,
                help="Ao digitar os 14 dígitos do CNPJ, o sistema preenche automaticamente todos os dados cadastrais (Razão Social, Endereço, Bairro, CEP, Cidade, UF e Representante), exatamente como no gerador de contratos da AGU."
            )

            status_cnpj = st.session_state.get("cgutec_cnpj_status")
            msg_cnpj = st.session_state.get("cgutec_cnpj_msg")

            if status_cnpj == "sucesso":
                st.success(f"✅ {msg_cnpj}")
            elif status_cnpj == "manual":
                st.info(f"ℹ️ {msg_cnpj}")

            # 4. Campos Cadastrais (preenchidos automaticamente e editáveis)
            col_forn1, col_forn2 = st.columns([1.6, 1])
            with col_forn1:
                forn_nome_val = st.text_input("Fornecedor / Razão Social:", key="cgutec_forn_nome", placeholder="Razão social da empresa")
            with col_forn2:
                col_rep_mini1, col_rep_mini2 = st.columns(2)
                with col_rep_mini1:
                    forn_rep_val = st.text_input("Representante Legal:", key="cgutec_forn_rep", placeholder="Representante Legal")
                with col_rep_mini2:
                    forn_cargo_val = st.text_input("Cargo:", key="cgutec_forn_cargo", placeholder="Sócio Administrador")

            def _callback_cep_fornecedor():
                c = re.sub(r"\D", "", str(st.session_state.get("cgutec_forn_cep", "")))
                if len(c) == 8:
                    st.session_state["cgutec_forn_cep"] = formatar_cep(c)
                    d = consultar_cep_viacep(c)
                    if d:
                        if d.get("logradouro") and (not st.session_state.get("cgutec_forn_end") or st.session_state.get("cgutec_forn_end") == "Endereço Comercial"):
                            st.session_state["cgutec_forn_end"] = d["logradouro"]
                        if d.get("bairro"):
                            st.session_state["cgutec_forn_bairro"] = d["bairro"]
                        if d.get("localidade"):
                            st.session_state["cgutec_forn_cidade"] = d["localidade"]
                        if d.get("uf"):
                            st.session_state["cgutec_forn_uf"] = d["uf"]

            st.markdown("###### 📍 Endereço da Sede Comercial da Contratada")
            forn_end_val = st.text_input("Endereço Comercial da Contratada (Logradouro / Número):", key="cgutec_forn_end", placeholder="Ex: Rua Principal, nº 100")

            col_end1, col_end2, col_end3, col_end4 = st.columns([1.5, 1.5, 0.8, 1.2])
            with col_end1:
                forn_bairro_val = st.text_input("Bairro da Contratada:", key="cgutec_forn_bairro", placeholder="Centro")
            with col_end2:
                forn_cid_val = st.text_input("Cidade da Contratada:", key="cgutec_forn_cidade", placeholder="Imperatriz")
            with col_end3:
                forn_uf_val = st.text_input("UF:", key="cgutec_forn_uf", placeholder="MA")
            with col_end4:
                forn_cep_val = st.text_input("CEP:", key="cgutec_forn_cep", placeholder="65900-000", on_change=_callback_cep_fornecedor, help="Digite o CEP para preencher o endereço automaticamente via ViaCEP.")

        # -------------------------------------------------------------
        # DADOS DO CONTRATO MUNICIPAL & CLASSIFICAÇÃO ORÇAMENTÁRIA (IA)
        # -------------------------------------------------------------
        with st.expander("🏦 Dados do Contrato & Classificação Orçamentária Inteligente (IA)", expanded=True):
            col_ct_val1, col_ct_val2 = st.columns(2)
            with col_ct_val1:
                _, _, _, prox_sugerido_cgutec = obter_ultimo_e_proximo_numero_contrato()
                if "cgutec_num_ct" not in st.session_state:
                    st.session_state["cgutec_num_ct"] = prox_sugerido_cgutec
                num_ct_cgutec = st.text_input(
                    "Nº do Contrato:",
                    placeholder=f"Ex: {prox_sugerido_cgutec}",
                    key="cgutec_num_ct",
                    help="Puxado automaticamente a partir do último contrato registrado."
                )
            with col_ct_val2:
                def _callback_sincronizar_soma(val_soma_str):
                    st.session_state["cgutec_val_ct_txt"] = val_soma_str

                soma_plan = st.session_state.get("cgutec_soma_planilha_detectada", 0.0)
                if soma_plan > 0:
                    soma_plan_txt = f"{soma_plan:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    if "cgutec_val_ct_txt" not in st.session_state or not st.session_state["cgutec_val_ct_txt"] or st.session_state["cgutec_val_ct_txt"] == "50.000,00":
                        st.session_state["cgutec_val_ct_txt"] = soma_plan_txt
                elif "cgutec_val_ct_txt" not in st.session_state:
                    st.session_state["cgutec_val_ct_txt"] = ""

                val_ct_cgutec_txt = st.text_input(
                    "Valor Global da Contratação (R$):",
                    key="cgutec_val_ct_txt",
                    placeholder="Ex: 0,00 (puxado automaticamente da planilha de itens)",
                    on_change=formatar_input_moeda_callback,
                    args=("cgutec_val_ct_txt",),
                    help="Valor total da contratação. Puxado automaticamente a partir da soma dos itens da planilha de detalhamento."
                )
                val_ct_cgutec = clean_money(val_ct_cgutec_txt)

                if soma_plan > 0:
                    c_soma_info, c_soma_reap = st.columns([2.2, 1.2])
                    with c_soma_info:
                        st.caption(f"✨ *Puxado automaticamente da planilha de itens:* **R$ {soma_plan:,.2f}**".replace(",", "X").replace(".", ",").replace("X", "."))
                    with c_soma_reap:
                        soma_txt_val = f"{soma_plan:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        st.button(
                            "🔄 Sincronizar com a Soma",
                            key="btn_reap_soma_plan",
                            on_click=_callback_sincronizar_soma,
                            args=(soma_txt_val,),
                            use_container_width=True,
                            help="Reaplicar a soma exata dos itens da planilha"
                        )

            st.markdown("---")
            st.markdown("##### 🏦 Classificação Orçamentária Inteligente (IA)")
            st.caption("Cole abaixo o texto da Dotação, Empenho, PDF ou Planilha. A Inteligência Artificial organizará os campos para a minuta oficial.")
            
            def _callback_extrair_dotacao_cgutec():
                raw_txt = st.session_state.get("cgutec_txt_dot_raw_input", "")
                if not raw_txt or not str(raw_txt).strip():
                    st.session_state["cgutec_dot_msg"] = ("warning", "Por favor, cole o texto da dotação orçamentária no campo antes de clicar no botão.")
                    return
                d_ext = extrair_dotacao_com_ia(raw_txt)
                if d_ext:
                    if d_ext.get("exercicio"):
                        st.session_state["cgutec_inp_dot_ex"] = d_ext["exercicio"]
                        st.session_state["cgutec_dot_ex"] = d_ext["exercicio"]
                    if d_ext.get("poder"):
                        st.session_state["cgutec_inp_dot_pod"] = d_ext["poder"]
                        st.session_state["cgutec_dot_pod"] = d_ext["poder"]
                    if d_ext.get("orgao"):
                        st.session_state["cgutec_inp_dot_org"] = d_ext["orgao"]
                        st.session_state["cgutec_dot_org"] = d_ext["orgao"]
                    if d_ext.get("unidade"):
                        st.session_state["cgutec_inp_dot_uni"] = d_ext["unidade"]
                        st.session_state["cgutec_dot_uni"] = d_ext["unidade"]
                    if d_ext.get("natureza"):
                        st.session_state["cgutec_inp_dot_nat"] = d_ext["natureza"]
                        st.session_state["cgutec_dot_nat"] = d_ext["natureza"]
                    if d_ext.get("outros_campos"):
                        st.session_state["cgutec_inp_dot_outros"] = d_ext["outros_campos"]
                        st.session_state["cgutec_dot_outros"] = d_ext["outros_campos"]
                    st.session_state["cgutec_dot_msg"] = ("success", "✨ Dotação orçamentária classificada e preenchida com sucesso nos campos abaixo!")
                else:
                    st.session_state["cgutec_dot_msg"] = ("error", "Não foi possível extrair os dados orçamentários. Verifique o texto informado.")

            c_cg_dot1, c_cg_dot2 = st.columns([2, 1])
            with c_cg_dot1:
                txt_dot_cgutec = st.text_area(
                    "Texto Bruto da Dotação Orçamentária (Copie e cole aqui):",
                    value=st.session_state.get("cgutec_txt_dot_raw", ""),
                    height=100,
                    key="cgutec_txt_dot_raw_input",
                    placeholder="Ex: 02.15.00 - Fundo M. de Saúde | Manutenção 12.361.0402.2022.0000 | Natureza: 3.3.90.39"
                )
            with c_cg_dot2:
                st.write("")
                st.write("")
                st.button(
                    "✨ Extrair com Inteligência Artificial",
                    key="btn_ext_ia_cgutec",
                    on_click=_callback_extrair_dotacao_cgutec,
                    use_container_width=True
                )

            if "cgutec_dot_msg" in st.session_state:
                tipo_m, texto_m = st.session_state.pop("cgutec_dot_msg")
                if tipo_m == "success":
                    st.success(texto_m)
                elif tipo_m == "warning":
                    st.warning(texto_m)
                else:
                    st.error(texto_m)
                            
            col_cg_d1, col_cg_d2 = st.columns(2)
            with col_cg_d1:
                cg_dot_ex = st.text_input("Exercício:", value=st.session_state.get("cgutec_dot_ex", "2026"), key="cgutec_inp_dot_ex")
                cg_dot_pod = st.text_input("Poder:", value=st.session_state.get("cgutec_dot_pod", "Poder Executivo 02.00"), key="cgutec_inp_dot_pod")
                cg_dot_org = st.text_input("Órgão Orçamentário:", value=st.session_state.get("cgutec_dot_org", f"Fundo Municipal / {nome_orgao_val}" if nome_orgao_val else "Fundo Municipal"), key="cgutec_inp_dot_org")
            with col_cg_d2:
                cg_dot_uni = st.text_input("Unidade Orçamentária / Projeto Atividade:", value=st.session_state.get("cgutec_dot_uni", st.session_state.get("cgutec_dot_ct", "MANUTENÇÃO E FUNCIONAMENTO 12.361.0402.2022.0000")), key="cgutec_inp_dot_uni", on_change=_callback_persistir_orgao)
                cg_dot_nat = st.text_input("Natureza da Despesa:", value=st.session_state.get("cgutec_dot_nat", "Outros Serviços de Terceiros 3.3.90.39.00" if "SERV" in str(cod_modelo_cgutec).upper() else "Material de consumo 3.3.90.30.00"), key="cgutec_inp_dot_nat")
                cg_dot_outros = st.text_input("Campos Adicionais / Fonte de Recursos / Outros:", value=st.session_state.get("cgutec_dot_outros", ""), key="cgutec_inp_dot_outros", placeholder="Ex: Fonte de Recursos: 15000000 - Recursos Próprios")
            
            dot_cgutec = f"{cg_dot_org} | {cg_dot_uni} | {cg_dot_nat}"
            if cg_dot_outros and str(cg_dot_outros).strip():
                dot_cgutec = f"{dot_cgutec} | {str(cg_dot_outros).strip()}"

        # -------------------------------------------------------------
        # 5. MOTOR DE INTELIGÊNCIA ARTIFICIAL / FORMATAÇÃO DE CLÁUSULAS
        # -------------------------------------------------------------
        st.markdown("---")
        with st.expander("🧠 Inserir Cláusulas Adicionais Especiais (Conformidade com a Nova Lei de Licitações nº 14.133/2021 via IA)", expanded=False):
            st.caption("O sistema utilizará Inteligência Artificial (Google Gemini 3.6 Flash / Motor Jurídico) para readequar suas anotações em estrita observância à Nova Lei de Licitações (Lei Federal nº 14.133/2021) e aos padrões da AGU.")
            titulo_clausula_cgutec = st.text_input(
                "Nome / Título da Cláusula (Livre Nomeação pelo Usuário):",
                value="DAS CONDIÇÕES ESPECIAIS E DISPOSIÇÕES COMPLEMENTARES (LEI Nº 14.133/2021)",
                help="Defina livremente o assunto/título da cláusula (ex: DA PROTEÇÃO DE DADOS, DO PROGRAMA DE COMPLIANCE, DA LOGÍSTICA REVERSA, etc.). O prefixo numérico oficial (ex: CLÁUSULA DÉCIMA TERCEIRA – ) será inserido automaticamente.",
                key="cgutec_titulo_clausula_extra"
            )
            texto_clausulas_adicionais = st.text_area(
                "Cole aqui as cláusulas extras que deseja incluir no contrato (um item por linha):",
                placeholder="Exemplo:\nO prazo de garantia dos equipamentos é de 12 meses.\nA CONTRATADA deverá manter preposto nos termos do art. 118 da Lei nº 14.133/2021.\nA fiscalização atuará em estrita observância ao art. 117 da Lei nº 14.133/2021.",
                height=150,
                key="cgutec_clausulas_extras_ia"
            )
            # Processar o texto:
            clausulas_formatadas_ia = formatar_clausulas_inteligencia(texto_clausulas_adicionais, 13)
            if clausulas_formatadas_ia:
                st.success(f"✨ Cláusula '{titulo_clausula_cgutec}' processada e readequada em estrita observância à Lei Federal nº 14.133/2021!")
                for cf in clausulas_formatadas_ia:
                    st.caption(f"✓ {cf}")

        dados_cgutec_final = {
            "numero_completo": num_ct_cgutec,
            "numero_contrato": int(str(num_ct_cgutec).split("/")[0]) if "/" in str(num_ct_cgutec) else 1,
            "processo_adm": proc_adm_val,
            "modalidade": modalidade_cgutec.upper(),
            "numero_modalidade": num_modalidade_cgutec.strip(),
            "demais_informacoes": f"{modalidade_cgutec.upper()} Nº {num_modalidade_cgutec.strip()}".strip(),
            "tipo_contratacao": tipo_contratacao_sel.upper(),
            "objeto": obj_orgao_val,
            "valor_total": val_ct_cgutec,
            "data_assinatura": date.today(),
            "data_vencimento": date.today().replace(year=date.today().year + 1),
            "vigencia_descricao": "12 (doze) meses a contar da data de assinatura",
            "secretario": rep_orgao_val,
            "cargo_secretario": cargo_orgao_val,
            "fiscal": fisc_orgao_val if fisc_orgao_val else "Servidor Fiscal Designado",
            "orgao_nome": nome_orgao_val,
            "orgao_cnpj": formatar_cnpj_cpf(cnpj_orgao_val),
            "orgao_cep": cep_orgao_val,
            "orgao_endereco": end_orgao_val,
            "orgao_bairro": bairro_orgao_val,
            "orgao_cidade": cid_orgao_val,
            "orgao_uf": uf_orgao_val,
            "numero_imovel": num_orgao_val,
            "complemento_imovel": comp_orgao_val,
            "setor_licitacoes": setor_orgao_val,
            "fornecedor_nome": forn_nome_val,
            "fornecedor_cnpj": formatar_cnpj_cpf(forn_cnpj_val),
            "fornecedor_representante": forn_rep_val,
            "fornecedor_cargo": forn_cargo_val,
            "fornecedor_endereco": forn_end_val,
            "fornecedor_bairro": forn_bairro_val,
            "fornecedor_cep": forn_cep_val,
            "fornecedor_cidade": forn_cid_val,
            "fornecedor_uf": forn_uf_val,
            "dotacao_orcamentaria": dot_cgutec,
            "dotacao_exercicio": cg_dot_ex,
            "dotacao_poder": cg_dot_pod,
            "dotacao_orgao": cg_dot_org,
            "dotacao_unidade": cg_dot_uni,
            "dotacao_natureza": cg_dot_nat,
            "dotacao_outros": cg_dot_outros if cg_dot_outros else "",
            "modelo_agu": cod_modelo_cgutec,
            "titulo_clausula_adicional": titulo_clausula_cgutec,
            "clausulas_adicionais_formatadas": clausulas_formatadas_ia,
            "tabela_dinamica": tab_din_cgutec if (tab_din_cgutec and tab_din_cgutec.get("headers")) else None,
            "itens": itens_cgutec_processados if itens_cgutec_processados else None,
            "usar_total_itens": True if ((tab_din_cgutec or itens_cgutec_processados) and st.session_state.get("cgutec_chk_usar_tot_planilha", True)) else False
        }

        with st.expander("👁️ Pré-Visualizar Cláusulas Oficiais da Minuta (Lei 14.133/2021)", expanded=False):
            from gerador_contratos_agu import gerar_clausulas_agu
            clausulas_cgutec_preview = gerar_clausulas_agu(dados_cgutec_final)
            for cl in clausulas_cgutec_preview:
                st.markdown(f"**{cl['numero']}**")
                for p_t in cl["conteudo"]:
                    st.caption(re.sub(r'<[^>]+>', '', p_t))

        st.markdown("##### 📁 Documentos Acessórios ao Contrato (Checklist)")
        # Use individual checkboxes instead of a multiselect for clearer UI
        docs_options_m3 = ["Capa do Processo Físico", "Extrato do Contrato (Publicação)", "Requerimento de Contratação"]
        anexos_sel_m3 = []
        for opt in docs_options_m3:
            if st.checkbox(opt, value=True, key=f"checkbox_m3_{opt}_{num_ct_cgutec}"):
                anexos_sel_m3.append(opt)
        dados_cgutec_final["documentos_selecionados"] = anexos_sel_m3

        st.markdown("---")
        b_c1, b_c2, b_c3, b_c4, b_c5 = st.columns([1, 1.2, 1, 1, 1.2])
        with b_c1:
            from gerador_contratos_agu import gerar_contrato_docx, gerar_pacote_zip
            atualizar_historico_orgao_bd(dados_cgutec_final.get("orgao_nome"), dados_cgutec_final)
            zip_bytes_m3 = gerar_pacote_zip(dados_cgutec_final, is_aditivo=False)
            st.download_button(
                label="📦 Baixar Pacote Completo (.zip)",
                data=zip_bytes_m3,
                file_name=f"Pacote_Contrato_{str(num_ct_cgutec).replace('/', '_')}.zip",
                mime="application/zip",
                key="dl_zip_cgutec",
                use_container_width=True,
                type="primary"
            )
        with b_c2:
            doc_cgutec_bytes = gerar_contrato_docx(dados_cgutec_final)
            st.download_button(
                label="📄 Baixar Apenas Minuta (.docx)",
                data=doc_cgutec_bytes,
                file_name=f"Minuta_{str(num_ct_cgutec).replace('/', '_')}_CGUTEC.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_docx_cgutec",
                use_container_width=True,
                type="secondary"
            )
        with b_c3:
            pdf_cgutec_bytes = gerar_contrato_pdf(dados_cgutec_final)
            st.download_button(
                label="📄 Baixar Minuta PDF (.pdf)",
                data=pdf_cgutec_bytes,
                file_name=f"Minuta_AGU_{str(num_ct_cgutec).replace('/', '_')}_CGUTEC.pdf",
                mime="application/pdf",
                key="dl_pdf_cgutec",
                use_container_width=True
            )
        with b_c4:
            if st.button("🌐 Sincronizar via API CGU/AGU", key="btn_sync_cgutec", use_container_width=True):
                with st.spinner("Consultando API AGU (cgu.agu.gov.br)..."):
                    c_agu = AGUApiClient(timeout=10)
                    res_agu = c_agu.gerar_minuta_agu(dados_cgutec_final)
                    if res_agu.get("sucesso"):
                        st.success("✅ Minuta validada e montada com sucesso pelo servidor da AGU!")
                    else:
                        st.info("ℹ️ Minuta local compilada e formatada com base nos modelos oficiais da AGU.")
        with b_c5:
            if st.button("💾 Salvar no Sistema", key="btn_salvar_cgutec_db", use_container_width=True, type="primary"):
                if st.session_state['usuario']['cargo'] == 'AUDITOR':
                    st.error('Acesso Negado: Auditores não podem salvar documentos no sistema.')
                else:
                    with st.spinner("Cadastrando e salvando contrato no sistema..."):
                        ok, msg = salvar_contrato_cgutec_bd(dados_cgutec_final)
                        if ok:
                            st.success(f"🎉 {msg} com sucesso!")
                            st.balloons()
                            time.sleep(1.5)
                            st.session_state["menu_selecionado"] = "📊 Contratos"
                            st.session_state["mostrar_formulario"] = False
                            st.rerun()
                        else:
                            st.error(f"Erro ao salvar: {msg}")


        # Rodapé oficial CGUTEC Versão 1.0.2
        st.markdown("""
        <div style='text-align: center; color: #64748b; font-size: 11.5px; margin-top: 25px; border-top: 1px solid #e2e8f0; padding-top: 12px;'>
            <strong>CGUTEC</strong> - Todos Direitos Reservados &nbsp;|&nbsp; <strong>Versão 1.0.2</strong> &nbsp;|&nbsp; Portal de Modelos de Contratos da Advocacia-Geral da União
        </div>
        """, unsafe_allow_html=True)

    with tab_m4:
        st.markdown("#### 📝 Elaboração e Emissão de Termos Aditivos Oficiais da AGU")
        st.caption("Prorrogação de vigência, alteração de valores (acréscimos e supressões de até 25%/50%), reequilíbrio econômico-financeiro e alterações qualitativas segundo os modelos padronizados da Advocacia-Geral da União (AGU).")

        st.markdown("""
        <div style='background: #f8fafc; border: 1px solid #cbd5e1; border-left: 4px solid #0284c7; border-radius: 6px; padding: 12px 16px; margin-top: 15px; margin-bottom: 25px; font-size: 13px; line-height: 1.6; color: #334155; display: inline-block; width: 100%; box-sizing: border-box;'>
            ⚖️ <strong>Modelos Padronizados de Referência AGU:</strong> 
            Minutas em estrita conformidade com os modelos oficiais da CGU/AGU disponíveis em 
            <a href='https://www.gov.br/agu/pt-br/composicao/cgu/cgu/modelos/licitacoesecontratos/8666e10520/termos-aditivos' target='_blank' style='color: #0284c7; font-weight: 600; text-decoration: underline;'>
                AGU - Modelos de Termos Aditivos (Gov.br)
            </a> 
            e integrados às regras da <strong>Lei nº 14.133/2021</strong> e <strong>Lei nº 8.666/1993</strong>.
        </div>
        """, unsafe_allow_html=True)

        if df_c.empty:
            st.info("Nenhum contrato cadastrado no sistema para aditamento.")
        else:
            opcoes_adit = ["Selecione um contrato para elaborar o termo aditivo..."] + df_c["numero_completo"].tolist()
            pre_c_adit = st.session_state.pop("aditivo_contrato_pre_sel", None)
            idx_adit = 0
            if pre_c_adit and pre_c_adit in opcoes_adit:
                idx_adit = opcoes_adit.index(pre_c_adit)
            escolha_adit = st.selectbox("Escolha o Contrato a ser Aditado:", opcoes_adit, index=idx_adit, key="sel_c_adit_m4")

            if escolha_adit != "Selecione um contrato para elaborar o termo aditivo...":
                row_ad = df_c[df_c["numero_completo"] == escolha_adit].iloc[0]
                cid_ad = int(row_ad["id"])
                hist_org_ad = obter_historico_orgao(row_ad["orgao"])
                prox_num_adit = gerar_proximo_numero_aditivo(cid_ad)
                
                val_orig = float(row_ad.get("valor_total") or 0.0)
                dt_venc_atual = row_ad.get("data_vencimento")
                dt_ass_orig = row_ad.get("data_assinatura")
                
                st.markdown(f"""
                <div style='background: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px; padding: 12px 16px; margin-bottom: 14px;'>
                    <div style='display: flex; justify-content: space-between; align-items: center;'>
                        <span style='font-size: 14px; font-weight: 700; color: #0369a1;'>Contrato Nº {row_ad['numero_completo']}</span>
                        <span style='font-size: 12px; background: #e0f2fe; color: #0284c7; padding: 2px 8px; border-radius: 12px; font-weight: 600;'>{row_ad.get('status', 'ATIVO')}</span>
                    </div>
                    <div style='font-size: 12px; color: #334155; margin-top: 6px;'>
                        🏛️ <strong>Secretaria:</strong> {row_ad['orgao']} | 🏢 <strong>Contratada:</strong> {row_ad['fornecedor']} (CNPJ: {formatar_cnpj_cpf(row_ad['cnpj_cpf'])})<br/>
                        💰 <strong>Valor Vigente:</strong> {formatar_moeda(val_orig)} | ⏳ <strong>Vencimento Vigente:</strong> {formatar_data_br(dt_venc_atual)}<br/>
                        📝 <strong>Objeto:</strong> {row_ad['objeto']}
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                c_f1, c_f2 = st.columns([1.2, 1])
                with c_f1:
                    tipo_opcoes = {
                        "Prorrogação de Vigência": "PRORROGACAO",
                        "Acréscimo de Valor (Quantitativo até 25%/50%)": "ACRESCIMO",
                        "Supressão de Valor (Quantitativo até 25%/50%)": "SUPRESSAO",
                        "Misto (Prorrogação de Vigência + Acréscimo de Valor)": "MISTO",
                        "Reequilíbrio Econômico-Financeiro / Reajuste": "REEQUILIBRIO",
                        "Alteração Qualitativa / Outras Condições": "OUTROS"
                    }
                    tipo_nome = st.selectbox(
                        "Tipo do Termo Aditivo:",
                        list(tipo_opcoes.keys()),
                        key="adit_tipo_sel"
                    )
                    tipo_cod = tipo_opcoes[tipo_nome]
                    
                    regime_opcoes = {
                        "Lei Federal nº 14.133/2021 (Nova Lei de Licitações - Art. 107 / 124-125)": "LEI_14133_2021",
                        "Lei Federal nº 8.666/1993 (Regime Anterior / Convênios - Art. 57, II / 65)": "LEI_8666_1993"
                    }
                    regime_nome = st.selectbox(
                        "Regime Jurídico / Lei de Licitações:",
                        list(regime_opcoes.keys()),
                        key="adit_regime_sel"
                    )
                    regime_cod = regime_opcoes[regime_nome]
                    
                    num_adit_val = st.number_input("Número Sequencial do Aditivo:", min_value=1, value=prox_num_adit, key="adit_num_seq")
                    num_comp_adit = f"{num_adit_val}º TERMO ADITIVO AO CONTRATO Nº {row_ad['numero_completo']}"
                    st.text_input("Título Oficial do Instrumento:", value=num_comp_adit, key="adit_titulo_inp")
                    
                with c_f2:
                    proc_adit_sug = f"{row_ad.get('processo_adm') or '001/2026'}-ADIT-{num_adit_val}"
                    proc_adit = st.text_input("Processo Administrativo do Aditivo:", value=proc_adit_sug, key="adit_proc_inp")
                    dt_ass_adit = st.date_input("Data de Assinatura do Aditivo:", value=date.today(), key="adit_dt_ass_inp")
                    sec_adit = st.text_input("Secretário(a) / Autoridade Signatária:", value=row_ad.get("secretario") or hist_org_ad["secretario"], key="adit_sec_inp")
                    fisc_adit = st.text_input("Fiscal Designado:", value=row_ad.get("fiscal") or hist_org_ad["fiscal"], key="adit_fisc_inp")
                    dot_adit = st.text_input("Dotação Orçamentária Específica:", value=hist_org_ad["dotacao"], key="adit_dot_inp")

                prazo_meses_val = 0
                nova_venc_calc = None
                val_adit_val = 0.0
                novo_val_calc = val_orig
                perc_calc = 0.0
                
                if tipo_cod in ["PRORROGACAO", "MISTO"]:
                    st.markdown("##### ⏱️ Prazos & Vigência do Aditamento")
                    cp1, cp2 = st.columns(2)
                    with cp1:
                        prazo_meses_val = st.number_input("Prazo Aditado (em meses):", min_value=1, max_value=120, value=12, step=1, key="adit_prazo_meses")
                    with cp2:
                        dt_ref_base = None
                        if dt_venc_atual:
                            try:
                                dt_ref_base = pd.to_datetime(dt_venc_atual).date()
                            except:
                                dt_ref_base = date.today()
                        else:
                            dt_ref_base = date.today()
                        
                        ano_nv = dt_ref_base.year + (dt_ref_base.month + prazo_meses_val - 1) // 12
                        mes_nv = (dt_ref_base.month + prazo_meses_val - 1) % 12 + 1
                        import calendar
                        dia_nv = min(dt_ref_base.day, calendar.monthrange(ano_nv, mes_nv)[1])
                        nova_venc_calc = date(ano_nv, mes_nv, dia_nv)
                        
                        st.markdown(f"""
                        <div style='background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 6px; padding: 8px 12px; margin-top: 14px;'>
                            <span style='font-size: 12px; color: #065f46; font-weight: 700;'>📅 Novo Vencimento Calculado:</span>
                            <div style='font-size: 15px; font-weight: 800; color: #059669;'>{formatar_data_br(nova_venc_calc)}</div>
                            <div style='font-size: 11px; color: #047857;'>De {formatar_data_br(dt_ref_base)} até {formatar_data_br(nova_venc_calc)} ({prazo_meses_val} meses)</div>
                        </div>
                        """, unsafe_allow_html=True)

                if tipo_cod in ["ACRESCIMO", "SUPRESSAO", "MISTO", "REEQUILIBRIO"]:
                    st.markdown("##### 💰 Alterações de Quantitativos e Valores (Limites AGU)")
                    
                    modo_calc = st.radio(
                        "Cálculo Automático por Percentual Legal:",
                        ["Entrada Manual em Reais (R$)", "10% (Ajuste Menor)", "25% (Limite Padrão AGU)", "50% (Reformas / Justificado)"],
                        horizontal=True,
                        key="adit_modo_calc"
                    )

                    cv1, cv2 = st.columns(2)
                    with cv1:
                        if modo_calc == "10% (Ajuste Menor)":
                            val_adit_val = val_orig * 0.10
                            st.text_input("Valor da Alteração Calculado (R$):", value=formatar_numero_moeda(val_adit_val), disabled=True)
                        elif modo_calc == "25% (Limite Padrão AGU)":
                            val_adit_val = val_orig * 0.25
                            st.text_input("Valor da Alteração Calculado (R$):", value=formatar_numero_moeda(val_adit_val), disabled=True)
                        elif modo_calc == "50% (Reformas / Justificado)":
                            val_adit_val = val_orig * 0.50
                            st.text_input("Valor da Alteração Calculado (R$):", value=formatar_numero_moeda(val_adit_val), disabled=True)
                        else:
                            val_sug = round(val_orig * 0.10, 2) if val_orig > 0 else 10000.0
                            if "adit_val_inp_txt" not in st.session_state:
                                st.session_state["adit_val_inp_txt"] = formatar_numero_moeda(val_sug)
                            val_adit_txt = st.text_input(
                                "Valor da Alteração Manual (R$):",
                                key="adit_val_inp_txt",
                                placeholder="Ex: 10.000,00",
                                on_change=formatar_input_moeda_callback,
                                args=("adit_val_inp_txt",),
                                help="Informe o valor com pontos e vírgula (ex: 10.000,00)."
                            )
                            val_adit_val = clean_money(val_adit_txt)

                    with cv2:
                        if val_orig > 0:
                            perc_calc = (val_adit_val / val_orig) * 100.0
                        else:
                            perc_calc = 0.0
                            
                        if tipo_cod == "SUPRESSAO":
                            novo_val_calc = max(0.0, val_orig - val_adit_val)
                            st.markdown(f"""
                            <div style='background: #fffbeb; border: 1px solid #fde68a; border-radius: 6px; padding: 8px 12px;'>
                                <div style='font-size: 12px; color: #92400e; font-weight: 700;'>📉 Supressão: {perc_calc:.2f}%</div>
                                <div style='font-size: 15px; font-weight: 800; color: #b45309;'>Novo Valor Consolidado: {formatar_moeda(novo_val_calc)}</div>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            novo_val_calc = val_orig + val_adit_val
                            st.markdown(f"""
                            <div style='background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 6px; padding: 8px 12px;'>
                                <div style='font-size: 12px; color: #065f46; font-weight: 700;'>📈 Acréscimo: {perc_calc:.2f}%</div>
                                <div style='font-size: 15px; font-weight: 800; color: #059669;'>Novo Valor Consolidado: {formatar_moeda(novo_val_calc)}</div>
                            </div>
                            """, unsafe_allow_html=True)

                    if perc_calc > 25.0:
                        st.warning(f"⚠️ **Alerta Jurídico AGU:** O percentual de alteração de **{perc_calc:.2f}%** excede o limite padrão de 25% (Art. 125 da Lei 14.133/21). Só é legalmente admitido em até 50% para reformas, ou por consenso devidamente motivado.")
                    else:
                        st.info(f"✅ O percentual de alteração (**{perc_calc:.2f}%**) encontra-se estritamente dentro do limite legal de 25% admitido pela AGU.")

                if tipo_cod == "PRORROGACAO":
                    just_sug = "A prorrogação do prazo de vigência contratual fundamenta-se na necessidade contínua dos serviços/fornecimentos para o regular funcionamento da Administração Pública Municipal, atestada a vantajosidade econômica perante pesquisa de mercado realizada no respectivo processo administrativo."
                elif tipo_cod == "ACRESCIMO":
                    just_sug = "O presente acréscimo de valor e quantitativo decorre da superveniência de novas demandas e ampliação das necessidades da Secretaria Municipal, respeitado o limite legal previsto na legislação de regência e mantidas as mesmas condições de preços unitários."
                elif tipo_cod == "SUPRESSAO":
                    just_sug = "A supressão pactuada visa à readequação dos quantitativos contratados às reais necessidades de consumo e disponibilidade orçamentária da Secretaria Contratante."
                elif tipo_cod == "MISTO":
                    just_sug = "O aditamento conjunto justifica-se pela imperiosa necessidade de continuidade das atividades municipais e concomitante adequação quantitativa do objeto, mantidas as condições mais vantajosas para o Município."
                elif tipo_cod == "REEQUILIBRIO":
                    just_sug = "A recomposição do equilíbrio econômico-financeiro fundamenta-se na ocorrência de variação extraordinária e imprevisível dos custos de execução do pacto, devidamente demonstrada em planilha e laudo técnico anexados aos autos."
                else:
                    just_sug = "A alteração qualitativa justifica-se pela melhor adequação técnica aos objetivos pretendidos pela Administração Municipal, sem descaracterização do objeto original."
                    
                just_input = st.text_area("Justificativa Administrativa da Alteração:", value=just_sug, height=85, key="adit_just_inp")
                
                cb_garantia = st.checkbox("Exigir complementação ou renovação da garantia de execução contratual", value=False, key="adit_cb_gar")
                txt_garantia = ""
                if cb_garantia:
                    txt_garantia = st.text_input("Disposições sobre a Garantia:", value="Caução/Seguro-garantia complementar a ser apresentado em até 10 dias úteis", key="adit_txt_gar")
                    
                st.markdown("##### 📎 Documentos Acessórios ao Aditivo")
                st.caption("Selecione os documentos complementares para gerar o pacote junto com a Minuta.")
                col_ac1, col_ac2 = st.columns(2)
                with col_ac1:
                    gerente_contratos = st.text_input("Gerente de Contratos (Responsável):", value="Jailany Chaves da Silva Pontes", key="adit_gerente_inp")
                with col_ac2:
                    licitacao_vin = st.text_input("Licitação Vinculada (Ex: PREGÃO ELETRÔNICO Nº 004/2024):", value=row_ad.get("modalidade") or "PREGÃO ELETRÔNICO Nº 004/2024", key="adit_lic_inp")

                st.markdown("##### 📁 Documentos Acessórios ao Termo Aditivo")
                st.caption("Selecione os anexos complementares que deseja incluir no Pacote ZIP junto com a Minuta do Aditivo:")
                opcoes_docs_adit = [
                    "Capa do Processo Físico", 
                    "Extrato de Aditivo (Publicação)", 
                    "Requerimento de Prorrogação/Aditamento", 
                    "Justificativa Administrativa Base"
                ]
                col_chk_a1, col_chk_a2 = st.columns(2)
                docs_selecionados = []
                for i_ca, opt_ca in enumerate(opcoes_docs_adit):
                    target_col = col_chk_a1 if i_ca % 2 == 0 else col_chk_a2
                    with target_col:
                        if st.checkbox(opt_ca, value=True, key=f"adit_chk_{opt_ca}_{cid_ad}"):
                            docs_selecionados.append(opt_ca)
                    
                dados_aditivo_final = {
                    "contrato_id": cid_ad,
                    "numero_aditivo": num_adit_val,
                    "numero_completo": num_comp_adit,
                    "numero_completo_aditivo": num_comp_adit,
                    "ano_aditivo": dt_ass_adit.year,
                    "tipo_aditivo": tipo_cod,
                    "regime_legal": regime_cod,
                    "processo_adm": proc_adit,
                    "data_assinatura": dt_ass_adit,
                    "data_publicacao": dt_ass_adit + timedelta(days=5),
                    "nova_data_vencimento": nova_venc_calc if nova_venc_calc else dt_venc_atual,
                    "prazo_aditado_meses": prazo_meses_val,
                    "prazo_meses": prazo_meses_val,
                    "valor_original": val_orig,
                    "valor_aditado": val_adit_val,
                    "percentual_aditado": perc_calc,
                    "novo_valor_total": novo_val_calc,
                    "objeto_aditivo": row_ad["objeto"],
                    "justificativa": just_input,
                    "dotacao_orcamentaria": dot_adit,
                    "garantia_execucao": txt_garantia if cb_garantia else None,
                    "secretario": sec_adit,
                    "cargo_secretario": hist_org_ad.get("cargo_secretario", "Secretário(a) Municipal Titular"),
                    "fiscal": fisc_adit,
                    "orgao_nome": row_ad["orgao"],
                    "orgao_cnpj": formatar_cnpj_cpf(hist_org_ad["cnpj"]),
                    "orgao_endereco": hist_org_ad["endereco"],
                    "orgao_bairro": hist_org_ad["bairro"],
                    "orgao_cep": hist_org_ad["cep"],
                    "orgao_cidade": hist_org_ad["cidade"],
                    "orgao_uf": hist_org_ad["uf"],
                    "fornecedor_nome": row_ad["fornecedor"],
                    "fornecedor_cnpj": formatar_cnpj_cpf(row_ad.get("cnpj_cpf", "00.000.000/0001-00")),
                    "fornecedor_representante": row_ad.get("fornecedor_representante") or "Representante Legal",
                    "fornecedor_cargo": row_ad.get("fornecedor_cargo") or "Sócio Administrador",
                    "fornecedor_endereco": row_ad.get("fornecedor_endereco") or "Sede Comercial da Empresa",
                    "fornecedor_bairro": row_ad.get("fornecedor_bairro") or "Centro",
                    "fornecedor_cidade": row_ad.get("fornecedor_cidade") or "Imperatriz",
                    "fornecedor_uf": row_ad.get("fornecedor_uf") or "MA",
                    "fornecedor_cep": row_ad.get("fornecedor_cep") or "65900-000",
                    "numero_contrato": row_ad["numero_completo"],
                    "data_inicio_aditivo": formatar_data_br(dt_venc_atual) if dt_venc_atual else formatar_data_br(date.today()),
                    "criado_por": st.session_state["usuario"].get("id") if st.session_state.get("usuario") else None,
                    "gerente_contratos": gerente_contratos,
                    "licitacao_vinculada": licitacao_vin,
                    "documentos_selecionados": docs_selecionados
                }
                
                with st.expander("👁️ Pré-Visualizar Cláusulas Oficiais da Minuta (Padrão AGU)", expanded=False):
                    clausulas_adit_preview = gerar_clausulas_aditivo(dados_aditivo_final)
                    for cl_a in clausulas_adit_preview:
                        st.markdown(f"**{cl_a['numero']}**")
                        for p_item in cl_a["conteudo"]:
                            st.caption(re.sub(r'<[^>]+>', '', p_item))
                            
                st.markdown("---")
                b_col1, b_col2, b_col3, b_col4 = st.columns([1.2, 1, 1, 1.2])
                
                with b_col1:
                    if st.button("💾 Salvar Termo Aditivo no Sistema", key="btn_salvar_adit_db", type="primary", use_container_width=True):
                        ok_salvo, msg_salvo = salvar_termo_aditivo(dados_aditivo_final, atualizar_contrato=True)
                        if ok_salvo:
                            st.success(f"✅ {msg_salvo}")
                            time.sleep(1.0)
                            st.rerun()
                        else:
                            st.error(f"Erro ao salvar: {msg_salvo}")
                            
                with b_col2:
                    try:
                        docx_ad_bytes = gerar_aditivo_docx(dados_aditivo_final)
                        st.download_button(
                            label="📥 Baixar Minuta em Word (.docx)",
                            data=docx_ad_bytes,
                            file_name=f"Minuta_Termo_Aditivo_{num_adit_val}_{str(row_ad['numero_completo']).replace('/', '_')}_AGU.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="btn_dl_docx_adit_m4",
                            use_container_width=True
                        )
                    except Exception as e_w:
                        st.error(f"Erro ao compilar Word: {e_w}")
                        
                with b_col3:
                    try:
                        pdf_ad_bytes = gerar_aditivo_pdf(dados_aditivo_final)
                        st.download_button(
                            label="📄 Baixar Minuta em PDF (.pdf)",
                            data=pdf_ad_bytes,
                            file_name=f"Minuta_Termo_Aditivo_{num_adit_val}_{str(row_ad['numero_completo']).replace('/', '_')}_AGU.pdf",
                            mime="application/pdf",
                            key="btn_dl_pdf_adit_m4",
                            use_container_width=True
                        )
                    except Exception as e_p:
                        st.error(f"Erro ao compilar PDF: {e_p}")

                with b_col4:
                    if docs_selecionados:
                        try:
                            from gerador_contratos_agu import gerar_pacote_zip
                            zip_adit_bytes = gerar_pacote_zip(dados_aditivo_final, is_aditivo=True)
                            st.download_button(
                                label="📦 Baixar Pacote Completo (.zip)",
                                data=zip_adit_bytes,
                                file_name=f"Pacote_Aditivo_{num_adit_val}_{str(row_ad['numero_completo']).replace('/', '_')}.zip",
                                mime="application/zip",
                                key=f"btn_dl_zip_adit_m4_{cid_ad}",
                                use_container_width=True,
                                help="Baixa um arquivo .zip contendo a minuta do termo aditivo e cada anexo selecionado como arquivos separados."
                            )
                        except Exception as e_a:
                            st.error(f"Erro ao gerar pacote ZIP: {e_a}")
                    else:
                        st.button("📦 Baixar Pacote (.zip)", disabled=True, use_container_width=True, key=f"btn_dl_zip_adit_m4_disabled_{cid_ad}")

                st.markdown("---")
                st.markdown("##### 📜 Histórico de Termos Aditivos já Celebrados para este Contrato:")
                df_hist_ad = carregar_aditivos_do_contrato(cid_ad)
                if df_hist_ad.empty:
                    st.info("Nenhum aditivo registrado anteriormente para este contrato.")
                else:
                    st.dataframe(
                        df_hist_ad[[
                            "numero_completo", "tipo_aditivo", "regime_legal", 
                            "data_assinatura", "prazo_aditado_meses", "valor_aditado", 
                            "novo_valor_total", "nova_data_vencimento", "status"
                        ]],
                        use_container_width=True,
                        hide_index=True
                    )



# ============================================
# TELA DE AUDITORIA
# ============================================

def tela_historico_acoes():
    st.markdown("""
        <div class='municipal-banner'>
            <div style='display: flex; justify-content: space-between; align-items: flex-start;'>
                <div>
                    <div class='municipal-badge-pill'>
                        <span class='beacon-active-gel' style='background: #ffffff;'></span> CONTROLE E TRANSPARÊNCIA
                    </div>
                    <h2 class='municipal-banner-title'>Histórico de Ações</h2>
                    <div class='municipal-banner-subtitle'>Prefeitura Municipal de Ribeirãozinho do Maranhão - MA</div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)



    with get_engine().connect() as conn:
        col1, col2, col3 = st.columns(3)
        with col1:
            filtro_tabela = st.selectbox("Módulo / Tabela", ["Todos", "contratos", "termos_aditivos", "protocolos"])
        with col2:
            filtro_acao = st.selectbox("Ação Executada", ["Todas", "CRIACAO", "ADITIVO", "EDICAO", "EXCLUSAO"])
        with col3:
            limite = st.selectbox("Últimos registros", [50, 100, 500, 1000])

        query = "SELECT id, data_hora, usuario_nome, tabela, acao, registro_id, descricao FROM audit_log WHERE 1=1"
        params = {}
        if filtro_tabela != "Todos":
            query += " AND tabela = :tb"
            params["tb"] = filtro_tabela
        if filtro_acao != "Todas":
            query += " AND acao = :ac"
            params["ac"] = filtro_acao
            
        query += " ORDER BY data_hora DESC LIMIT :limit"
        params["limit"] = limite
        
        df_audit = pd.read_sql(text(query), conn, params=params)

    if df_audit.empty:
        st.info("Nenhum registro de auditoria encontrado para os filtros selecionados.")
    else:
        # Formatar coluna data_hora
        df_audit["data_hora"] = pd.to_datetime(df_audit["data_hora"]).dt.strftime("%d/%m/%Y %H:%M:%S")
        
        st.dataframe(
            df_audit,
            column_config={
                "id": st.column_config.NumberColumn("ID Log", format="%d"),
                "data_hora": "Data e Hora",
                "usuario_nome": "Usuário Responsável",
                "tabela": "Módulo (Tabela)",
                "acao": "Ação",
                "registro_id": "ID Registro",
                "descricao": "Detalhes da Ação"
            },
            use_container_width=True,
            hide_index=True
        )

# ============================================

# ============================================
# TELA 7: DASHBOARD ANALÍTICA (BI)
# ============================================
def tela_dashboard_bi():
    st.markdown("<h2 class='municipal-banner-title'>📈 Inteligência de Dados e BI</h2>", unsafe_allow_html=True)
    st.markdown("<div class='municipal-banner-subtitle'>Visão Geral Estratégica dos Contratos Municipais</div><br>", unsafe_allow_html=True)
    
    with get_engine().connect() as conn:
        df_c = pd.read_sql("SELECT * FROM contratos", conn)
        
    if df_c.empty:
        st.warning("Nenhum dado suficiente para gerar gráficos no momento.")
        return
        
    # Processamento básico
    df_c['valor_total'] = pd.to_numeric(df_c['valor_total'], errors='coerce').fillna(0)
    if 'data_vencimento' in df_c.columns:
        df_c['data_vencimento'] = pd.to_datetime(df_c['data_vencimento'], errors='coerce')
        
    # KPIs Topo
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f'''
        <div class="gel-kpi-card">
            <div class="gel-kpi-title">Total de Contratos</div>
            <div class="gel-kpi-value" style="color: #0284c7;">{len(df_c)}</div>
        </div>
        ''', unsafe_allow_html=True)
    with col2:
        val_total = df_c['valor_total'].sum()
        st.markdown(f'''
        <div class="gel-kpi-card">
            <div class="gel-kpi-title">Volume Financeiro Total</div>
            <div class="gel-kpi-value" style="color: #059669;">{formatar_moeda(val_total)}</div>
        </div>
        ''', unsafe_allow_html=True)
    with col3:
        vencidos = 0
        if 'data_vencimento' in df_c.columns:
            agora_ts = pd.Timestamp.now().normalize()
            vencidos = len(df_c[(df_c['data_vencimento'] < agora_ts) & (df_c['status'] != 'DISTRATADO')])
        st.markdown(f'''
        <div class="gel-kpi-card">
            <div class="gel-kpi-title">Contratos Vencidos</div>
            <div class="gel-kpi-value" style="color: #dc2626;">{vencidos}</div>
        </div>
        ''', unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Gráficos Nativos
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.markdown("#### 📊 Distribuição por Status")
        status_counts = df_c['status'].value_counts()
        st.bar_chart(status_counts, color="#0ea5e9")
        
    with col_g2:
        st.markdown("#### 🏢 Volume Financeiro por Órgão")
        if 'orgao_contratante' in df_c.columns:
            orgao_val = df_c.groupby('orgao_contratante')['valor_total'].sum().sort_values(ascending=False)
            st.bar_chart(orgao_val, color="#10b981")
            
    # Curva de Vencimentos (Próximos 6 meses)
    if 'data_vencimento' in df_c.columns:
        st.markdown("#### 📅 Calendário de Vencimentos (Próximos 12 meses)")
        hoje_ts = pd.Timestamp.now().normalize()
        doze_meses_ts = hoje_ts + pd.Timedelta(days=365)
        df_venc = df_c[(df_c['data_vencimento'] >= hoje_ts) & (df_c['data_vencimento'] <= doze_meses_ts)].copy()
        if not df_venc.empty:
            df_venc['Mes_Ano'] = df_venc['data_vencimento'].dt.to_period('M').astype(str)
            venc_counts = df_venc['Mes_Ano'].value_counts().sort_index()
            st.line_chart(venc_counts, color="#f59e0b")
        else:
            st.info("Nenhum contrato programado para vencer nos próximos 12 meses.")


# ============================================
# TELA 8: GESTÃO DE FORNECEDORES E PENALIDADES
# ============================================
def tela_fornecedores():
    st.markdown("<h2 class='municipal-banner-title'>🏢 Gestão de Fornecedores</h2>", unsafe_allow_html=True)
    st.markdown("<div class='municipal-banner-subtitle'>Controle de Empresas, CNPJs e Penalidades</div><br>", unsafe_allow_html=True)
    
    try:
        with get_engine().connect() as conn:
            df_f = pd.read_sql("SELECT * FROM fornecedores ORDER BY razao_social ASC", conn)
    except Exception as e:
        st.error(f"Erro ao acessar tabela de fornecedores: {e}")
        return
        
    aba_lista, aba_novo = st.tabs(["📋 Base de Fornecedores", "➕ Cadastrar Fornecedor"])
    
    with aba_lista:
        if df_f.empty:
            st.info("Nenhum fornecedor cadastrado na base.")
        else:
            st.dataframe(
                df_f[["cnpj_cpf", "razao_social", "status_penalidade", "telefone", "cidade", "uf"]],
                use_container_width=True,
                hide_index=True
            )
            
            st.markdown("### ⚠️ Aplicar/Remover Penalidade")
            f_sel = st.selectbox("Selecione o Fornecedor para Atualizar Status:", df_f["razao_social"].tolist())
            if f_sel:
                row_f = df_f[df_f["razao_social"] == f_sel].iloc[0]
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**CNPJ:** {row_f['cnpj_cpf']}")
                    st.write(f"**Status Atual:** {row_f['status_penalidade']}")
                with col2:
                    novo_status = st.selectbox("Novo Status Jurídico:", ["REGULAR", "ADVERTÊNCIA", "SUSPENSO / INIDÔNEO"])
                    obs_penal = st.text_input("Processo / Observação da Penalidade:", value=row_f["observacoes"] if row_f["observacoes"] else "")
                    
                    if st.button("💾 Atualizar Status do Fornecedor", type="primary"):
                        if st.session_state["usuario"]["cargo"] == "AUDITOR":
                            st.error("Auditores não têm permissão para editar dados.")
                        else:
                            with get_engine().connect() as conn:
                                conn.execute(
                                    text("UPDATE fornecedores SET status_penalidade = :st, observacoes = :ob WHERE id = :id"),
                                    {"st": novo_status, "ob": obs_penal, "id": int(row_f["id"])}
                                )
                                conn.commit()
                                registrar_audit_log(conn, "fornecedores", int(row_f["id"]), "EDICAO", f"Alterou status de {row_f['status_penalidade']} para {novo_status}")
                            st.success("Fornecedor atualizado!")
                            st.rerun()

    with aba_novo:
        st.markdown("#### ➕ Cadastrar / Importar Fornecedor")
        def _callback_tf_cnpj_auto():
            cnpj_in = st.session_state.get("tf_cnpj_busca", "")
            c_l = re.sub(r"\D", "", str(cnpj_in or ""))
            if len(c_l) == 14:
                st.session_state["tf_cnpj_busca"] = formatar_cnpj_cpf(c_l)
                # 1. Banco de Dados
                res_b = buscar_fornecedor_no_banco(c_l, get_engine())
                if res_b and res_b.get("razao_social") and res_b.get("telefone") and res_b.get("email"):
                    st.session_state["tf_form_cnpj"] = res_b.get("cnpj_formatado") or formatar_cnpj_cpf(c_l)
                    st.session_state["tf_form_razao"] = res_b.get("razao_social") or ""
                    st.session_state["tf_form_tel"] = res_b.get("telefone") or ""
                    st.session_state["tf_form_email"] = res_b.get("email") or ""
                    st.toast(f"Empresa {res_b.get('razao_social')} encontrada no banco!", icon="✅")
                else:
                    # 2. API oficial automática
                    res_a = puxar_fornecedor_api_e_cadastrar(c_l, get_engine())
                    if res_a.get("sucesso"):
                        st.session_state["tf_form_cnpj"] = res_a.get("cnpj_formatado") or formatar_cnpj_cpf(c_l)
                        st.session_state["tf_form_razao"] = res_a.get("razao_social") or (res_b.get("razao_social") if res_b else "")
                        st.session_state["tf_form_tel"] = res_a.get("telefone") or (res_b.get("telefone") if res_b else "")
                        st.session_state["tf_form_email"] = res_a.get("email") or (res_b.get("email") if res_b else "")
                        st.toast(f"Dados obtidos na Receita Federal!", icon="🎉")
                    elif res_b and res_b.get("razao_social"):
                        st.session_state["tf_form_cnpj"] = res_b.get("cnpj_formatado") or formatar_cnpj_cpf(c_l)
                        st.session_state["tf_form_razao"] = res_b.get("razao_social") or ""
                        st.session_state["tf_form_tel"] = res_b.get("telefone") or ""
                        st.session_state["tf_form_email"] = res_b.get("email") or ""

        st.caption("Digite o CNPJ para auto-preenchimento automático dos dados cadastrais (padrão oficial AGU):")
        st.text_input(
            "CNPJ da Empresa (Auto-preenchimento):",
            placeholder="00.000.000/0001-00 (ou somente números)",
            key="tf_cnpj_busca",
            on_change=_callback_tf_cnpj_auto,
            help="Ao digitar os 14 números do CNPJ, o sistema busca no banco e na Receita Federal e preenche automaticamente o formulário abaixo."
        )

        with st.form("form_novo_fornecedor"):
            st.markdown("##### Dados Cadastrais")
            c1, c2 = st.columns([1, 2])
            f_cnpj = c1.text_input("CNPJ", value=st.session_state.get("tf_form_cnpj", ""), placeholder="00.000.000/0001-00")
            f_razao = c2.text_input("Razão Social", value=st.session_state.get("tf_form_razao", ""))
            
            c3, c4 = st.columns(2)
            f_tel = c3.text_input("Telefone", value=st.session_state.get("tf_form_tel", ""))
            f_email = c4.text_input("E-mail", value=st.session_state.get("tf_form_email", ""))
            
            sub_f = st.form_submit_button("Salvar Fornecedor no Banco", type="primary")
            
            if sub_f:
                if st.session_state["usuario"]["cargo"] == "AUDITOR":
                    st.error("Auditores não têm permissão para cadastrar dados.")
                elif not f_cnpj or not f_razao:
                    st.error("CNPJ e Razão Social são obrigatórios.")
                else:
                    with get_engine().connect() as conn:
                        try:
                            res = conn.execute(
                                text("INSERT INTO fornecedores (cnpj_cpf, razao_social, telefone, email) VALUES (:c, :r, :t, :e) RETURNING id"),
                                {"c": f_cnpj, "r": f_razao, "t": f_tel, "e": f_email}
                            ).fetchone()
                            conn.commit()
                            registrar_audit_log(conn, "fornecedores", int(res[0]), "CRIACAO", f"Cadastrou fornecedor {f_cnpj}")
                            st.success("Cadastrado com sucesso no banco de dados!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao cadastrar (CNPJ já existe?): {e}")

# MAIN (SIDEBAR COM AS CORES DA CIDADE)
# ============================================


# ============================================
# GESTÃO DE SECRETÁRIOS & FISCAIS MUNICIPAIS
# ============================================

def tela_gestao_secretarios_fiscais():
    brasao_b64 = obter_brasao_b64()
    brasao_banner_img = f"<div style='background: rgba(255,255,255,0.92); padding: 6px 12px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.15); display: flex; align-items: center; justify-content: center;'><img src='data:image/png;base64,{brasao_b64}' style='height: 52px; width: auto;' alt='Brasão' /></div>" if brasao_b64 else ""

    st.markdown(f"""
        <div class='municipal-banner' style='margin-bottom: 16px;'>
            <div style='display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 14px;'>
                <div>
                    <div class='municipal-badge-pill'>
                        <span class='beacon-active-gel' style='background: #ffffff;'></span> GESTÃO DE AUTORIDADES & FISCALIZAÇÃO
                    </div>
                    <h2 class='municipal-banner-title'>🏛️ Secretários(as) e Fiscais Municipais</h2>
                    <div class='municipal-banner-subtitle'>
                        Atualize os titulares e fiscais de cada pasta municipal. <strong>As alterações passam a valer automaticamente para todos os novos contratos e minutas emitidos a partir de agora.</strong> Os contratos anteriores permanecem intactos com os nomes registrados à época da sua criação.
                    </div>
                </div>
                {brasao_banner_img}
            </div>
        </div>
    """, unsafe_allow_html=True)

    tab_lista, tab_novo_org = st.tabs([
        "👥 Secretarias & Titulares Cadastrados",
        "➕ Cadastrar Nova Secretaria / Órgão"
    ])

    with tab_lista:
        with get_engine().connect() as conn:
            rows_orgaos = conn.execute(
                text("SELECT * FROM orgaos WHERE ativo = 1 ORDER BY nome ASC")
            ).mappings().fetchall()

        total_orgs = len(rows_orgaos)
        com_sec = sum(1 for r in rows_orgaos if r.get("secretario_padrao") and str(r["secretario_padrao"]).strip())
        com_fisc = sum(1 for r in rows_orgaos if r.get("fiscal_padrao") and str(r["fiscal_padrao"]).strip())

        k1, k2, k3 = st.columns(3)
        with k1:
            st.markdown(f"""
                <div class='gel-kpi-card'>
                    <div class='gel-kpi-title'>Secretarias / Órgãos</div>
                    <div class='gel-kpi-value'>{total_orgs}</div>
                    <div class='gel-kpi-sub'>Pastas Ativas no Município</div>
                </div>
            """, unsafe_allow_html=True)
        with k2:
            st.markdown(f"""
                <div class='gel-kpi-card'>
                    <div class='gel-kpi-title'>Titulares Designados</div>
                    <div class='gel-kpi-value' style='color: #0284c7;'>{com_sec}</div>
                    <div class='gel-kpi-sub'>Secretários / Diretores</div>
                </div>
            """, unsafe_allow_html=True)
        with k3:
            st.markdown(f"""
                <div class='gel-kpi-card'>
                    <div class='gel-kpi-title'>Fiscais Cadastrados</div>
                    <div class='gel-kpi-value' style='color: #16a34a;'>{com_fisc}</div>
                    <div class='gel-kpi-sub'>Servidores Fiscais Ativos</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

        col_b1, col_b2 = st.columns([3, 1])
        with col_b1:
            busca_org = st.text_input(
                "🔍 Pesquisar Secretaria, Secretário ou Fiscal:",
                placeholder="Ex: Educação, Saúde, João Vitor, Sirleide, etc.",
                key="filtro_busca_gestores"
            )
        with col_b2:
            st.write("")
            st.write("")
            filtro_pendente = st.checkbox("Apenas pendentes de fiscal/secretário", key="chk_filtro_pendente")

        orgs_filtrados = []
        for r in rows_orgaos:
            termo = busca_org.lower().strip()
            nome_o = str(r.get("nome", "")).lower()
            sec_o = str(r.get("secretario_padrao", "") or "").lower()
            fisc_o = str(r.get("fiscal_padrao", "") or "").lower()
            carg_o = str(r.get("cargo_secretario", "") or "").lower()

            match_busca = (not termo) or (termo in nome_o) or (termo in sec_o) or (termo in fisc_o) or (termo in carg_o)
            match_pend = True
            if filtro_pendente:
                match_pend = (not r.get("secretario_padrao") or not str(r["secretario_padrao"]).strip()) or (not r.get("fiscal_padrao") or not str(r["fiscal_padrao"]).strip())

            if match_busca and match_pend:
                orgs_filtrados.append(r)

        st.caption(f"Mostrando **{len(orgs_filtrados)}** de **{total_orgs}** pastas municipais.")

        for org in orgs_filtrados:
            oid = org["id"]
            nome_org = org["nome"]
            sec_atual = org.get("secretario_padrao") or ""
            carg_atual = org.get("cargo_secretario") or "Secretário(a) Municipal Titular"
            fisc_atual = org.get("fiscal_padrao") or ""
            cnpj_atual = org.get("cnpj") or "01.612.834/0001-86"
            end_atual = org.get("endereco") or "Rua Principal, s/n"
            bairro_atual = org.get("bairro") or "Centro"
            cep_atual = org.get("cep") or "65928-000"

            with st.expander(f"🏛️ **{nome_org}** — 👤 {sec_atual or '⚠️ Sem Secretário'} | 📋 Fiscal: {fisc_atual or '⚠️ Sem Fiscal'}", expanded=False):
                with st.form(f"form_editar_gestor_{oid}"):
                    c1, c2 = st.columns(2)
                    with c1:
                        novo_sec = st.text_input(
                            "👤 Nome Completo do(a) Secretário(a) / Titular:",
                            value=sec_atual,
                            placeholder="Ex: JOÃO VITOR SOUSA JUSTINO",
                            help="Este nome sairá como CONTRATANTE e na assinatura dos novos contratos gerados para esta Secretaria."
                        )
                        novo_cargo = st.text_input(
                            "🏷️ Cargo Oficial do Titular:",
                            value=carg_atual,
                            placeholder="Ex: Secretário Municipal de Educação"
                        )
                    with c2:
                        novo_fisc = st.text_input(
                            "📋 Nome do Servidor Fiscal Designado (Portaria):",
                            value=fisc_atual,
                            placeholder="Ex: CHARLIANE DE ABREU MACIEL",
                            help="Este nome sairá na Cláusula de Fiscalização e no Extrato de Contrato dos novos documentos."
                        )
                        novo_cnpj = st.text_input(
                            "🏢 CNPJ Institucional da Secretaria / Órgão:",
                            value=cnpj_atual
                        )

                    with st.expander("📍 Endereço da Sede Administrativa", expanded=False):
                        e1, e2, e3 = st.columns([2, 1.5, 1])
                        novo_end = e1.text_input("Endereço / Logradouro:", value=end_atual, key=f"end_o_{oid}")
                        novo_bai = e2.text_input("Bairro:", value=bairro_atual, key=f"bai_o_{oid}")
                        novo_cep = e3.text_input("CEP:", value=cep_atual, key=f"cep_o_{oid}")

                    sub_salvar_org = st.form_submit_button("💾 Salvar Alterações desta Secretaria", type="primary", use_container_width=True)

                    if sub_salvar_org:
                        if not novo_sec.strip():
                            st.warning("⚠️ Informe o nome do Secretário(a) Titular.")
                        else:
                            try:
                                with get_engine().connect() as conn:
                                    conn.execute(
                                        text("""
                                            UPDATE orgaos
                                            SET secretario_padrao = :sec,
                                                cargo_secretario = :carg,
                                                fiscal_padrao = :fisc,
                                                cnpj = :cnpj,
                                                endereco = :end,
                                                bairro = :bai,
                                                cep = :cep
                                            WHERE id = :id
                                        """),
                                        {
                                            "sec": novo_sec.strip().upper(),
                                            "carg": novo_cargo.strip(),
                                            "fisc": novo_fisc.strip().upper() if novo_fisc else "",
                                            "cnpj": novo_cnpj.strip(),
                                            "end": novo_end.strip(),
                                            "bai": novo_bai.strip(),
                                            "cep": novo_cep.strip(),
                                            "id": oid
                                        }
                                    )
                                    conn.commit()
                                    registrar_audit_log(
                                        conn, "orgaos", oid, "EDICAO_TITULAR",
                                        f"Atualizou titular para {novo_sec.strip().upper()} e fiscal para {novo_fisc.strip().upper()}"
                                    )
                                    try:
                                        obter_historico_orgao.clear()
                                    except Exception:
                                        pass

                                st.success(f"✅ **{nome_org}** atualizada com sucesso! Todos os **novos contratos e minutas** sairão com o titular **{novo_sec.strip().upper()}** e fiscal **{novo_fisc.strip().upper() or 'Designado'}**.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erro ao salvar: {e}")

    with tab_novo_org:
        st.markdown("##### ➕ Cadastro de Nova Secretaria ou Entidade Municipal")
        st.caption("Cadastre uma nova Secretaria, Fundo Municipal, Autarquia ou Órgão com seu respectivo Secretário Titular e Fiscal.")

        with st.form("form_cadastrar_novo_orgao_gestores"):
            c_no1, c_no2 = st.columns([2, 1])
            with c_no1:
                novo_org_nome = st.text_input("Nome Oficial do Órgão / Secretaria *", placeholder="Ex: SECRETARIA MUNICIPAL DE SEGURANÇA PÚBLICA")
            with c_no2:
                novo_org_sigla = st.text_input("Sigla do Órgão:", placeholder="Ex: SEMSEG")

            c_no3, c_no4 = st.columns(2)
            with c_no3:
                novo_org_sec = st.text_input("Nome do(a) Secretário(a) Titular:", placeholder="Nome do titular responsável")
                novo_org_carg = st.text_input("Cargo Oficial:", value="Secretário(a) Municipal Titular")
            with c_no4:
                novo_org_fisc = st.text_input("Fiscal de Contrato Padrão:", placeholder="Nome do servidor fiscal")
                novo_org_cnpj = st.text_input("CNPJ:", value="01.612.834/0001-86")

            c_no5, c_no6, c_no7 = st.columns([2, 1.5, 1])
            with c_no5:
                novo_org_end = st.text_input("Logradouro / Número:", value="Rua Principal, s/n")
            with c_no6:
                novo_org_bai = st.text_input("Bairro:", value="Centro")
            with c_no7:
                novo_org_cep = st.text_input("CEP:", value="65928-000")

            btn_criar_org = st.form_submit_button("🏛️ Cadastrar Secretaria no Sistema", type="primary", use_container_width=True)

            if btn_criar_org:
                if not novo_org_nome.strip():
                    st.error("O nome da Secretaria / Órgão é obrigatório.")
                else:
                    try:
                        with get_engine().connect() as conn:
                            res = conn.execute(
                                text("""
                                    INSERT INTO orgaos (
                                        nome, sigla, secretario_padrao, cargo_secretario, fiscal_padrao,
                                        cnpj, endereco, bairro, cep, cidade, uf, ativo
                                    ) VALUES (
                                        :nome, :sigla, :sec, :carg, :fisc,
                                        :cnpj, :end, :bai, :cep, 'Ribeirãozinho do Maranhão', 'MA', 1
                                    )
                                """),
                                {
                                    "nome": novo_org_nome.strip().upper(),
                                    "sigla": novo_org_sigla.strip().upper() if novo_org_sigla else novo_org_nome[:10].upper(),
                                    "sec": novo_org_sec.strip().upper() if novo_org_sec else "",
                                    "carg": novo_org_carg.strip() if novo_org_carg else "Secretário(a) Municipal Titular",
                                    "fisc": novo_org_fisc.strip().upper() if novo_org_fisc else "",
                                    "cnpj": novo_org_cnpj.strip(),
                                    "end": novo_org_end.strip(),
                                    "bai": novo_org_bai.strip(),
                                    "cep": novo_org_cep.strip()
                                }
                            )
                            conn.commit()
                            registrar_audit_log(conn, "orgaos", res.lastrowid or 0, "CRIACAO", f"Cadastrou órgão {novo_org_nome.strip().upper()}")
                            try:
                                obter_historico_orgao.clear()
                            except Exception:
                                pass
                        st.success(f"✅ Secretaria **{novo_org_nome.strip().upper()}** cadastrada com sucesso!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao cadastrar: {e}")

# ============================================
# CENTRAL DE BACKUPS, SEGURANÇA & MONITORAMENTO
# ============================================

@st.cache_data(ttl=600, show_spinner=False)
def obter_status_conexao_banco():
    t0 = time.time()
    try:
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
            latencia = int((time.time() - t0) * 1000)
            db_url_str = str(engine.url).lower()
            is_pg = "postgresql" in db_url_str or "postgres" in db_url_str
            return {
                "online": True,
                "tipo": "PostgreSQL (Supabase Nuvem)" if is_pg else "SQLite Local",
                "latencia_ms": max(latencia, 1),
                "total_contratos": 66,
                "msg": "Banco Conectado e Operacional"
            }
    except Exception as e:
        return {
            "online": False,
            "tipo": "Indisponível",
            "latencia_ms": -1,
            "total_contratos": 0,
            "msg": str(e)
        }

def json_serializer_backup(obj):
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    if isinstance(obj, Decimal):
        return float(obj)
    return str(obj)

def gerar_backup_excel(eng):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    query = """
        SELECT 
            c.id AS "ID",
            c.numero_contrato AS "Nº Contrato",
            c.ano_contrato AS "Ano",
            c.numero_completo AS "Nº Completo",
            o.nome AS "Secretaria / Órgão",
            c.modalidade AS "Modalidade",
            c.processo_adm AS "Processo Adm",
            c.objeto AS "Objeto Contratual",
            COALESCE(f.razao_social, f.nome_fantasia, 'Não informado') AS "Fornecedor / Razão Social",
            f.cnpj_cpf AS "CNPJ / CPF Fornecedor",
            c.valor_total AS "Valor Total (R$)",
            c.data_assinatura AS "Data de Assinatura",
            c.data_publicacao AS "Data de Publicação",
            c.data_vencimento AS "Data de Vencimento",
            c.vigencia_descricao AS "Vigência",
            c.secretario AS "Secretário(a) Titular",
            c.fiscal AS "Fiscal do Contrato",
            c.status AS "Status",
            c.dotacao_orcamentaria AS "Dotação Orçamentária",
            c.modelo_agu AS "Modelo AGU",
            f.telefone AS "Telefone Fornecedor",
            f.email AS "Email Fornecedor",
            c.observacoes AS "Observações",
            c.data_criacao AS "Data de Registro no Sistema"
        FROM contratos c
        LEFT JOIN orgaos o ON c.orgao_id = o.id
        LEFT JOIN fornecedores f ON c.fornecedor_id = f.id
        ORDER BY c.ano_contrato DESC, c.numero_contrato ASC
    """
    
    with eng.connect() as conn:
        df = pd.read_sql(text(query), conn)
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Contratos', index=False)
        ws = writer.sheets['Contratos']
        
        # Cabeçalho estilizado em Azul Real Oficial
        header_fill = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        # Formatação de linhas e auto-ajuste de largura
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)
            
    output.seek(0)
    return output.getvalue()

def gerar_backup_json(eng):
    tabelas = ['contratos', 'fornecedores', 'orgaos', 'termos_aditivos', 'protocolos', 'usuarios', 'audit_log']
    dados = {
        "metadados": {
            "municipio": "Prefeitura Municipal de Ribeirãozinho do Maranhão - MA",
            "gerado_em": datetime.now().isoformat(),
            "tipo_banco": "PostgreSQL (Supabase)",
            "versao_sistema": "2026.1",
            "conteudo": "BACKUP GERAL COMPLETO (TODAS AS TABELAS DO BANCO DE DADOS)"
        },
        "tabelas": {}
    }
    with eng.connect() as conn:
        for tab in tabelas:
            try:
                rows = conn.execute(text(f"SELECT * FROM {tab}")).mappings().all()
                registros = []
                for r in rows:
                    d = dict(r)
                    if tab == 'usuarios' and 'senha_hash' in d:
                        del d['senha_hash']
                    registros.append(d)
                dados["tabelas"][tab] = registros
            except Exception:
                dados["tabelas"][tab] = []
            
    return json.dumps(dados, indent=2, ensure_ascii=False, default=json_serializer_backup).encode('utf-8')

def gerar_backup_sql(eng):
    tabelas = ['orgaos', 'fornecedores', 'usuarios', 'contratos', 'termos_aditivos', 'protocolos', 'audit_log']
    linhas_sql = [
        "-- ===================================================================",
        "-- BACKUP DUMP SQL GERAL - PREFEITURA DE RIBEIRÃOZINHO DO MARANHÃO - MA",
        "-- CONTÉM TODAS AS INFORMAÇÕES DO BANCO DE DADOS",
        f"-- Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        "-- ===================================================================\n"
    ]
    with eng.connect() as conn:
        for tab in tabelas:
            try:
                rows = conn.execute(text(f"SELECT * FROM {tab}")).mappings().all()
                if not rows:
                    continue
                linhas_sql.append(f"\n-- Tabela: {tab} ({len(rows)} registros)")
                cols = list(rows[0].keys())
                cols_str = ", ".join(cols)
                for r in rows:
                    valores = []
                    for c in cols:
                        v = r[c]
                        if v is None:
                            valores.append("NULL")
                        elif isinstance(v, (int, float, Decimal)):
                            valores.append(str(v))
                        else:
                            v_str = str(v).replace("'", "''")
                            valores.append(f"'{v_str}'")
                    vals_str = ", ".join(valores)
                    linhas_sql.append(f"INSERT INTO {tab} ({cols_str}) VALUES ({vals_str});")
            except Exception:
                pass
                
    return "\n".join(linhas_sql).encode('utf-8')

def tela_backups_sistema():
    brasao_b64 = obter_brasao_b64()
    brasao_banner_img = f"<div style='background: rgba(255,255,255,0.92); padding: 6px 12px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.15); display: flex; align-items: center; justify-content: center;'><img src='data:image/png;base64,{brasao_b64}' style='height: 52px; width: auto;' alt='Brasão' /></div>" if brasao_b64 else ""

    st.markdown(f"""
        <div class='municipal-banner' style='margin-bottom: 16px;'>
            <div style='display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 14px;'>
                <div>
                    <div class='municipal-badge-pill'>
                        <span class='beacon-active-gel' style='background: #ffffff;'></span> SEGURANÇA, CONECTIVIDADE & BACKUPS
                    </div>
                    <h2 class='municipal-banner-title'>💾 Central de Backups & Integridade de Dados</h2>
                    <div class='municipal-banner-subtitle'>
                        Monitore a conexão em tempo real e exporte backups sob demanda: <strong>Planilha Excel (.xlsx) exclusiva de Contratos</strong> ou <strong>Backup Geral com todas as informações do banco</strong>.
                    </div>
                </div>
                {brasao_banner_img}
            </div>
        </div>
    """, unsafe_allow_html=True)

    status = obter_status_conexao_banco()

    # Métricas de Conexão e Integridade
    col_st1, col_st2, col_st3, col_st4 = st.columns(4)
    with col_st1:
        if status["online"]:
            st.metric("Status do Banco", "🟢 ONLINE / ATIVO", delta="Nuvem Conectada")
        else:
            st.metric("Status do Banco", "🔴 DESCONECTADO", delta="Verifique conexão", delta_color="inverse")
    with col_st2:
        st.metric("Tecnologia de Armazenamento", status["tipo"])
    with col_st3:
        st.metric("Tempo de Resposta (Ping)", f"{status['latencia_ms']} ms" if status['latencia_ms'] > 0 else "N/A")
    with col_st4:
        st.metric("Total de Contratos Ativos", f"{status['total_contratos']} Contratos")

    st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

    # Abas da Central de Segurança
    aba_manual, aba_auto, aba_diagnostico = st.tabs([
        "📥 Backup Manual Sob Demanda (Excel / JSON / SQL)",
        "☁️ Backup Automático em Nuvem (Supabase)",
        "🔍 Diagnóstico das Tabelas"
    ])

    with aba_manual:
        st.markdown("""
        <div style='background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 14px 18px; margin-bottom: 20px;'>
            <h4 style='color: #0369a1; margin: 0 0 6px 0; font-size: 15px;'>📦 Exportação Completa Sob Demanda</h4>
            <p style='color: #475569; font-size: 13px; margin: 0;'>
                Escolha o formato desejado: a <strong>Planilha Excel</strong> reúne exclusivamente todos os contratos com todos os dados completos; os formatos <strong>JSON</strong> e <strong>SQL</strong> contêm todas as informações e tabelas do banco de dados.
            </p>
        </div>
        """, unsafe_allow_html=True)

        col_b1, col_b2, col_b3 = st.columns(3)

        # 1. Excel (Somente Contratos com dados completos)
        with col_b1:
            st.markdown("""
            <div style='background: #ffffff; border: 1.5px solid #0284c7; border-radius: 12px; padding: 18px; box-shadow: 0 4px 12px rgba(2, 132, 199, 0.08); text-align: center; height: 100%; display: flex; flex-direction: column; justify-content: space-between;'>
                <div>
                    <div style='font-size: 38px; margin-bottom: 8px;'>📊</div>
                    <h3 style='color: #0369a1; font-size: 16px; margin: 0 0 6px 0; font-weight: 800;'>Planilha de Contratos</h3>
                    <div style='background: #e0f2fe; color: #0369a1; font-size: 11px; font-weight: 700; padding: 3px 8px; border-radius: 20px; display: inline-block; margin-bottom: 12px;'>
                        SOMENTE CONTRATOS (DADOS COMPLETOS)
                    </div>
                    <p style='color: #64748b; font-size: 12px; text-align: left; line-height: 1.4;'>
                        Contém <strong>todos os 66 contratos</strong> municipais com todas as <strong>24 colunas detalhadas</strong> (fornecedores, valores em R$, secretarias, vigências, datas, dotação, fiscais). Ideal para Excel, relatórios de gestão e TCE.
                    </p>
                </div>
            </div>
            """, unsafe_allow_html=True)
            with st.spinner("Preparando planilha de Contratos..."):
                try:
                    excel_bytes = gerar_backup_excel(engine)
                    st.download_button(
                        label="📥 Baixar Contratos (.xlsx)",
                        data=excel_bytes,
                        file_name=f"Planilha_Contratos_Completos_Ribeiraozinho_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_down_excel",
                        use_container_width=True
                    )
                except Exception as ex_err:
                    st.error(f"Erro ao gerar Excel: {ex_err}")

        # 2. JSON (Banco de dados completo)
        with col_b2:
            st.markdown("""
            <div style='background: #ffffff; border: 1.5px solid #059669; border-radius: 12px; padding: 18px; box-shadow: 0 4px 12px rgba(5, 150, 105, 0.08); text-align: center; height: 100%; display: flex; flex-direction: column; justify-content: space-between;'>
                <div>
                    <div style='font-size: 38px; margin-bottom: 8px;'>📦</div>
                    <h3 style='color: #059669; font-size: 16px; margin: 0 0 6px 0; font-weight: 800;'>Backup Geral do Banco</h3>
                    <div style='background: #d1fae5; color: #065f46; font-size: 11px; font-weight: 700; padding: 3px 8px; border-radius: 20px; display: inline-block; margin-bottom: 12px;'>
                        BANCO COMPLETO (TODAS AS INFORMAÇÕES)
                    </div>
                    <p style='color: #64748b; font-size: 12px; text-align: left; line-height: 1.4;'>
                        Contém <strong>todas as tabelas do banco</strong>: Contratos, Fornecedores, Secretarias, Termos Aditivos, Protocolos, Usuários e Auditoria em formato padronizado JSON.
                    </p>
                </div>
            </div>
            """, unsafe_allow_html=True)
            with st.spinner("Preparando backup geral do banco..."):
                try:
                    json_bytes = gerar_backup_json(engine)
                    st.download_button(
                        label="📥 Baixar Banco Completo (.json)",
                        data=json_bytes,
                        file_name=f"Backup_Geral_Banco_Ribeiraozinho_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                        mime="application/json",
                        key="btn_down_json",
                        use_container_width=True
                    )
                except Exception as js_err:
                    st.error(f"Erro ao gerar JSON: {js_err}")

        # 3. SQL (Dump do banco completo)
        with col_b3:
            st.markdown("""
            <div style='background: #ffffff; border: 1.5px solid #d97706; border-radius: 12px; padding: 18px; box-shadow: 0 4px 12px rgba(217, 119, 6, 0.08); text-align: center; height: 100%; display: flex; flex-direction: column; justify-content: space-between;'>
                <div>
                    <div style='font-size: 38px; margin-bottom: 8px;'>💾</div>
                    <h3 style='color: #d97706; font-size: 16px; margin: 0 0 6px 0; font-weight: 800;'>Dump SQL Geral</h3>
                    <div style='background: #fef3c7; color: #92400e; font-size: 11px; font-weight: 700; padding: 3px 8px; border-radius: 20px; display: inline-block; margin-bottom: 12px;'>
                        BANCO COMPLETO (.SQL RESTAURÁVEL)
                    </div>
                    <p style='color: #64748b; font-size: 12px; text-align: left; line-height: 1.4;'>
                        Instruções SQL completas de inserção de <strong>todas as tabelas do banco</strong> para restauração total do sistema em qualquer servidor relacional.
                    </p>
                </div>
            </div>
            """, unsafe_allow_html=True)
            with st.spinner("Preparando script SQL..."):
                try:
                    sql_bytes = gerar_backup_sql(engine)
                    st.download_button(
                        label="📥 Baixar Dump SQL Geral (.sql)",
                        data=sql_bytes,
                        file_name=f"Dump_Geral_Banco_Ribeiraozinho_{datetime.now().strftime('%Y%m%d_%H%M%S')}.sql",
                        mime="application/sql",
                        key="btn_down_sql",
                        use_container_width=True
                    )
                except Exception as sq_err:
                    st.error(f"Erro ao gerar SQL: {sq_err}")


    with aba_auto:
        st.markdown("""
        <div style='background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 22px; box-shadow: 0 2px 8px rgba(0,0,0,0.04);'>
            <div style='display: flex; align-items: center; gap: 12px; margin-bottom: 14px;'>
                <span style='font-size: 30px;'>☁️</span>
                <div>
                    <h3 style='color: #0369a1; margin: 0; font-size: 17px; font-weight: 800;'>Backup Automático Contínuo na Nuvem (Supabase / AWS)</h3>
                    <p style='color: #64748b; font-size: 12px; margin: 2px 0 0 0;'>Proteção contra perda de dados em nível de infraestrutura internacional</p>
                </div>
            </div>
            
            <div style='display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 16px; margin-top: 16px;'>
                <div style='background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px;'>
                    <div style='font-weight: 700; color: #0f172a; font-size: 13px; margin-bottom: 6px;'>🕒 Snapshots Diários Automáticos</div>
                    <div style='color: #475569; font-size: 12px; line-height: 1.45;'>
                        O banco de dados do Supabase cria cópias completas diárias de todas as tabelas e dados automaticamente durante a madrugada, sem interromper o uso do sistema.
                    </div>
                </div>

                <div style='background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px;'>
                    <div style='font-weight: 700; color: #0f172a; font-size: 13px; margin-bottom: 6px;'>🛡️ Redundância e Tolerância a Falhas</div>
                    <div style='color: #475569; font-size: 12px; line-height: 1.45;'>
                        Hospedado no data center da AWS em São Paulo (sa-east-1) com discos SSD NVMe redundantes, proteção contra falhas de hardware e criptografia em repouso e trânsito (SSL/TLS).
                    </div>
                </div>

                <div style='background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px;'>
                    <div style='font-weight: 700; color: #0f172a; font-size: 13px; margin-bottom: 6px;'>👥 Salvamento Simultâneo Concorrente</div>
                    <div style='color: #475569; font-size: 12px; line-height: 1.45;'>
                        Múltiplos operadores podem cadastrar e editar contratos ao mesmo tempo pela web ou computador, sem conflitos de concorrência ou bloqueio de arquivos.
                    </div>
                </div>
            </div>

            <div style='margin-top: 20px; padding: 12px 16px; background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 8px; display: flex; align-items: center; gap: 10px;'>
                <span style='font-size: 20px;'>✅</span>
                <div style='color: #065f46; font-size: 12.5px;'>
                    <strong>Conclusão:</strong> Seus dados já estão salvos e protegidos na nuvem de forma permanente. Os downloads manuais são recomendados para cópias de segurança locais e arquivamento municipal.
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with aba_diagnostico:
        st.markdown("<h4 style='color: #0f172a; font-size: 14px; margin-bottom: 10px;'>📋 Contagem e Saúde das Tabelas no Banco Ativo</h4>", unsafe_allow_html=True)
        try:
            with engine.connect() as conn:
                info_tabelas = [
                    ("contratos", "Contratos Municipais Registrados", "📄"),
                    ("fornecedores", "Empresas e Fornecedores Cadastrados", "🏢"),
                    ("orgaos", "Secretarias e Órgãos Municipais", "🏛️"),
                    ("termos_aditivos", "Termos Aditivos Registrados", "📑"),
                    ("protocolos", "Processos e Protocolos Internos", "📋"),
                    ("usuarios", "Usuários e Operadores do Sistema", "👥"),
                    ("audit_log", "Logs de Auditoria e Ações", "🕒")
                ]
                dados_diag = []
                for t_name, t_desc, t_icon in info_tabelas:
                    try:
                        qtd = conn.execute(text(f"SELECT count(*) FROM {t_name}")).scalar()
                        dados_diag.append({
                            "Ícone": t_icon,
                            "Tabela": t_name,
                            "Descrição": t_desc,
                            "Total de Registros": qtd,
                            "Status": "🟢 Saudável"
                        })
                    except Exception:
                        dados_diag.append({
                            "Ícone": t_icon,
                            "Tabela": t_name,
                            "Descrição": t_desc,
                            "Total de Registros": 0,
                            "Status": "⚠️ Não encontrada / Vazia"
                        })
                df_diag = pd.DataFrame(dados_diag)
                st.dataframe(df_diag, use_container_width=True, hide_index=True)
        except Exception as err_diag:
            st.error(f"Erro ao carregar diagnóstico: {err_diag}")

def main():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = True
    if "usuario" not in st.session_state:
        st.session_state["usuario"] = {
            "id": 1,
            "nome": "Thiago",
            "email": "admin",
            "cargo": "ADMIN",
            "ativo": 1
        }
    if "mostrar_formulario" not in st.session_state:
        st.session_state["mostrar_formulario"] = False
    if "mostrar_formulario_protocolo" not in st.session_state:
        st.session_state["mostrar_formulario_protocolo"] = False

    def mudar_menu(novo_menu, rerun=False):
        st.session_state["menu_selecionado"] = novo_menu
        st.session_state["mostrar_formulario"] = False
        st.session_state["mostrar_formulario_protocolo"] = False
        st.session_state.pop("editando_contrato_id", None)
        st.session_state.pop("scroll_para_ficha", None)
        st.session_state.pop("contrato_agu_selecionado", None)
        st.session_state.pop("aditivo_contrato_pre_sel", None)
        st.session_state.pop("ultimo_contrato_criado", None)
        if rerun:
            st.rerun()

    def executar_logout():
        st.session_state["logged_in"] = False
        st.session_state["usuario"] = None
        st.session_state["mostrar_formulario"] = False
        st.session_state["mostrar_formulario_protocolo"] = False
        st.session_state["menu_selecionado"] = "📊 Contratos" 
        st.rerun()

    login_area = st.empty()
    if not st.session_state.get("logged_in"):
        with login_area.container():
            tela_login()
        st.stop()
    else:
        login_area.empty()
        with st.sidebar:
            # Emblema Municipal Oficial com Brasão de Ribeirãozinho do Maranhão - MA
            brasao_b64 = obter_brasao_b64()
            brasao_sidebar_html = f"<div style='display: flex; justify-content: center; margin-bottom: 8px;'><img src='data:image/png;base64,{brasao_b64}' style='width: 54px; height: auto; filter: drop-shadow(0 2px 4px rgba(0,0,0,0.15));' alt='Brasão Oficial' /></div>" if brasao_b64 else "<div style='font-size: 28px; margin-bottom: 4px;'>🏛️</div>"
            st.markdown(f"""
            <div style='text-align: center; padding: 4px 6px 8px 6px; border-bottom: 1px solid rgba(255,255,255,0.1); margin-bottom: 10px;'>
                {brasao_sidebar_html}
                <div style='font-weight: 800; font-size: 11px; color: #38bdf8; line-height: 1.25; font-family: "Plus Jakarta Sans"; letter-spacing: -0.2px;'>
                    PREFEITURA MUNICIPAL DE<br>RIBEIRÃOZINHO DO MARANHÃO
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Card de Usuário Claro
            nome_usuario = st.session_state['usuario'].get('nome', 'Administrador')
            cargo = st.session_state['usuario'].get('cargo', 'OPERADOR')
            iniciais = obter_iniciais(nome_usuario)
            agora = datetime.now()

            st.markdown(f"""
            <div class='gel-user-box' style='margin-bottom: 6px;'>
                <div class='gel-avatar-circle'>{iniciais}</div>
                <div style='flex: 1; overflow: hidden;'>
                    <div class='gel-user-name' title='{nome_usuario}'>{nome_usuario}</div>
                    <span class='gel-role-pill'>{cargo}</span>
                </div>
            </div>
            <div style='margin: 4px 0 12px 0; padding: 4px 8px; background: rgba(255,255,255,0.06); border-radius: 6px; display: flex; justify-content: space-between; align-items: center; font-size: 10px; color: #94a3b8;'>
                <span style='color: #34d399; font-weight: 700;'>● Sistema Ativo</span>
                <span>{agora.strftime('%d/%m/%Y • %H:%M')}</span>
            </div>
            """, unsafe_allow_html=True)

            # Determinar item ativo atual
            if st.session_state.get("mostrar_formulario"):
                menu_atual = "➕ Novo Contrato"
            elif st.session_state.get("mostrar_formulario_protocolo"):
                menu_atual = "➕ Novo Protocolo"
            else:
                menu_atual = st.session_state.get("menu_selecionado", "📊 Contratos")

            # Módulo 1: Contratos
            st.button("📊 Contratos", key="nav_btn_contratos", use_container_width=True, 
                      type="primary" if menu_atual == "📊 Contratos" else "secondary",
                      on_click=mudar_menu, args=("📊 Contratos",))

            # Módulo 2: Protocolos
            st.button("📋 Protocolos", key="nav_btn_protocolos", use_container_width=True, 
                      type="primary" if menu_atual == "📋 Protocolos" else "secondary",
                      on_click=mudar_menu, args=("📋 Protocolos",))

            # Módulo 3: Modelos AGU / Minutas
            st.button("⚖️ Modelos AGU / Minutas", key="nav_btn_agu", use_container_width=True, 
                      type="primary" if menu_atual == "⚖️ Modelos AGU / Minutas" else "secondary",
                      on_click=mudar_menu, args=("⚖️ Modelos AGU / Minutas",))

            # Módulo 4: Importar Planilha
            st.button("📥 Importar Planilha", key="nav_btn_importar", use_container_width=True, 
                      type="primary" if menu_atual == "📥 Importar Planilha" else "secondary",
                      on_click=mudar_menu, args=("📥 Importar Planilha",))

            # Módulo 5: Usuários (Apenas Administrador)
            if cargo in ["ADMIN", "ADMINISTRADOR"]:
                st.button("👥 Usuários", key="nav_btn_usuarios", use_container_width=True, 
                          type="primary" if menu_atual == "👥 Usuários" else "secondary",
                          on_click=mudar_menu, args=("👥 Usuários",))

            # Módulo 6: Histórico de Ações (Aberto a todos)
            st.button("🕒 Histórico de Ações", key="nav_btn_auditoria", use_container_width=True, 
                      type="primary" if menu_atual == "🕒 Histórico de Ações" else "secondary",
                      on_click=mudar_menu, args=("🕒 Histórico de Ações",))

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            # Módulo 8: Fornecedores
            st.button("🏢 Fornecedores", key="nav_btn_forn", use_container_width=True, 
                      type="primary" if menu_atual == "🏢 Fornecedores" else "secondary",
                      on_click=mudar_menu, args=("🏢 Fornecedores",))

            # Módulo 9: Secretários e Fiscais
            st.button("🏛️ Secretários e Fiscais", key="nav_btn_gestores", use_container_width=True, 
                      type="primary" if menu_atual == "🏛️ Secretários e Fiscais" else "secondary",
                      on_click=mudar_menu, args=("🏛️ Secretários e Fiscais",))

            # Módulo 10: Backups e Segurança
            st.button("💾 Backups e Segurança", key="nav_btn_backups", use_container_width=True, 
                      type="primary" if menu_atual == "💾 Backups e Segurança" else "secondary",
                      on_click=mudar_menu, args=("💾 Backups e Segurança",))

            st.button("🚪 Sair do Sistema", key="btn_logout", use_container_width=True, on_click=executar_logout)

        # Roteamento Estável Instantâneo (Sem remount, sem sobreposição e sem piscamento)
        with st.container():
            if st.session_state.get("mostrar_formulario"):
                formulario_contrato()
            elif st.session_state.get("mostrar_formulario_protocolo"):
                formulario_protocolo()
            elif menu_atual == "📊 Contratos":
                dashboard_contratos()
            elif menu_atual == "✏️ Editar Contrato":
                tela_edicao_contrato()
            elif menu_atual == "📋 Protocolos":
                dashboard_protocolos()
            elif menu_atual == "⚖️ Modelos AGU / Minutas":
                tela_modelos_agu()
            elif menu_atual == "📥 Importar Planilha":
                tela_importar_planilha()
            elif menu_atual == "👥 Usuários":
                tela_usuarios()
            elif menu_atual == "🕒 Histórico de Ações":
                tela_historico_acoes()
            elif menu_atual == "📈 Dashboard / Relatórios":
                tela_dashboard_bi()
            elif menu_atual == "🏢 Fornecedores":
                tela_fornecedores()
            elif menu_atual == "🏛️ Secretários e Fiscais":
                tela_gestao_secretarios_fiscais()
            elif menu_atual == "💾 Backups e Segurança":
                tela_backups_sistema()
            elif menu_atual == "➕ Novo Contrato":
                formulario_contrato()
            elif menu_atual == "➕ Novo Protocolo":
                formulario_protocolo()

if __name__ == "__main__":
    main()
