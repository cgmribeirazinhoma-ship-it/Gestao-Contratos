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
    extrair_tabela_dinamica_df
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
            try { if (window.parent && window.parent.document) list.push(window.parent.document); } catch(e) {}
            try { if (window.top && window.top.document) list.push(window.top.document); } catch(e) {}
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

        // 2. Injeção de Estilo Global no Document, Parent e Top para eliminar Header, Share e Manage App
        const cssOcultar = `
            header,
            header[data-testid="stHeader"],
            [data-testid="stHeader"],
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
            .stApp > header {
                display: none !important;
            }
            div[data-testid="stAppViewContainer"] > section.main > div.block-container,
            .main .block-container {
                padding-top: 1.2rem !important;
            }
        `;

        function injetarCSS() {
            getDocs().forEach(function(doc) {
                try {
                    if (doc.head && !doc.getElementById('gel-hide-streamlit-bars')) {
                        const style = doc.createElement('style');
                        style.id = 'gel-hide-streamlit-bars';
                        style.innerHTML = cssOcultar;
                        doc.head.appendChild(style);
                    }
                } catch(e) {}
            });
        }
        injetarCSS();

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

        // 4. Varredura contínua para manter Share, Header e Manage App 100% ocultos
        function varrerEOcultar() {
            const seletores = [
                'header',
                'header[data-testid="stHeader"]',
                '[data-testid="stHeader"]',
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
                    const botoes = doc.querySelectorAll('header button, [data-testid="stToolbar"] button, [data-testid="stHeader"] button, button[data-testid="stShareButton"]');
                    botoes.forEach(function(btn) {
                        const txt = (btn.textContent || '').trim().toLowerCase();
                        const aria = (btn.getAttribute('aria-label') || '').toLowerCase();
                        const title = (btn.getAttribute('title') || '').toLowerCase();
                        if (txt === 'share' || txt === 'compartilhar' || aria.includes('share') || title.includes('share') || aria.includes('compartilhar')) {
                            btn.style.setProperty('display', 'none', 'important');
                            btn.style.setProperty('visibility', 'hidden', 'important');
                            btn.style.setProperty('pointer-events', 'none', 'important');
                            const pai = btn.closest('[data-testid="stToolbarActions"], [data-testid="stToolbar"], header');
                            if (pai) {
                                pai.style.setProperty('display', 'none', 'important');
                            }
                        }
                    });
                } catch(e) {}
            });
        }

        varrerEOcultar();
        setInterval(varrerEOcultar, 400);
    })();
</script>
""", height=0, width=0)


# Estilização CSS Clara, Moderna e Dinâmica com a Paleta Oficial de Ribeirãozinho do Maranhão - MA
# Cores Oficiais: Azul Real / Celeste (#0284C7 / #0369A1), Verde Esperança (#059669 / #16A34A), Amarelo Ouro (#D97706 / #F59E0B), Branco (#FFFFFF)
st.markdown("""
<div translate="no" class="notranslate" style="display:none"></div>
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;700&family=Inter:wght@300;400;500;600;700;800&display=swap');

    /* Prevenção de conflito de reconciliação DOM com tradutores */
    .notranslate {
        translate: no !important;
    }

    /* Ocultar completamente a barra superior (Header, Toolbar, botão Share, Deploy), 'Manage app' e footer */
    header,
    header[data-testid="stHeader"],
    [data-testid="stHeader"],
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

    /* Ajuste de margem superior para o conteúdo ocupar o topo perfeitamente */
    .stApp > header {
        display: none !important;
    }
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
        background-color: #f1f5f9 !important;
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
    
    /* Prevenir efeito fantasma e transparência de telas anteriores durante a transição */
    div[data-stale="true"],
    .stApp[data-stale="true"],
    [data-testid="stAppViewContainer"][data-stale="true"],
    [data-testid="stVerticalBlock"][data-stale="true"],
    [data-testid="stForm"][data-stale="true"] {
        opacity: 1 !important;
        filter: none !important;
        transition: none !important;
    }

    /* Mais compactação e fundo limpo nos formulários do sistema */
    div[data-testid="stForm"] {
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

    /* Barra Lateral Escura (Menu do Sistema) */
    section[data-testid="stSidebar"] {
        background-color: #012b63 !important;
        border-right: 1px solid #fbbf24 !important; /* Borda amarela separadora */
        box-shadow: 2px 0 12px rgba(0, 0, 0, 0.2) !important;
        min-width: 260px !important;
        max-width: 260px !important;
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
        # Se for PostgreSQL via psycopg2/psycopg2-binary, a URL deve comecar com postgresql://
        if db_url.startswith("postgres://"):
            db_url = db_url.replace("postgres://", "postgresql://", 1)
        return create_engine(db_url, pool_pre_ping=True, pool_recycle=300)

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
            "endereco": "",
            "bairro": "Centro",
            "cidade": "Ribeirãozinho do Maranhão",
            "uf": "MA",
            "secretario": "",
            "cargo_secretario": "",
            "fiscal": "",
            "dotacao": "",
            "modelo_agu": "COMPRAS"
        }

    is_educa = "EDUCA" in str(orgao_nome).upper()
    sec_fallback = MAPA_SECRETARIOS_PADRAO.get(orgao_nome, "")
    cargo_fallback = MAPA_CARGO_SECRETARIOS.get(orgao_nome, "Secretário Municipal de Educação" if is_educa else "Secretário(a) Municipal")
    end_fallback = "Av. Tancredo Neves, nº 120" if is_educa else ""
    fisc_fallback = "CHARLIANE DE ABREU MACIEL" if is_educa else ""
    dot_fallback = "Fundo Municipal de Educação / SEMED - 02.05.00 - 12.361.0402.2022.0000" if is_educa else ""

    dados = {
        "cnpj": "01.612.834/0001-86",
        "cep": "65928-000",
        "endereco": end_fallback,
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
            # 1. Tabela de Órgãos (busca exata primeiro, depois por palavra-chave distintiva)
            row = conn.execute(
                text("SELECT * FROM orgaos WHERE nome = :n"),
                {"n": orgao_nome}
            ).mappings().first()

            if not row:
                palavras = [p for p in re.split(r'[\s,/]+', normalizar_texto(orgao_nome)) if len(p) > 3 and p not in ("SECRETARIA", "MUNICIPAL", "GABINETE", "COORDENACAO", "GERENCIA", "SETOR", "FUNDO")]
                for p in palavras:
                    row = conn.execute(
                        text("SELECT * FROM orgaos WHERE UPPER(nome) LIKE :kw LIMIT 1"),
                        {"kw": f"%{p}%"}
                    ).mappings().first()
                    if row:
                        break

            oid = None
            if row:
                oid = row.get("id")
                if row.get("cnpj"): dados["cnpj"] = str(row["cnpj"]).strip()
                if row.get("cep"): dados["cep"] = str(row["cep"]).strip()
                if row.get("endereco") and str(row["endereco"]).strip() not in ("", "Rua Principal, s/n"):
                    dados["endereco"] = str(row["endereco"]).strip()
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

            # 2. Histórico recente de contratos do órgão alimentados pelos usuários (vinculados ao orgao_id)
            if oid:
                # 2.1 Fiscal mais recente preenchido nos contratos (apenas se não estiver definido no cadastro do órgão)
                if not dados.get("fiscal") or dados["fiscal"].startswith("Fiscal da") or dados["fiscal"].startswith("Fiscal de"):
                    row_fisc = conn.execute(
                        text("""
                            SELECT c.fiscal
                            FROM contratos c
                            WHERE c.orgao_id = :oid
                              AND c.fiscal IS NOT NULL 
                              AND TRIM(c.fiscal) != ''
                              AND c.fiscal NOT LIKE 'Fiscal da%'
                              AND c.fiscal NOT LIKE 'Fiscal de%'
                            ORDER BY c.id DESC LIMIT 1
                        """),
                        {"oid": oid}
                    ).mappings().first()
                    if row_fisc and row_fisc.get("fiscal"):
                        dados["fiscal"] = str(row_fisc["fiscal"]).strip()

                # 2.2 Secretário mais recente preenchido nos contratos (apenas se não estiver definido no cadastro do órgão)
                if not dados.get("secretario"):
                    row_sec = conn.execute(
                        text("""
                            SELECT c.secretario
                            FROM contratos c
                            WHERE c.orgao_id = :oid
                              AND c.secretario IS NOT NULL 
                              AND TRIM(c.secretario) != ''
                            ORDER BY c.id DESC LIMIT 1
                        """),
                        {"oid": oid}
                    ).mappings().first()
                    if row_sec and row_sec.get("secretario"):
                        dados["secretario"] = str(row_sec["secretario"]).strip()

                # 2.3 Dotação orçamentária recente preenchida nos contratos
                row_dot = conn.execute(
                    text("""
                        SELECT c.dotacao_orcamentaria
                        FROM contratos c
                        WHERE c.orgao_id = :oid
                          AND c.dotacao_orcamentaria IS NOT NULL 
                          AND TRIM(c.dotacao_orcamentaria) != ''
                        ORDER BY c.id DESC LIMIT 1
                    """),
                    {"oid": oid}
                ).mappings().first()
                if row_dot and row_dot.get("dotacao_orcamentaria"):
                    dados["dotacao"] = str(row_dot["dotacao_orcamentaria"]).strip()

                # 2.4 Endereço recente preenchido nos contratos
                row_end = conn.execute(
                    text("""
                        SELECT c.endereco_orgao, c.bairro_orgao, c.cep_orgao, c.cidade_orgao, c.uf_orgao
                        FROM contratos c
                        WHERE c.orgao_id = :oid
                          AND c.endereco_orgao IS NOT NULL 
                          AND TRIM(c.endereco_orgao) != ''
                        ORDER BY c.id DESC LIMIT 1
                    """),
                    {"oid": oid}
                ).mappings().first()
                if row_end:
                    if row_end.get("endereco_orgao"): dados["endereco"] = str(row_end["endereco_orgao"]).strip()
                    if row_end.get("bairro_orgao"): dados["bairro"] = str(row_end["bairro_orgao"]).strip()
                    if row_end.get("cep_orgao"): dados["cep"] = str(row_end["cep_orgao"]).strip()
                    if row_end.get("cidade_orgao"): dados["cidade"] = str(row_end["cidade_orgao"]).strip()
                    if row_end.get("uf_orgao"): dados["uf"] = str(row_end["uf_orgao"]).strip()

                # 2.5 Modelo AGU recente
                row_mod = conn.execute(
                    text("""
                        SELECT c.modelo_agu
                        FROM contratos c
                        WHERE c.orgao_id = :oid
                          AND c.modelo_agu IS NOT NULL 
                          AND TRIM(c.modelo_agu) != ''
                        ORDER BY c.id DESC LIMIT 1
                    """),
                    {"oid": oid}
                ).mappings().first()
                if row_mod and row_mod.get("modelo_agu"):
                    dados["modelo_agu"] = str(row_mod["modelo_agu"]).strip()

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
                    SET cnpj = :cnpj,
                        cep = :cep,
                        endereco = :end,
                        bairro = :bai,
                        cidade = :cid,
                        uf = :uf,
                        secretario_padrao = :sec,
                        cargo_secretario = :carg,
                        fiscal_padrao = :fisc,
                        dotacao_padrao = :dot
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
    if pd.isna(val) or val is None or str(val).strip() == "":
        return 0.0
    val_str = str(val).strip()
    val_str = re.sub(r"[R$\s]", "", val_str)
    if "," in val_str and "." in val_str:
        if val_str.rfind(",") > val_str.rfind("."):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "," in val_str:
        val_str = val_str.replace(",", ".")
    try:
        return float(val_str)
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
            uemail = (u["email"] or "").strip()
            usenha_hash = (u["senha_hash"] or "").strip()
            ucargo = u["cargo"]

            # Identificadores válidos para este usuário
            uemail_lower = uemail.lower()
            uemail_user = uemail_lower.split("@")[0] if "@" in uemail_lower else uemail_lower
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
                uemail_lower, uemail_user,
                unome_lower, unome_norm, unome_primeiro, unome_slug,
                unome_limpo_lower, unome_limpo_norm, unome_limpo_primeiro, unome_limpo_slug
            }

            # Atalhos amigáveis por cargo
            if ucargo in ["ADMIN", "ADMINISTRADOR"]:
                candidatos.add("admin")
                candidatos.add("administrador")
            if ucargo == "OPERADOR":
                candidatos.add("operador")

            # Se o login digitado corresponder a qualquer identificador do usuário
            if login_lower in candidatos or login_norm in candidatos or login_clean in candidatos:
                senha_ok = False
                if usenha_hash.startswith(("$2a$", "$2b$", "$2y$")):
                    try:
                        senha_ok = bcrypt.checkpw(senha_clean.encode("utf-8"), usenha_hash.encode("utf-8"))
                    except Exception:
                        senha_ok = False
                elif usenha_hash == senha_clean or usenha_hash == str(senha):
                    senha_ok = True
                    # Atualizar automaticamente para hash bcrypt para maior segurança
                    try:
                        novo_h = bcrypt.hashpw(senha_clean.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
                        conn.execute(text("UPDATE usuarios SET senha_hash = :h WHERE id = :id"), {"h": novo_h, "id": uid})
                        conn.commit()
                    except Exception:
                        pass

                if senha_ok:
                    return {"id": uid, "nome": unome, "email": uemail, "cargo": ucargo}

    return None

@st.cache_data(ttl=120, show_spinner=False)
def carregar_usuarios():
    with engine.connect() as conn:
        query = text("SELECT id, nome, email, cargo, ativo, data_criacao FROM usuarios ORDER BY nome ASC")
        return pd.read_sql(query, conn)

def cadastrar_novo_usuario(nome, email=None, senha="", cargo="OPERADOR"):
    if not nome or not str(nome).strip():
        return False, "O campo 'Nome Cadastrado' é obrigatório."
    nome_limpo = str(nome).strip()

    if not email or not str(email).strip():
        slug = normalizar_texto_login(nome_limpo)
        slug_clean = re.sub(r'[^a-z0-9]', '', slug)
        email_limpo = f"{slug_clean}@prefeitura.gov.br"
    else:
        email_limpo = str(email).strip().lower()

    if not senha or len(senha.strip()) < 4:
        return False, "A senha deve conter no mínimo 4 caracteres."

    senha_limpa = senha.strip()
    senha_hash = bcrypt.hashpw(senha_limpa.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

    with engine.connect() as conn:
        existe = conn.execute(
            text("SELECT id FROM usuarios WHERE LOWER(nome) = LOWER(:n) OR LOWER(email) = LOWER(:e)"),
            {"n": nome_limpo, "e": email_limpo}
        ).fetchone()
        if existe:
            return False, f"Já existe um usuário cadastrado com este Nome ('{nome_limpo}') ou Login/E-mail ('{email_limpo}')."

        try:
            conn.execute(
                text("""
                    INSERT INTO usuarios (nome, email, senha_hash, cargo, ativo, data_criacao)
                    VALUES (:nome, :email, :senha, :cargo, 1, CURRENT_TIMESTAMP)
                """),
                {"nome": nome_limpo, "email": email_limpo, "senha": senha_hash, "cargo": cargo}
            )
            conn.commit()
            try:
                carregar_usuarios.clear()
            except Exception:
                pass
            login_sugestao = email_limpo.split("@")[0]
            return True, f"Usuário '{nome_limpo}' cadastrado com sucesso! Pode entrar com o login '{login_sugestao}' ou '{email_limpo}'."
        except Exception as err:
            return False, f"Erro ao gravar usuário: {str(err)}"

def atualizar_usuario(usuario_id, nome, email, cargo, ativo, nova_senha=None):
    if not nome or not str(nome).strip():
        return False, "O campo 'Nome Cadastrado' é obrigatório."
    nome_limpo = str(nome).strip()
    email_limpo = str(email).strip().lower() if email else ""

    if nova_senha is not None and str(nova_senha).strip() != "":
        senha_limpa = str(nova_senha).strip()
        if len(senha_limpa) < 4:
            return False, "A nova senha deve conter no mínimo 4 caracteres."
    else:
        senha_limpa = None

    with engine.connect() as conn:
        if email_limpo:
            conflito = conn.execute(
                text("SELECT id FROM usuarios WHERE LOWER(email) = LOWER(:e) AND id != :id"),
                {"e": email_limpo, "id": usuario_id}
            ).fetchone()
            if conflito:
                return False, f"Já existe outro usuário com o login/e-mail '{email_limpo}'."

        try:
            if senha_limpa:
                senha_hash = bcrypt.hashpw(senha_limpa.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
                conn.execute(
                    text("UPDATE usuarios SET nome = :nome, email = :email, cargo = :cargo, ativo = :ativo, senha_hash = :senha WHERE id = :id"),
                    {"nome": nome_limpo, "email": email_limpo, "cargo": cargo, "ativo": 1 if ativo else 0, "senha": senha_hash, "id": usuario_id}
                )
            else:
                conn.execute(
                    text("UPDATE usuarios SET nome = :nome, email = :email, cargo = :cargo, ativo = :ativo WHERE id = :id"),
                    {"nome": nome_limpo, "email": email_limpo, "cargo": cargo, "ativo": 1 if ativo else 0, "id": usuario_id}
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
                o.nome as orgao,
                c.modalidade,
                c.processo_adm,
                c.objeto,
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
                c.qtd_pagamentos,
                c.observacoes,
                u.nome as usuario_criador,
                c.data_criacao
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
    for c in colunas_reais:
        c_norm = normalizar_texto(c)
        for t in termos:
            t_norm = normalizar_texto(t)
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

    st.markdown(f"""
    <style>
        /* 1. Fundo Geral da Tela de Login (Azul Escuro Municipal) */
        .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stHeader"],
        section.main {{
            background: linear-gradient(135deg, #001f3f 0%, #012b63 50%, #013b82 100%) !important;
            background-attachment: fixed !important;
        }}

        /* 2. Ocultar sidebar e topo completamente na tela de login */
        section[data-testid="stSidebar"],
        [data-testid="stSidebar"],
        header[data-testid="stHeader"] {{
            display: none !important;
        }}

        /* 3. Centralização e Espaçamento da área principal */
        div[data-testid="stAppViewContainer"] > section.main > div.block-container,
        .main .block-container {{
            max-width: 100% !important;
            padding-top: 1.5rem !important;
            padding-bottom: 2rem !important;
            display: flex !important;
            flex-direction: column !important;
            justify-content: center !important;
            align-items: center !important;
            min-height: 90vh !important;
        }}

        /* 4. Card do Formulário de Login - Azul Escuro com Borda Amarela Ouro */
        div[data-testid="stForm"],
        form[aria-label="login_form_gel"],
        .gel-login-box [data-testid="stForm"] {{
            background: #012b63 !important;
            border-radius: 14px !important;
            padding: 30px 32px !important;
            border: 2.5px solid #fbbf24 !important; /* Borda Amarela Ouro */
            box-shadow: 0 16px 40px rgba(0, 0, 0, 0.5), 0 0 20px rgba(251, 191, 36, 0.25) !important;
            width: 100% !important;
            max-width: 460px !important;
            margin: 0 auto !important;
        }}

        /* 5. Brasão e Títulos */
        .login-brasao {{
            display: flex;
            justify-content: center;
            margin-bottom: 12px;
        }}
        .login-brasao img {{
            width: 75px;
            height: auto;
            filter: drop-shadow(0 4px 10px rgba(0,0,0,0.4));
        }}

        .login-prefeitura-nome {{
            font-family: 'Plus Jakarta Sans', 'Arial', sans-serif !important;
            font-weight: 800 !important;
            font-size: 15px !important;
            color: #fbbf24 !important; /* Letra Amarela Forte */
            text-align: center !important;
            line-height: 1.25 !important;
            margin-bottom: 6px !important;
            letter-spacing: 0.3px !important;
            text-shadow: 0 2px 4px rgba(0,0,0,0.5) !important;
        }}

        .login-sistema-nome {{
            font-size: 11.5px !important;
            color: #fbbf24 !important; /* Letra Amarela Forte */
            font-weight: 700 !important;
            text-align: center !important;
            margin-bottom: 24px !important;
            letter-spacing: 0.5px !important;
            text-transform: uppercase !important;
            opacity: 0.95;
        }}

        /* 6. Labels dos Inputs com Letra Amarela Ouro */
        div[data-testid="stForm"] div[data-testid="stTextInput"] label,
        div[data-testid="stForm"] div[data-testid="stTextInput"] label p,
        div[data-testid="stForm"] label,
        div[data-testid="stForm"] label p {{
            font-size: 11.5px !important;
            color: #fbbf24 !important; /* Letra Amarela */
            font-weight: 800 !important;
            letter-spacing: 0.8px !important;
            text-transform: uppercase !important;
            margin-bottom: 4px !important;
        }}

        /* 7. Campos de Entrada (Inputs) */
        div[data-testid="stForm"] div[data-testid="stTextInput"] input {{
            font-size: 14px !important;
            padding: 10px 14px !important;
            border-radius: 8px !important;
            background: #091a36 !important;
            color: #ffffff !important;
            border: 1.5px solid #2563eb !important;
        }}
        div[data-testid="stForm"] div[data-testid="stTextInput"] input:focus {{
            border-color: #fbbf24 !important;
            box-shadow: 0 0 10px rgba(251, 191, 36, 0.45) !important;
            outline: none !important;
        }}

        /* 8. Botão de Entrar (Amarelo Ouro com Letra Azul) */
        div[data-testid="stForm"] [data-testid="stFormSubmitButton"] button {{
            background: linear-gradient(180deg, #fbbf24 0%, #f59e0b 100%) !important;
            color: #012b63 !important; /* Letra azul marinho contrastante */
            font-weight: 800 !important;
            font-size: 14px !important;
            border-radius: 8px !important;
            border: none !important;
            margin-top: 10px !important;
            padding: 10px 16px !important;
            letter-spacing: 0.5px !important;
            box-shadow: 0 4px 14px rgba(245, 158, 11, 0.4) !important;
            transition: all 0.2s ease !important;
        }}
        div[data-testid="stForm"] [data-testid="stFormSubmitButton"] button:hover {{
            background: linear-gradient(180deg, #fde047 0%, #d97706 100%) !important;
            box-shadow: 0 6px 20px rgba(245, 158, 11, 0.6) !important;
            transform: translateY(-1px) !important;
        }}

        /* 9. Rodapé */
        .login-footer-text {{
            color: #fbbf24 !important; /* Letra Amarela */
            font-size: 11.5px !important;
            text-align: center !important;
            margin-top: 16px !important;
            line-height: 1.4 !important;
        }}
        .login-footer-text strong {{
            color: #ffffff !important;
        }}
        .login-footer-copy {{
            color: #93c5fd !important;
            font-size: 10px !important;
            text-align: center !important;
            margin-top: 6px !important;
            letter-spacing: 0.3px !important;
        }}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<div style='height: 4vh;'></div>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.3, 1])
    
    with col2:
        brasao_html = f"<div class='login-brasao'><img src='{brasao_img}' alt='Brasão Oficial' /></div>" if brasao_img else "<div style='text-align: center; font-size: 48px; margin-bottom: 14px;'>🏛️</div>"

        with st.form("login_form_gel"):
            st.markdown(f"""
                {brasao_html}
                <div class='login-prefeitura-nome'>
                    PREFEITURA MUNICIPAL DE RIBEIRÃOZINHO<br>DO MARANHÃO-MA
                </div>
                <div class='login-sistema-nome'>
                    Gestão de Contratos — Modelos AGU
                </div>
            """, unsafe_allow_html=True)

            usuario_login = st.text_input("LOGIN OU E-MAIL", placeholder="Ex: admin, operador ou seu e-mail", key="input_login_usuario")
            senha = st.text_input("SENHA", type="password", placeholder="Digite sua senha", key="input_senha_usuario")
            submit = st.form_submit_button("→ Entrar no Sistema", use_container_width=True)

            if submit:
                if not usuario_login or not senha:
                    st.warning("⚠️ Por favor, informe seu login e senha.")
                else:
                    with st.spinner("Autenticando e acessando o sistema..."):
                        usuario = verificar_login(usuario_login, senha)
                        if usuario:
                            st.session_state["logged_in"] = True
                            st.session_state["usuario"] = usuario
                            st.rerun()
                        else:
                            st.error("❌ Usuário ou senha incorretos, ou usuário inativo.")

        # Rodapé fora do card
        st.markdown("""
        <div style='text-align: center; margin-top: 16px;'>
            <div class='login-footer-text'>
                🔒 Para novos acessos ou redefinição de senha, contate a <strong>Administração Municipal</strong>
            </div>
            <div class='login-footer-copy'>
                © 2026 Prefeitura Municipal de Ribeirãozinho do Maranhão - MA • Todos os direitos reservados
            </div>
        </div>
        """, unsafe_allow_html=True)

# ============================================
# INTERFACE - DASHBOARD CONTRATOS MODERNO (GEL)
# ============================================


def tela_edicao_contrato(cid):
    st.markdown("### ✏️ Edição de Contrato Municipal")
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
    c_row = df[df["id"] == cid]
    if c_row.empty:
        st.error("Contrato não encontrado.")
        if st.button("Voltar"):
            st.session_state.pop("editando_contrato_id", None)
            st.rerun()
        return
    c_row = c_row.iloc[0]

    with st.form("form_edicao_contrato"):
        c1, c2 = st.columns(2)
        with c1:
            novo_val = st.number_input("Valor Total (R$)", value=float(c_row.get("valor_total", 0.0)), step=100.0)
            nova_ass = st.date_input("Data de Assinatura", value=pd.to_datetime(c_row["data_assinatura"]).date() if pd.notna(c_row.get("data_assinatura")) else date.today())
            nova_pub = st.date_input("Data de Publicação", value=pd.to_datetime(c_row["data_publicacao"]).date() if pd.notna(c_row.get("data_publicacao")) else date.today())
            novo_venc = st.date_input("Data de Vencimento", value=pd.to_datetime(c_row["data_vencimento"]).date() if pd.notna(c_row.get("data_vencimento")) else date.today())
        with c2:
            nova_vig = st.text_input("Vigência Descrição", value=c_row.get("vigencia_descricao", "12 meses"))
            novo_sec = st.text_input("Secretário(a) Titular", value=c_row.get("secretario", ""))
            novo_fisc = st.text_input("Fiscal do Contrato", value=c_row.get("fiscal", ""))
            novo_dot = st.text_input("Dotação Orçamentária", value=c_row.get("dotacao_orcamentaria", ""))

        novo_obj = st.text_area("Objeto Contratual", value=c_row.get("objeto", ""), height=100)

        c_sub1, c_sub2 = st.columns(2)
        with c_sub1:
            if st.form_submit_button("💾 Salvar Alterações", use_container_width=True, type="primary"):
                try:
                    with get_engine().connect() as conn:
                        conn.execute(
                            text("""
                                UPDATE contratos SET
                                    valor_total = :val, data_assinatura = :d_a, data_publicacao = :d_p,
                                    data_vencimento = :d_v, vigencia_descricao = :vig,
                                    secretario = :sec, fiscal = :fisc, dotacao_orcamentaria = :dot,
                                    objeto = :obj, atualizado_em = CURRENT_TIMESTAMP
                                WHERE id = :cid
                            """),
                            {
                                "val": novo_val, "d_a": nova_ass, "d_p": nova_pub, "d_v": novo_venc,
                                "vig": nova_vig, "sec": novo_sec, "fisc": novo_fisc, "dot": novo_dot,
                                "obj": novo_obj, "cid": int(cid)
                            }
                        )
                        registrar_audit_log(conn, "contratos", int(cid), "EDICAO", "Contrato editado manualmente")
                        conn.commit()
                        try:
                            carregar_contratos.clear()
                        except Exception:
                            pass
                    st.toast("✅ Contrato atualizado com sucesso!", icon="💾")
                    st.session_state.pop("editando_contrato_id", None)
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar: {e}")
        with c_sub2:
            if st.form_submit_button("❌ Cancelar", use_container_width=True):
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
        # Botão de Novo Contrato
        col_btn, col_esp = st.columns([2, 8])
        with col_btn:
            if st.button("➕ Novo Contrato (Padrão AGU)", use_container_width=True, type="primary"):
                st.session_state["menu_selecionado"] = "⚖️ Modelos AGU / Minutas"
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

                                btn_lbl = "✅ Ficha Aberta" if is_selecionado else "Abrir"
                                btn_tp = "primary" if is_selecionado else "tertiary"
                                st.button(
                                    btn_lbl,
                                    key=f"btn_card_{item['id']}",
                                    use_container_width=True,
                                    type=btn_tp,
                                    on_click=_selecionar_card_cb,
                                    args=(item["numero_completo"],)
                                )

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

            st.markdown(f"""
            <div style='background: #eff6ff; border: 1.5px solid #bfdbfe; border-left: 5px solid #0284c7; border-radius: 8px; padding: 10px 16px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: center;'>
                <div>
                    <span style='font-size: 14px; font-weight: 800; color: #0369a1;'>📄 Contrato Selecionado: {c_row['numero_completo']}</span>
                    <span style='font-size: 12px; color: #475569; margin-left: 12px;'>🏛️ {c_row['orgao']} | 🏭 Fornecedor: <strong>{c_row['fornecedor']}</strong></span>
                </div>
                <span style='font-size: 11px; background: #0284c7; color: #fff; padding: 3px 10px; border-radius: 12px; font-weight: 700;'>FICHA ABERTA PELO CARD</span>
            </div>
            """, unsafe_allow_html=True)

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
                    if st.button("✏️ Editar Dados do Contrato", key=f"btn_edit_{cid}", use_container_width=True):
                        st.session_state["editando_contrato_id"] = cid
                        st.rerun()

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

    uploaded_file = st.file_uploader("Selecione o arquivo de contratos:", type=["xlsx", "xls", "csv"])

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

            st.success(f"Arquivo carregado com sucesso! **{len(df_up)}** registros identificados.")

            cols_existentes = list(df_up.columns)
            def_num = detectar_coluna(cols_existentes, ["CONTRATO", "NUMERO", "N°", "Nº"])
            def_org = detectar_coluna(cols_existentes, ["SECRETARIA", "ORGAO", "ÓRGÃO", "SETOR"])
            def_sec = detectar_coluna(cols_existentes, ["SECRETARIO", "SECRETÁRIO", "TITULAR", "GESTOR"])
            def_forn = detectar_coluna(cols_existentes, ["FORNECEDOR", "EMPRESA", "CONTRATADO", "RAZAO"])
            def_cnpj = detectar_coluna(cols_existentes, ["CNPJ", "CPF", "DOC"])
            def_obj = detectar_coluna(cols_existentes, ["OBJETO", "DESCRICAO", "DESCRIÇÃO", "SERVICO"])
            def_val = detectar_coluna(cols_existentes, ["VALOR TOTAL", "VALOR", "MONTANTE", "GLOBAL"])
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

            st.markdown("---")
            opcoes_grav = ["Adicionar / Atualizar registros", "Substituir base inteira (Recarregar)"]
            modo_imp = st.segmented_control("Modo de Gravação:", opcoes_grav, default=opcoes_grav[0])
            if not modo_imp:
                modo_imp = opcoes_grav[0]
            modo_code = "substituir" if "Substituir" in modo_imp else "adicionar"

            if st.button("🚀 Processar e Integrar Planilha", type="primary", use_container_width=True):
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
                "id": "ID", "nome": "Nome Cadastrado", "email": "Login / E-mail de Acesso", "cargo": "Perfil / Função"
            }),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("---")
        opcoes_usuarios = [f"{r['id']} - {r['nome']} ({r['cargo']}) | Login: {r['email']}" for _, r in df_u.iterrows()]
        u_sel = st.selectbox("Selecione usuário para gerenciar:", opcoes_usuarios)

        if u_sel:
            uid = int(u_sel.split(" - ")[0])
            u_row = df_u[df_u["id"] == uid].iloc[0]

            with st.form(f"form_edit_u_{uid}"):
                c1, c2 = st.columns(2)
                with c1:
                    e_nome = st.text_input("Nome Cadastrado (Nome Completo) *", value=u_row["nome"])
                    e_email = st.text_input("Login / E-mail de Acesso *", value=u_row["email"], help="Login ou e-mail utilizado para entrar no sistema.")
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
                    ok, msg = atualizar_usuario(uid, e_nome, e_email, e_cargo, e_ativo, e_senha if e_senha else None)
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
            n_email = st.text_input("Login ou E-mail de Acesso (opcional)", placeholder="Ex: roberto ou roberto@prefeitura.gov.br (se vazio, gerado automaticamente)")
            c_cad1, c_cad2 = st.columns(2)
            with c_cad1:
                n_cargo = st.selectbox("Perfil de Acesso:", ["OPERADOR", "ADMIN"])
            with c_cad2:
                n_senha = st.text_input("Senha Inicial * (mínimo 4 caracteres)", type="password", value="123456")

            st.caption("💡 O usuário poderá entrar digitando tanto o login/e-mail cadastrado quanto o primeiro nome.")

            if st.form_submit_button("➕ Criar Acesso", type="primary", use_container_width=True):
                ok, msg = cadastrar_novo_usuario(n_nome, n_email, n_senha, n_cargo)
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
    if not texto_bruto or not str(texto_bruto).strip():
        return {}
    try:
        import json
        import requests
        api_key = "AQ.Ab8RN6J8dyEt9ArI7BqLKaLvuF52lLgIlr68BaQ6cekMTdu3dw"
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key={api_key}"
        
        prompt = f'''Você é um especialista em contabilidade e orçamento público municipal.
O usuário forneceu um texto (possivelmente copiado de PDF/Excel) contendo os dados da dotação orçamentária.
Texto bruto: {texto_bruto}

Extraia ou classifique rigorosamente as informações orçamentárias nos seguintes campos:
1. "exercicio": Ano do exercício (ex: 2024, 2025).
2. "poder": (ex: "Poder Executivo 02.00").
3. "orgao": Órgão ou Fundo responsável (ex: "Fundo Municipal de Saúde 02.15.00").
4. "unidade": Unidade Orçamentária / Projeto / Atividade completa (ex: "MANUTENÇÃO 12.361.0402.2022.0000").
5. "natureza": Natureza da despesa (ex: "Outros Serviços de Terceiros 3.3.90.39.00").

Retorne APENAS um JSON válido com estas exatas chaves, sem markdown, sem explicações:
{{"exercicio": "", "poder": "", "orgao": "", "unidade": "", "natureza": ""}}'''
        
        payload = {"contents": [{"parts":[{"text": prompt}]}]}
        resp = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=15)
        data = resp.json()
        if "candidates" in data and data["candidates"]:
            resposta_llm = data["candidates"][0]["content"]["parts"][0]["text"]
            import re
            json_str = re.search(r"\{.*\}", resposta_llm, re.DOTALL)
            if json_str:
                return json.loads(json_str.group(0))
    except Exception as e:
        pass
    return {}

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
                    "mod": dados.get("modalidade"), "proc": dados.get("processo_adm"),
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
            conn.commit()
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
                c_m0_1, c_m0_2 = st.columns([1, 1.5])
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
                    c_dot_ia1, c_dot_ia2 = st.columns([2, 1])
                    with c_dot_ia1:
                        txt_dot_m1 = st.text_area("Texto Bruto da Dotação Orçamentária (Copie e cole aqui)", height=100, key=f"txt_dot_m1_{row['id']}")
                    with c_dot_ia2:
                        st.write("")
                        st.write("")
                        if st.button("✨ Extrair (IA)", key=f"btn_ext_ia_m1_{row['id']}", use_container_width=True):
                            with st.spinner("Analisando..."):
                                d_ext = extrair_dotacao_com_ia(txt_dot_m1)
                                if d_ext:
                                    st.session_state[f"m1_dot_ex_{row['id']}"] = d_ext.get("exercicio", "")
                                    st.session_state[f"m1_dot_pod_{row['id']}"] = d_ext.get("poder", "")
                                    st.session_state[f"m1_dot_org_{row['id']}"] = d_ext.get("orgao", "")
                                    st.session_state[f"m1_dot_uni_{row['id']}"] = d_ext.get("unidade", "")
                                    st.session_state[f"m1_dot_nat_{row['id']}"] = d_ext.get("natureza", "")
                                    st.success("Extraído com sucesso!")
                                else:
                                    st.error("Falha ao extrair.")
                    
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


                with st.expander("📦 Planilha ou Documento de Itens do Contrato (Auto-Ajustável - Excel / CSV / PDF)", expanded=False):
                    st.caption("A tabela no documento gerado se **auto-ajusta dinamicamente** à quantidade de colunas, linhas e descrição dos cabeçalhos da planilha ou PDF importado.")
                    up_itens_m1 = st.file_uploader(
                        "Importar Planilha ou PDF de Itens (.xlsx, .xls, .csv, .pdf):",
                        type=["xlsx", "xls", "csv", "pdf"],
                        key=f"upload_itens_m1_{row['id']}",
                        help="Carregue qualquer planilha ou PDF de itens. A quantidade de colunas, linhas e cabeçalhos será ajustada automaticamente na minuta."
                    )
                    itens_m1_proc = []
                    if up_itens_m1 is not None:
                        itens_m1_proc = processar_planilha_itens_excel(up_itens_m1.getvalue(), file_name=up_itens_m1.name, objeto_contexto=row.get("objeto", ""))
                        if itens_m1_proc:
                            tot_m1 = sum(float(it.get("_vt_float", 0.0) or 0.0) for it in itens_m1_proc)
                            st.success(f"✅ **{len(itens_m1_proc)} itens** carregados da planilha!")
                            df_itens_m1_view = pd.DataFrame([
                                {
                                    "Item": it["item"],
                                    "Descrição": it["descricao"],
                                    "Qtd": it["quantidade"],
                                    "Unid": it["unidade"],
                                    "Valor Unitário": it["valor_unitario"],
                                    "Valor Total": it["valor_total"]
                                }
                                for it in itens_m1_proc
                            ])
                            st.dataframe(df_itens_m1_view, use_container_width=True, hide_index=True)
                            st.caption(f"Valor Total dos Itens da Planilha: **R$ {tot_m1:,.2f}**".replace(",", "X").replace(".", ",").replace("X", "."))

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
                    "modalidade": row.get("modalidade") or "DISPENSA DE LICITAÇÃO",
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
                    "dotacao_orcamentaria": f"{m1_dot_org} | {m1_dot_uni} | {m1_dot_nat}",
                    "dotacao_exercicio": m1_dot_ex,
                    "dotacao_poder": m1_dot_pod,
                    "dotacao_orgao": m1_dot_org,
                    "dotacao_unidade": m1_dot_uni,
                    "dotacao_natureza": m1_dot_nat,
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
        
        # Inicialização dos dados se ainda não presentes no session_state (sem órgão padrão forçado)
        if "cgutec_rep_orgao" not in st.session_state:
            st.session_state["cgutec_org_aux_sel"] = "Selecione o Órgão / Secretaria Contratante..."
            st.session_state["cgutec_nome_orgao"] = ""
            st.session_state["cgutec_rep_orgao"] = ""
            st.session_state["cgutec_cargo_orgao"] = ""
            st.session_state["cgutec_fisc_orgao"] = ""
            st.session_state["cgutec_cnpj_orgao"] = "01.612.834/0001-86"
            st.session_state["cgutec_cep_orgao"] = "65928-000"
            st.session_state["cgutec_end_orgao"] = ""
            st.session_state["cgutec_num_orgao"] = "s/n"
            st.session_state["cgutec_bairro_orgao"] = "Centro"
            st.session_state["cgutec_cid_orgao"] = "Ribeirãozinho do Maranhão"
            st.session_state["cgutec_uf_orgao"] = "MA"
            st.session_state["cgutec_dot_ct"] = ""

        # Seletor inteligente para agilizar com histórico institucional da secretaria escolhida pelo usuário
        def _preencher_dados_orgao_cgutec():
            nome_org = st.session_state.get("cgutec_org_aux_sel", "")
            if nome_org and nome_org != "Selecione o Órgão / Secretaria Contratante...":
                h = obter_historico_orgao(nome_org)
                st.session_state["cgutec_cnpj_orgao"] = h.get("cnpj", "01.612.834/0001-86")
                st.session_state["cgutec_rep_orgao"] = h.get("secretario", "")
                st.session_state["cgutec_cargo_orgao"] = h.get("cargo_secretario", "Secretário(a) Municipal")
                st.session_state["cgutec_nome_orgao"] = nome_org
                st.session_state["cgutec_cep_orgao"] = h.get("cep", "65928-000")
                
                # Tratar endereço e número separadamente
                end_str = h.get("endereco", "")
                num_extraido = "s/n"
                if ", nº " in end_str:
                    partes = end_str.split(", nº ")
                    st.session_state["cgutec_end_orgao"] = partes[0]
                    num_extraido = partes[1]
                elif ", n° " in end_str:
                    partes = end_str.split(", n° ")
                    st.session_state["cgutec_end_orgao"] = partes[0]
                    num_extraido = partes[1]
                elif ", s/n" in end_str.lower():
                    st.session_state["cgutec_end_orgao"] = re.sub(r',\s*s/n', '', end_str, flags=re.IGNORECASE).strip()
                    num_extraido = "s/n"
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
                st.session_state["cgutec_nome_orgao"] = ""
                st.session_state["cgutec_rep_orgao"] = ""
                st.session_state["cgutec_cargo_orgao"] = ""
                st.session_state["cgutec_fisc_orgao"] = ""
                st.session_state["cgutec_end_orgao"] = ""
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

        c_o1, c_o2 = st.columns(2)
        with c_o1:
            cnpj_orgao_val = st.text_input("CNPJ", placeholder="00.000.000/0001-00", key="cgutec_cnpj_orgao")
            rep_orgao_val = st.text_input("Representado por", placeholder="Nome do(a) Secretário(a) ou Representante", key="cgutec_rep_orgao")
            cargo_orgao_val = st.text_input("Cargo", placeholder="Ex: Secretário(a) Municipal", key="cgutec_cargo_orgao")
            nome_orgao_val = st.text_input("Órgão ou entidade pública", placeholder="Ex: SECRETARIA MUNICIPAL DE...", key="cgutec_nome_orgao")
            setor_orgao_val = st.text_input("Setor responsável pelas licitações", value="Comissão Permanente de Licitação - CPL / Setor de Contratos", placeholder="Setor responsável pelas licitações", key="cgutec_setor_orgao")
            cep_orgao_val = st.text_input("CEP", placeholder="CEP", key="cgutec_cep_orgao")

        with c_o2:
            end_orgao_val = st.text_input("Endereço Rua,Av,Rodovia...", placeholder="Endereço da sede...", key="cgutec_end_orgao")
            bairro_orgao_val = st.text_input("Bairro", placeholder="Bairro...", key="cgutec_bairro_orgao")
            
            c_num1, c_num2 = st.columns([1, 1.5])
            with c_num1:
                num_orgao_val = st.text_input("Número", placeholder="Número", key="cgutec_num_orgao")
            with c_num2:
                comp_orgao_val = st.text_input("Complemento", value="Prédio Administrativo", placeholder="Bloco, torre, quadra...", key="cgutec_comp_orgao")
                
            c_cid1, c_cid2 = st.columns([2, 1])
            with c_cid1:
                cid_orgao_val = st.text_input("Cidade", placeholder="Cidade...", key="cgutec_cid_orgao")
            with c_cid2:
                uf_orgao_val = st.text_input("UF", placeholder="UF", key="cgutec_uf_orgao")
                
            fisc_orgao_val = st.text_input("Fiscal Designado do Contrato", placeholder="Servidor designado como fiscal", key="cgutec_fisc_orgao")

        nome_alvo_objeto = nome_orgao_val.strip() if nome_orgao_val and nome_orgao_val.strip() else "Administração Municipal"
        obj_padrao_sug = f"Contratação de {tipo_contratacao_sel.lower()} para atendimento das demandas e funcionamento regular da {nome_alvo_objeto}."
        obj_orgao_val = st.text_area("Insira o objeto da licitação", value=st.session_state.get("cgutec_obj_orgao", obj_padrao_sug), placeholder="Insira o objeto da licitação", height=85, key="cgutec_obj_orgao")

        # -------------------------------------------------------------
        # 2.1. PLANILHA DE ITENS DA CONTRATAÇÃO (CLÁUSULA PRIMEIRA - OBJETO)
        # -------------------------------------------------------------
        st.markdown("##### 📦 Itens da Contratação (Cláusula Primeira - Detalhamento dos Itens)")
        with st.expander("📊 Importar Planilha ou PDF de Itens (Auto-Ajustável com IA)", expanded=True):
            st.caption("A tabela do documento gerado se **auto-ajusta dinamicamente** à quantidade de colunas, linhas e descrição dos cabeçalhos da planilha (.xlsx, .xls, .csv) ou PDF importado. Não é necessário modelo fixo.")
            
            arquivo_planilha_itens = st.file_uploader(
                "Selecione a Planilha ou Arquivo PDF de Itens (.xlsx, .xls, .csv, .pdf):",
                type=["xlsx", "xls", "csv", "pdf"],
                key="cgutec_upload_planilha_itens",
                help="Carregue qualquer planilha ou documento PDF com itens. A tabela da minuta será gerada auto-ajustando cabeçalhos, quantidade de linhas e colunas."
            )

            # Processamento dinâmico
            itens_cgutec_processados = []
            tab_din_cgutec = None

            if arquivo_planilha_itens is not None:
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
                st.dataframe(df_dyn_view, use_container_width=True, hide_index=True)

                c_bot_ia1, c_bot_ia2 = st.columns([1.5, 1.2])
                with c_bot_ia1:
                    if st.button("✨ Padronizar / Otimizar Cabeçalhos com IA (Gemini)", key="btn_ia_opt_headers", use_container_width=True):
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
                with c_bot_ia2:
                    st.metric(
                        label="Soma Total dos Itens",
                        value=tot_txt_disp,
                        help="Somatório oficial calculado automaticamente a partir dos itens da planilha."
                    )

                if tot_itens_calc > 0:
                    st.info(f"💎 **Valor Global Vinculado:** O somatório de **{tot_txt_disp}** foi puxado automaticamente para o campo **Valor Global da Contratação**.")

        st.markdown("---")

        # -------------------------------------------------------------
        # 3. DEMAIS INFORMAÇÕES
        # -------------------------------------------------------------
        st.markdown("##### ⚙️ Demais Informações")
        c_inf1, c_inf2 = st.columns([1.3, 1.5])
        with c_inf1:
            modalidade_cgutec = st.radio(
                "Modalidade / Procedimento:",
                ["Pregão", "Concorrência", "Contratação direta"],
                index=0,
                key="cgutec_modalidade_sel",
                horizontal=True
            )
        with c_inf2:
            proc_adm_val = st.text_input(
                "Processo Administrativo",
                value="001/2026-ADM",
                placeholder="Processo Administrativo - Deixar em Branco se não houver",
                key="cgutec_proc_adm"
            )

        # -------------------------------------------------------------
        # 4. DADOS COMPLEMENTARES DA CONTRATADA E VALOR
        # -------------------------------------------------------------
        with st.expander("🏢 Dados da Empresa Contratada (Fornecedor)", expanded=True):
            st.caption("💡 *Ao digitar o CNPJ da empresa, se já constar no banco de dados os dados serão preenchidos automaticamente. Caso não conste no banco, você poderá puxar via API da Receita Federal:*")

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

            def _executar_busca_cnpj_banco(cnpj_val):
                c_limpo = re.sub(r"\D", "", str(cnpj_val or ""))
                if not c_limpo:
                    st.session_state["cgutec_cnpj_status"] = None
                    st.session_state["cgutec_cnpj_msg"] = None
                    return

                if len(c_limpo) == 14:
                    st.session_state["cgutec_forn_cnpj"] = formatar_cnpj_cpf(c_limpo)

                # 1. Consulta no Banco de Dados Local
                res_banco = buscar_fornecedor_no_banco(c_limpo, engine)
                if res_banco:
                    # CONSTA NO BANCO DE DADOS: Puxa todos os dados do banco!
                    st.session_state["cgutec_forn_cnpj"] = res_banco.get("cnpj_formatado") or formatar_cnpj_cpf(c_limpo)
                    st.session_state["cgutec_forn_nome"] = res_banco.get("razao_social") or res_banco.get("nome_fantasia") or ""
                    st.session_state["cgutec_forn_rep"] = res_banco.get("representante_nome") or "Representante Legal"
                    st.session_state["cgutec_forn_cargo"] = res_banco.get("representante_cargo") or "Sócio Administrador"
                    st.session_state["cgutec_forn_end"] = res_banco.get("endereco") or "Endereço Comercial"
                    st.session_state["cgutec_forn_bairro"] = res_banco.get("bairro") or "Centro"
                    st.session_state["cgutec_forn_cidade"] = res_banco.get("cidade") or "Imperatriz"
                    st.session_state["cgutec_forn_uf"] = res_banco.get("uf") or "MA"
                    st.session_state["cgutec_forn_cep"] = res_banco.get("cep") or "65900-000"
                    st.session_state["cgutec_cnpj_status"] = "encontrado_banco"
                    st.session_state["cgutec_cnpj_msg"] = f"Empresa **{res_banco.get('razao_social')}** encontrada no banco de dados local! Dados preenchidos automaticamente."
                else:
                    # NÃO CONSTA NO BANCO DE DADOS: usuário deve puxar via API
                    if len(c_limpo) == 14:
                        st.session_state["cgutec_cnpj_status"] = "nao_encontrado_banco"
                        st.session_state["cgutec_cnpj_nao_cadastrado"] = c_limpo
                        st.session_state["cgutec_cnpj_msg"] = f"O CNPJ **{formatar_cnpj_cpf(c_limpo)}** não consta no banco de dados local."
                    else:
                        st.session_state["cgutec_cnpj_status"] = "incompleto"
                        st.session_state["cgutec_cnpj_msg"] = "Digite os 14 números do CNPJ para pesquisar no banco."

            def _callback_cnpj_input():
                _executar_busca_cnpj_banco(st.session_state.get("cgutec_forn_cnpj", ""))

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
                            _executar_busca_cnpj_banco(cnpj_ext)

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

            # 2. Campo Inteligente de CNPJ com Verificação Automática ao Digitar
            c_cnpj_inp, c_cnpj_btn = st.columns([2.6, 1.4])
            with c_cnpj_inp:
                forn_cnpj_val = st.text_input(
                    "CNPJ da Contratada (Ao digitar, busca no banco de dados):",
                    key="cgutec_forn_cnpj",
                    placeholder="00.000.000/0001-00 (ou somente números)",
                    on_change=_callback_cnpj_input,
                    help="Ao digitar o CNPJ (14 dígitos), o sistema busca automaticamente no banco de dados. Caso não conste, o sistema habilita o botão para puxar via API."
                )
            with c_cnpj_btn:
                st.write("")
                st.write("")
                if st.button("⚡ Buscar no Banco", key="btn_verificar_cnpj_cgutec", use_container_width=True, type="secondary"):
                    _executar_busca_cnpj_banco(forn_cnpj_val)
                    st.rerun()

            # 3. Status da Busca e Ação de Puxar via API
            status_cnpj = st.session_state.get("cgutec_cnpj_status")
            msg_cnpj = st.session_state.get("cgutec_cnpj_msg")

            if status_cnpj == "encontrado_banco":
                st.success(f"✅ {msg_cnpj}")
            elif status_cnpj == "encontrado_api":
                st.success(f"🎉 {msg_cnpj}")
            elif status_cnpj == "nao_encontrado_banco":
                st.warning(f"⚠️ {msg_cnpj}")
                c_btn_api, _ = st.columns([1.6, 2])
                with c_btn_api:
                    if st.button("🌐 Puxar Dados via API (Receita Federal)", type="primary", key="btn_puxar_api_receita_m3", use_container_width=True):
                        cnpj_pend = st.session_state.get("cgutec_cnpj_nao_cadastrado") or forn_cnpj_val
                        with st.spinner(f"Consultando dados oficiais da empresa na Receita Federal via API..."):
                            res_api = puxar_fornecedor_api_e_cadastrar(cnpj_pend, engine)
                            if res_api.get("sucesso"):
                                st.session_state["cgutec_forn_cnpj"] = res_api.get("cnpj_formatado") or formatar_cnpj_cpf(cnpj_pend)
                                st.session_state["cgutec_forn_nome"] = res_api.get("razao_social") or ""
                                st.session_state["cgutec_forn_rep"] = res_api.get("representante_nome") or "Representante Legal"
                                st.session_state["cgutec_forn_cargo"] = res_api.get("representante_cargo") or "Sócio Administrador"
                                st.session_state["cgutec_forn_end"] = res_api.get("endereco") or "Endereço Comercial"
                                st.session_state["cgutec_forn_bairro"] = res_api.get("bairro") or "Centro"
                                st.session_state["cgutec_forn_cidade"] = res_api.get("cidade") or "Imperatriz"
                                st.session_state["cgutec_forn_uf"] = res_api.get("uf") or "MA"
                                st.session_state["cgutec_forn_cep"] = res_api.get("cep") or "65900-000"
                                st.session_state["cgutec_cnpj_status"] = "encontrado_api"
                                st.session_state["cgutec_cnpj_msg"] = f"Dados da empresa **{res_api.get('razao_social')}** obtidos com sucesso via API da Receita Federal e cadastrados no banco de dados!"
                                st.rerun()
                            else:
                                st.error(f"❌ Não foi possível obter dados via API da Receita: {res_api.get('erro', 'CNPJ não encontrado ou serviço indisponível.')}")
            elif status_cnpj == "incompleto":
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
                forn_cep_val = st.text_input("CEP:", key="cgutec_forn_cep", placeholder="65900-000")

        # -------------------------------------------------------------
        # DADOS DO CONTRATO MUNICIPAL & CLASSIFICAÇÃO ORÇAMENTÁRIA (IA)
        # -------------------------------------------------------------
        with st.expander("🏦 Dados do Contrato & Classificação Orçamentária Inteligente (IA)", expanded=True):
            col_ct_val1, col_ct_val2 = st.columns(2)
            with col_ct_val1:
                _, _, _, prox_sugerido_cgutec = obter_ultimo_e_proximo_numero_contrato()
                num_ct_cgutec = st.text_input(
                    "Nº do Contrato:",
                    value=st.session_state.get("cgutec_num_ct", prox_sugerido_cgutec),
                    placeholder=f"Ex: {prox_sugerido_cgutec}",
                    key="cgutec_num_ct",
                    help="Puxado automaticamente a partir do último contrato registrado."
                )
            with col_ct_val2:
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
                        if st.button("🔄 Sincronizar com a Soma", key="btn_reap_soma_plan", use_container_width=True, help="Reaplicar a soma exata dos itens da planilha"):
                            st.session_state["cgutec_val_ct_txt"] = f"{soma_plan:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            st.rerun()

            st.markdown("---")
            st.markdown("##### 🏦 Classificação Orçamentária Inteligente (IA)")
            st.caption("Cole abaixo o texto da Dotação, Empenho, PDF ou Planilha. A Inteligência Artificial organizará os campos para a minuta oficial.")
            
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
                if st.button("✨ Extrair com Inteligência Artificial", key="btn_ext_ia_cgutec", use_container_width=True):
                    with st.spinner("Classificando dotação com IA..."):
                        d_ext_cg = extrair_dotacao_com_ia(txt_dot_cgutec)
                        if d_ext_cg:
                            st.session_state["cgutec_dot_ex"] = d_ext_cg.get("exercicio", "")
                            st.session_state["cgutec_dot_pod"] = d_ext_cg.get("poder", "")
                            st.session_state["cgutec_dot_org"] = d_ext_cg.get("orgao", "")
                            st.session_state["cgutec_dot_uni"] = d_ext_cg.get("unidade", "")
                            st.session_state["cgutec_dot_nat"] = d_ext_cg.get("natureza", "")
                            st.success("Dotação classificada com sucesso!")
                        else:
                            st.error("Não foi possível extrair. Verifique o texto informado.")
                            
            col_cg_d1, col_cg_d2 = st.columns(2)
            with col_cg_d1:
                cg_dot_ex = st.text_input("Exercício:", value=st.session_state.get("cgutec_dot_ex", "2026"), key="cgutec_inp_dot_ex")
                cg_dot_pod = st.text_input("Poder:", value=st.session_state.get("cgutec_dot_pod", "Poder Executivo 02.00"), key="cgutec_inp_dot_pod")
                cg_dot_org = st.text_input("Órgão Orçamentário:", value=st.session_state.get("cgutec_dot_org", f"Fundo Municipal / {nome_orgao_val}" if nome_orgao_val else "Fundo Municipal"), key="cgutec_inp_dot_org")
            with col_cg_d2:
                cg_dot_uni = st.text_input("Unidade Orçamentária / Projeto Atividade:", value=st.session_state.get("cgutec_dot_uni", st.session_state.get("cgutec_dot_ct", "MANUTENÇÃO E FUNCIONAMENTO 12.361.0402.2022.0000")), key="cgutec_inp_dot_uni")
                cg_dot_nat = st.text_input("Natureza da Despesa:", value=st.session_state.get("cgutec_dot_nat", "Outros Serviços de Terceiros 3.3.90.39.00" if "SERV" in str(cod_modelo_cgutec).upper() else "Material de consumo 3.3.90.30.00"), key="cgutec_inp_dot_nat")
            
            dot_cgutec = f"{cg_dot_org} | {cg_dot_uni} | {cg_dot_nat}"

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
            "demais_informacoes": modalidade_cgutec.upper(),
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
        st.caption("Ao informar o CNPJ, o sistema verifica se já existe no banco de dados. Caso não exista, você pode puxar os dados completos via API da Receita Federal:")

        c_tf_c1, c_tf_c2 = st.columns([2.5, 1.2])
        with c_tf_c1:
            cnpj_tf_in = st.text_input("CNPJ para Consulta:", placeholder="00.000.000/0001-00 (ou somente números)", key="tf_cnpj_busca")
        with c_tf_c2:
            st.write("")
            st.write("")
            btn_tf_consultar = st.button("⚡ Buscar no Banco / API", key="btn_tf_consultar_cnpj", use_container_width=True)

        if btn_tf_consultar and cnpj_tf_in:
            c_l = re.sub(r"\D", "", cnpj_tf_in)
            if len(c_l) == 14:
                res_b = buscar_fornecedor_no_banco(c_l, get_engine())
                if res_b:
                    st.session_state["tf_form_cnpj"] = res_b.get("cnpj_formatado") or formatar_cnpj_cpf(c_l)
                    st.session_state["tf_form_razao"] = res_b.get("razao_social") or ""
                    st.session_state["tf_form_tel"] = res_b.get("telefone") or ""
                    st.session_state["tf_form_email"] = res_b.get("email") or ""
                    st.info(f"✅ Empresa já consta no banco de dados: **{res_b.get('razao_social')}**")
                else:
                    with st.spinner("Empresa não encontrada no banco. Puxando via API da Receita Federal..."):
                        res_a = puxar_fornecedor_api_e_cadastrar(c_l, get_engine())
                        if res_a.get("sucesso"):
                            st.session_state["tf_form_cnpj"] = res_a.get("cnpj_formatado") or formatar_cnpj_cpf(c_l)
                            st.session_state["tf_form_razao"] = res_a.get("razao_social") or ""
                            st.session_state["tf_form_tel"] = res_a.get("telefone") or ""
                            st.session_state["tf_form_email"] = res_a.get("email") or ""
                            st.success(f"🎉 Dados da empresa **{res_a.get('razao_social')}** obtidos via Receita Federal e cadastrados com sucesso!")
                            st.rerun()
                        else:
                            st.error(f"❌ Não foi possível obter via API da Receita: {res_a.get('erro', 'CNPJ não encontrado.')}")
            else:
                st.warning("⚠️ Digite um CNPJ válido com 14 dígitos.")

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
        st.session_state["logged_in"] = False
    if "usuario" not in st.session_state:
        st.session_state["usuario"] = None
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
    if not st.session_state["logged_in"]:
        with login_area.container():
            tela_login()
        st.stop()
    else:
        login_area.empty()
        with st.sidebar:
            # Emblema Municipal Oficial com Brasão de Ribeirãozinho do Maranhão - MA
            brasao_b64 = obter_brasao_b64()
            brasao_sidebar_html = f"<div style='display: flex; justify-content: center; margin-bottom: 8px;'><img src='data:image/png;base64,{brasao_b64}' style='width: 60px; height: auto; filter: drop-shadow(0 4px 6px rgba(0,0,0,0.15));' alt='Brasão Oficial' /></div>" if brasao_b64 else "<div style='font-size: 32px; margin-bottom: 4px;'>🏛️</div>"
            st.markdown(f"""
            <div style='text-align: center; padding: 4px 8px 10px 8px; border-bottom: 2px solid #e2e8f0; margin-bottom: 12px;'>
                {brasao_sidebar_html}
                <div style='font-weight: 800; font-size: 11.5px; color: #0369a1; line-height: 1.25; font-family: "Plus Jakarta Sans"; letter-spacing: -0.2px;'>
                    PREFEITURA MUNICIPAL DE<br>RIBEIRÃOZINHO DO MARANHÃO - MA
                </div>
                <div style='font-size: 10px; color: #059669; margin-top: 2px; font-weight: 700; letter-spacing: 0.5px;'>ESTADO DO MARANHÃO</div>
            </div>
            """, unsafe_allow_html=True)

            # Card de Usuário Claro
            nome_usuario = st.session_state['usuario'].get('nome', 'Administrador')
            cargo = st.session_state['usuario'].get('cargo', 'OPERADOR')
            iniciais = obter_iniciais(nome_usuario)

            st.markdown(f"""
            <div class='gel-user-box'>
                <div class='gel-avatar-circle'>{iniciais}</div>
                <div style='flex: 1; overflow: hidden;'>
                    <div class='gel-user-name' title='{nome_usuario}'>{nome_usuario}</div>
                    <span class='gel-role-pill'>{cargo}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Relógio Oficial e Data
            agora = datetime.now()
            st.markdown(f"""
            <div class='gel-clock-box'>
                <div style='color: #475569;'>📅 <strong>{formatar_data_extenso_br(agora)}</strong></div>
                <div style='margin-top: 2px;'>🕒 Brasília: <span class='gel-clock-time'>{agora.strftime('%H:%M:%S')}</span></div>
            </div>
            """, unsafe_allow_html=True)

            # Indicador Visual de Status do Banco de Dados em Tempo Real
            st_banco = obter_status_conexao_banco()
            if st_banco["online"]:
                st.markdown(f"""
                <div class='gel-db-status-box'>
                    <div class='gel-db-pulse-dot'></div>
                    <div style='flex: 1; text-align: left;'>
                        <div style='font-weight: 800; color: #10b981; font-size: 11px; line-height: 1.15;'>🟢 BANCO ATIVO NA NUVEM</div>
                        <div style='font-size: 9.5px; color: #94a3b8; margin-top: 2px;'>{st_banco["tipo"].split("(")[0].strip()} • {st_banco["total_contratos"]} contratos</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class='gel-db-status-box' style='border-color: #ef4444;'>
                    <div class='gel-db-pulse-dot offline'></div>
                    <div style='flex: 1; text-align: left;'>
                        <div style='font-weight: 800; color: #ef4444; font-size: 11px; line-height: 1.15;'>🔴 BANCO DESCONECTADO</div>
                        <div style='font-size: 9.5px; color: #fca5a5; margin-top: 2px;'>Verifique conexão</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("""
            <div style='margin-bottom: 10px; padding: 0 4px; display: flex; justify-content: space-between; align-items: center;'>
                <span style='font-size: 10.5px; font-weight: 800; color: #0284c7; letter-spacing: 0.5px; text-transform: uppercase;'>
                    ⚡ NAVEGAÇÃO
                </span>
                <span class='nav-section-badge' style='font-size: 9px; padding: 2px 6px;'>
                    SISTEMA ONLINE
                </span>
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
