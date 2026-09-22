import os
import io
import re
from datetime import datetime, date
from typing import Dict, Any, Optional

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether, HRFlowable
)
from reportlab.pdfgen import canvas

from agu_api_client import AGUApiClient

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
HEADER_OFICIAL_PATH = os.path.join(ASSETS_DIR, "header_oficial_rib.png")
FOOTER_OFICIAL_PATH = os.path.join(ASSETS_DIR, "footer_oficial_rib.png")
LOGO_OFICIAL_PATH = os.path.join(ASSETS_DIR, "logo_oficial_rib.png")
BRASAO_PATH = os.path.join(ASSETS_DIR, "brasao_edison_lobao.png")
if not os.path.exists(BRASAO_PATH):
    BRASAO_PATH = os.path.join(ASSETS_DIR, "brasao_web.png")

def valor_por_extenso(valor: float) -> str:
    """Converte valor numérico em reais para extenso simplificado"""
    try:
        val = float(valor)
        inteiro = int(val)
        centavos = int(round((val - inteiro) * 100))
        
        # Dicionários de números
        unidades = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
        especiais = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"]
        dezenas = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"]
        centenas = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"]

        def converter_grupo(n):
            if n == 0:
                return ""
            if n == 100:
                return "cem"
            c = n // 100
            resto = n % 100
            d = resto // 10
            u = resto % 10
            partes = []
            if c > 0:
                partes.append(centenas[c])
            if d == 1:
                partes.append(especiais[u])
            else:
                if d > 1:
                    partes.append(dezenas[d])
                if u > 0:
                    partes.append(unidades[u])
            return " e ".join(partes)

        if inteiro == 0:
            str_int = "zero reais"
        else:
            milhoes = inteiro // 1_000_000
            milhares = (inteiro % 1_000_000) // 1000
            resto_mil = inteiro % 1000
            grupos = []
            if milhoes > 0:
                grupos.append(f"{converter_grupo(milhoes)} {'milhão' if milhoes == 1 else 'milhões'}")
            if milhares > 0:
                if milhares == 1 and not milhoes:
                    grupos.append("mil")
                else:
                    grupos.append(f"{converter_grupo(milhares)} mil")
            if resto_mil > 0:
                grupos.append(converter_grupo(resto_mil))
            str_int = " e ".join(grupos) + (" real" if inteiro == 1 else " reais")

        if centavos > 0:
            str_cent = f" e {converter_grupo(centavos)} centavo" if centavos == 1 else f" e {converter_grupo(centavos)} centavos"
            return f"{str_int}{str_cent}"
        return str_int
    except Exception:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_moeda_br(valor: float) -> str:
    try:
        val = float(valor)
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"

def formatar_cnpj_br(val) -> str:
    """Formata CNPJ no padrão oficial XX.XXX.XXX/XXXX-XX ou CPF XXX.XXX.XXX-XX"""
    if not val:
        return "00.000.000/0001-00"
    s = re.sub(r"\D", "", str(val)).strip()
    if len(s) == 14:
        return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:]}"
    elif len(s) == 11:
        return f"{s[:3]}.{s[3:6]}.{s[6:9]}-{s[9:]}"
    return str(val).strip()

def gerar_planilha_itens_excel_formatada(headers: list, rows: list, total_geral: float = 0.0, sheet_name: str = "Itens do Contrato") -> bytes:
    """
    Gera uma planilha Excel (.xlsx) altamente profissional com:
    - Quebra de texto automática ativada em todas as células (wrap_text=True)
    - Bordas finas completas em todas as células (ajuste de texto às bordas)
    - Auto-ajuste proporcional da largura das colunas
    - Cabeçalho padronizado em azul marinho (#003366) com texto branco em negrito
    - Alinhamento vertical centralizado e horizontal específico (texto à esquerda, valores à direita, itens/quantidades ao centro)
    - Linha de Totalizador final destacada
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    # Estilos
    thin_border_side = Side(border_style='thin', color='CBD5E1')
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    header_fill = PatternFill(start_color='003366', end_color='003366', fill_type='solid')
    header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')

    body_font = Font(name='Arial', size=9.5)
    zebra_fill = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
    white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

    total_fill = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
    total_font = Font(name='Arial', size=10, bold=True, color='003366')

    # Detecta alinhamento de cada coluna
    col_alignments = []
    for h in headers:
        h_up = str(h).upper()
        if any(k in h_up for k in ['VALOR', 'PREÇO', 'PRECO', 'TOTAL', 'UNIT', 'R$', 'CUSTO']):
            col_alignments.append('right')
        elif any(k in h_up for k in ['ITEM', 'Nº', 'NUM', 'ORDEM', 'UNID', 'UND', 'QTD', 'QUANT', 'QUANTIDADE', 'UF', 'LOTE']):
            col_alignments.append('center')
        else:
            col_alignments.append('left')

    # 1. Escreve Cabeçalho
    ws.append(headers)
    ws.row_dimensions[1].height = 26
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = cell_border

    # 2. Escreve Linhas com Quebra de Texto e Ajuste de Bordas
    for row_idx, r_vals in enumerate(rows, start=2):
        row_data = [str(r_vals[c]) if c < len(r_vals) else "" for c in range(len(headers))]
        ws.append(row_data)
        ws.row_dimensions[row_idx].height = 24

        row_fill = zebra_fill if (row_idx % 2 == 0) else white_fill
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = body_font
            cell.fill = row_fill
            cell.border = cell_border
            align_h = col_alignments[col_idx - 1] if (col_idx - 1) < len(col_alignments) else 'left'
            cell.alignment = Alignment(horizontal=align_h, vertical='center', wrap_text=True)

    # 3. Linha de Totalizador (se aplicável)
    last_row = len(rows) + 1
    if total_geral > 0:
        tot_row_idx = last_row + 1
        ws.row_dimensions[tot_row_idx].height = 26

        tot_col_idx = len(headers)
        for i, h in enumerate(headers):
            h_up = str(h).upper()
            if any(k in h_up for k in ['TOTAL', 'VALOR TOTAL', 'PREÇO TOTAL', 'VALOR GLOBAL']):
                tot_col_idx = i + 1
                break

        for c_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=tot_row_idx, column=c_idx)
            cell.fill = total_fill
            cell.font = total_font
            cell.border = cell_border
            cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)

        if tot_col_idx > 1:
            ws.merge_cells(start_row=tot_row_idx, start_column=1, end_row=tot_row_idx, end_column=tot_col_idx - 1)
            cell_lbl = ws.cell(row=tot_row_idx, column=1)
            cell_lbl.value = "VALOR TOTAL GERAL:"
            cell_lbl.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)

        cell_val = ws.cell(row=tot_row_idx, column=tot_col_idx)
        cell_val.value = f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        cell_val.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)

    # 4. Ajuste proporcional da largura das colunas
    for col_idx, col in enumerate(ws.columns, start=1):
        col_header = str(headers[col_idx - 1]).upper() if col_idx <= len(headers) else ""
        if any(k in col_header for k in ['DESC', 'ESPEC', 'OBJET', 'PROD', 'MATERIAL', 'SERVIÇO']):
            col_width = 48
        elif any(k in col_header for k in ['ITEM', 'Nº', 'UND', 'UNID']):
            col_width = 10
        elif any(k in col_header for k in ['QTD', 'QUANT']):
            col_width = 14
        elif any(k in col_header for k in ['VALOR', 'PREÇO', 'UNIT', 'TOTAL', 'R$']):
            col_width = 18
        else:
            max_len = max(len(str(cell.value or '')) for cell in col[:15]) if len(col) > 0 else 12
            col_width = min(max(max_len + 4, 12), 35)

        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_width

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

def renderizar_tabela_html_formatada(headers: list, rows: list, total_geral: float = 0.0, max_height: int = 480) -> str:
    """
    Renderiza tabela HTML responsiva com:
    - Quebra de texto automática em todas as células (word-wrap, white-space normal)
    - Ajuste de texto às bordas (padding 8px 12px, border 1px solid #cbd5e1)
    - Cabeçalho fixo no topo (sticky) com visual executivo em azul petróleo (#003366) e texto branco
    - Linhas zebradas suaves (#f8fafc / #ffffff) e efeito hover
    - Linha de totalizador em destaque
    """
    if not headers or not rows:
        return ""

    col_align = []
    col_widths_css = []
    for h in headers:
        h_up = str(h).upper()
        if any(k in h_up for k in ['DESC', 'ESPEC', 'OBJET', 'PROD', 'MATERIAL', 'SERVIÇO']):
            col_align.append('left')
            col_widths_css.append('42%')
        elif any(k in h_up for k in ['ITEM', 'Nº', 'NUM', 'ORDEM', 'UNID', 'UND']):
            col_align.append('center')
            col_widths_css.append('8%')
        elif any(k in h_up for k in ['QTD', 'QUANT']):
            col_align.append('center')
            col_widths_css.append('10%')
        elif any(k in h_up for k in ['VALOR', 'PREÇO', 'PRECO', 'UNIT', 'TOTAL', 'R$']):
            col_align.append('right')
            col_widths_css.append('15%')
        else:
            col_align.append('left')
            col_widths_css.append('15%')

    html = [
        f'<div style="max-height: {max_height}px; overflow-y: auto; overflow-x: auto; '
        f'border: 1px solid #cbd5e1; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 12px;">',
        '<table style="width: 100%; border-collapse: collapse; font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif; font-size: 12.5px; color: #1e293b;">',
        '<thead>',
        '<tr style="background-color: #003366; color: #ffffff; position: sticky; top: 0; z-index: 2;">'
    ]

    for ci, h in enumerate(headers):
        w_style = f"width: {col_widths_css[ci]};" if ci < len(col_widths_css) else ""
        html.append(
            f'<th style="{w_style} padding: 10px 12px; border: 1px solid #0f2c4d; text-align: center; font-weight: 700; letter-spacing: 0.3px; text-transform: uppercase; font-size: 11.5px;">'
            f'{h}</th>'
        )
    html.append('</tr></thead><tbody>')

    for ri, r in enumerate(rows):
        bg = "#ffffff" if (ri % 2 == 0) else "#f8fafc"
        html.append(f'<tr style="background-color: {bg}; transition: background 0.15s ease;">')
        for ci, h in enumerate(headers):
            val = str(r[ci]) if ci < len(r) else ""
            align = col_align[ci] if ci < len(col_align) else 'left'
            is_item_col = (ci == 0 and any(k in str(h).upper() for k in ['ITEM', 'Nº']))
            extra_bold = "font-weight: 600;" if is_item_col else ""
            html.append(
                f'<td style="padding: 7px 12px; border: 1px solid #cbd5e1; text-align: {align}; '
                f'vertical-align: middle; word-wrap: break-word; white-space: normal; line-height: 1.45; {extra_bold}">'
                f'{val}</td>'
            )
        html.append('</tr>')

    if total_geral > 0:
        tot_fmt = f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        tot_col_idx = len(headers) - 1
        for i, h in enumerate(headers):
            h_up = str(h).upper()
            if any(k in h_up for k in ['TOTAL', 'VALOR TOTAL', 'PREÇO TOTAL', 'VALOR GLOBAL']):
                tot_col_idx = i
                break

        html.append('<tr style="background-color: #f1f5f9; font-weight: 700; border-top: 2px solid #003366; position: sticky; bottom: 0; z-index: 1;">')
        if tot_col_idx > 0:
            html.append(
                f'<td colspan="{tot_col_idx}" style="padding: 9px 12px; border: 1px solid #cbd5e1; text-align: right; color: #003366; font-size: 12px;">'
                f'VALOR TOTAL:</td>'
            )
        html.append(
            f'<td style="padding: 9px 12px; border: 1px solid #cbd5e1; text-align: right; color: #003366; font-size: 13px; font-weight: 800;">'
            f'{tot_fmt}</td>'
        )
        if tot_col_idx < len(headers) - 1:
            sobra = len(headers) - 1 - tot_col_idx
            html.append(f'<td colspan="{sobra}" style="padding: 9px 12px; border: 1px solid #cbd5e1;"></td>')
        html.append('</tr>')

    html.append('</tbody></table></div>')
    return ''.join(html)

def aplicar_estilo_tabela_oficial_docx(t_it, col_widths_pt=None):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn, nsdecls
    from docx.shared import Pt
    from docx.oxml import parse_xml
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    
    t_it.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_it.autofit = False

    # Bordas elegantes em todas as células (Azul escuro do modelo)
    tblPr = t_it._tbl.tblPr
    tblBorders = parse_xml(
        r'<w:tblBorders {} >'
        r'<w:top w:val="single" w:sz="8" w:space="0" w:color="224061"/>'
        r'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="224061"/>'
        r'<w:insideH w:val="single" w:sz="8" w:space="0" w:color="224061"/>'
        r'<w:insideV w:val="single" w:sz="8" w:space="0" w:color="224061"/>'
        r'<w:left w:val="single" w:sz="8" w:space="0" w:color="224061"/>'
        r'<w:right w:val="single" w:sz="8" w:space="0" w:color="224061"/>'
        r'</w:tblBorders>'.format(nsdecls('w'))
    )
    tblPr.append(tblBorders)

    # Margens internas
    cell_mar = parse_xml(
        r'<w:tblCellMar {} >'
        r'<w:top w:w="80" w:type="dxa"/>'
        r'<w:bottom w:w="80" w:type="dxa"/>'
        r'<w:left w:w="120" w:type="dxa"/>'
        r'<w:right w:w="120" w:type="dxa"/>'
        r'</w:tblCellMar>'.format(nsdecls('w'))
    )
    tblPr.append(cell_mar)

    if len(t_it.rows) > 0:
        t_it.rows[0]._tr.get_or_add_trPr().append(parse_xml(r'<w:tblHeader {}/>'.format(nsdecls('w'))))

    num_cols = len(t_it.columns)
    if not col_widths_pt or len(col_widths_pt) != num_cols:
        col_widths_pt = [480.0 / max(num_cols, 1)] * num_cols
    else:
        s_pts = sum(col_widths_pt)
        if s_pts > 0:
            col_widths_pt = [(w / s_pts) * 480.0 for w in col_widths_pt]

    for row in t_it.rows:
        row._tr.get_or_add_trPr().append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w'))))
        for ci, cell in enumerate(row.cells):
            if ci < len(col_widths_pt):
                cell.width = Pt(col_widths_pt[ci])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.1

def aplicar_fundo_celula_docx(cell, color_hex):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def criar_modelo_excel_itens() -> bytes:
    """Gera uma planilha Excel (.xlsx) modelo padrão pré-formatada com quebra de texto e bordas."""
    headers = ["ITEM", "DESCRIÇÃO", "QUANTIDADE", "UNID.", "VALOR UNITÁRIO", "VALOR TOTAL"]
    rows = [
        ["1", "Oxford Preto", "945", "Metros", "14,13", "13.352,85"],
        ["2", "Oxford Preto Premium", "300", "Metros", "14,13", "4.239,00"],
        ["3", "Oxford Laranja", "90", "Metros", "17,86", "1.607,40"],
        ["4", "Oxford Azul Royal", "50", "Metros", "9,02", "451,00"],
        ["5", "Oxford Azul Royal", "250", "Metros", "9,02", "2.255,00"]
    ]
    return gerar_planilha_itens_excel_formatada(headers, rows, total_geral=21905.25)

class ResultadoPlanilhaItens(list):
    """Lista de itens para total retrocompatibilidade, enriquecida com metadados da tabela dinâmica."""
    def __init__(self, itens, tabela_dinamica=None):
        super().__init__(itens)
        self.tabela_dinamica = tabela_dinamica or {}

    @property
    def itens(self):
        return list(self)

def extrair_tabela_dinamica_df(df) -> dict:
    """
    Extrai e normaliza dinamicamente a estrutura tabular completa a partir de um DataFrame:
    - Auto-ajusta quantidade e descrição dos cabeçalhos da planilha
    - Auto-ajusta quantidade de linhas e células
    - Detecta alinhamentos adequados (esquerda, centro, direita)
    - Calcula larguras proporcionais para folha A4 em PDF (largura útil ~523.27 pt)
    - Identifica a coluna de Total e calcula o somatório oficial
    """
    import pandas as pd
    try:
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        if df.empty:
            return {}

        headers = [str(c).strip() for c in df.columns]
        num_cols = len(headers)

        rows = []
        for _, r in df.iterrows():
            row_vals = []
            for c in df.columns:
                val = r[c]
                if pd.isna(val) or val is None:
                    row_vals.append("")
                elif isinstance(val, float):
                    if val.is_integer():
                        row_vals.append(str(int(val)))
                    else:
                        c_up = str(c).upper()
                        if any(k in c_up for k in ['VALOR', 'PREÇO', 'UNIT', 'TOTAL', 'R$']):
                            row_vals.append(formatar_moeda_br(val))
                        else:
                            row_vals.append(f"{val:.2f}".replace('.', ','))
                elif isinstance(val, int):
                    row_vals.append(str(val))
                else:
                    row_vals.append(str(val).strip())
            rows.append(row_vals)

        def _parse_num_local(val):
            if val is None or (isinstance(val, float) and str(val) == 'nan'):
                return 0.0
            if isinstance(val, (int, float)):
                return float(val)
            s = str(val).strip()
            if not s or s.lower() in ['nan', 'none', '-', '', 'null']:
                return 0.0
            is_neg = False
            if s.startswith('(') and s.endswith(')'):
                is_neg = True
                s = s[1:-1].strip()
            elif s.startswith('-'):
                is_neg = True
                s = s[1:].strip()
            s = re.sub(r'[R$\s\xa0BRL]', '', s)
            if not s:
                return 0.0
            if ',' in s and '.' in s:
                if s.rfind(',') > s.rfind('.'):
                    s = s.replace('.', '').replace(',', '.')
                else:
                    s = s.replace(',', '')
            elif ',' in s:
                partes = s.split(',')
                if len(partes) > 2:
                    s = s.replace(',', '')
                else:
                    s = s.replace(',', '.')
            elif '.' in s:
                partes = s.split('.')
                if len(partes) > 2:
                    s = s.replace('.', '')
                else:
                    if len(partes[1]) == 3 and len(partes[0]) <= 3:
                        s = s.replace('.', '')
            try:
                res = float(s)
                return -res if is_neg else res
            except Exception:
                return 0.0

        total_col_idx = None
        qtd_col_idx = None
        vu_col_idx = None

        descartes_total = ['QTD', 'QUANT', 'QUANTIDADE', 'DIAS', 'HORAS', 'PARCELA', 'MESES', 'PAGAMENTO']

        # 1. Prioridade máxima: colunas financeiras de total explícito
        for i, h in enumerate(headers):
            h_up = h.upper()
            if any(k in h_up for k in ['VALOR TOTAL', 'V. TOTAL', 'VL. TOTAL', 'PREÇO TOTAL', 'PRECO TOTAL', 'VALOR GLOBAL', 'TOTAL (R$)', 'TOTAL R$', 'VALOR HOMOLOGADO', 'SUBTOTAL']):
                if not any(d in h_up for d in descartes_total):
                    total_col_idx = i
                    break

        # 2. Se não encontrou, busca por "TOTAL" sem termos de descarte
        if total_col_idx is None:
            for i, h in enumerate(headers):
                h_up = h.upper()
                if 'TOTAL' in h_up and not any(d in h_up for d in descartes_total):
                    total_col_idx = i
                    break

        for i, h in enumerate(headers):
            h_up = h.upper()
            if any(k in h_up for k in ['PREÇO UNID', 'PRECO UNID', 'VALOR UNID', 'P. UNID', 'V. UNID', 'VL. UNID', 'UNIT', 'UNITARIO', 'V. UNIT', 'VL. UNIT', 'PREÇO UNIT', 'PRECO UNIT', 'VALOR UNITÁRIO', 'VALOR UNIT']) and vu_col_idx is None:
                vu_col_idx = i
            elif any(k in h_up for k in ['QTD', 'QUANT', 'QUANTIDADE']) and not any(k in h_up for k in ['VALOR', 'PREÇO', 'PRECO']) and qtd_col_idx is None:
                qtd_col_idx = i

        def _is_summary_row(row_vals):
            for cell in row_vals[:4]:
                c_str = str(cell or '').upper().strip()
                if c_str in ['TOTAL', 'TOTAL GERAL', 'SUBTOTAL', 'VALOR TOTAL', 'SOMA', 'TOTAL DA PROPOSTA', 'VALOR GLOBAL']:
                    return True
                if c_str.startswith('TOTAL ') or c_str.startswith('TOTAL:'):
                    return True
            txt_all = ' '.join([str(c or '').upper().strip() for c in row_vals])
            words = txt_all.split()
            if 'TOTAL' in words or 'SUBTOTAL' in words or 'TOTAL:' in words:
                return True
            return False

        # Isola linhas de itens da linha de rodapé de total da planilha para não somar duas vezes
        resumo_row_val = None
        item_rows = []
        for r in rows:
            if _is_summary_row(r):
                # Extrai valor financeiro do resumo da célula de total ou da última célula numérica preenchida
                val_res = None
                if total_col_idx is not None and total_col_idx < len(r) and _parse_num_local(r[total_col_idx]) > 0:
                    val_res = _parse_num_local(r[total_col_idx])
                else:
                    for c in reversed(r):
                        v_num = _parse_num_local(c)
                        if v_num > 0:
                            val_res = v_num
                            break
                if val_res and val_res > 0:
                    resumo_row_val = val_res
            else:
                item_rows.append(r)

        total_geral = 0.0
        if total_col_idx is not None:
            for r in item_rows:
                if total_col_idx < len(r):
                    val_txt = r[total_col_idx]
                    total_geral += _parse_num_local(val_txt)
        elif qtd_col_idx is not None and vu_col_idx is not None:
            for r in item_rows:
                q = _parse_num_local(r[qtd_col_idx]) if qtd_col_idx < len(r) else 1.0
                v = _parse_num_local(r[vu_col_idx]) if vu_col_idx < len(r) else 0.0
                total_geral += (q * v)
        else:
            for i in reversed(range(num_cols)):
                h_up = headers[i].upper()
                if any(k in h_up for k in ['VALOR', 'PREÇO', 'PRECO']) and not any(d in h_up for d in descartes_total):
                    total_col_idx = i
                    break
            if total_col_idx is not None:
                for r in item_rows:
                    if total_col_idx < len(r):
                        total_geral += _parse_num_local(r[total_col_idx])

        # Se a soma dos itens foi zero ou se houver valor explícito no rodapé da planilha, adota o valor oficial do rodapé
        if total_geral == 0.0 and resumo_row_val and resumo_row_val > 0:
            total_geral = resumo_row_val
        elif resumo_row_val and resumo_row_val > 0:
            if abs(total_geral - resumo_row_val) < 0.05:
                total_geral = resumo_row_val
            elif total_geral > 0 and abs(total_geral - resumo_row_val) < (resumo_row_val * 0.01):
                total_geral = resumo_row_val

        col_alignments = []
        col_weights = []
        for i, h in enumerate(headers):
            h_up = h.upper()
            if any(k in h_up for k in ['ITEM', 'Nº', 'NUM', 'ORDEM', 'POSIÇÃO', 'LOTE', 'CÓDIGO', 'COD', 'UNIDADE', 'UNID', 'UND', 'QTD', 'QUANT', 'QUANTIDADE', 'UF']):
                col_alignments.append('CENTER')
                col_weights.append(1.0 if any(k in h_up for k in ['ITEM', 'Nº', 'UNID', 'UND']) else 1.3)
            elif any(k in h_up for k in ['VALOR', 'PREÇO', 'UNIT', 'TOTAL', 'CUSTO', 'R$']):
                col_alignments.append('RIGHT')
                col_weights.append(1.75)
            else:
                col_alignments.append('LEFT')
                col_weights.append(4.2 if any(k in h_up for k in ['DESC', 'ESPEC', 'OBJET', 'PROD', 'MATERIAL', 'SERVIÇO']) else 2.6)

        LARGURA_UTIL_PDF = 523.27
        soma_pesos = sum(col_weights) if sum(col_weights) > 0 else float(num_cols)
        col_widths = [round((w / soma_pesos) * LARGURA_UTIL_PDF, 2) for w in col_weights]
        
        diff = round(LARGURA_UTIL_PDF - sum(col_widths), 2)
        if col_widths:
            col_widths[-1] += diff

        return {
            "headers": headers,
            "rows": item_rows,
            "col_alignments": col_alignments,
            "col_widths": col_widths,
            "total_col_idx": total_col_idx,
            "total_geral": total_geral,
            "total_formatado": formatar_moeda_br(total_geral),
            "qtd_colunas": num_cols,
            "qtd_linhas": len(item_rows)
        }
    except Exception as e:
        print("Erro ao extrair tabela dinâmica do DataFrame:", e)
        return {}

def estruturar_tabela_com_ia(conteudo_bruto: str, objeto_contexto: str = "") -> dict:
    """
    Utiliza Inteligência Artificial (Google Gemini 3.6 Flash via REST) para interpretar,
    estruturar e padronizar qualquer tabela/listagem de itens da contratação, gerando cabeçalhos
    claros e administrativos, ajustando linhas, quantidades, unidades e valores monetários.
    """
    if not conteudo_bruto or not str(conteudo_bruto).strip():
        return {}

    try:
        import requests
        import json
        api_key = "AQ.Ab8RN6J8dyEt9ArI7BqLKaLvuF52lLgIlr68BaQ6cekMTdu3dw"
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key={api_key}"
        
        prompt = f"""
Você é um Auditor e Procurador Municipal especialista em Licitações e Contratos (Lei nº 14.133/2021) e nos Modelos Oficiais da Consultoria-Geral da União (AGU).
Sua tarefa é analisar a listagem ou dados brutos fornecidos e convertê-los em uma tabela oficial estruturada para a Cláusula Primeira (Detalhamento do Objeto) de uma minuta contratual.

CONTEXTO DA CONTRATAÇÃO:
Objeto: {objeto_contexto or 'Contratação de aquisições / fornecimento para atendimento das demandas municipais'}

CONTEÚDO BRUTO:
{conteudo_bruto}

DIRETRIZES DE FORMATAÇÃO:
1. Extraia e padronize os cabeçalhos de coluna em CAIXA ALTA (ex: "ITEM", "DESCRIÇÃO DO PRODUTO / MATERIAL", "QUANTIDADE", "UNID.", "VALOR UNITÁRIO (R$)", "VALOR TOTAL (R$)").
2. Se houver outras colunas úteis como "CÓDIGO", "MARCA" ou "ESPECIFICAÇÃO", preserve-as com nomes claros.
3. Formate valores monetários no padrão brasileiro (ex: R$ 14,13 ou 13.352,85).
4. Assegure que as quantidades e unidades estejam corretas.
5. Calcule a soma total exata dos itens.
6. Retorne ESTRITAMENTE um objeto JSON válido, sem markdown ```json``` e sem comentários, no formato exato:
{{
  "headers": ["ITEM", "DESCRIÇÃO", "QUANTIDADE", "UNID.", "VALOR UNITÁRIO", "VALOR TOTAL"],
  "rows": [
    ["1", "Oxford Preto", "945", "Metros", "R$ 14,13", "R$ 13.352,85"]
  ]
}}
"""
        payload = {"contents": [{"parts": [{"text": prompt}]}]}
        headers_req = {"Content-Type": "application/json"}
        response = requests.post(url, headers=headers_req, json=payload, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            raw_res = data["candidates"][0]["content"]["parts"][0]["text"].strip()
            raw_res = re.sub(r'^```(?:json)?\s*', '', raw_res)
            raw_res = re.sub(r'\s*```$', '', raw_res)
            
            parsed = json.loads(raw_res)
            if "headers" in parsed and "rows" in parsed and len(parsed["headers"]) > 0:
                import pandas as pd
                df_res = pd.DataFrame(parsed["rows"], columns=parsed["headers"])
                return extrair_tabela_dinamica_df(df_res)
    except Exception as e:
        print("Erro na estruturação de tabela com IA:", e)
    
    return {}

def processar_planilha_itens_excel(file_bytes_or_buffer, file_name: str = "", objeto_contexto: str = "") -> ResultadoPlanilhaItens:
    """
    Processa arquivo Excel (.xlsx / .xls), CSV ou PDF contendo itens da contratação.
    Auto-ajusta quantidade de colunas, linhas e cabeçalhos com total flexibilidade.
    Retorna ResultadoPlanilhaItens (compatível com lista de itens e com .tabela_dinamica).
    """
    import pandas as pd
    try:
        raw_bytes = None
        if isinstance(file_bytes_or_buffer, (bytes, bytearray)):
            raw_bytes = bytes(file_bytes_or_buffer)
            buf = io.BytesIO(raw_bytes)
        elif hasattr(file_bytes_or_buffer, 'read'):
            raw_bytes = file_bytes_or_buffer.read()
            if hasattr(file_bytes_or_buffer, 'seek'):
                file_bytes_or_buffer.seek(0)
            buf = io.BytesIO(raw_bytes)
        else:
            buf = file_bytes_or_buffer

        is_pdf = False
        fn = (file_name or getattr(file_bytes_or_buffer, 'name', '') or "").lower()
        if fn.endswith('.pdf') or (raw_bytes and raw_bytes[:4] == b'%PDF'):
            is_pdf = True

        df = None
        if is_pdf:
            try:
                import pdfplumber
                with pdfplumber.open(io.BytesIO(raw_bytes) if raw_bytes else buf) as pdf:
                    all_raw_tables = []
                    for page in pdf.pages:
                        # Extrai com snap_tolerance e join_tolerance para evitar divisão em colunas fantasmas
                        t = page.extract_tables(table_settings={
                            'vertical_strategy': 'lines',
                            'horizontal_strategy': 'lines',
                            'snap_tolerance': 6,
                            'join_tolerance': 6
                        })
                        if not t:
                            t = page.extract_tables()
                        for tbl in t:
                            all_raw_tables.append(tbl)
                    
                    if all_raw_tables:
                        headers = None
                        all_data_rows = []
                        for tbl in all_raw_tables:
                            for r in tbl:
                                cells = [str(cell or '').strip() for cell in r]
                                if not any(cells):
                                    continue
                                c_up = [c.upper() for c in cells]
                                # Linha de cabeçalho
                                if any('ITEM' in c for c in c_up) and any(('ESPEC' in c or 'DESC' in c or 'MATERIAL' in c or 'PROD' in c) for c in c_up):
                                    if headers is None:
                                        headers = [' '.join(c.split()) for c in cells]
                                    continue
                                if headers is None:
                                    headers = [c if c else f"COL_{i+1}" for i, c in enumerate(cells)]
                                    continue
                                # Repetição de cabeçalhos em páginas seguintes
                                if any('ITEM' in c for c in c_up) and any(('ESPEC' in c or 'DESC' in c or 'MATERIAL' in c or 'PROD' in c) for c in c_up):
                                    continue
                                
                                # Continuação de descrição de linha anterior (item vazio e sem valor financeiro)
                                if not cells[0] and len(cells) > 1 and cells[1] and all(not c for c in cells[2:]):
                                    if all_data_rows:
                                        all_data_rows[-1][1] += ' ' + cells[1]
                                    continue

                                all_data_rows.append(cells)
                        
                        if headers and all_data_rows:
                            num_cols = len(headers)
                            norm_rows = []
                            for r in all_data_rows:
                                if len(r) == num_cols:
                                    norm_rows.append(r)
                                elif len(r) < num_cols:
                                    norm_rows.append(r + [''] * (num_cols - len(r)))
                                else:
                                    norm_rows.append(r[:num_cols])
                            df = pd.DataFrame(norm_rows, columns=headers)
                    else:
                        texto_pdf = "\n".join([p.extract_text() or "" for p in pdf.pages])
                        if texto_pdf.strip():
                            tab_ia = estruturar_tabela_com_ia(texto_pdf, objeto_contexto)
                            if tab_ia and "headers" in tab_ia and "rows" in tab_ia:
                                df = pd.DataFrame(tab_ia["rows"], columns=tab_ia["headers"])
            except Exception as e_pdf:
                print("Erro ao extrair dados de PDF:", e_pdf)

        if df is None or df.empty:
            if hasattr(buf, 'seek'):
                buf.seek(0)
            try:
                df = pd.read_excel(buf)
            except Exception:
                if hasattr(buf, 'seek'):
                    buf.seek(0)
                try:
                    df = pd.read_csv(buf, sep=None, engine='python')
                except Exception:
                    if hasattr(buf, 'seek'):
                        buf.seek(0)
                    df = pd.read_csv(buf)

        tab_din = extrair_tabela_dinamica_df(df)

        col_item = None
        col_desc = None
        col_qtd = None
        col_unid = None
        col_vu = None
        col_vt = None

        descartes_vt = ['QTD', 'QUANT', 'QUANTIDADE', 'DIAS', 'HORAS', 'PARCELA', 'MESES', 'PAGAMENTO']
        for col in df.columns:
            c_upper = str(col).strip().upper()
            if any(k in c_upper for k in ['PREÇO UNID', 'PRECO UNID', 'VALOR UNID', 'P. UNID', 'V. UNID', 'VL. UNID', 'UNIT', 'UNITARIO', 'V. UNIT', 'VL. UNIT', 'PREÇO UNIT', 'PRECO UNIT', 'VALOR UNITÁRIO', 'VALOR UNIT']):
                col_vu = col
            elif any(k in c_upper for k in ['VALOR TOTAL', 'V. TOTAL', 'VL. TOTAL', 'PREÇO TOTAL', 'PRECO TOTAL', 'VALOR GLOBAL', 'TOTAL (R$)', 'TOTAL R$', 'SUBTOTAL']):
                if not any(d in c_upper for d in descartes_vt):
                    col_vt = col
            elif 'TOTAL' in c_upper and not any(d in c_upper for d in descartes_vt) and col_vt is None:
                col_vt = col
            elif any(k in c_upper for k in ['QTD', 'QUANT', 'QUANTIDADE']) and not any(k in c_upper for k in ['VALOR', 'PREÇO', 'PRECO']):
                col_qtd = col
            elif any(k in c_upper for k in ['UNIDADE', 'UNID', 'UND', 'MEDIDA']) and not any(k in c_upper for k in ['PREÇO', 'PRECO', 'VALOR', 'UNIT', 'CUSTO', 'P.', 'V.']):
                col_unid = col
            elif any(k in c_upper for k in ['DESC', 'PROD', 'ESPEC', 'SERV', 'OBJET', 'MATERIAL', 'DISCRIMINAÇÃO']):
                col_desc = col
            elif any(k in c_upper for k in ['ITEM', 'Nº', 'NUMERO', 'POSI']):
                col_item = col

        itens = []
        for idx, row in df.iterrows():
            # Ignora linha de resumo ou rodapé da planilha para não duplicar itens e não somar o total duas vezes
            row_vals_clean = [str(v or '').strip() for v in row.values]
            row_str_full = " ".join([v.upper() for v in row_vals_clean if v])
            cells_up = [v.upper() for v in row_vals_clean if v]
            words = row_str_full.split()

            if any(k in row_str_full for k in ["TOTAL GERAL", "VALOR TOTAL", "SUBTOTAL", "TOTAL DA PROPOSTA", "SOMA TOTAL", "VALOR GLOBAL"]) or \
               any(c in ["TOTAL", "SUBTOTAL", "SOMA", "RESUMO", "TOTAL:"] for c in cells_up[:3]) or \
               'TOTAL' in words or 'SUBTOTAL' in words or 'TOTAL:' in words:
                continue

            def limpar_num(val):
                if val is None or (isinstance(val, float) and str(val) == 'nan'):
                    return 0.0
                if isinstance(val, (int, float)):
                    return float(val)
                s = str(val).strip()
                if not s or s.lower() in ['nan', 'none', '-', '', 'null']:
                    return 0.0
                is_neg = False
                if s.startswith('(') and s.endswith(')'):
                    is_neg = True
                    s = s[1:-1].strip()
                elif s.startswith('-'):
                    is_neg = True
                    s = s[1:].strip()
                s = re.sub(r'[R$\s\xa0BRL]', '', s)
                if not s:
                    return 0.0
                if ',' in s and '.' in s:
                    if s.rfind(',') > s.rfind('.'):
                        s = s.replace('.', '').replace(',', '.')
                    else:
                        s = s.replace(',', '')
                elif ',' in s:
                    partes = s.split(',')
                    if len(partes) > 2:
                        s = s.replace(',', '')
                    else:
                        s = s.replace(',', '.')
                elif '.' in s:
                    partes = s.split('.')
                    if len(partes) > 2:
                        s = s.replace('.', '')
                    else:
                        if len(partes[1]) == 3 and len(partes[0]) <= 3:
                            s = s.replace('.', '')
                try:
                    res = float(s)
                    return -res if is_neg else res
                except Exception:
                    return 0.0

            item_num = int(limpar_num(row[col_item])) if col_item and not pd.isna(row[col_item]) else (len(itens) + 1)
            desc = str(row[col_desc]).strip() if col_desc and not pd.isna(row[col_desc]) else f"Item {item_num}"
            qtd = limpar_num(row[col_qtd]) if col_qtd else 1.0
            unid = str(row[col_unid]).strip() if col_unid and not pd.isna(row[col_unid]) else "Unid."
            vu = limpar_num(row[col_vu]) if col_vu else 0.0
            vt = limpar_num(row[col_vt]) if col_vt else 0.0
            if vt == 0.0 and qtd > 0 and vu > 0:
                vt = qtd * vu
            elif vu == 0.0 and qtd > 0 and vt > 0:
                vu = vt / qtd

            if qtd == 0.0 and vu == 0.0 and vt == 0.0:
                continue

            qtd_str = f"{int(qtd)}" if qtd == int(qtd) else f"{qtd:.2f}".replace('.', ',')

            itens.append({
                "item": item_num,
                "descricao": desc,
                "quantidade": qtd_str,
                "unidade": unid,
                "valor_unitario": formatar_moeda_br(vu),
                "valor_total": formatar_moeda_br(vt),
                "_vu_float": vu,
                "_vt_float": vt,
                "_qtd_float": qtd
            })
            
        soma_total_itens = sum(it.get("_vt_float", 0.0) for it in itens)
        if isinstance(tab_din, dict):
            if tab_din.get("total_geral", 0.0) > 0 and soma_total_itens > 0:
                if abs(tab_din["total_geral"] - soma_total_itens) < 0.05:
                    soma_total_itens = tab_din["total_geral"]
            tab_din["total_geral"] = soma_total_itens if soma_total_itens > 0 else tab_din.get("total_geral", 0.0)
            tab_din["total_formatado"] = formatar_moeda_br(tab_din["total_geral"])
            
        return ResultadoPlanilhaItens(itens, tabela_dinamica=tab_din)
    except Exception as e:
        print("Erro ao processar planilha de itens:", e)
        return ResultadoPlanilhaItens([])

def formatar_data_extenso(data_obj) -> str:
    try:
        if isinstance(data_obj, str):
            dt = datetime.strptime(data_obj[:10], "%Y-%m-%d")
        elif isinstance(data_obj, (date, datetime)):
            dt = data_obj
        else:
            dt = datetime.now()
        meses = [
            "janeiro", "fevereiro", "março", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
        ]
        return f"{dt.day} de {meses[dt.month - 1]} de {dt.year}"
    except Exception:
        return datetime.now().strftime("%d/%m/%Y")

class NumberedCanvas(canvas.Canvas):
    """
    Canvas oficial com cabeçalho (sol e faixas de Ribeirãozinho do Maranhão / Ribeirãozinho)
    e rodapé institucional oficial em todas as páginas.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        page_w, page_h = A4
        
        # 1. Cabeçalho Oficial (Topo da página, sangria total)
        if os.path.exists(HEADER_OFICIAL_PATH):
            header_h = 88.3
            self.drawImage(
                HEADER_OFICIAL_PATH,
                0,
                page_h - header_h,
                width=page_w,
                height=header_h,
                preserveAspectRatio=False,
                mask='auto'
            )
        else:
            self.setFont("Helvetica-Bold", 10)
            self.setFillColor(colors.HexColor("#0369a1"))
            self.drawCentredString(page_w / 2.0, page_h - 1.5 * cm, "PREFEITURA MUNICIPAL DE RIBEIRÃOZINHO DO MARANHÃO - MA")
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#059669"))
            self.drawCentredString(page_w / 2.0, page_h - 2.0 * cm, "ESTADO DO MARANHÃO — CNPJ: 01.597.627/0001-34")

        # 2. Rodapé Oficial (Base da página, sangria total)
        if os.path.exists(FOOTER_OFICIAL_PATH):
            footer_h = 63.3
            self.drawImage(
                FOOTER_OFICIAL_PATH,
                0,
                0,
                width=page_w,
                height=footer_h,
                preserveAspectRatio=False,
                mask='auto'
            )
        else:
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#64748b"))
            texto_rodape = f"Prefeitura Municipal de Ribeirãozinho do Maranhão - MA | CNPJ: 01.597.627/0001-34 | Página {self._pageNumber} de {page_count}"
            self.drawCentredString(page_w / 2.0, 1.2 * cm, texto_rodape)

        self.restoreState()

def preparar_dados_contrato(dados: Dict[str, Any]) -> Dict[str, Any]:
    """Garante que todos os campos necessários existam com valores padronizados"""
    d = dict(dados)
    d["municipio"] = "PREFEITURA MUNICIPAL DE RIBEIRÃOZINHO DO MARANHÃO - MA"
    d["estado"] = "ESTADO DO MARANHÃO"
    d["orgao_nome"] = d.get("orgao_nome") or d.get("orgao") or "SECRETARIA MUNICIPAL DE ADMINISTRAÇÃO"
    d["orgao_cnpj"] = formatar_cnpj_br(d.get("orgao_cnpj") or d.get("cnpj_orgao") or "01.612.834/0001-86")
    d["orgao_cep"] = d.get("orgao_cep") or d.get("cep_orgao") or "65928-000"
    d["orgao_endereco"] = d.get("orgao_endereco") or d.get("endereco_orgao") or "Rua Principal, s/n"
    d["orgao_bairro"] = d.get("orgao_bairro") or d.get("bairro_orgao") or "Centro"
    d["orgao_cidade"] = d.get("orgao_cidade") or d.get("cidade_orgao") or "Ribeirãozinho do Maranhão"
    d["orgao_uf"] = d.get("orgao_uf") or d.get("uf_orgao") or "MA"
    d["secretario"] = d.get("secretario_nome") or d.get("secretario") or "SECRETÁRIO MUNICIPAL TITULAR"
    d["cargo_secretario"] = d.get("secretario_cargo") or d.get("cargo_secretario") or "Secretário(a) Municipal Titular"
    d["secretario_cpf"] = d.get("secretario_cpf") or d.get("cpf_secretario") or "23*.***.**3-78"
    d["secretario_portaria"] = d.get("secretario_portaria") or d.get("portaria_secretario") or "Portaria Municipal de Nomeação"
    d["fiscal"] = d.get("fiscal_nome") or d.get("fiscal") or "Servidor Fiscal Designado"
    d["fiscal_cargo"] = d.get("fiscal_cargo") or "Fiscal do Contrato - Portaria Municipal"
    d["setor_licitacoes"] = d.get("setor_licitacoes") or d.get("setor") or "Comissão Permanente de Licitação - CPL"
    d["numero_imovel"] = d.get("numero_imovel") or d.get("orgao_numero") or "s/n"
    d["complemento_imovel"] = d.get("complemento_imovel") or d.get("orgao_complemento") or ""
    d["tipo_contratacao"] = d.get("tipo_contratacao") or ("AQUISIÇÕES" if "COMPRA" in str(d.get("modelo_agu", "")).upper() else "SERVIÇOS")
    
    d["fornecedor_nome"] = d.get("fornecedor_nome") or d.get("fornecedor") or "FORNECEDOR CONTRATADO LTDA"
    d["fornecedor_cnpj"] = formatar_cnpj_br(d.get("fornecedor_cnpj") or d.get("cnpj_cpf") or d.get("cnpj") or "00.000.000/0001-00")
    d["fornecedor_representante"] = d.get("fornecedor_representante") or d.get("fornecedor_rep") or d.get("nome_representante") or "Representante Legal da Contratada"
    d["fornecedor_cargo"] = d.get("fornecedor_cargo") or d.get("fornecedor_rep_cargo") or d.get("cargo_representante") or "Sócio Administrador"
    d["fornecedor_cpf"] = d.get("fornecedor_cpf") or d.get("fornecedor_rep_cpf") or d.get("cpf_representante") or "610.166.283-74"
    d["fornecedor_endereco"] = d.get("fornecedor_endereco") or d.get("endereco") or "Endereço Comercial da Contratada"
    d["fornecedor_num"] = d.get("fornecedor_num") or d.get("numero_fornecedor") or "s/n"
    d["fornecedor_bairro"] = d.get("fornecedor_bairro") or d.get("bairro") or "Centro"
    d["fornecedor_cep"] = d.get("fornecedor_cep") or d.get("cep") or "65900-000"
    d["fornecedor_cidade"] = d.get("fornecedor_cidade") or d.get("cidade") or "Imperatriz"
    d["fornecedor_uf"] = d.get("fornecedor_uf") or d.get("uf") or "MA"

    num_val = d.get("numero_completo") or d.get("numero") or d.get("numero_contrato") or "001/2026"
    d["numero_completo"] = str(num_val) if "/" in str(num_val) else f"{num_val}/2026"
    d["modalidade"] = d.get("modalidade") or "DISPENSA DE LICITAÇÃO"
    d["numero_modalidade"] = d.get("numero_modalidade") or "013/2026"
    d["processo_adm"] = d.get("processo_adm") or d.get("processo") or "001/2026-SEMED"
    d["objeto"] = d.get("objeto") or "Contratação administrativa conforme especificações técnicas do Termo de Referência."
    
    valor_float = float(d.get("valor_total") or 0.0)
    d["valor_total_float"] = valor_float
    d["valor_total_formatado"] = formatar_moeda_br(valor_float)
    d["valor_extenso"] = valor_por_extenso(valor_float)
    
    d["data_assinatura"] = d.get("data_assinatura") or date.today()
    d["data_vencimento"] = d.get("data_vencimento") or (date.today().replace(year=date.today().year + 1))
    d["data_assinatura_extenso"] = formatar_data_extenso(d["data_assinatura"])
    d["vigencia_descricao"] = d.get("vigencia_descricao") or "12 (doze) meses a contar da data de assinatura"
    d["dotacao_orcamentaria"] = d.get("dotacao_orcamentaria") or "Unidade Orçamentária: 02 - Recursos Ordinários / Próprios do Município"
    d["modelo_agu"] = (d.get("modelo_agu") or "COMPRAS").upper()
    d["foro"] = d.get("foro") or "Comarca de Imperatriz - MA"

    # Metadados de Dotação e Itens para Tabelas Visuais
    exercicio_padrao = str(d["data_assinatura"])[:4] if isinstance(d["data_assinatura"], (str, date, datetime)) else "2026"
    d["dotacao_exercicio"] = d.get("dotacao_exercicio") or exercicio_padrao
    d["dotacao_poder"] = d.get("dotacao_poder") or "Poder Executivo 02.00"
    d["dotacao_orgao"] = d.get("dotacao_orgao") or f"Fundo Municipal / {d['orgao_nome']} 02.15.00"
    d["dotacao_unidade"] = d.get("dotacao_unidade") or d.get("dotacao_orcamentaria") or "MANUTENÇÃO E FUNCIONAMENTO 12.361.0402.2022.0000"
    d["dotacao_natureza"] = d.get("dotacao_natureza") or ("Material de consumo 33.90.30.00" if "COMPRA" in str(d.get("tipo_contratacao", "")).upper() else "Outros Serviços de Terceiros 33.90.39.00")
    
    itens_in = d.get("itens")
    if not itens_in or not isinstance(itens_in, list) or len(itens_in) == 0:
        d["itens"] = [
            {
                "item": 1,
                "descricao": d["objeto"],
                "quantidade": "1",
                "unidade": "Global",
                "valor_unitario": d["valor_total_formatado"],
                "valor_total": d["valor_total_formatado"]
            }
        ]
    else:
        norm_itens = []
        tot_itens = 0.0
        for idx, it in enumerate(itens_in, start=1):
            num_it = it.get("item", idx)
            desc_it = it.get("descricao") or it.get("desc") or d["objeto"]
            qtd_it = it.get("quantidade") or it.get("qtd") or "1"
            unid_it = it.get("unidade") or it.get("un") or "Unid."
            vu = it.get("valor_unitario") or it.get("unitario") or d["valor_total_formatado"]
            vt = it.get("valor_total") or it.get("total") or d["valor_total_formatado"]
            
            vt_float = 0.0
            if "_vt_float" in it and it["_vt_float"]:
                try:
                    vt_float = float(it["_vt_float"])
                except Exception:
                    vt_float = 0.0
            elif isinstance(vt, (int, float)):
                vt_float = float(vt)
            else:
                s_vt = str(vt).replace("R$", "").replace(" ", "").strip()
                if "," in s_vt and "." in s_vt:
                    s_vt = s_vt.replace(".", "").replace(",", ".")
                elif "," in s_vt:
                    s_vt = s_vt.replace(",", ".")
                try:
                    vt_float = float(s_vt)
                except Exception:
                    vt_float = 0.0

            tot_itens += vt_float

            if isinstance(vu, (int, float)):
                vu = formatar_moeda_br(vu)
            if isinstance(vt, (int, float)):
                vt = formatar_moeda_br(vt)

            norm_itens.append({
                "item": num_it,
                "descricao": desc_it,
                "quantidade": str(qtd_it),
                "unidade": str(unid_it),
                "valor_unitario": str(vu),
                "valor_total": str(vt),
                "_vt_float": vt_float
            })
        d["itens"] = norm_itens

        tab_din = d.get("tabela_dinamica")
        if not tab_din and hasattr(itens_in, "tabela_dinamica"):
            tab_din = itens_in.tabela_dinamica
            d["tabela_dinamica"] = tab_din

        if tab_din and tab_din.get("total_geral", 0.0) > 0:
            tot_din = tab_din["total_geral"]
            if valor_float == 0.0 or d.get("usar_total_itens", False):
                valor_float = tot_din
                d["valor_total_float"] = valor_float
                d["valor_total_formatado"] = formatar_moeda_br(valor_float)
                d["valor_extenso"] = valor_por_extenso(valor_float)
                d["tabela_dinamica"]["total_formatado"] = formatar_moeda_br(valor_float)
        elif tot_itens > 0 and (valor_float == 0.0 or d.get("usar_total_itens", False)):
            valor_float = tot_itens
            d["valor_total_float"] = valor_float
            d["valor_total_formatado"] = formatar_moeda_br(valor_float)
            d["valor_extenso"] = valor_por_extenso(valor_float)

    # Higienização de campos para garantir sempre Ribeirãozinho do Maranhão
    for k, v in list(d.items()):
        if isinstance(v, str) and ("Edison" in v or "Lobão" in v or "Lobao" in v):
            d[k] = re.sub(r'Governador\s+Edison\s+Lob[ãa]o', 'Ribeirãozinho do Maranhão', v, flags=re.IGNORECASE)

    return d


def parse_dotacao_item(k, v):
    import re
    if k == "Exercício":
        return (k, v, "")
    m = re.match(r'^(.*?)\s+([\d\.\-]+)$', v.strip())
    if m:
        return (k, m.group(1), m.group(2))
    return (k, v, "")

def gerar_clausulas_agu(dados: Dict[str, Any]) -> list:
    """Monta a lista estruturada das cláusulas contratuais oficiais da AGU (Lei 14.133/2021) em estrita paridade com https://cgu.agu.gov.br/contrato/"""
    dados = preparar_dados_contrato(dados)
    mod = dados["modelo_agu"]

    raw_obj = str(dados.get('objeto', '')).strip()
    if raw_obj.lower().startswith(('o ', 'a ', 'os ', 'as ')):
        obj_txt = raw_obj
    elif raw_obj.lower().startswith(('fornecimento', 'aquisição', 'prestação', 'execução', 'contratação', 'locação')):
        obj_txt = f"o {raw_obj[0].lower()}{raw_obj[1:]}"
    else:
        obj_txt = f"o fornecimento/execução de {raw_obj}"

    # Cláusulas oficiais da AGU (Lei nº 14.133/2021)
    clausulas_base = [
        {
            "numero": "CLÁUSULA PRIMEIRA – DO OBJETO (art. 92, I e II)",
            "conteudo": [
                f"1.1. O objeto do presente instrumento é {obj_txt}, nas condições estabelecidas no Termo de Referência e na proposta da CONTRATADA.",
                "1.2. Descrição detalhada do objeto da contratação:",
                "1.3. O regime de execução contratual adotado observa integralmente as normas da Lei Federal nº 14.133, de 1º de abril de 2021, e seus regulamentos correlatos no âmbito da Administração Pública.",
                "1.4. Vinculam esta contratação, independentemente de transcrição, o Termo de Referência, o Edital da licitação (ou ato de contratação direta) e a Proposta de Preços formulada pela CONTRATADA."
            ]
        },
        {
            "numero": "CLÁUSULA SEGUNDA – DA VIGÊNCIA E PRORROGAÇÃO (arts. 105 a 114)",
            "conteudo": [
                f"2.1. O prazo de vigência deste Contrato é de {dados['vigencia_descricao']}, com início na data de sua assinatura e término previsto em {dados['data_vencimento']}.",
                "2.2. A vigência poderá ser prorrogada mediante termo aditivo formal, nos casos e condições expressamente autorizados pelos artigos 105 a 114 da Lei nº 14.133/2021, desde que demonstrada a vantajosa economicidade para a Administração Pública e a regularidade fiscal da CONTRATADA."
            ]
        },
        {
            "numero": "CLÁUSULA TERCEIRA – DOS MODELOS DE EXECUÇÃO E GESTÃO CONTRATUAIS (art. 117 e art. 140)",
            "conteudo": [
                f"3.1. A execução do presente contrato será acompanhada e fiscalizada pelo servidor {dados['fiscal']}, formalmente designado como Fiscal do Contrato pela autoridade competente, nos termos do art. 117 da Lei nº 14.133/2021.",
                "3.2. O recebimento provisório dar-se-á pelo fiscal no ato da entrega do objeto ou término da etapa, para verificação de conformidade de especificações e quantidades.",
                "3.3. O recebimento definitivo será formalizado no prazo de até 15 (quinze) dias após o recebimento provisório, mediante termo detalhado emitido pela comissão ou gestor competente, após a comprovação do atendimento integral das exigências contratuais.",
                "3.4. A fiscalização de que trata esta cláusula não exclui nem reduz a responsabilidade da CONTRATADA por quaisquer irregularidades, inclusive perante terceiros."
            ]
        },
        {
            "numero": "CLÁUSULA QUARTA – DA SUBCONTRATAÇÃO (art. 122)",
            "conteudo": [
                "4.1. É vedada a subcontratação total do objeto do presente contrato.",
                "4.2. A subcontratação parcial somente será admitida se expressamente autorizada pela CONTRATANTE, nos limites e condições estipulados no Termo de Referência e no art. 122 da Lei Federal nº 14.133/2021, mantendo-se a CONTRATADA com a responsabilidade integral pela perfeita execução deste instrumento.",
                "4.3. É vedada a subcontratação de pessoa física ou jurídica que tenha participado do procedimento de licitação ou que se encontre impedida ou declarada inidônea para licitar ou contratar com o Poder Público."
            ]
        },
        {
            "numero": "CLÁUSULA QUINTA – DOS PREÇOS E DO VALOR GLOBAL (art. 92, V)",
            "conteudo": [
                f"5.1. O valor total global deste Contrato é de {dados['valor_total_formatado']} ({dados['valor_extenso']}).",
                "5.2. No valor acima estão incluídas todas as despesas ordinárias diretas e indiretas decorrentes da execução do objeto, inclusive tributos, encargos sociais, trabalhistas, previdenciários, fiscais e comerciais, fretes, seguros e quaisquer outros que incidam ou venham a incidir sobre o cumprimento do ajuste."
            ]
        },
        {
            "numero": "CLÁUSULA SEXTA – DA DOTAÇÃO ORÇAMENTÁRIA (art. 92, VIII e art. 150)",
            "conteudo": [
                f"6.1. As despesas decorrentes da presente contratação correrão à conta dos créditos orçamentários específicos da(o) {dados['orgao_nome']}, conforme a seguinte classificação orçamentária:",
                f"    {dados['dotacao_orcamentaria']}",
                "6.2. A dotação relativa aos exercícios financeiros subsequentes será indicada após a aprovação da respectiva Lei Orçamentária Anual, mediante simples apostilamento."
            ]
        },
        {
            "numero": "CLÁUSULA SÉTIMA – DO PAGAMENTO E LIQUIDAÇÃO (arts. 141 a 146)",
            "conteudo": [
                "7.1. O pagamento será realizado no prazo de até 30 (trinta) dias consecutivos contados do recebimento definitivo do objeto ou serviço e da apresentação da respectiva Nota Fiscal/Fatura devidamente atestada pelo Fiscal do Contrato.",
                "7.2. A liquidação da despesa observará rigorosamente a ordem cronológica para cada fonte diferenciada de recursos, nos moldes do art. 141 da Lei Federal nº 14.133/2021.",
                "7.3. Nenhum pagamento será efetuado enquanto pendente de liquidação qualquer obrigação financeira ou apresentação de regularidade fiscal, previdenciária e trabalhista por parte da CONTRATADA.",
                "7.4. Em caso de atraso injustificado no pagamento imputável à Administração, os valores devidos serão atualizados monetariamente pelo IPCA/IBGE pro rata tempore, desde o vencimento até a data do efetivo adimplemento."
            ]
        },
        {
            "numero": "CLÁUSULA OITAVA – DO REAJUSTE DE PREÇOS (art. 92, V e § 3º)",
            "conteudo": [
                "8.1. Os preços são fixos e irreajustáveis pelo período de 12 (doze) meses contados da data limite para apresentação da proposta ou do orçamento estimado.",
                "8.2. Após o interregno de 1 (um) ano, os preços poderão ser reajustados mediante aplicação do Índice Nacional de Preços ao Consumidor Amplo (IPCA/IBGE) ou índice setorial oficial correspondente, exclusivamente para compensar a variação efetiva dos custos de produção."
            ]
        },
        {
            "numero": "CLÁUSULA NONA – DA GARANTIA CONTRATUAL (art. 96)",
            "conteudo": [
                f"9.1. {dados.get('garantia_texto') or 'Não será exigida prestação de garantia de execução contratual para a presente contratação, em consonância com o art. 96 da Lei Federal nº 14.133/2021 e as justificativas técnicas constantes dos autos do processo administrativo.'}"
            ]
        },
        {
            "numero": "CLÁUSULA DÉCIMA – DAS OBRIGAÇÕES DA CONTRATANTE (art. 92, X)",
            "conteudo": [
                "10.1. Exigir o cumprimento de todas as obrigações assumidas pela CONTRATADA, de acordo com as cláusulas contratuais e os termos de sua proposta.",
                "10.2. Proporcionar todas as condições necessárias para que a CONTRATADA possa desempenhar regularmente o fornecimento ou a prestação dos serviços contratados.",
                "10.3. Efetuar o pagamento à CONTRATADA no valor e prazo pactuados, após o regular atesto da execução.",
                "10.4. Notificar formalmente a CONTRATADA sobre quaisquer irregularidades, imperfeições ou falhas constatadas na execução do objeto, fixando prazo razoável para a sua correção."
            ]
        },
        {
            "numero": "CLÁUSULA DÉCIMA PRIMEIRA – DAS OBRIGAÇÕES DA CONTRATADA (art. 92, XIV, XVI e XVII)",
            "conteudo": [
                "11.1. Executar o objeto contratado com presteza, perfeição e em estrita consonância com as especificações da proposta e do Termo de Referência.",
                "11.2. Manter, durante toda a vigência do contrato, em compatibilidade com as obrigações por ela assumidas, todas as condições de habilitação e qualificação exigidas na contratação.",
                "11.3. Reparar, corrigir, remover, reconstruir ou substituir, às suas expensas, no total ou em parte, o objeto do contrato em que se verificarem vícios, defeitos ou incorreções.",
                "11.4. Responsabilizar-se pelos vícios e danos decorrentes da execução do objeto, bem como por todo e qualquer dano causado à Administração Pública ou a terceiros.",
                "11.5. Cumprir integralmente as normas da Lei Federal nº 13.709/2018 (Lei Geral de Proteção de Dados Pessoais - LGPD), implementando medidas técnicas e administrativas aptas a resguardar dados pessoais acessados em decorrência do contrato.",
                "11.6. Observar a reserva de vagas para pessoas com deficiência, reabilitados da Previdência Social e mulheres vítimas de violência doméstica, nos termos do art. 116 da Lei nº 14.133/2021, quando aplicável."
            ]
        },
        {
            "numero": "CLÁUSULA DÉCIMA SEGUNDA – DAS INFRAÇÕES E SANÇÕES ADMINISTRATIVAS (art. 156)",
            "conteudo": [
                "12.1. O descumprimento total ou parcial das obrigações assumidas ensejará a aplicação das sanções administrativas previstas no art. 156 da Lei Federal nº 14.133/2021, garantida a prévia e ampla defesa:",
                "    a) Advertência por escrito;",
                "    b) Multa moratória ou compensatória de até 20% (vinte por cento) sobre o valor do contrato;",
                "    c) Impedimento de licitar e contratar com a Administração Pública Municipal pelo prazo de até 3 (três) anos;",
                "    d) Declaração de inidoneidade para licitar ou contratar com a Administração Pública de todos os entes federativos pelo prazo de 3 (três) a 6 (seis) anos."
            ]
        },
        {
            "numero": "CLÁUSULA DÉCIMA TERCEIRA – DA EXTINÇÃO CONTRATUAL (arts. 137 a 139)",
            "conteudo": [
                "13.1. O presente contrato poderá ser extinto nas hipóteses e condições previstas nos artigos 137, 138 e 139 da Lei Federal nº 14.133/2021, sem prejuízo da incidência das sanções cabíveis.",
                "13.2. A extinção determinada por ato unilateral da Administração Pública será precedida de autorização escrita e fundamentada da autoridade competente e de regular processo administrativo."
            ]
        },
        {
            "numero": "CLÁUSULA DÉCIMA QUARTA – DA PUBLICAÇÃO NO PORTAL NACIONAL DE CONTRATAÇÕES PÚBLICAS (PNCP) (art. 94)",
            "conteudo": [
                "14.1. Incumbirá à CONTRATANTE divulgar e manter o inteiro teor deste Contrato no Portal Nacional de Contratações Públicas (PNCP), na forma preconizada pelo art. 94 da Lei Federal nº 14.133/2021, no prazo de até 20 (vinte) dias úteis contados da data de sua assinatura, condição indispensável para a sua eficácia plena perante terceiros."
            ]
        }
    ]

    # Definição dos ordinais para numeração dinâmica
    ordinais = ["PRIMEIRA", "SEGUNDA", "TERCEIRA", "QUARTA", "QUINTA", "SEXTA", "SÉTIMA", "OITAVA", "NONA", "DÉCIMA", 
                "DÉCIMA PRIMEIRA", "DÉCIMA SEGUNDA", "DÉCIMA TERCEIRA", "DÉCIMA QUARTA", "DÉCIMA QUINTA", "DÉCIMA SEXTA",
                "DÉCIMA SÉTIMA", "DÉCIMA OITAVA", "DÉCIMA NONA", "VIGÉSIMA"]

    # Verifica se há cláusulas adicionais formatadas enviadas do frontend
    clausulas_extras = dados.get("clausulas_adicionais_formatadas")
    if clausulas_extras and isinstance(clausulas_extras, list) and len(clausulas_extras) > 0:
        idx_nova = len(clausulas_base)
        nome_ordinal = ordinais[idx_nova] if idx_nova < len(ordinais) else f"{idx_nova + 1}ª"
        
        titulo_usuario = dados.get("titulo_clausula_adicional")
        if not titulo_usuario or not str(titulo_usuario).strip():
            titulo_final = "DAS CONDIÇÕES ESPECIAIS E DISPOSIÇÕES COMPLEMENTARES (LEI Nº 14.133/2021)"
        else:
            titulo_limpo = str(titulo_usuario).strip()
            titulo_limpo = re.sub(r'^(?:CL[AÁ]USULA\s+[A-Z0-9ªº\-–\s]+[–\-:]\s*)+', '', titulo_limpo, flags=re.IGNORECASE).strip()
            titulo_final = titulo_limpo.upper() if titulo_limpo else "DAS CONDIÇÕES ESPECIAIS E DISPOSIÇÕES COMPLEMENTARES (LEI Nº 14.133/2021)"

        clausula_extra_dict = {
            "numero": f"CLÁUSULA {nome_ordinal} – {titulo_final}",
            "conteudo": clausulas_extras
        }
        clausulas_base.append(clausula_extra_dict)

    # Adicionar sempre a Cláusula do Foro no final
    idx_foro = len(clausulas_base)
    nome_foro = ordinais[idx_foro] if idx_foro < len(ordinais) else f"{idx_foro + 1}ª"
    clausulas_base.append({
        "numero": f"CLÁUSULA {nome_foro} – DO FORO (art. 92, § 1º)",
        "conteudo": [
            "Fica eleito o Foro da Comarca de Imperatriz - MA, com expressa renúncia a qualquer outro, por mais privilegiado que seja, para dirimir quaisquer dúvidas ou litígios decorrentes da execução do presente instrumento contratual que não puderem ser compostos por conciliação, conforme prevê o art. 92, § 1º, da Lei Federal nº 14.133/2021."
        ]
    })

    return clausulas_base

def gerar_contrato_docx(dados_brutos: Dict[str, Any]) -> io.BytesIO:
    """Gera o arquivo oficial Word (.docx) padronizado da AGU / Prefeitura"""
    dados = preparar_dados_contrato(dados_brutos)
    doc = docx.Document()

    # Configurar margens compatíveis com o layout oficial e sangria
    for sec in doc.sections:
        sec.top_margin = Inches(1.3)
        sec.bottom_margin = Inches(1.0)
        sec.left_margin = Inches(0.8)
        sec.right_margin = Inches(0.8)
        
        # Cabeçalho Oficial no Word
        if os.path.exists(HEADER_OFICIAL_PATH):
            header = sec.header
            header.is_linked_to_previous = False
            p_head = header.paragraphs[0]
            p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_head = p_head.add_run()
            r_head.add_picture(HEADER_OFICIAL_PATH, width=Inches(6.9))

        # Rodapé Oficial no Word
        if os.path.exists(FOOTER_OFICIAL_PATH):
            footer = sec.footer
            footer.is_linked_to_previous = False
            p_foot = footer.paragraphs[0]
            p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_foot = p_foot.add_run()
            r_foot.add_picture(FOOTER_OFICIAL_PATH, width=Inches(6.9))

    # 1. Bloco de Identificação Superior (Alinhado à esquerda como no modelo original)
    p_id = doc.add_paragraph()
    p_id.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_id.paragraph_format.line_spacing = 1.25
    p_id.paragraph_format.space_before = Pt(4)
    p_id.paragraph_format.space_after = Pt(12)

    r1 = p_id.add_run(f"TERMO DE CONTRATO Nº {dados['numero_completo']}\n")
    r1.bold = True
    r1.font.name = "Arial"
    r1.font.size = Pt(10)

    num_mod_limpo = str(dados.get('numero_modalidade') or '010/2026').strip()
    num_mod_limpo = re.sub(r'^(n[ºo°\. ]+|n\.\s*)', '', num_mod_limpo, flags=re.IGNORECASE).strip()
    mod_nome = str(dados.get('modalidade') or 'PREGÃO').strip()

    r2 = p_id.add_run(f"{mod_nome.upper()} Nº {num_mod_limpo}\n")
    r2.bold = True
    r2.font.name = "Arial"
    r2.font.size = Pt(10)

    r3 = p_id.add_run(f"PROCESSO ADMINISTRATIVO Nº {dados['processo_adm']}")
    r3.bold = True
    r3.font.name = "Arial"
    r3.font.size = Pt(10)

    # 2. Ementa Oficial do Contrato (Abaixo da identificação, recuada à direita)
    p_em = doc.add_paragraph()
    p_em.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_em.paragraph_format.left_indent = Inches(3.2)
    p_em.paragraph_format.line_spacing = 1.15
    p_em.paragraph_format.space_before = Pt(8)
    p_em.paragraph_format.space_after = Pt(18)

    nome_mun_ementa = "MUNICÍPIO DE RIBEIRÃOZINHO DO MARANHÃO - MA"
    r_em = p_em.add_run(
        f"CONTRATO ADMINISTRATIVO Nº {dados['numero_completo']} QUE FAZEM ENTRE SI "
        f"O {nome_mun_ementa}, POR INTERMÉDIO DA {dados['orgao_nome'].upper()} "
        f"E A EMPRESA {dados['fornecedor_nome'].upper()}."
    )
    r_em.bold = True
    r_em.font.name = "Arial"
    r_em.font.size = Pt(9.5)

    # Preâmbulo Jurídico: ajuste gramatical de preposição conforme a modalidade
    mod_lower = mod_nome.lower()
    if "concorr" in mod_lower:
        decorrente_texto = f"decorrente da {mod_nome} nº {num_mod_limpo}"
    elif "dispensa" in mod_lower or "inexig" in mod_lower or "contrata" in mod_lower:
        decorrente_texto = f"decorrente de {mod_nome} nº {num_mod_limpo}"
    else:
        decorrente_texto = f"decorrente do {mod_nome} nº {num_mod_limpo}"

    preambulo_texto = (
        f"Pelo presente contrato, celebram entre si, de um lado, o Município de Ribeirãozinho do Maranhão – MA, "
        f"por intermédio da {dados['orgao_nome']}, com sede na {dados['orgao_endereco']}, nº {dados['numero_imovel']}, "
        f"Bairro {dados['orgao_bairro']}, Ribeirãozinho do Maranhão – MA, inscrita no CNPJ sob o nº {dados['orgao_cnpj']}, "
        f"neste ato representada pelo {dados['cargo_secretario']}, {dados['secretario']}, "
        f"portador (a) do CPF nº {dados['secretario_cpf']}, doravante denominado CONTRATANTE, e, de outro lado, a "
        f"{dados['fornecedor_nome']}, inscrita no CNPJ nº {dados['fornecedor_cnpj']}, com sede na "
        f"{dados['fornecedor_endereco']}, nº {dados['fornecedor_num']}, Bairro {dados['fornecedor_bairro']}, "
        f"{dados['fornecedor_cidade']}/{dados['fornecedor_uf']}, CEP {dados['fornecedor_cep']}, na qualidade de Fornecedor Registrado, "
        f"neste ato representada por seu representante legal, {dados['fornecedor_representante']}, "
        f"CPF nº {dados['fornecedor_cpf']}. Tendo em vista o Processo Administrativo nº {dados['processo_adm']} e a "
        f"Lei nº 14.133, de 1º de abril de 2021, resolvem celebrar o presente Termo de Contrato, {decorrente_texto}, "
        f"mediante as cláusulas e condições a seguir enunciadas."
    )
    p_pre = doc.add_paragraph()
    p_pre.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_pre.paragraph_format.first_line_indent = Inches(0.5)
    p_pre.paragraph_format.line_spacing = 1.15
    p_pre.paragraph_format.space_after = Pt(10)
    r_pre = p_pre.add_run(preambulo_texto)
    r_pre.font.name = "Times New Roman"
    r_pre.font.size = Pt(10.5)

    # 3. Cláusulas Padronizadas AGU
    clausulas = gerar_clausulas_agu(dados)
    for cl in clausulas:
        p_c_tit = doc.add_paragraph()
        p_c_tit.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_c_tit = p_c_tit.add_run(cl["numero"])
        r_c_tit.bold = True
        r_c_tit.font.name = "Arial"
        r_c_tit.font.size = Pt(10)
        p_c_tit.paragraph_format.space_before = Pt(8)
        p_c_tit.paragraph_format.space_after = Pt(2)

        is_clausula_objeto = "OBJETO" in cl["numero"].upper()
        is_clausula_dotacao = "DOTAÇÃO" in cl["numero"].upper() or "DOTACAO" in cl["numero"].upper()

        if is_clausula_objeto:
            itens_list = dados.get("itens", [])
            if not itens_list:
                itens_list = [{
                    "item": 1,
                    "descricao": dados.get("objeto", "Item contratado"),
                    "quantidade": "1",
                    "unidade": "Unid.",
                    "valor_unitario": dados.get("valor_total_formatado", "R$ 0,00"),
                    "valor_total": dados.get("valor_total_formatado", "R$ 0,00"),
                    "_vt_float": 0.0
                }]

            tot_float = sum(float(it.get("_vt_float", 0.0) or 0.0) for it in itens_list)
            if tot_float > 0:
                tot_display = f"R$ {tot_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                tot_display = dados.get("valor_total_formatado", "R$ 0,00")

            for idx_p, par in enumerate(cl["conteudo"]):
                p_c = doc.add_paragraph()
                p_c.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_c.paragraph_format.first_line_indent = Inches(0.4)
                p_c.paragraph_format.line_spacing = 1.15
                p_c.paragraph_format.space_after = Pt(3)
                r_c = p_c.add_run(par)
                r_c.font.name = "Times New Roman"
                r_c.font.size = Pt(10)

                if par.startswith("1.2.") or (len(cl["conteudo"]) == 2 and idx_p == 0):
                    # Inserir Tabela de Itens no Word (Auto-ajustável para qualquer quantidade de colunas e cabeçalhos)
                    tab_din = dados.get("tabela_dinamica")
                    if not tab_din and hasattr(dados.get("itens"), "tabela_dinamica"):
                        tab_din = dados["itens"].tabela_dinamica

                    if tab_din and tab_din.get("headers") and tab_din.get("rows"):
                        dyn_headers = tab_din["headers"]
                        dyn_rows = tab_din["rows"]
                        dyn_alignments = tab_din.get("col_alignments", ["LEFT"] * len(dyn_headers))
                        tot_col = tab_din.get("total_col_idx")
                        tot_val_txt = tab_din.get("total_formatado") or tot_display

                        n_cols = len(dyn_headers)
                        n_rows = len(dyn_rows) + 1 + (1 if tot_col is not None else 0)
                        t_it = doc.add_table(rows=n_rows, cols=n_cols)
                        aplicar_estilo_tabela_oficial_docx(t_it)

                        # Cabeçalho dinâmico
                        hdr_cells = t_it.rows[0].cells
                        for c_idx, h_text in enumerate(dyn_headers):
                            hdr_cells[c_idx].text = str(h_text)
                            aplicar_fundo_celula_docx(hdr_cells[c_idx], 'E2EBF4')
                            p = hdr_cells[c_idx].paragraphs[0]
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for r in p.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(8.5)
                                r.bold = True
                                r.font.color.rgb = RGBColor(34, 64, 97)

                        # Linhas de dados
                        for r_idx, row_vals in enumerate(dyn_rows):
                            row_cells = t_it.rows[r_idx + 1].cells
                            for c_idx, cell_val in enumerate(row_vals):
                                if c_idx < len(row_cells):
                                    row_cells[c_idx].text = str(cell_val)
                                    p = row_cells[c_idx].paragraphs[0]
                                    al = dyn_alignments[c_idx] if c_idx < len(dyn_alignments) else "LEFT"
                                    if al == "RIGHT":
                                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                                    elif al == "CENTER":
                                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    else:
                                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                    for r in p.runs:
                                        r.font.name = "Arial"
                                        r.font.size = Pt(8)

                        # Linha de total se aplicável
                        if tot_col is not None:
                            tot_row_cells = t_it.rows[-1].cells
                            # Mescla da primeira coluna até a penúltima antes do total
                            merge_end = tot_col - 1 if tot_col > 0 else 0
                            a = tot_row_cells[0]
                            b = tot_row_cells[merge_end]
                            if merge_end > 0:
                                a.merge(b)
                            a.text = "TOTAL GERAL:"
                            p_lbl = a.paragraphs[0]
                            p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            for r in p_lbl.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(8.5)
                                r.bold = True
                                r.font.color.rgb = RGBColor(0, 51, 102)

                            tot_row_cells[tot_col].text = tot_val_txt
                            p_val = tot_row_cells[tot_col].paragraphs[0]
                            p_val.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            for r in p_val.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(8.5)
                                r.bold = True
                                r.font.color.rgb = RGBColor(0, 51, 102)
                    else:
                        # Fallback padrão de tabela com 6 colunas fixas
                        t_it = doc.add_table(rows=len(itens_list) + 2, cols=6)
                        col_w = [Inches(0.5), Inches(3.2), Inches(0.6), Inches(0.6), Inches(1.0), Inches(1.0)]
                        aplicar_estilo_tabela_oficial_docx(t_it, col_widths_pt=[c.pt for c in col_w])

                        headers = ["Item", "Descrição do Objeto / Especificação", "Qtd", "Unid.", "Valor Unit.", "Valor Total"]
                        hdr_cells = t_it.rows[0].cells
                        for i_h, h in enumerate(headers):
                            hdr_cells[i_h].text = h
                            aplicar_fundo_celula_docx(hdr_cells[i_h], 'E2EBF4')
                            p_h = hdr_cells[i_h].paragraphs[0]
                            p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for r_h in p_h.runs:
                                r_h.font.name = "Arial"
                                r_h.font.size = Pt(8.5)
                                r_h.bold = True
                                r_h.font.color.rgb = RGBColor(34, 64, 97)

                        for idx_it, row_it in enumerate(itens_list):
                            r_cells = t_it.rows[idx_it + 1].cells
                            vals = [
                                str(row_it.get("item", idx_it + 1)),
                                str(row_it.get("descricao", "")),
                                str(row_it.get("quantidade", "1")),
                                str(row_it.get("unidade", "Unid.")),
                                str(row_it.get("valor_unitario", dados["valor_total_formatado"])),
                                str(row_it.get("valor_total", dados["valor_total_formatado"]))
                            ]
                            for c_i, v in enumerate(vals):
                                r_cells[c_i].text = v
                                p_c = r_cells[c_i].paragraphs[0]
                                p_c.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_i == 1 else (WD_ALIGN_PARAGRAPH.RIGHT if c_i >= 4 else WD_ALIGN_PARAGRAPH.CENTER)
                                for r in p_c.runs:
                                    r.font.name = "Arial"
                                    r.font.size = Pt(8)

                        # Linha Total
                        tot_cells = t_it.rows[-1].cells
                        a = tot_cells[0]
                        b = tot_cells[4]
                        a.merge(b)
                        a.text = "TOTAL:"
                        p_tot_lbl = a.paragraphs[0]
                        p_tot_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        rtotl = p_tot_lbl.runs[0]
                        rtotl.bold = True
                        rtotl.font.name = "Arial"
                        rtotl.font.size = Pt(8.5)

                        tot_cells[5].text = tot_display
                        p_tot_val = tot_cells[5].paragraphs[0]
                        p_tot_val.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        rtotv = p_tot_val.runs[0]
                        rtotv.bold = True
                        rtotv.font.name = "Arial"
                        rtotv.font.size = Pt(8.5)
        elif is_clausula_dotacao:
            for idx_d, par in enumerate(cl["conteudo"]):
                p_c = doc.add_paragraph()
                p_c.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_c.paragraph_format.first_line_indent = Inches(0.4)
                p_c.paragraph_format.line_spacing = 1.15
                p_c.paragraph_format.space_after = Pt(3)
                r_c = p_c.add_run(par)
                r_c.font.name = "Times New Roman"
                r_c.font.size = Pt(10)

                if idx_d == 0:
                    dot_items = [
                        ("Exercício", str(dados.get("dotacao_exercicio") or "2026")),
                        ("Poder", str(dados.get("dotacao_poder") or "Poder Executivo 02.00")),
                        ("Órgão", str(dados.get("dotacao_orgao") or f"Fundo Municipal / {dados['orgao_nome']} 02.15.00")),
                        ("Unidade Orçamentária/atividade", str(dados.get("dotacao_unidade") or dados.get("dotacao_orcamentaria") or "MANUTENÇÃO E FUNCIONAMENTO 12.361.0402.2022.0000")),
                        ("Natureza da despesa", str(dados.get("dotacao_natureza") or ("Material de consumo 33.90.30.00" if "COMPRA" in str(dados.get("tipo_contratacao", "")).upper() else "Outros Serviços de Terceiros 33.90.39.00")))
                    ]
                    if dados.get("dotacao_fonte") or dados.get("fonte_recursos") or dados.get("fonte"):
                        dot_items.append(("Fonte de recursos", str(dados.get("dotacao_fonte") or dados.get("fonte_recursos") or dados.get("fonte"))))
                    if dados.get("dotacao_nota_empenho") or dados.get("nota_empenho") or dados.get("empenho"):
                        dot_items.append(("Nota de empenho", str(dados.get("dotacao_nota_empenho") or dados.get("nota_empenho") or dados.get("empenho"))))

                    t_dot = doc.add_table(rows=len(dot_items), cols=3)
                    t_dot.alignment = WD_TABLE_ALIGNMENT.CENTER
                    aplicar_estilo_tabela_oficial_docx(t_dot, col_widths_pt=[120.0, 240.0, 120.0])
                    for r_i, (lbl, val) in enumerate(dot_items):
                        k, v1, v2 = parse_dotacao_item(lbl, val)
                        c0 = t_dot.cell(r_i, 0)
                        c1 = t_dot.cell(r_i, 1)
                        c2 = t_dot.cell(r_i, 2)
                        aplicar_fundo_celula_docx(c0, 'E2EBF4')
                        
                        r0 = c0.paragraphs[0].add_run(k)
                        r0.bold = True
                        r0.font.name = "Arial"
                        r0.font.size = Pt(8.5)
                        r0.font.color.rgb = RGBColor(34, 64, 97)
                        
                        r1 = c1.paragraphs[0].add_run(v1)
                        r1.font.name = "Arial"
                        r1.font.size = Pt(8.5)
                        
                        r2 = c2.paragraphs[0].add_run(v2)
                        r2.font.name = "Arial"
                        r2.font.size = Pt(8.5)
                        c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            for par in cl["conteudo"]:
                p_c_corpo = doc.add_paragraph()
                p_c_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_c_corpo.paragraph_format.first_line_indent = Inches(0.4)
                p_c_corpo.paragraph_format.line_spacing = 1.15
                p_c_corpo.paragraph_format.space_after = Pt(4)
                r_par = p_c_corpo.add_run(par)
                r_par.font.size = Pt(10)
                r_par.font.name = "Times New Roman"

    # Fecho do Contrato
    p_fecho = doc.add_paragraph()
    p_fecho.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_fecho.paragraph_format.first_line_indent = Inches(0.5)
    p_fecho.paragraph_format.space_before = Pt(12)
    p_fecho.paragraph_format.space_after = Pt(12)
    r_f = p_fecho.add_run(
        "E, por estarem assim justos e contratados, assinam o presente instrumento em 2 (duas) vias de igual teor e forma, "
        "na presença de 2 (duas) testemunhas instrumentárias, para que produza todos os seus jurídicos e legais efeitos."
    )
    r_f.font.size = Pt(11)
    r_f.font.name = "Times New Roman"

    p_loc = doc.add_paragraph()
    p_loc.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_loc.paragraph_format.space_before = Pt(8)
    p_loc.paragraph_format.space_after = Pt(20)
    r_loc = p_loc.add_run(f"Ribeirãozinho do Maranhão – MA, {dados['data_assinatura_extenso']}.\n\n")
    r_loc.font.size = Pt(10.5)
    r_loc.font.name = "Times New Roman"

    # Assinaturas em Tabela 2 Colunas
    t_ass = doc.add_table(rows=2, cols=2)
    t_ass.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_ass.autofit = False

    c00 = t_ass.cell(0, 0)
    c01 = t_ass.cell(0, 1)
    c10 = t_ass.cell(1, 0)
    c11 = t_ass.cell(1, 1)

    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p00.add_run(f"_____________________________________________\n{dados['secretario']}\n{dados['cargo_secretario']}\n{dados['secretario_portaria']}\nContratante\n\n")

    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p01.add_run(f"_____________________________________________\n{dados['fornecedor_nome']}\n{dados['fornecedor_representante']}\n{dados['fornecedor_cargo']}\nContratada\n\n")

    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p10.add_run(f"_____________________________________________\n{dados['fiscal']}\nFiscal do Contrato - Portaria Municipal")

    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p11.add_run("TESTEMUNHAS:\n\n1. ___________________________ CPF:\n2. ___________________________ CPF:")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_contrato_pdf(dados_brutos: Dict[str, Any]) -> io.BytesIO:
    """Gera o arquivo oficial PDF (.pdf) padronizado da AGU / Prefeitura"""
    dados = preparar_dados_contrato(dados_brutos)
    buffer = io.BytesIO()

    # Dimensões A4: 595.27 x 841.89 pt.
    # Cabeçalho: 88.3 pt. Rodapé: 63.3 pt.
    # Margens calculadas para garantir que o conteúdo nunca colida com os banners oficiais:
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=45,
        rightMargin=45,
        topMargin=96,
        bottomMargin=70
    )

    styles = getSampleStyleSheet()
    page_content_w = 595.27 - 90  # 505.27 pt

    # Estilos tipográficos oficiais
    style_top_left = ParagraphStyle(
        'TopLeft',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor("#0f172a"),
        alignment=0
    )
    style_top_right = ParagraphStyle(
        'TopRight',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#0f172a"),
        alignment=4
    )
    style_preamble = ParagraphStyle(
        'Preamble',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=9.5,
        leading=13.5,
        alignment=4,
        firstLineIndent=15,
        spaceBefore=8,
        spaceAfter=10
    )
    style_clause_title = ParagraphStyle(
        'ClauseTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=13,
        alignment=0,
        textColor=colors.HexColor("#0f172a"),
        spaceBefore=9,
        spaceAfter=3,
        keepWithNext=True
    )
    style_clause_body = ParagraphStyle(
        'ClauseBody',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=9,
        leading=13,
        alignment=4,
        firstLineIndent=14,
        spaceAfter=3
    )
    style_bullet = ParagraphStyle(
        'ClauseBullet',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=9,
        leading=12.5,
        alignment=4,
        leftIndent=15,
        spaceAfter=2
    )
    style_tbl_th = ParagraphStyle(
        'TblTH',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.white,
        alignment=1
    )
    style_tbl_td = ParagraphStyle(
        'TblTD',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#1e293b"),
        alignment=0
    )
    style_tbl_td_c = ParagraphStyle(
        'TblTDC',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#1e293b"),
        alignment=1
    )
    style_tbl_td_r = ParagraphStyle(
        'TblTDR',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#1e293b"),
        alignment=2
    )
    style_sig_sub = ParagraphStyle(
        'SigSub',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=colors.HexColor("#334155")
    )

    story = []

    # 1. Bloco de Identificação Superior (Alinhado à esquerda como no modelo original)
    style_identificacao = ParagraphStyle(
        'TopIdentificacao',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5,
        leading=14,
        textColor=colors.HexColor("#0f172a"),
        alignment=0,
        spaceAfter=12
    )
    num_mod_limpo_pdf = str(dados.get('numero_modalidade') or '010/2026').strip()
    num_mod_limpo_pdf = re.sub(r'^(n[ºo°\. ]+|n\.\s*)', '', num_mod_limpo_pdf, flags=re.IGNORECASE).strip()
    mod_nome_pdf = str(dados.get('modalidade') or 'PREGÃO').strip()

    texto_id = (
        f"<b>TERMO DE CONTRATO Nº {dados['numero_completo']}</b><br/>"
        f"<b>{mod_nome_pdf.upper()} Nº {num_mod_limpo_pdf}</b><br/>"
        f"<b>PROCESSO ADMINISTRATIVO Nº {dados['processo_adm']}</b>"
    )
    story.append(Paragraph(texto_id, style_identificacao))

    # 2. Ementa Oficial do Contrato (Abaixo da identificação, recuada à direita)
    style_ementa_oficial = ParagraphStyle(
        'EmentaOficial',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor("#0f172a"),
        alignment=4,  # Justified
        leftIndent=220,  # Recuo de 220 pt posicionando o bloco na metade direita
        spaceAfter=16
    )
    nome_mun_ementa = "MUNICÍPIO DE RIBEIRÃOZINHO DO MARANHÃO - MA"
    texto_ementa = (
        f"<b>CONTRATO ADMINISTRATIVO Nº {dados['numero_completo']} QUE FAZEM ENTRE SI "
        f"O {nome_mun_ementa}, POR INTERMÉDIO DA {dados['orgao_nome'].upper()} "
        f"E A EMPRESA {dados['fornecedor_nome'].upper()}.</b>"
    )
    story.append(Paragraph(texto_ementa, style_ementa_oficial))

    # 2. Preâmbulo Jurídico: ajuste gramatical de preposição conforme a modalidade
    mod_lower_pdf = mod_nome_pdf.lower()
    if "concorr" in mod_lower_pdf:
        decorrente_texto_pdf = f"decorrente da {mod_nome_pdf} nº {num_mod_limpo_pdf}"
    elif "dispensa" in mod_lower_pdf or "inexig" in mod_lower_pdf or "contrata" in mod_lower_pdf:
        decorrente_texto_pdf = f"decorrente de {mod_nome_pdf} nº {num_mod_limpo_pdf}"
    else:
        decorrente_texto_pdf = f"decorrente do {mod_nome_pdf} nº {num_mod_limpo_pdf}"

    preambulo_texto = (
        f"Pelo presente contrato, celebram entre si, de um lado, o <b>Município de Ribeirãozinho do Maranhão – MA</b>, "
        f"por intermédio da <b>{dados['orgao_nome']}</b>, com sede na {dados['orgao_endereco']}, nº {dados['numero_imovel']}, "
        f"Bairro {dados['orgao_bairro']}, Ribeirãozinho do Maranhão – MA, inscrita no CNPJ sob o nº <b>{dados['orgao_cnpj']}</b>, "
        f"neste ato representada pelo(a) {dados['cargo_secretario']}, <b>{dados['secretario']}</b>, "
        f"portador(a) do CPF nº {dados['secretario_cpf']}, doravante denominado <b>CONTRATANTE</b>, e, de outro lado, a "
        f"<b>{dados['fornecedor_nome']}</b>, inscrita no CNPJ nº <b>{dados['fornecedor_cnpj']}</b>, com sede na "
        f"{dados['fornecedor_endereco']}, nº {dados['fornecedor_num']}, Bairro {dados['fornecedor_bairro']}, "
        f"{dados['fornecedor_cidade']}/{dados['fornecedor_uf']}, CEP {dados['fornecedor_cep']}, na qualidade de Fornecedor Registrado, "
        f"neste ato representada por seu representante legal, <b>{dados['fornecedor_representante']}</b>, "
        f"CPF nº {dados['fornecedor_cpf']}. Tendo em vista o Processo Administrativo nº <b>{dados['processo_adm']}</b> e a "
        f"<b>Lei nº 14.133, de 1º de abril de 2021</b>, resolvem celebrar o presente Termo de Contrato, <b>{decorrente_texto_pdf}</b>, "
        f"mediante as cláusulas e condições a seguir enunciadas."
    )
    story.append(Paragraph(preambulo_texto, style_preamble))

    # 3. Cláusulas Padronizadas AGU
    clausulas = gerar_clausulas_agu(dados)

    for cl in clausulas:
        story.append(Paragraph(cl["numero"], style_clause_title))
        
        is_clausula_objeto = "OBJETO" in cl["numero"].upper()
        is_clausula_dotacao = "DOTAÇÃO" in cl["numero"].upper() or "DOTACAO" in cl["numero"].upper()

        # Inserir tabela de itens após 1.2 na Cláusula Primeira
        if is_clausula_objeto:
            itens_list = dados.get("itens", [])
            if not itens_list:
                itens_list = [{
                    "item": 1,
                    "descricao": dados.get("objeto", "Item contratado"),
                    "quantidade": "1",
                    "unidade": "Unid.",
                    "valor_unitario": dados.get("valor_total_formatado", "R$ 0,00"),
                    "valor_total": dados.get("valor_total_formatado", "R$ 0,00"),
                    "_vt_float": 0.0
                }]

            tot_float = sum(float(it.get("_vt_float", 0.0) or 0.0) for it in itens_list)
            if tot_float > 0:
                tot_display = f"R$ {tot_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                tot_display = dados.get("valor_total_formatado", "R$ 0,00")

            for idx_p, p_txt in enumerate(cl["conteudo"]):
                story.append(Paragraph(p_txt, style_clause_body))
                if p_txt.startswith("1.2.") or (len(cl["conteudo"]) == 2 and idx_p == 0):
                    story.append(Spacer(1, 3))
                    
                    tab_din = dados.get("tabela_dinamica")
                    if not tab_din and hasattr(dados.get("itens"), "tabela_dinamica"):
                        tab_din = dados["itens"].tabela_dinamica

                    if tab_din and tab_din.get("headers") and tab_din.get("rows"):
                        dyn_headers = tab_din["headers"]
                        dyn_rows = tab_din["rows"]
                        dyn_alignments = tab_din.get("col_alignments", ["LEFT"] * len(dyn_headers))
                        col_widths = tab_din.get("col_widths") or tab_din.get("col_widths_pt")
                        tot_col = tab_din.get("total_col_idx")
                        tot_val_txt = tab_din.get("total_formatado") or tot_display
                        
                        table_itens_data = [
                            [Paragraph(str(h), style_tbl_th) for h in dyn_headers]
                        ]
                        for r_vals in dyn_rows:
                            row_cells = []
                            for ci in range(len(dyn_headers)):
                                val_str = str(r_vals[ci]) if ci < len(r_vals) else ""
                                align_code = dyn_alignments[ci] if ci < len(dyn_alignments) else "LEFT"
                                style_cell = style_tbl_td_c if align_code == "CENTER" else (
                                    style_tbl_td_r if align_code == "RIGHT" else style_tbl_td
                                )
                                row_cells.append(Paragraph(val_str, style_cell))
                            table_itens_data.append(row_cells)
                            
                        has_tot = tot_col is not None and tot_col < len(dyn_headers)
                        if has_tot:
                            footer_row = ["" for _ in range(len(dyn_headers))]
                            footer_row[0] = Paragraph("<b>TOTAL</b>", ParagraphStyle('TblTotL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=2))
                            footer_row[tot_col] = Paragraph(f"<b>{tot_val_txt}</b>", ParagraphStyle('TblTotV', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=2))
                            table_itens_data.append(footer_row)
                            
                        t_itens = Table(table_itens_data, colWidths=col_widths, repeatRows=1)
                        t_styles = [
                            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E2EBF4")),
                            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                            ('GRID', (0, 0), (-1, -2 if has_tot else -1), 0.5, colors.HexColor("#cbd5e1")),
                            ('TOPPADDING', (0, 0), (-1, -1), 3),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                            ('LEFTPADDING', (0, 0), (-1, -1), 4),
                            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                        ]
                        if has_tot:
                            if tot_col > 0:
                                t_styles.append(('SPAN', (0, -1), (tot_col - 1, -1)))
                            t_styles.extend([
                                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#f1f5f9")),
                                ('LINEABOVE', (0, -1), (-1, -1), 1.0, colors.HexColor("#003366")),
                                ('LINEBELOW', (0, -1), (-1, -1), 1.0, colors.HexColor("#003366")),
                            ])
                        t_itens.setStyle(TableStyle(t_styles))
                        story.append(t_itens)
                        story.append(Spacer(1, 4))
                    else:
                        # Fallback padrão
                        col_w_item = 28
                        col_w_desc = 247
                        col_w_qtd = 55
                        col_w_un = 45
                        col_w_vu = 65
                        col_w_vt = 65.27
                        
                        table_itens_data = [
                            [
                                Paragraph("ITEM", style_tbl_th),
                                Paragraph("DESCRIÇÃO", style_tbl_th),
                                Paragraph("QUANTIDADE", style_tbl_th),
                                Paragraph("UNID.", style_tbl_th),
                                Paragraph("VALOR UNITÁRIO", style_tbl_th),
                                Paragraph("VALOR TOTAL", style_tbl_th)
                            ]
                        ]
                        
                        for row_it in itens_list:
                            table_itens_data.append([
                                Paragraph(str(row_it.get("item", "1")), style_tbl_td_c),
                                Paragraph(str(row_it.get("descricao", dados["objeto"])), style_tbl_td),
                                Paragraph(str(row_it.get("quantidade", "1")), style_tbl_td_c),
                                Paragraph(str(row_it.get("unidade", "Unid.")), style_tbl_td_c),
                                Paragraph(str(row_it.get("valor_unitario", dados["valor_total_formatado"])), style_tbl_td_r),
                                Paragraph(str(row_it.get("valor_total", dados["valor_total_formatado"])), style_tbl_td_r)
                            ])
                            
                        table_itens_data.append([
                            Paragraph("<b>TOTAL</b>", ParagraphStyle('TblTotL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=2)),
                            "", "", "", "",
                            Paragraph(f"<b>{tot_display}</b>", ParagraphStyle('TblTotV', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=2))
                        ])
                        
                        t_itens = Table(table_itens_data, colWidths=[col_w_item, col_w_desc, col_w_qtd, col_w_un, col_w_vu, col_w_vt], repeatRows=1)
                        t_itens.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E2EBF4")),
                            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#224061"))
                            ,('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#224061")),
                            ('SPAN', (0, -1), (4, -1)),
                            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#f1f5f9")),
                            ('LINEABOVE', (0, -1), (-1, -1), 1.0, colors.HexColor("#003366")),
                            ('LINEBELOW', (0, -1), (-1, -1), 1.0, colors.HexColor("#003366")),
                            ('TOPPADDING', (0, 0), (-1, -1), 2.5),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
                            ('LEFTPADDING', (0, 0), (-1, -1), 3),
                            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                        ]))
                        story.append(t_itens)
                        story.append(Spacer(1, 4))
        # Inserir tabela de dotação orçamentária na Cláusula Quarta
        elif is_clausula_dotacao:
            for idx_d, d_txt in enumerate(cl["conteudo"]):
                story.append(Paragraph(d_txt, style_clause_body))
                if idx_d == 0:
                    story.append(Spacer(1, 3))
                    dot_items = [
                        ("Exercício", str(dados.get("dotacao_exercicio") or "2026")),
                        ("Poder", str(dados.get("dotacao_poder") or "Poder Executivo 02.00")),
                        ("Órgão", str(dados.get("dotacao_orgao") or f"Fundo Municipal / {dados.get('orgao_nome', '')} 02.15.00")),
                        ("Unidade Orçamentária/atividade", str(dados.get("dotacao_unidade") or dados.get("dotacao_orcamentaria") or "MANUTENÇÃO E FUNCIONAMENTO 12.361.0402.2022.0000")),
                        ("Natureza da despesa", str(dados.get("dotacao_natureza") or ("Material de consumo 33.90.30.00" if "COMPRA" in str(dados.get("tipo_contratacao", "")).upper() else "Outros Serviços de Terceiros 33.90.39.00")))
                    ]
                    if dados.get("dotacao_fonte") or dados.get("fonte_recursos") or dados.get("fonte"):
                        dot_items.append(("Fonte de recursos", str(dados.get("dotacao_fonte") or dados.get("fonte_recursos") or dados.get("fonte"))))
                    if dados.get("dotacao_nota_empenho") or dados.get("nota_empenho") or dados.get("empenho"):
                        dot_items.append(("Nota de empenho", str(dados.get("dotacao_nota_empenho") or dados.get("nota_empenho") or dados.get("empenho"))))

                    t_dot_data = [
                        [Paragraph(f"<b>{k}</b>", style_tbl_td), Paragraph(v, style_tbl_td)]
                        for k, v in dot_items
                    ]
                    t_dot = Table(t_dot_data, colWidths=[160, page_content_w - 160])
                    t_dot.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor("#f8fafc")),
                        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#224061")),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('TOPPADDING', (0, 0), (-1, -1), 3),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                        ('LEFTPADDING', (0, 0), (-1, -1), 6),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                    ]))
                    story.append(t_dot)
                    story.append(Spacer(1, 4))
        else:
            for p_txt in cl["conteudo"]:
                if p_txt.startswith("    ") or p_txt.startswith("a)") or p_txt.startswith("b)") or p_txt.startswith("i.") or p_txt.startswith("ii."):
                    story.append(Paragraph(p_txt, style_bullet))
                else:
                    story.append(Paragraph(p_txt, style_clause_body))

    # 4. Fecho e Data
    story.append(Spacer(1, 8))
    data_fecho_texto = f"Ribeirãozinho do Maranhão – MA, {dados['data_assinatura_extenso']}."
    style_date = ParagraphStyle('DateFinal', parent=styles['Normal'], fontName='Times-Roman', fontSize=10, alignment=2, spaceAfter=14)
    story.append(Paragraph(data_fecho_texto, style_date))

    # 5. Assinaturas em Tabela 2 Colunas
    bloco_ass = []
    bloco_ass.append(Spacer(1, 6))

    t_ass_1_data = [
        [
            Paragraph(f"____________________________________________<br/><b>{dados['secretario']}</b><br/>{dados['cargo_secretario']}<br/>{dados['secretario_portaria']}<br/><b>Contratante</b>", style_sig_sub),
            Paragraph(f"____________________________________________<br/><b>{dados['fornecedor_nome']}</b><br/>{dados['fornecedor_representante']}<br/>{dados['fornecedor_cargo']}<br/><b>Contratada</b>", style_sig_sub)
        ]
    ]
    t_ass_1 = Table(t_ass_1_data, colWidths=[page_content_w / 2.0, page_content_w / 2.0])
    t_ass_1.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 16),
    ]))
    bloco_ass.append(t_ass_1)

    t_test_data = [
        [
            Paragraph(f"____________________________________________<br/><b>{dados['fiscal']}</b><br/>Fiscal do Contrato - Portaria Municipal", style_sig_sub),
            Paragraph("<b>TESTEMUNHAS:</b><br/><br/>1- ________________________________ CPF:<br/>2- ________________________________ CPF:", style_sig_sub)
        ]
    ]
    t_test = Table(t_test_data, colWidths=[page_content_w / 2.0, page_content_w / 2.0])
    t_test.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    bloco_ass.append(t_test)

    story.append(KeepTogether(bloco_ass))

    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer

def gerar_comprovante_protocolo_pdf(prot: Dict[str, Any]) -> io.BytesIO:
    """
    Gera o Comprovante Oficial de Protocolo em PDF (1 ou 2 vias) com Brasão Municipal,
    tabela de metadados, campos para carimbo/assinatura e linha tracejada para corte.
    """
    buffer = io.BytesIO()
    vias = prot.get("qtd_vias", 1)
    
    # Se forem 2 vias na mesma folha, margens menores para caber as duas em 1 página A4 (29.7cm)
    margin = 1.0 * cm if vias == 2 else 1.4 * cm
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin
    )
    
    styles = getSampleStyleSheet()
    
    style_header_title = ParagraphStyle(
        'ProtHeaderTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9.5 if vias == 2 else 11,
        leading=11.5 if vias == 2 else 14,
        textColor=colors.HexColor("#0369a1"),
        alignment=0
    )
    style_header_sub = ParagraphStyle(
        'ProtHeaderSub',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7 if vias == 2 else 8.5,
        leading=8.5 if vias == 2 else 10,
        textColor=colors.HexColor("#059669"),
        alignment=0
    )
    style_meta_via = ParagraphStyle(
        'ProtMetaVia',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7.5 if vias == 2 else 9,
        leading=9.5,
        textColor=colors.HexColor("#64748b"),
        alignment=2
    )
    style_meta_num = ParagraphStyle(
        'ProtMetaNum',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11 if vias == 2 else 14,
        leading=13 if vias == 2 else 16,
        textColor=colors.HexColor("#0284c7"),
        alignment=2
    )
    style_tbl_label = ParagraphStyle(
        'ProtTblLabel',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7 if vias == 2 else 8.5,
        leading=9 if vias == 2 else 10.5,
        textColor=colors.HexColor("#1e293b")
    )
    style_tbl_val = ParagraphStyle(
        'ProtTblVal',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7 if vias == 2 else 8.5,
        leading=9 if vias == 2 else 10.5,
        textColor=colors.HexColor("#334155")
    )
    style_sig = ParagraphStyle(
        'ProtSig',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7 if vias == 2 else 8.5,
        leading=9 if vias == 2 else 10.5,
        alignment=1,
        textColor=colors.HexColor("#334155")
    )
    style_footer = ParagraphStyle(
        'ProtFooter',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=6 if vias == 2 else 7.5,
        leading=7.5 if vias == 2 else 9,
        alignment=1,
        textColor=colors.HexColor("#94a3b8")
    )

    story = []
    agora_str = datetime.now().strftime("%d/%m/%Y às %H:%M")
    page_w = A4[0] - 2 * margin

    def criar_bloco_via(via_label: str):
        bloco = []
        
        # 1. Cabeçalho (Brasão + Título + Número)
        img_element = ""
        if os.path.exists(BRASAO_PATH):
            img_size = 1.2 * cm if vias == 2 else 1.6 * cm
            img_element = Image(BRASAO_PATH, width=img_size, height=img_size)
            
        header_text = [
            Paragraph("PREFEITURA MUNICIPAL DE<br/>RIBEIRÃOZINHO DO MARANHÃO - MA", style_header_title),
            Spacer(1, 0.8 * mm),
            Paragraph("ESTADO DO MARANHÃO", style_header_sub)
        ]
        
        header_num = [
            Paragraph(via_label, style_meta_via),
            Paragraph(f"PROTOCOLO Nº {prot.get('numero', '—')}", style_meta_num)
        ]
        
        t_header = Table(
            [[img_element, header_text, header_num]],
            colWidths=[1.5 * cm if vias == 2 else 1.9 * cm, page_w - 7.5 * cm, 6.0 * cm if vias == 2 else 5.6 * cm]
        )
        t_header.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (0,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 1),
            ('TOPPADDING', (0,0), (-1,-1), 1),
        ]))
        bloco.append(t_header)
        bloco.append(Spacer(1, 1.5 * mm))
        bloco.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#0284c7"), spaceAfter=3 if vias == 2 else 6))
        
        # 2. Tabela de Dados
        t_data = [
            [
                Paragraph("<b>Data de Recebimento:</b>", style_tbl_label),
                Paragraph(str(prot.get("data_rec", "—")), style_tbl_val),
                Paragraph("<b>Secretaria Origem:</b>", style_tbl_label),
                Paragraph(str(prot.get("orgao", "—")), style_tbl_val)
            ],
            [
                Paragraph("<b>Solicitante:</b>", style_tbl_label),
                Paragraph(str(prot.get("solicitante", "—")), style_tbl_val),
                "", ""
            ],
            [
                Paragraph("<b>Assunto / Demanda:</b>", style_tbl_label),
                Paragraph(str(prot.get("assunto", "—")), style_tbl_val),
                "", ""
            ],
            [
                Paragraph("<b>Contrato Vinculado:</b>", style_tbl_label),
                Paragraph(str(prot.get("contrato", "—")), style_tbl_val),
                Paragraph("<b>Fornecedor:</b>", style_tbl_label),
                Paragraph(str(prot.get("fornecedor", "—")), style_tbl_val)
            ],
            [
                Paragraph("<b>Secretário(a):</b>", style_tbl_label),
                Paragraph(str(prot.get("secretario", "—")), style_tbl_val),
                Paragraph("<b>Registrado por:</b>", style_tbl_label),
                Paragraph(str(prot.get("usuario", "—")), style_tbl_val)
            ],
            [
                Paragraph("<b>Observações:</b>", style_tbl_label),
                Paragraph(str(prot.get("observacoes", "—")), style_tbl_val),
                "", ""
            ]
        ]
        
        col_w1 = 3.4 * cm if vias == 2 else 3.8 * cm
        col_w2 = (page_w - (col_w1 * 2)) / 2
        col_w3 = col_w1
        col_w4 = col_w2
        
        t_tabela = Table(t_data, colWidths=[col_w1, col_w2, col_w3, col_w4])
        t_tabela.setStyle(TableStyle([
            ('SPAN', (1, 1), (3, 1)),
            ('SPAN', (1, 2), (3, 2)),
            ('SPAN', (1, 5), (3, 5)),
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor("#f8fafc")),
            ('BACKGROUND', (2, 0), (2, 0), colors.HexColor("#f8fafc")),
            ('BACKGROUND', (2, 3), (2, 3), colors.HexColor("#f8fafc")),
            ('BACKGROUND', (2, 4), (2, 4), colors.HexColor("#f8fafc")),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#224061")),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 2.5 if vias == 2 else 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5 if vias == 2 else 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ]))
        bloco.append(t_tabela)
        
        # 3. Assinaturas
        bloco.append(Spacer(1, 3 * mm if vias == 2 else 7 * mm))
        sig_data = [
            [
                Paragraph("____________________________________________________<br/><b>Servidor(a) Responsável pelo Recebimento</b><br/>Carimbo / Matrícula", style_sig),
                Paragraph("____________________________________________________<br/><b>Entregue por (Assinatura do Recebedor)</b><br/>Data: ____/____/________", style_sig)
            ]
        ]
        sig_table = Table(sig_data, colWidths=[page_w / 2, page_w / 2])
        sig_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 1),
        ]))
        bloco.append(sig_table)
        
        # 4. Rodapé
        bloco.append(Spacer(1, 1.5 * mm if vias == 2 else 3 * mm))
        bloco.append(Paragraph(
            f"Documento gerado em {agora_str} pelo Sistema Integrado de Gestão — Prefeitura Municipal de Ribeirãozinho do Maranhão - MA",
            style_footer
        ))
        
        # Borda externa do comprovante
        t_box = Table([[bloco]], colWidths=[page_w])
        t_box.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#0284c7")),
            ('BACKGROUND', (0, 0), (-1, -1), colors.white),
            ('TOPPADDING', (0, 0), (-1, -1), 5 if vias == 2 else 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5 if vias == 2 else 9),
            ('LEFTPADDING', (0, 0), (-1, -1), 7 if vias == 2 else 11),
            ('RIGHTPADDING', (0, 0), (-1, -1), 7 if vias == 2 else 11),
        ]))
        
        return t_box

    # Monta 1 ou 2 vias
    story.append(criar_bloco_via("1ª VIA — PROTOCOLO"))
    
    if vias == 2:
        story.append(Spacer(1, 2.5 * mm))
        story.append(Paragraph("✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ✂", style_footer))
        story.append(Spacer(1, 2.5 * mm))
        story.append(criar_bloco_via("2ª VIA — RECEBEDOR"))
        
    doc.build(story)
    buffer.seek(0)
    return buffer


def gerar_clausulas_aditivo(dados: Dict[str, Any]) -> list:
    """
    Gera o rol oficial de cláusulas padronizadas de Termos Aditivos da AGU (CGU/AGU)
    conforme as Leis Federais nº 14.133/2021 e nº 8.666/1993.
    """
    tipo = str(dados.get("tipo_aditivo", "PRORROGACAO")).upper()
    regime = str(dados.get("regime_legal", "LEI_14133_2021"))
    is_14133 = "14133" in regime or "14.133" in regime
    
    lei_ref = "Lei Federal nº 14.133, de 1º de abril de 2021" if is_14133 else "Lei Federal nº 8.666, de 21 de junho de 1993"
    art_prorrogacao = "art. 107 da Lei nº 14.133/2021" if is_14133 else "art. 57, inciso II, da Lei nº 8.666/1993"
    art_acrescimo = "art. 124, inciso I, alínea 'b' e art. 125 da Lei nº 14.133/2021" if is_14133 else "art. 65, inciso I, alínea 'b' e § 1º da Lei nº 8.666/1993"
    art_reequilibrio = "art. 124, inciso II, alínea 'd' da Lei nº 14.133/2021" if is_14133 else "art. 65, inciso II, alínea 'd' da Lei nº 8.666/1993"
    
    val_orig = float(dados.get("valor_original", 0.0) or 0.0)
    val_orig_str = formatar_moeda_br(val_orig)
    val_orig_ext = valor_por_extenso(val_orig)
    
    val_adit = float(dados.get("valor_aditado", 0.0) or 0.0)
    val_adit_str = formatar_moeda_br(val_adit)
    val_adit_ext = valor_por_extenso(val_adit)
    
    novo_val = float(dados.get("novo_valor_total", val_orig) or val_orig)
    novo_val_str = formatar_moeda_br(novo_val)
    novo_val_ext = valor_por_extenso(novo_val)
    
    perc = float(dados.get("percentual_aditado", 0.0) or 0.0)
    meses = int(dados.get("prazo_meses", 12) or 12)
    dt_inicio = str(dados.get("data_inicio_aditivo", ""))
    dt_fim = str(dados.get("nova_data_vencimento", ""))
    
    clausulas = []
    
    # CLÁUSULA PRIMEIRA - DO OBJETO
    obj_paragraphs = []
    if tipo == "PRORROGACAO":
        obj_paragraphs.append(
            f"1.1. O presente Termo Aditivo tem por objeto a <b>PRORROGAÇÃO DO PRAZO DE VIGÊNCIA</b> do Contrato originário por mais "
            f"<b>{meses} meses</b>, com termo inicial em <b>{dt_inicio}</b> "
            f"e termo final em <b>{dt_fim}</b>, com fundamento no <b>{art_prorrogacao}</b>, tendo em vista o manifesto interesse público da Administração "
            f"e a manutenção da vantajosidade econômica atestada nos autos do processo."
        )
        if dados.get("justificativa"):
            obj_paragraphs.append(f"1.2. <b>Justificativa Técnica da Prorrogação:</b> {dados['justificativa']}")
    elif tipo == "ACRESCIMO":
        obj_paragraphs.append(
            f"1.1. O presente Termo Aditivo tem por objeto o <b>ACRÉSCIMO QUANTITATIVO E DE VALOR</b> correspondente a "
            f"<b>{perc:.2f}%</b> do valor inicial atualizado do Contrato, importando no acréscimo de <b>{val_adit_str} ({val_adit_ext})</b>, "
            f"com fundamento no <b>{art_acrescimo}</b>, mantidas as mesmas condições de preços unitários e exigências contratuais."
        )
        if dados.get("justificativa"):
            obj_paragraphs.append(f"1.2. <b>Justificativa Técnica da Alteração:</b> {dados['justificativa']}")
    elif tipo == "SUPRESSAO":
        obj_paragraphs.append(
            f"1.1. O presente Termo Aditivo tem por objeto a <b>SUPRESSÃO QUANTITATIVA E DE VALOR</b> correspondente a "
            f"<b>{perc:.2f}%</b> do valor inicial atualizado do Contrato, no montante de <b>{val_adit_str} ({val_adit_ext})</b>, "
            f"com fundamento no <b>{art_acrescimo}</b>, em razão da adequação da demanda da Secretaria Contratante."
        )
        if dados.get("justificativa"):
            obj_paragraphs.append(f"1.2. <b>Justificativa da Supressão:</b> {dados['justificativa']}")
    elif tipo == "MISTO":
        obj_paragraphs.append(
            f"1.1. O presente Termo Aditivo tem por objeto conjunto: "
            f"<br/>a) A <b>PRORROGAÇÃO DO PRAZO DE VIGÊNCIA</b> por mais <b>{meses} meses</b>, fixando o novo termo final em <b>{dt_fim}</b>, com amparo no {art_prorrogacao}; e"
            f"<br/>b) A <b>ALTERAÇÃO DE VALOR</b> no montante de <b>{val_adit_str} ({val_adit_ext})</b>, equivalente a <b>{perc:.2f}%</b> do valor contratado, com fulcro no {art_acrescimo}."
        )
        if dados.get("justificativa"):
            obj_paragraphs.append(f"1.2. <b>Justificativa Técnica:</b> {dados['justificativa']}")
    elif tipo == "REEQUILIBRIO":
        obj_paragraphs.append(
            f"1.1. O presente Termo Aditivo tem por objeto a <b>RECOMPOSIÇÃO DO EQUILÍBRIO ECONÔMICO-FINANCEIRO / REAJUSTAMENTO</b> dos preços contratuais, "
            f"nos termos do <b>{art_reequilibrio}</b>, em decorrência da alteração superveniente dos custos de execução devidamente comprovada nos autos, "
            f"resultando na variação de <b>{val_adit_str} ({val_adit_ext})</b>."
        )
        if dados.get("justificativa"):
            obj_paragraphs.append(f"1.2. <b>Justificativa Técnica:</b> {dados['justificativa']}")
    else:
        obj_paragraphs.append(
            f"1.1. O presente Termo Aditivo tem por objeto a alteração qualitativa e adequação das condições do Contrato: "
            f"{dados.get('objeto_aditivo', 'Alteração consensual nos termos da legislação regente')}."
        )
        if dados.get("justificativa"):
            obj_paragraphs.append(f"1.2. <b>Justificativa da Modificação:</b> {dados['justificativa']}")

    clausulas.append({
        "numero": "CLÁUSULA PRIMEIRA – DO OBJETO",
        "conteudo": obj_paragraphs
    })

    # CLÁUSULA SEGUNDA - DO VALOR
    if tipo in ["ACRESCIMO", "MISTO", "REEQUILIBRIO"]:
        clausulas.append({
            "numero": "CLÁUSULA SEGUNDA – DO VALOR CONSOLIDADO",
            "conteudo": [
                f"2.1. Em virtude do presente aditamento, o valor global do Contrato fica acrescido de <b>{val_adit_str} ({val_adit_ext})</b>, "
                f"passando o valor total atualizado do Contrato de {val_orig_str} ({val_orig_ext}) para o montante global de "
                f"<b>{novo_val_str} ({novo_val_ext})</b>.",
                "2.2. A CONTRATADA declara expressamente que os preços unitários e globais praticados encontram-se compatíveis com os preços de mercado vigentes e que aceita os novos quantitativos sem qualquer ressalva."
            ]
        })
    elif tipo == "SUPRESSAO":
        clausulas.append({
            "numero": "CLÁUSULA SEGUNDA – DO VALOR SUPRIMIDO",
            "conteudo": [
                f"2.1. Em decorrência da supressão pactuada, o valor contratual fica reduzido em <b>{val_adit_str} ({val_adit_ext})</b>, "
                f"restando o novo valor total do Contrato fixado em <b>{novo_val_str} ({novo_val_ext})</b>.",
                "2.2. A CONTRATADA declara concordar expressamente com a supressão quantitativa ora formalizada."
            ]
        })
    elif tipo == "PRORROGACAO":
        clausulas.append({
            "numero": "CLÁUSULA SEGUNDA – DO VALOR PARA O NOVO PERÍODO",
            "conteudo": [
                f"2.1. O valor estimado para a execução contratual durante o período de prorrogação ora aditado é de "
                f"<b>{val_adit_str if val_adit > 0 else val_orig_str} ({val_adit_ext if val_adit > 0 else val_orig_ext})</b>, "
                "mantidos integralmente os preços unitários inicialmente acordados, ressalvado o direito a futuro reajuste nos termos contratuais.",
                f"2.2. O valor total consolidado do Contrato, somados os períodos executados e o ora prorrogado, perfaz a quantia de <b>{novo_val_str} ({novo_val_ext})</b>."
            ]
        })
    else:
        clausulas.append({
            "numero": "CLÁUSULA SEGUNDA – DO PREÇO",
            "conteudo": [
                f"2.1. Permanece em vigor o valor pactuado de <b>{novo_val_str} ({novo_val_ext})</b>, inalteradas as condições de medição e faturamento."
            ]
        })

    # CLÁUSULA TERCEIRA - DA VIGÊNCIA (se aplicável)
    if tipo in ["PRORROGACAO", "MISTO"]:
        clausulas.append({
            "numero": "CLÁUSULA TERCEIRA – DA VIGÊNCIA E EFICÁCIA",
            "conteudo": [
                f"3.1. A vigência do Contrato fica estendida até a data de <b>{dt_fim}</b>, aplicando-se o método de contagem estabelecido na "
                f"Orientação Normativa AGU nº 38/2011 e Parecer nº 85/2019/DECOR/CGU/AGU.",
                "3.2. A eficácia das disposições deste instrumento produzirá efeitos a partir da data de sua assinatura, com eficácia plena após regular publicação de seu extrato."
            ]
        })

    # CLÁUSULA QUARTA - DA DOTAÇÃO ORÇAMENTÁRIA
    dotacao = dados.get("dotacao_orcamentaria") or "Dotação orçamentária própria consignada na Lei Orçamentária Anual"
    clausulas.append({
        "numero": "CLÁUSULA QUARTA – DA DOTAÇÃO ORÇAMENTÁRIA",
        "conteudo": [
            f"4.1. As despesas decorrentes da execução do presente Termo Aditivo correrão à conta dos recursos orçamentários específicos da "
            f"<b>{dados.get('orgao_nome', 'Secretaria Contratante')}</b>, sob a seguinte classificação programática e dotação:",
            f"<b>{dotacao}</b>",
            "4.2. No caso de prorrogação que alcance exercícios financeiros subsequentes, a CONTRATANTE consignará as dotações necessárias nas respectivas Leis Orçamentárias Anuais."
        ]
    })

    # CLÁUSULA QUINTA - DA GARANTIA CONTRATUAL
    garantia = dados.get("garantia_execucao")
    if garantia:
        clausulas.append({
            "numero": "CLÁUSULA QUINTA – DA COMPLEMENTAÇÃO DA GARANTIA",
            "conteudo": [
                f"5.1. A CONTRATADA obriga-se a complementar/renovar a garantia de execução prestada no contrato originário: {garantia}, "
                "no prazo improrrogável de até 10 (dez) dias úteis contados da assinatura deste Termo Aditivo, sob pena de rescisão contratual e aplicação das penalidades cabíveis."
            ]
        })

    # CLÁUSULA SEXTA - DA RATIFICAÇÃO
    clausulas.append({
        "numero": "CLÁUSULA SEXTA – DA RATIFICAÇÃO DAS DEMAIS CLÁUSULAS",
        "conteudo": [
            "6.1. Ficam expressamente <b>RATIFICADAS E CONFIRMADAS</b> todas as demais cláusulas, condições e estipulações estabelecidas no "
            f"Contrato originário nº <b>{dados.get('numero_contrato', '')}</b> e em eventuais termos aditivos anteriores que não tenham sido "
            "expressamente alteradas ou revogadas pelo presente instrumento, permanecendo as mesmas em pleno vigor e efeito."
        ]
    })

    # CLÁUSULA SÉTIMA - DA PUBLICAÇÃO
    clausulas.append({
        "numero": "CLÁUSULA SÉTIMA – DA PUBLICAÇÃO E TRANSPARÊNCIA",
        "conteudo": [
            f"7.1. Incumbirá ao MUNICÍPIO DE RIBEIRÃOZINHO DO MARANHÃO - MA providenciar a publicação do extrato do presente Termo Aditivo "
            f"no Diário Oficial do Município e sua divulgação no Portal Nacional de Contratações Públicas (PNCP), nos termos e prazos "
            f"previstos na <b>{lei_ref}</b>."
        ]
    })

    return clausulas


def gerar_aditivo_docx(dados: Dict[str, Any]) -> io.BytesIO:
    """
    Gera a Minuta Oficial de Termo Aditivo em formato Word (.docx)
    segundo os modelos padronizados da AGU e com cabeçalho/rodapé institucionais.
    """
    dados = dict(dados)
    dados["fornecedor_cnpj"] = formatar_cnpj_br(dados.get("fornecedor_cnpj"))
    dados["orgao_cnpj"] = formatar_cnpj_br(dados.get("orgao_cnpj"))
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1.3)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

        if os.path.exists(HEADER_OFICIAL_PATH):
            header = section.header
            header.is_linked_to_previous = False
            p_head = header.paragraphs[0]
            p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_head = p_head.add_run()
            r_head.add_picture(HEADER_OFICIAL_PATH, width=Inches(6.9))

        if os.path.exists(FOOTER_OFICIAL_PATH):
            footer = section.footer
            footer.is_linked_to_previous = False
            p_foot = footer.paragraphs[0]
            p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_foot = p_foot.add_run()
            r_foot.add_picture(FOOTER_OFICIAL_PATH, width=Inches(6.9))
        
    p_mun = doc.add_paragraph()
    p_mun.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_mun = p_mun.add_run("PREFEITURA MUNICIPAL DE RIBEIRÃOZINHO DO MARANHÃO - MA\n")
    r_mun.bold = True
    r_mun.font.name = "Arial"
    r_mun.font.size = Pt(11)
    r_mun.font.color.rgb = RGBColor(3, 105, 161)
    
    r_org = p_mun.add_run(f"{dados.get('orgao_nome', 'SECRETARIA MUNICIPAL')}\n")
    r_org.bold = True
    r_org.font.name = "Arial"
    r_org.font.size = Pt(10)
    
    r_est = p_mun.add_run("ESTADO DO MARANHÃO")
    r_est.bold = True
    r_est.font.name = "Arial"
    r_est.font.size = Pt(9)
    r_est.font.color.rgb = RGBColor(5, 150, 105)
    
    # Título do Termo Aditivo
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.space_before = Pt(8)
    p_tit.paragraph_format.space_after = Pt(2)
    r_tit = p_tit.add_run(f"{dados.get('numero_completo_aditivo', 'TERMO ADITIVO')}\n")
    r_tit.bold = True
    r_tit.font.name = "Arial"
    r_tit.font.size = Pt(12)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    regime_str = "LEI Nº 14.133/2021" if "14133" in str(dados.get("regime_legal", "")) else "LEI Nº 8.666/1993"
    r_sub = p_sub.add_run(f"PROCESSO ADMINISTRATIVO Nº {dados.get('processo_adm', '001/2026')} | REGIME LEGAL: {regime_str}")
    r_sub.italic = True
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(9)
    r_sub.font.color.rgb = RGBColor(71, 85, 105)
    
    # Preâmbulo
    p_pre = doc.add_paragraph()
    p_pre.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_pre.paragraph_format.first_line_indent = Inches(0.5)
    p_pre.paragraph_format.line_spacing = 1.15
    p_pre.paragraph_format.space_after = Pt(10)
    
    lei_completa = "Lei Federal nº 14.133, de 1º de abril de 2021" if "14133" in str(dados.get("regime_legal", "")) else "Lei Federal nº 8.666, de 21 de junho de 1993"
    preambulo_texto = (
        f"O MUNICÍPIO DE RIBEIRÃOZINHO DO MARANHÃO - ESTADO DO MARANHÃO, pessoa jurídica de direito público interno, "
        f"inscrito no CNPJ sob o nº {dados.get('orgao_cnpj', '01.612.834/0001-86')}, por intermédio da(o) {dados.get('orgao_nome')}, "
        f"com sede na {dados.get('orgao_endereco', 'Rua Principal, s/n')}, Bairro {dados.get('orgao_bairro', 'Centro')}, "
        f"CEP {dados.get('orgao_cep', '65928-000')}, na cidade de {dados.get('orgao_cidade', 'Ribeirãozinho do Maranhão')} - {dados.get('orgao_uf', 'MA')}, "
        f"neste ato representada(o) pelo(a) {dados.get('cargo_secretario', 'Secretário(a) Municipal Titular')}, "
        f"Senhor(a) {dados.get('secretario', 'SECRETÁRIO MUNICIPAL')}, doravante denominado simplesmente CONTRATANTE, "
        f"e, de outro lado, a empresa {dados.get('fornecedor_nome')}, inscrita no CNPJ/MF sob o nº {dados.get('fornecedor_cnpj')}, "
        f"com sede na {dados.get('fornecedor_endereco', 'Sede Comercial')}, representada por seu {dados.get('fornecedor_cargo', 'Representante Legal')}, "
        f"Senhor(a) {dados.get('fornecedor_representante', 'REPRESENTANTE LEGAL')} doravante denominada CONTRATADA, "
        f"tendo em vista o constante dos autos do Processo Administrativo nº {dados.get('processo_adm')} e com amparo na {lei_completa}, "
        f"têm entre si justo e acordado o presente {dados.get('numero_completo_aditivo', 'TERMO ADITIVO')}, mediante as cláusulas seguintes:"
    )
    r_pre = p_pre.add_run(preambulo_texto)
    r_pre.font.name = "Times New Roman"
    r_pre.font.size = Pt(10.5)
    
    clausulas = gerar_clausulas_aditivo(dados)
    for cl in clausulas:
        p_c_tit = doc.add_paragraph()
        p_c_tit.paragraph_format.space_before = Pt(8)
        p_c_tit.paragraph_format.space_after = Pt(2)
        r_c_tit = p_c_tit.add_run(cl["numero"])
        r_c_tit.bold = True
        r_c_tit.font.name = "Arial"
        r_c_tit.font.size = Pt(10)
        
        for item_txt in cl["conteudo"]:
            p_item = doc.add_paragraph()
            p_item.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_item.paragraph_format.first_line_indent = Inches(0.3)
            p_item.paragraph_format.line_spacing = 1.15
            p_item.paragraph_format.space_after = Pt(3)
            texto_limpo = re.sub(r'<[^>]+>', '', item_txt)
            r_item = p_item.add_run(texto_limpo)
            r_item.font.name = "Times New Roman"
            r_item.font.size = Pt(10)
            
    p_fecho = doc.add_paragraph()
    p_fecho.paragraph_format.space_before = Pt(10)
    p_fecho.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_f = p_fecho.add_run(
        "E, por estarem assim justos e acordados, assinam o presente Termo Aditivo em 2 (duas) vias de igual teor e forma, "
        "na presença de 2 (duas) testemunhas instrumentárias, para que produza seus regulares efeitos de direito."
    )
    r_f.font.name = "Times New Roman"
    r_f.font.size = Pt(10.5)
    
    dt_ext = dados.get("data_assinatura_extenso") or formatar_data_extenso(dados.get("data_assinatura") or date.today())
    p_dt = doc.add_paragraph()
    p_dt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_dt.paragraph_format.space_after = Pt(20)
    r_dt = p_dt.add_run(f"Ribeirãozinho do Maranhão - MA, {dt_ext}.")
    r_dt.italic = True
    r_dt.font.name = "Times New Roman"
    r_dt.font.size = Pt(10)
    
    table_ass = doc.add_table(rows=2, cols=2)
    table_ass.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_ass.autofit = False
    
    c00 = table_ass.cell(0, 0)
    c01 = table_ass.cell(0, 1)
    c10 = table_ass.cell(1, 0)
    c11 = table_ass.cell(1, 1)
    
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p00.add_run(f"_____________________________________________\nCONTRATANTE\n{dados.get('secretario')}\n{dados.get('cargo_secretario')}\n{dados.get('orgao_nome')}\n\n")
    
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p01.add_run(f"_____________________________________________\nCONTRATADA\n{dados.get('fornecedor_representante')}\n{dados.get('fornecedor_cargo')}\n{dados.get('fornecedor_nome')}\n\n")
    
    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p10.add_run(f"_____________________________________________\nFISCAL DO CONTRATO\n{dados.get('fiscal', 'Servidor Designado')}\nFiscal Titular")
    
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p11.add_run("TESTEMUNHAS:\n\n1. ___________________________ CPF:\n2. ___________________________ CPF:")
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def gerar_aditivo_pdf(dados: Dict[str, Any]) -> io.BytesIO:
    """
    Gera a Minuta Oficial de Termo Aditivo em PDF de alta resolução com ReportLab,
    utilizando o NumberedCanvas institucional com faixas e dados da Prefeitura.
    """
    dados = dict(dados)
    dados["fornecedor_cnpj"] = formatar_cnpj_br(dados.get("fornecedor_cnpj"))
    dados["orgao_cnpj"] = formatar_cnpj_br(dados.get("orgao_cnpj"))
    buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=45,
        rightMargin=45,
        topMargin=96,
        bottomMargin=70
    )
    
    styles = getSampleStyleSheet()
    page_content_w = 595.27 - 90
    
    style_h_mun = ParagraphStyle('AditHMun', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13, alignment=1, textColor=colors.HexColor("#0369a1"))
    style_h_org = ParagraphStyle('AditHOrg', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, leading=12, alignment=1, textColor=colors.HexColor("#0f172a"))
    style_h_est = ParagraphStyle('AditHEst', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=1, textColor=colors.HexColor("#059669"))
    style_title = ParagraphStyle('AditTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=14, alignment=1, textColor=colors.HexColor("#0284c7"), spaceBefore=6, spaceAfter=2)
    style_sub = ParagraphStyle('AditSub', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=8.5, leading=11, alignment=1, textColor=colors.HexColor("#475569"), spaceAfter=8)
    style_preamble = ParagraphStyle('AditPreamble', parent=styles['Normal'], fontName='Times-Roman', fontSize=9.5, leading=13.5, alignment=4, firstLineIndent=15, spaceAfter=6)
    style_c_title = ParagraphStyle('AditCTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5, leading=13, textColor=colors.HexColor("#0f172a"), spaceBefore=6, spaceAfter=2, keepWithNext=True)
    style_c_body = ParagraphStyle('AditCBody', parent=styles['Normal'], fontName='Times-Roman', fontSize=9, leading=12.5, alignment=4, firstLineIndent=14, spaceAfter=3)
    style_sig_box = ParagraphStyle('AditSigBox', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=1, textColor=colors.HexColor("#0f172a"))

    story = []
    
    story.append(Paragraph("PREFEITURA MUNICIPAL DE RIBEIRÃOZINHO DO MARANHÃO - MA", style_h_mun))
    story.append(Paragraph(str(dados.get("orgao_nome", "SECRETARIA MUNICIPAL")), style_h_org))
    story.append(Paragraph("ESTADO DO MARANHÃO", style_h_est))
    story.append(Spacer(1, 2 * mm))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#0284c7"), spaceAfter=6))
    
    story.append(Paragraph(str(dados.get("numero_completo_aditivo", "TERMO ADITIVO")), style_title))
    regime_str = "LEI Nº 14.133/2021" if "14133" in str(dados.get("regime_legal", "")) else "LEI Nº 8.666/1993"
    story.append(Paragraph(f"PROCESSO ADMINISTRATIVO Nº {dados.get('processo_adm')} | REGIME LEGAL: {regime_str}", style_sub))
    
    lei_completa = "Lei Federal nº 14.133, de 1º de abril de 2021" if "14133" in str(dados.get("regime_legal", "")) else "Lei Federal nº 8.666, de 21 de junho de 1993"
    preambulo_texto = (
        f"O <b>MUNICÍPIO DE RIBEIRÃOZINHO DO MARANHÃO - ESTADO DO MARANHÃO</b>, pessoa jurídica de direito público interno, "
        f"inscrito no CNPJ sob o nº <b>{dados.get('orgao_cnpj', '01.612.834/0001-86')}</b>, por intermédio da(o) <b>{dados.get('orgao_nome')}</b>, "
        f"com sede na {dados.get('orgao_endereco', 'Rua Principal, s/n')}, Bairro {dados.get('orgao_bairro', 'Centro')}, "
        f"CEP {dados.get('orgao_cep', '65928-000')}, na cidade de {dados.get('orgao_cidade', 'Ribeirãozinho do Maranhão')} - {dados.get('orgao_uf', 'MA')}, "
        f"neste ato representada(o) pelo(a) {dados.get('cargo_secretario', 'Secretário(a) Municipal Titular')}, "
        f"Senhor(a) <b>{dados.get('secretario', 'SECRETÁRIO MUNICIPAL')}</b>, doravante denominado simplesmente CONTRATANTE, "
        f"e, de outro lado, a empresa <b>{dados.get('fornecedor_nome')}</b>, inscrita no CNPJ/MF sob o nº <b>{dados.get('fornecedor_cnpj')}</b>, "
        f"com sede na {dados.get('fornecedor_endereco', 'Sede Comercial')}, representada por seu {dados.get('fornecedor_cargo', 'Representante Legal')}, "
        f"Senhor(a) <b>{dados.get('fornecedor_representante', 'REPRESENTANTE LEGAL')}</b>, doravante denominada CONTRATADA, "
        f"tendo em vista o que consta no Processo Administrativo nº <b>{dados.get('processo_adm')}</b> e em observância às disposições da <b>{lei_completa}</b>, "
        f"resolvem celebrar o presente <b>{dados.get('numero_completo_aditivo', 'TERMO ADITIVO')}</b>, mediante as cláusulas seguintes:"
    )
    story.append(Paragraph(preambulo_texto, style_preamble))
    
    clausulas = gerar_clausulas_aditivo(dados)
    for cl in clausulas:
        story.append(Paragraph(cl["numero"], style_c_title))
        for item_p in cl["conteudo"]:
            story.append(Paragraph(item_p, style_c_body))
            
    story.append(Spacer(1, 4 * mm))
    fecho_texto = (
        "E, por estarem assim justos e acordados, assinam o presente Termo Aditivo em 2 (duas) vias de igual teor e forma, "
        "na presença de 2 (duas) testemunhas instrumentárias, para que produza todos os seus jurídicos e legais efeitos."
    )
    story.append(Paragraph(fecho_texto, style_preamble))
    
    dt_ext = dados.get("data_assinatura_extenso") or formatar_data_extenso(dados.get("data_assinatura") or date.today())
    story.append(Paragraph(f"Ribeirãozinho do Maranhão - MA, {dt_ext}.", ParagraphStyle('DateAdit', parent=styles['Normal'], fontName='Times-Italic', fontSize=9.5, alignment=2, spaceAfter=10)))
    
    story.append(Spacer(1, 5 * mm))
    t_ass_data = [
        [
            Paragraph(f"____________________________________________<br/><b>CONTRATANTE</b><br/>{dados.get('secretario')}<br/>{dados.get('cargo_secretario')}<br/>{dados.get('orgao_nome')}", style_sig_box),
            Paragraph(f"____________________________________________<br/><b>CONTRATADA</b><br/>{dados.get('fornecedor_representante')}<br/>{dados.get('fornecedor_cargo')}<br/>{dados.get('fornecedor_nome')}", style_sig_box)
        ],
        ["", ""],
        [
            Paragraph(f"____________________________________________<br/><b>FISCAL DO CONTRATO</b><br/>{dados.get('fiscal', 'Servidor Designado')}<br/>Fiscal Titular", style_sig_box),
            Paragraph("____________________________________________<br/><b>TESTEMUNHAS</b><br/>1. Nome: _______________________ CPF:<br/>2. Nome: _______________________ CPF:", style_sig_box)
        ]
    ]
    t_ass = Table(t_ass_data, colWidths=[page_content_w / 2.0, page_content_w / 2.0])
    t_ass.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(KeepTogether([t_ass]))
    
    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer







def aplicar_estilo_oficial_docx(doc):
    for sec in doc.sections:
        # Aumentar a margem superior para evitar que o cabecalho sobreponha o texto
        sec.top_margin = Inches(1.8)
        sec.bottom_margin = Inches(1.5)
        sec.left_margin = Inches(1.18) # 3 cm
        sec.right_margin = Inches(0.78) # 2 cm
        
        if os.path.exists(HEADER_OFICIAL_PATH):
            header = sec.header
            header.is_linked_to_previous = False
            p_head = header.paragraphs[0]
            p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_head = p_head.add_run()
            r_head.add_picture(HEADER_OFICIAL_PATH, width=Inches(6.5))

        if os.path.exists(FOOTER_OFICIAL_PATH):
            footer = sec.footer
            footer.is_linked_to_previous = False
            p_foot = footer.paragraphs[0]
            p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_foot = p_foot.add_run()
            r_foot.add_picture(FOOTER_OFICIAL_PATH, width=Inches(6.5))

def gerar_anexos_aditivo_docx(dados: Dict[str, Any]) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = docx.Document()
    aplicar_estilo_oficial_docx(doc)
    
    def set_font(run, bold=False, size=12):
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        run.bold = bold

    docs_sel = dados.get("documentos_selecionados", [])
    if not docs_sel:
        doc.add_paragraph("Nenhum documento selecionado.")
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    def add_header_padrao():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(24)
        t_h = f"ESTADO DO MARANHÃO\nPREFEITURA MUNICIPAL DE {dados.get('orgao_cidade', 'RIBEIRÃOZINHO DO MARANHÃO').upper()}\n{str(dados.get('orgao_nome', '')).upper()}\nCNPJ: {formatar_cnpj_br(dados.get('orgao_cnpj', ''))}"
        run = p.add_run(t_h)
        set_font(run, bold=True, size=12)

    first = True
    
    if any(k.lower() in d.lower() for d in docs_sel for k in ["Capa", "Rosto"]):
        if not first: doc.add_page_break()
        first = False
        add_header_padrao()
        
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2.paragraph_format.space_after = Pt(24)
        t2 = f"PROCESSO ADMINISTRATIVO Nº {dados.get('processo_adm', '')}\nCONTRATO Nº {dados.get('numero_completo', '')}\n{str(dados.get('licitacao_vinculada', '')).upper()}"
        run = p2.add_run(t2)
        set_font(run, bold=True, size=12)
        
        p3 = doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p3.paragraph_format.space_after = Pt(24)
        r3a = p3.add_run("OBJETO: ")
        set_font(r3a, bold=True)
        r3b = p3.add_run(f"O presente instrumento tem por objeto a prorrogação do prazo de vigência contratual do contrato em epígrafe que tem como objeto {dados.get('objeto', '').lower()}.")
        set_font(r3b)

        p4 = doc.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p4.paragraph_format.space_after = Pt(48)
        r4a = p4.add_run("CONTRATADA: ")
        set_font(r4a, bold=True)
        r4b = p4.add_run(f"{dados.get('fornecedor_nome', '')}, inscrita no CNPJ sob o n.º {formatar_cnpj_br(dados.get('fornecedor_cnpj', ''))}, com sede na {dados.get('fornecedor_endereco', '')}, {dados.get('fornecedor_bairro', '')} CEP: {dados.get('fornecedor_cep', '')} no município de {dados.get('fornecedor_cidade', '')}.")
        set_font(r4b)

        p5 = doc.add_paragraph()
        p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r5 = p5.add_run(f"{dados.get('orgao_endereco', '')}, {dados.get('orgao_bairro', '')} – {dados.get('orgao_cidade', '')}/{dados.get('orgao_uf', '')} CEP {dados.get('orgao_cep', '')}")
        set_font(r5, bold=True, size=11)

    if any(k.lower() in d.lower() for d in docs_sel for k in ["Extrato"]):
        if not first: doc.add_page_break()
        first = False
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(24)
        run = p.add_run(f"EXTRATO DO {str(dados.get('numero_completo_aditivo', '')).upper()}")
        set_font(run, bold=True, size=12)

        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        txt = f"EXTRATO DO {str(dados.get('numero_completo_aditivo', '')).upper()}; PROCESSO ADMINISTRATIVO Nº {dados.get('processo_adm', '')}; {dados.get('licitacao_vinculada', '')}.\n\nCONTRATANTE: {dados.get('orgao_nome', '')}.\nCONTRATADA: {dados.get('fornecedor_nome', '')}, CNPJ: {formatar_cnpj_br(dados.get('fornecedor_cnpj', ''))}.\nOBJETO: {dados.get('objeto_aditivo', '')}.\nVALOR: {formatar_moeda_br(dados.get('valor_aditivo', 0.0))}.\nDATA DE ASSINATURA: {formatar_data_extenso(dados.get('data_assinatura', date.today()))}.\nASSINAM: {dados.get('secretario', '')} (Contratante) e {dados.get('fornecedor_representante', '')} (Contratada)."
        run2 = p2.add_run(txt)
        set_font(run2)

    if any(k.lower() in d.lower() for d in docs_sel for k in ["Requerimento"]):
        if not first: doc.add_page_break()
        first = False
        add_header_padrao()
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(24)
        r = p.add_run("REQUERIMENTO DE PRORROGAÇÃO DE PRAZO")
        set_font(r, bold=True, size=12)
        
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2.paragraph_format.space_after = Pt(24)
        r2 = p2.add_run(f"A Senhora\n{dados.get('fiscal', 'Gerente de Contratos')}\nFiscal/Gerente de contratos.\n\nASSUNTO: Prorrogação de prazo de vigência do contrato nº {dados.get('numero_completo', '')}")
        set_font(r2, size=12)
        
        p3 = doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p3.paragraph_format.space_after = Pt(24)
        r3 = p3.add_run(f"Considerando que o contrato Nº {dados.get('numero_completo', '')} tem previsão de encerramento em {formatar_data_extenso(dados.get('data_vencimento_atual', date.today()))}, é imprescindível a continuidade na contratação de empresa especializada para {dados.get('objeto', '').lower()}, uma vez que tais serviços são fundamentais para a {dados.get('orgao_nome', 'Administração')}.")
        set_font(r3)
        
        p4 = doc.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p4.paragraph_format.space_after = Pt(48)
        r4 = p4.add_run(f"{dados.get('orgao_cidade', 'Ribeirãozinho do Maranhão')}, {formatar_data_extenso(date.today())}.")
        set_font(r4)
        
        p5 = doc.add_paragraph()
        p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p5.paragraph_format.space_after = Pt(36)
        r5 = p5.add_run(f"____________________________________________________\n{dados.get('secretario', '')}\n{dados.get('cargo_secretario', '')}\n{dados.get('orgao_nome', '')}")
        set_font(r5)
        
        p6 = doc.add_paragraph()
        p6.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r6 = p6.add_run("RECEBIDO EM: ____/___/_____")
        set_font(r6)

    if any(k.lower() in d.lower() for d in docs_sel for k in ["Justificativa"]):
        if not first: doc.add_page_break()
        first = False
        add_header_padrao()
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(24)
        r = p.add_run("JUSTIFICATIVA")
        set_font(r, bold=True, size=12)
        
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2.paragraph_format.space_after = Pt(24)
        r2 = p2.add_run(f"Assunto: PRORROGAÇÃO DE PRAZO CONTRATUAL\nContrato nº: {dados.get('numero_completo', '')}\nContratada: {dados.get('fornecedor_nome', '')}\n\nExma. Sr (a).\n{dados.get('secretario', '')}\n{dados.get('cargo_secretario', '')}.")
        set_font(r2)
        
        p3 = doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p3.paragraph_format.space_after = Pt(12)
        r3 = p3.add_run(f"O Contrato nº {dados.get('numero_completo', '')}, o presente instrumento tem por objeto a prorrogação do prazo de vigência contratual do contrato em epígrafe que tem como objeto {dados.get('objeto', '').lower()}. Ocorre que a secretaria requer a prorrogação do prazo de Vigência Contratual tendo em vista que a validade é até {formatar_data_extenso(dados.get('data_vencimento_atual', date.today()))}, e por se tratar de um serviço contínuo e essencial, em razão das necessidades da {dados.get('orgao_nome', '')}, necessita a Administração prorrogar o Contrato vigente por igual período, na forma dos artigos 106 e 107 da lei n° 14.133, de 2021.\n\nAssim, apresentamos a seguir as razões que nos levam a entender viável e justificada a prorrogação da vigência do supracitado contrato:\n\na) os serviços prestados são contínuos e essenciais;\nb) considerando que os serviços foram prestados de maneira satisfatória;\nc) os serviços vêm sendo prestados de modo regular e tem produzido os efeitos desejados;\nd) sob o ponto de vista legal.")
        set_font(r3)

        p_art = doc.add_paragraph()
        p_art.paragraph_format.left_indent = Inches(1.5)
        p_art.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_art = p_art.add_run("Art. 107. Os contratos de serviços e fornecimentos contínuos poderão ser prorrogados sucessivamente, respeitada a vigência máxima decenal, desde que haja previsão em edital e que a autoridade competente ateste que as condições e os preços permanecem vantajosos para a Administração, permitida a negociação com o contratado ou a extinção contratual sem ônus para qualquer das partes.")
        set_font(r_art, bold=True, size=11)

    doc.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_anexos_contrato_docx(dados: Dict[str, Any]) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = docx.Document()
    aplicar_estilo_oficial_docx(doc)
    
    def set_font(run, bold=False, size=12):
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        run.bold = bold

    docs_sel = dados.get("documentos_selecionados", [])
    if not docs_sel:
        doc.add_paragraph("Nenhum documento selecionado.")
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    def add_header_padrao():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(24)
        t_h = f"ESTADO DO MARANHÃO\nPREFEITURA MUNICIPAL DE {dados.get('orgao_cidade', 'RIBEIRÃOZINHO DO MARANHÃO').upper()}\n{str(dados.get('orgao_nome', '')).upper()}\nCNPJ: {formatar_cnpj_br(dados.get('orgao_cnpj', ''))}"
        run = p.add_run(t_h)
        set_font(run, bold=True, size=12)

    first = True
    
    if any(k.lower() in d.lower() for d in docs_sel for k in ["Capa", "Rosto"]):
        if not first: doc.add_page_break()
        first = False
        add_header_padrao()
        
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2.paragraph_format.space_after = Pt(24)
        t2 = f"PROCESSO ADMINISTRATIVO Nº {dados.get('processo_adm', '')}\nCONTRATO Nº {dados.get('numero_completo', '')}\nMODALIDADE: {str(dados.get('modalidade', '')).upper()}"
        run = p2.add_run(t2)
        set_font(run, bold=True, size=12)
        
        p3 = doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p3.paragraph_format.space_after = Pt(24)
        r3a = p3.add_run("OBJETO: ")
        set_font(r3a, bold=True)
        r3b = p3.add_run(f"O presente instrumento tem por objeto {dados.get('objeto', '').lower()}.")
        set_font(r3b)

        p4 = doc.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p4.paragraph_format.space_after = Pt(48)
        r4a = p4.add_run("CONTRATADA: ")
        set_font(r4a, bold=True)
        r4b = p4.add_run(f"{dados.get('fornecedor_nome', '')}, inscrita no CNPJ sob o n.º {formatar_cnpj_br(dados.get('fornecedor_cnpj', ''))}, com sede na {dados.get('fornecedor_endereco', '')}, {dados.get('fornecedor_bairro', '')} CEP: {dados.get('fornecedor_cep', '')} no município de {dados.get('fornecedor_cidade', '')}.")
        set_font(r4b)

        p5 = doc.add_paragraph()
        p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r5 = p5.add_run(f"{dados.get('orgao_endereco', '')}, {dados.get('orgao_bairro', '')} – {dados.get('orgao_cidade', '')}/{dados.get('orgao_uf', '')} CEP {dados.get('orgao_cep', '')}")
        set_font(r5, bold=True, size=11)

    if any(k.lower() in d.lower() for d in docs_sel for k in ["Extrato"]):
        if not first: doc.add_page_break()
        first = False
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(24)
        run = p.add_run(f"EXTRATO DO CONTRATO Nº {str(dados.get('numero_completo', '')).upper()}")
        set_font(run, bold=True, size=12)

        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        txt = f"EXTRATO DO CONTRATO Nº {str(dados.get('numero_completo', '')).upper()}; PROCESSO ADMINISTRATIVO Nº {dados.get('processo_adm', '')}; MODALIDADE: {dados.get('modalidade', '')}.\n\nCONTRATANTE: {dados.get('orgao_nome', '')}.\nCONTRATADA: {dados.get('fornecedor_nome', '')}, CNPJ: {formatar_cnpj_br(dados.get('fornecedor_cnpj', ''))}.\nOBJETO: {dados.get('objeto', '')}.\nVALOR GLOBAL: {formatar_moeda_br(dados.get('valor_total', 0.0))}.\nDATA DE ASSINATURA: {formatar_data_extenso(dados.get('data_assinatura', date.today()))}.\nASSINAM: {dados.get('secretario', '')} (Contratante) e {dados.get('fornecedor_representante', '')} (Contratada)."
        run2 = p2.add_run(txt)
        set_font(run2)

    if any(k.lower() in d.lower() for d in docs_sel for k in ["Requerimento"]):
        if not first: doc.add_page_break()
        first = False
        add_header_padrao()
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(24)
        r = p.add_run("REQUERIMENTO DE CONTRATAÇÃO")
        set_font(r, bold=True, size=12)
        
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2.paragraph_format.space_after = Pt(24)
        r2 = p2.add_run(f"A Senhora\n{dados.get('secretario', '')}\n{dados.get('cargo_secretario', '')}.\n\nASSUNTO: Contratação de empresa para {dados.get('objeto', '').lower()}")
        set_font(r2, size=12)
        
        p3 = doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p3.paragraph_format.space_after = Pt(24)
        r3 = p3.add_run(f"Considerando a necessidade contínua e inadiável desta Administração, requeremos a formalização do CONTRATO Nº {dados.get('numero_completo', '')}, vinculado ao Processo Administrativo nº {dados.get('processo_adm', '')} ({str(dados.get('modalidade', '')).upper()}).\n\nA empresa vencedora do certame é {dados.get('fornecedor_nome', '')} (CNPJ: {formatar_cnpj_br(dados.get('fornecedor_cnpj', ''))}), no valor global de {formatar_moeda_br(dados.get('valor_total', 0.0))}.\n\nJustifica-se o presente pleito visando a manutenção dos serviços de interesse público, de acordo com o planejamento estratégico municipal.")
        set_font(r3)
        
        p4 = doc.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p4.paragraph_format.space_after = Pt(48)
        r4 = p4.add_run(f"{dados.get('orgao_cidade', 'Ribeirãozinho do Maranhão')}, {formatar_data_extenso(date.today())}.")
        set_font(r4)
        
        p5 = doc.add_paragraph()
        p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p5.paragraph_format.space_after = Pt(24)
        r5 = p5.add_run(f"____________________________________________________\n{dados.get('solicitante', 'Setor Requisitante')}\nSetor Solicitante")
        set_font(r5)
        
        p6 = doc.add_paragraph()
        p6.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r6 = p6.add_run("RECEBIDO EM: ____/___/_____")
        set_font(r6)

    doc.save(buffer)
    buffer.seek(0)
    return buffer


import zipfile

def gerar_pacote_zip(dados: Dict[str, Any], is_aditivo=False) -> io.BytesIO:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        
        # 1. Gerar e adicionar a Minuta (Contrato ou Aditivo)
        safe_num = str(dados.get('numero_completo', 'S_N')).replace('/', '_')
        if is_aditivo:
            minuta_bytes = gerar_aditivo_docx(dados)
            safe_num_adit = str(dados.get('numero_aditivo', 'S_N'))
            zip_file.writestr(f"1_Minuta_Aditivo_{safe_num_adit}_{safe_num}.docx", minuta_bytes.getvalue())
        else:
            minuta_bytes = gerar_contrato_docx(dados)
            zip_file.writestr(f"1_Minuta_Contrato_{safe_num}.docx", minuta_bytes.getvalue())

        # 2. Gerar Anexos Separadamente
        docs_sel = dados.get("documentos_selecionados", [])
        for doc_name in docs_sel:
            # Cria um dicionário falso selecionando só 1 documento para a função de anexos gerar só ele
            dados_temp = dados.copy()
            dados_temp["documentos_selecionados"] = [doc_name]
            
            if is_aditivo:
                anexo_bytes = gerar_anexos_aditivo_docx(dados_temp)
            else:
                anexo_bytes = gerar_anexos_contrato_docx(dados_temp)
                
            safe_doc_name = doc_name.replace(" ", "_").replace("/", "_")
            zip_file.writestr(f"Anexo_{safe_doc_name}_{safe_num}.docx", anexo_bytes.getvalue())
            
    zip_buffer.seek(0)
    return zip_buffer
