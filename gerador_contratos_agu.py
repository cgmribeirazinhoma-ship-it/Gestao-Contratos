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

def is_summary_row(row_vals: list) -> bool:
    """
    Detecta de forma inteligente se uma linha representa um rodapé / linha de total / resumo (ex: TOTAL GERAL).
    Não confunde descrições de produtos normais que contêm a palavra 'total' como texto.
    """
    if not row_vals:
        return False
    non_empty = [str(c or '').strip() for c in row_vals if str(c or '').strip()]
    if not non_empty:
        return False
    c0 = non_empty[0].upper()
    c1 = non_empty[1].upper() if len(non_empty) > 1 else ""

    summary_keywords = [
        'TOTAL', 'TOTAL GERAL', 'SUBTOTAL', 'VALOR TOTAL', 'VALOR GLOBAL',
        'SOMA', 'SOMA TOTAL', 'TOTAL DA PROPOSTA', 'TOTAL DO OBJETO', 'TOTAL ESTIMADO', 'TOTAL HOMOLOGADO'
    ]
    if c0 in summary_keywords or c1 in summary_keywords:
        return True
    summary_prefixes = [
        'TOTAL:', 'TOTAL ', 'TOTAL-', 'VALOR TOTAL:', 'VALOR TOTAL ', 'VALOR GLOBAL:', 'VALOR GLOBAL ',
        'SUBTOTAL:', 'SUBTOTAL ', 'SUB-TOTAL', 'SUB TOTAL', 'SOMA:', 'SOMA ', 'SOMA TOTAL'
    ]
    if any(c0.startswith(k) for k in summary_prefixes):
        return True
    if any(c1.startswith(k) for k in summary_prefixes):
        return True
    return False

def gerar_planilha_itens_excel_formatada(headers: list, rows: list, total_geral: float = 0.0, sheet_name: str = "Itens do Contrato", is_dotacao: bool = False, append_total_row: bool = False, has_real_headers: bool = None) -> bytes:
    """
    Gera uma planilha Excel (.xlsx) altamente profissional com:
    - Quebra de texto automática ativada em todas as células (wrap_text=True)
    - Bordas finas completas em todas as células (ajuste de texto às bordas)
    - Auto-ajuste proporcional da largura das colunas
    - Preserva 100% a fidelidade de linhas e colunas da planilha modelo subida
    - Se is_dotacao: padrão oficial com bordas e texto azul marinho (#224061),
      fundo azul suave (#E2EBF4) para títulos/cabeçalhos e fundo branco (#FFFFFF) para dados.
    - Se itens: cabeçalho padronizado em azul marinho (#003366) com texto branco em negrito.
      Linhas de resumo/total já existentes na planilha modelo recebem destaque em negrito sem duplicação.
    """
    import io
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

    def _clean_opx(val):
        if val is None:
            return ""
        s = str(val).encode('utf-8', errors='ignore').decode('utf-8')
        return ILLEGAL_CHARACTERS_RE.sub('', s)

    wb = openpyxl.Workbook()
    ws = wb.active

    # Auto-detecta se é dotação exclusivamente pelo nome da aba ou parâmetro explícito
    if not is_dotacao:
        sn_up = sheet_name.upper()
        if "DOTAÇ" in sn_up or "DOTAC" in sn_up:
            is_dotacao = True
        else:
            headers, rows = reconciliar_tabela_irregular(headers, rows, is_dotacao=False)

    if is_dotacao:
        ws.title = "Dotação Orçamentária"[:31]

        # Paleta oficial extraída do padrão
        navy_side = Side(border_style='thin', color='224061')
        cell_border = Border(left=navy_side, right=navy_side, top=navy_side, bottom=navy_side)
        
        bg_blue = PatternFill(start_color='E2EBF4', end_color='E2EBF4', fill_type='solid')
        bg_white = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        zebra_blue_fill = PatternFill(start_color='EBF3FB', end_color='EBF3FB', fill_type='solid')
        font_bold = Font(name='Arial', size=9.5, bold=True, color='224061')
        font_regular = Font(name='Arial', size=9.5, color='224061')

        # Verifica se é estrutura de dotação oficial de 3 colunas
        clean_rows = []
        for r_i, r in enumerate(rows):
            c1_chk = str(r[1] or '').strip() if len(r) > 1 else ''
            c2_chk = str(r[2] or '').strip() if len(r) > 2 else ''
            if r_i > 0 and not c1_chk and not c2_chk:
                continue
            clean_rows.append(r)
        rows = clean_rows
        is_st, start_unid, end_unid, has_ex = is_structural_budget_table(headers, rows)
        if is_st:
            for r_idx, r_vals in enumerate(rows, start=1):
                c1_val = str(r_vals[1] if len(r_vals) > 1 else "")
                num_lines = max((len(c1_val) // 42) + c1_val.count('\n') + 1, 1)
                ws.row_dimensions[r_idx].height = max(26, num_lines * 19)

                for c_idx in range(1, 4):
                    val = r_vals[c_idx - 1] if (c_idx - 1) < len(r_vals) else ""
                    val_str = _clean_opx(str(val).replace("<br>", "\n").replace("<br/>", "\n"))
                    cell = ws.cell(row=r_idx, column=c_idx, value=val_str)
                    cell.border = cell_border
                    
                    if c_idx == 1 or r_idx == 1:
                        cell.fill = bg_blue
                        cell.font = font_bold
                    else:
                        cell.fill = zebra_blue_fill if (r_idx % 2 == 0) else bg_white
                        cell.font = font_regular

                    if c_idx == 1:
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                    elif r_idx == 1:
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    elif c_idx == 2:
                        cell.alignment = Alignment(horizontal='justify', vertical='center', wrap_text=True)
                    else:
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            # Mescla Exercício na linha 1 entre coluna 2 e 3
            ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=3)
            cell_ex = ws.cell(row=1, column=2)
            cell_ex.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell_ex.fill = bg_blue
            cell_ex.font = font_bold

            # Mescla vertical na coluna A de start_unid+1 a end_unid+1 (ZERO LINHAS NA COLUNA A!)
            if start_unid != -1 and end_unid > start_unid:
                ws.merge_cells(start_row=start_unid + 1, start_column=1, end_row=end_unid + 1, end_column=1)
                cell_u = ws.cell(row=start_unid + 1, column=1)
                cell_u.value = "Unidade Orçamentária\n/Atividade"
                cell_u.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell_u.fill = bg_blue
                cell_u.font = font_bold
            elif start_unid != -1 and end_unid == start_unid:
                cell_u = ws.cell(row=start_unid + 1, column=1)
                cell_u.value = "Unidade Orçamentária\n/Atividade"
                cell_u.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell_u.fill = bg_blue
                cell_u.font = font_bold

            # Garante que nenhuma célula da coluna A fique em branco se houver dados na coluna B
            for r_idx in range(1, len(rows) + 1):
                cell_a = ws.cell(row=r_idx, column=1)
                cell_b = ws.cell(row=r_idx, column=2)
                if type(cell_a).__name__ != 'MergedCell':
                    if not str(cell_a.value or '').strip() and str(cell_b.value or '').strip():
                        cell_a.value = "Unidade Orçamentária\n/Atividade"
                        cell_a.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                        cell_a.fill = bg_blue
                        cell_a.font = font_bold

            ws.column_dimensions['A'].width = 28
            ws.column_dimensions['B'].width = 58
            ws.column_dimensions['C'].width = 26

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return out.getvalue()
        else:
            # Dotação adaptativa (outras configurações de colunas)
            if has_real_headers is not None:
                has_real_hdrs_dt = bool(has_real_headers)
            else:
                has_real_hdrs_dt = headers and not any(str(h).strip().startswith("COLUNA ") for h in headers) and not all(str(h).strip().startswith("COLUNA ") for h in headers)
            if has_real_hdrs_dt:
                clean_hdrs_dt = [_clean_opx("" if (str(h).startswith("COLUNA ") or str(h).startswith("_")) else str(h)) for h in headers]
                ws.append(clean_hdrs_dt)
                ws.row_dimensions[1].height = 26
                for col_idx in range(1, len(headers) + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.fill = bg_blue
                    cell.font = font_bold
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.border = cell_border
                start_row_dt = 2
            else:
                start_row_dt = 1

            for row_idx, r_vals in enumerate(rows, start=start_row_dt):
                row_data = [_clean_opx(str(r_vals[c]).replace("<br>", "\n").replace("<br/>", "\n")) if c < len(r_vals) else "" for c in range(len(headers))]
                ws.append(row_data)
                ws.row_dimensions[row_idx].height = 24

                is_banner = (len(r_vals) > 1 and str(r_vals[0] or '').strip() and not any(str(x or '').strip() for x in r_vals[1:]))
                if is_banner:
                    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=len(headers))
                    cell_b = ws.cell(row=row_idx, column=1)
                    cell_b.fill = bg_blue
                    cell_b.font = font_bold
                    cell_b.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    for col_idx in range(1, len(headers) + 1):
                        ws.cell(row=row_idx, column=col_idx).border = cell_border
                    continue

                for col_idx in range(1, len(headers) + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.border = cell_border
                    if col_idx == 1:
                        cell.fill = bg_blue
                        cell.font = font_bold
                    else:
                        cell.fill = zebra_blue_fill if (row_idx % 2 == 0) else bg_white
                        cell.font = font_regular
                    align_h = 'center' if col_idx == len(headers) or col_idx == 1 else 'left'
                    cell.alignment = Alignment(horizontal=align_h, vertical='center', wrap_text=True)

            for col_idx, col in enumerate(ws.columns, start=1):
                max_len = max(len(str(cell.value or '')) for cell in col[:15]) if len(col) > 0 else 12
                col_width = min(max(max_len + 4, 14), 45)
                ws.column_dimensions[get_column_letter(col_idx)].width = col_width

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return out.getvalue()

    # Fluxo normal para itens de compras / serviços
    ws.title = sheet_name[:31]

    # Estilos
    thin_border_side = Side(border_style='thin', color='CBD5E1')
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    header_fill = PatternFill(start_color='003366', end_color='003366', fill_type='solid')
    header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')

    body_font = Font(name='Arial', size=9.5)
    zebra_fill = PatternFill(start_color='EBF3FB', end_color='EBF3FB', fill_type='solid')
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
            col_alignments.append('justify')

    # 1. Escreve Cabeçalho (apenas se a tabela tiver cabeçalhos autênticos)
    if has_real_headers is not None:
        has_real_hdrs = bool(has_real_headers)
    else:
        has_real_hdrs = headers and not any(str(h).strip().startswith("COLUNA ") for h in headers) and not all(str(h).strip().startswith("COLUNA ") for h in headers)
    if has_real_hdrs:
        clean_hdrs_it = [_clean_opx("" if (str(h).startswith("COLUNA ") or str(h).startswith("_")) else str(h)) for h in headers]
        ws.append(clean_hdrs_it)
        has_multi_line_hdr = any('\n' in str(h) or '<br>' in str(h).lower() or len(str(h).strip()) > 10 for h in headers)
        ws.row_dimensions[1].height = 36 if has_multi_line_hdr else 26
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = cell_border
        start_row_data = 2
    else:
        start_row_data = 1

    # 2. Escreve Linhas com Quebra de Texto e Ajuste de Bordas (Preservação 100% de linhas da planilha modelo)
    has_existing_summary = False
    for row_idx, r_vals in enumerate(rows, start=start_row_data):
        row_data = [_clean_opx(r_vals[c]) if c < len(r_vals) else "" for c in range(len(headers))]
        ws.append(row_data)
        ws.row_dimensions[row_idx].height = 24

        is_res = is_summary_row(r_vals)
        if is_res:
            has_existing_summary = True

        row_fill = total_fill if is_res else (zebra_fill if (row_idx % 2 == 0) else white_fill)
        row_font = total_font if is_res else body_font

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = row_font
            cell.fill = row_fill
            cell.border = cell_border
            align_h = col_alignments[col_idx - 1] if (col_idx - 1) < len(col_alignments) else 'left'
            if is_res and any(k in str(headers[col_idx - 1]).upper() for k in ['VALOR', 'PREÇO', 'PRECO', 'TOTAL', 'R$']):
                align_h = 'right'
            cell.alignment = Alignment(horizontal=align_h, vertical='center', wrap_text=True)

    # 3. Linha de Totalizador adicional (APENAS se explicitamente solicitado e se a planilha modelo não contiver linha de resumo)
    last_row = len(rows) + 1
    if append_total_row and (not has_existing_summary) and total_geral > 0:
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
        elif any(k in col_header for k in ['QTD', 'QUANT', 'QUANTIDADE']):
            if any(k in col_header for k in ['DISP', 'DISPONIVEL', 'DISPONÍVEL', 'SOLIC', 'ESTIM', 'CONTRAT', 'SALDO']) or len(col_header) > 12:
                col_width = 18
            else:
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


def detectar_se_eh_cabecalho(row: list, is_dotacao: bool = False) -> bool:
    """
    Detecta de forma inteligente se uma linha representa um cabeçalho oficial ou se é dado.
    Retorna True se for cabeçalho, False se for dado.
    """
    if not row or not any(str(x).strip() for x in row):
        return False
    row_strs = [str(x or "").strip() for x in row if str(x or "").strip()]
    if not row_strs:
        return False
    row_up = " ".join(row_strs).upper()
    
    # 1. Se contiver códigos orçamentários pontuados (ex: 04.122.0052.2006.0000, 33.90.30) ou sequências longas de dígitos, NÃO é cabeçalho
    if any(re.search(r'\d{2}\.\d{3}', s) or re.search(r'\d{2}\.\d{2}\.\d{2}', s) or re.search(r'\d{5,}', s) for s in row_strs):
        return False
    
    # 2. Se contiver valores monetários com R$ ou formato numérico de preço (ex: 1.250,00), NÃO é cabeçalho
    if any('R$' in s or re.search(r'^\d{1,3}(\.\d{3})*,\d{2}$', s) for s in row_strs):
        return False
    
    # 3. Se a linha for linha de dados da tabela estrutural (Exercício, Poder, Órgão, Natureza), NÃO é cabeçalho
    c0_up = str(row_strs[0]).upper()
    if any(k in c0_up for k in ['EXERC', 'EXERCÍCIO', 'EXERCICIO']) and len(row_strs) > 1 and any(re.search(r'\d{4}', s) for s in row_strs[1:]):
        return False
    if 'PODER' in c0_up and len(row_strs) > 1 and any(k in str(s).upper() for s in row_strs[1:] for k in ['EXECUTIVO', 'LEGISLATIVO', 'JUDICIÁRIO', 'JUDICIARIO']):
        return False
    if any(k in c0_up for k in ['ÓRGÃO', 'ORGAO']) and len(row_strs) > 1 and any(k in str(s).upper() for s in row_strs[1:] for k in ['FUNDO', 'PREFEITURA', 'SECRETARIA', 'GABINETE', 'SEC.']):
        return False
    if any(k in c0_up for k in ['NATUREZA', 'ELEMENTO']) and len(row_strs) > 1 and any(k in str(s).upper() for s in row_strs[1:] for k in ['MATERIAL', 'SERVIÇO', 'SERVICO', 'CONSUMO', 'TERCEIROS']):
        return False
        
    # 4. Termos específicos de dados operacionais / nomes de secretarias concretas (com limites de palavras)
    organ_patterns = [
        r'\bMANUTENÇÃO\b', r'\bMANUTENCAO\b', r'\bMANUT\.', r'\bGESTÃO DA\b', r'\bGESTAO DA\b',
        r'\bAÇÕES DA\b', r'\bACOES DA\b', r'\bSERVIÇOS URBANOS\b', r'\bSERVICOS URBANOS\b',
        r'\bFUNDO MUNICIPAL DE\b', r'\bSECRETARIA MUNICIPAL DE\b', r'\bSECRETARIA MUNICIPAL DA\b', r'\bSECRETARIA MUNICIPAL DO\b',
        r'\bPREFEITURA MUNICIPAL DE\b', r'\bPREFEITURA MUNICIPAL DA\b', r'\bGABINETE DO PREFEITO\b',
        r'\bSECRETARIA DE\b', r'\bSECRETARIA DA\b', r'\bSECRETARIA DO\b',
        r'\bFUNDO DE\b', r'\bFUNDO DA\b', r'\bFUNDO DO\b'
    ]
    if any(re.search(pat, row_up) for pat in organ_patterns):
        return False

    # 5. Se for dotação, limite é 50 caracteres. Se for tabela de itens/objeto, especificações podem ter até 160 caracteres
    max_len_lim = 50 if is_dotacao else 160
    if any(len(s) > max_len_lim for s in row_strs):
        return False

    # 6. Termos típicos de cabeçalho
    header_terms = [
        'ITEM', 'DISCRIMINAÇÃO', 'DISCRIMINACAO', 'ESPECIFICAÇÃO', 'ESPECIFICACAO',
        'DESCRIÇÃO', 'DESCRICAO', 'UNIDADE', 'UNID', 'UND', 'QUANT', 'QUANTIDADE', 'QTD',
        'VALOR', 'PREÇO', 'PRECO', 'UNITÁRIO', 'UNITARIO', 'TOTAL', 'SUBTOTAL', 'MARCA',
        'CÓDIGO', 'CODIGO', 'COD', 'DOTAÇÃO', 'DOTACAO', 'CLASSIFICAÇÃO', 'CLASSIFICACAO',
        'ÓRGÃO', 'ORGAO', 'PROGRAMA', 'PROJETO', 'ATIVIDADE', 'FONTE', 'FICHA', 'NATUREZA',
        'ELEMENTO', 'RUBRICA', 'EXERCÍCIO', 'EXERCICIO', 'PODER', 'CONTA', 'DESTINAÇÃO', 'AÇÃO', 'ACAO',
        'SECRETARIA', 'SECRETARIAS', 'SETOR', 'DEPARTAMENTO', 'VENCIMENTO', 'DATA', 'PRAZO', 'MODALIDADE',
        'DEMANDANTE', 'CATMAT', 'CATSER', 'REF', 'REFERÊNCIA', 'REFERENCIA', 'MODELO', 'FABRICANTE',
        'OBSERVAÇÃO', 'OBS', 'LOTE', 'LOCAL'
    ]
    matches = sum(1 for ht in header_terms if ht in row_up)
    if matches >= 1:
        return True
    
    # Se todas as células forem palavras sem números compostos (ex: nomes de colunas), é cabeçalho
    if all(not re.search(r'\d{2,}', s) for s in row_strs):
        return True
    
    return False

def gerar_cabecalhos_adaptativos(sample_rows: list, num_cols: int, is_dotacao: bool = False) -> list:
    """
    Gera cabeçalhos inteligentes e adaptados ao modelo de dados contido nas colunas.
    """
    if num_cols <= 0:
        return []
    
    import re
    headers = []
    for ci in range(num_cols):
        col_samples = [str(r[ci]).strip() for r in sample_rows if ci < len(r) and str(r[ci]).strip()][:15]
        col_text = " ".join(col_samples).upper()
        
        has_codes = any(re.search(r'\d{2}\.\d{3}\.\d{4}', s) or re.search(r'\d{2}\.\d{2}\.\d{2}', s) for s in col_samples)
        has_currency = any('R$' in s or re.search(r'^\d{1,3}(\.\d{3})*,\d{2}$', s) for s in col_samples)
        has_short_int = all(s.isdigit() and len(s) <= 4 for s in col_samples if s) and len(col_samples) > 0
        has_unid = all(s in ['UN', 'UNID', 'UND', 'KG', 'M', 'L', 'CX', 'PCT', 'PAR', 'HR', 'DIA', 'MES', 'SV', 'GL', 'LOTE'] for s in col_samples if s) and len(col_samples) > 0
        has_org = any(k in col_text for k in ['MANUTEN', 'FUNDO', 'SECRETARIA', 'GABINETE', 'SEC.', 'PREFEITURA', 'SEMED'])
        
        if is_dotacao:
            if has_codes:
                headers.append('CÓDIGO / DOTAÇÃO')
            elif has_currency:
                headers.append('VALOR ESTIMADO (R$)')
            elif has_org:
                headers.append('UNIDADE ORÇAMENTÁRIA / ATIVIDADE')
            elif any(k in col_text for k in ['EXERC', '2025', '2026', '2027']) and ci == 0:
                headers.append('EXERCÍCIO')
            elif ci == 0 and num_cols == 2:
                headers.append('UNIDADE ORÇAMENTÁRIA / ATIVIDADE')
            elif ci == 1 and num_cols == 2:
                headers.append('CÓDIGO / DOTAÇÃO')
            elif ci == 0 and num_cols == 3:
                headers.append('CLASSIFICAÇÃO')
            elif ci == 1 and num_cols == 3:
                headers.append('DESCRIÇÃO / ESPECIFICAÇÃO')
            elif ci == 2 and num_cols == 3:
                headers.append('CÓDIGO')
            else:
                headers.append(f'COLUNA {ci + 1}')
        else:
            if has_currency or any(k in col_text for k in ['TOTAL', 'VALOR', 'PREÇO']):
                headers.append('VALOR TOTAL (R$)' if ci == num_cols - 1 else 'VALOR UNITÁRIO (R$)')
            elif has_short_int and ci == 0:
                headers.append('ITEM')
            elif has_short_int:
                headers.append('QUANTIDADE')
            elif has_unid:
                headers.append('UNIDADE')
            elif ci == 1 or (ci == 0 and num_cols <= 2):
                headers.append('DESCRIÇÃO DO OBJETO')
            else:
                headers.append(f'COLUNA {ci + 1}')
                
    vistos = {}
    final_headers = []
    for h in headers:
        if h in vistos:
            vistos[h] += 1
            final_headers.append(f"{h} {vistos[h]}")
        else:
            vistos[h] = 1
            final_headers.append(h)
    return final_headers

def is_structural_budget_table(headers: list, rows: list):
    """
    Detecta se uma tabela representa estritamente a estrutura de dotação orçamentária oficial
    (Exercício, Poder, Órgão, Unidades Orçamentárias/Atividades e Natureza da despesa).
    Retorna: (is_structural: bool, start_unid: int, end_unid: int, has_exercicio: bool)
    """
    if not rows or len(rows) < 2:
        return False, -1, -1, False
    
    n_cols = len(headers) if headers else (len(rows[0]) if rows else 0)
    if n_cols != 3 or max(len(r) for r in rows) != 3:
        return False, -1, -1, False
    
    keywords_found = 0
    start_unid = -1
    end_unid = -1
    has_exercicio = False
    
    for idx, r in enumerate(rows):
        c0 = str(r[0] or "").strip().upper()
        c1 = str(r[1] or "").strip() if len(r) > 1 else ""
        c2 = str(r[2] or "").strip() if len(r) > 2 else ""

        c1_up = c1.upper()
        is_natureza = any(k in c0 for k in ["NATUREZA", "ELEMENTO"]) or (
            (re.match(r'^\d{1,2}\.\d{1,2}\.\d{1,2}', c1) or re.match(r'^\d{1,2}\.\d{1,2}\.\d{1,2}', c2)) and
            any(k in (c1_up + " " + str(r[2] or "").upper()) for k in ["MATERIAL", "SERVIÇO", "SERVICO", "CONSUMO", "OBRAS", "EQUIPAMENTO", "PESSOAL", "TERCEIROS"])
        )

        if any(k in c0 for k in ["EXERC", "EXERCÍCIO", "EXERCICIO"]):
            keywords_found += 1
            has_exercicio = True
        elif "PODER" in c0:
            keywords_found += 1
        elif any(k in c0 for k in ["ÓRGÃO", "ORGAO"]):
            keywords_found += 1
        elif is_natureza:
            keywords_found += 1
        elif any(k in c0 for k in ["UNIDADE", "ATIVIDADE", "PROJETO"]):
            keywords_found += 1
            if start_unid == -1:
                start_unid = idx
                end_unid = idx
            else:
                end_unid = idx
        elif any(k in (c0 + " " + c1_up) for k in ['MANUTEN', 'SECRETARIA', 'FUNDO', 'GABINETE', 'SEC.', 'PREFEITURA', 'SERVIÇO', 'AÇÕES', 'ACOES', 'GESTÃO', 'GESTAO', 'SEMED', 'SMS', 'FMS', 'FMAS', 'FME', 'ADMINISTRAÇÃO', 'ADMINISTRACAO']) or (
            re.search(r'\d{2}\.\d{3}', c2) or re.search(r'\d{2}\.\d{2}', c2) or re.search(r'\d{4}\.\d{4}', c2)
        ):
            # Linha de secretaria / unidade orçamentária / atividade (mesmo com c0 vazio)
            if not any(k in c0 for k in ["EXERC", "PODER", "ÓRGÃO", "ORGAO"]):
                if start_unid == -1:
                    start_unid = idx
                    end_unid = idx
                else:
                    end_unid = idx
        elif start_unid != -1 and idx > start_unid and not is_natureza:
            if not any(k in c0 for k in ["EXERC", "PODER", "ÓRGÃO", "ORGAO", "FONTE", "FICHA"]):
                if c1 or c2:
                    end_unid = idx
    
    is_structural = (keywords_found >= 2 or (start_unid != -1 and end_unid >= start_unid))
    return is_structural, start_unid, end_unid, has_exercicio

def calcular_larguras_colunas_adaptativas(headers: list, rows: list, total_width_pt: float = 485.0, is_dotacao: bool = False) -> tuple:
    """
    Calcula dinamicamente a largura proporcional ideal e o alinhamento adequado
    para cada coluna de qualquer planilha ou tabela subida (2, 3, 4, 5, 6, 7, 8+ colunas),
    garantindo que colunas de descrição/especificação recebam espaço generoso para quebra de texto,
    enquanto itens, códigos, unidades, quantidades e valores monetários fiquem compactos e proporcionais.
    Retorna: (col_widths_pt: list[float], col_alignments: list[str])
    """
    import re
    if not headers:
        return [], []
    
    n_cols = len(headers)
    if n_cols == 1:
        return [round(total_width_pt, 2)], ["LEFT"]

    col_alignments = []
    col_raw_weights = []

    for ci, h in enumerate(headers):
        h_str = str(h or "").strip()
        h_up = h_str.upper()
        
        sample_vals = [str(r[ci] or "").strip() for r in rows if ci < len(r) and str(r[ci] or "").strip()]
        avg_len = (sum(len(s) for s in sample_vals) / len(sample_vals)) if sample_vals else float(len(h_str))
        max_len = max([len(s) for s in sample_vals] or [len(h_str)])

        if is_dotacao:
            # Dotação orçamentária
            if any(k in h_up for k in ['EXERC', 'ANO', 'PODER', 'COD', 'CÓDIGO', 'ITEM', 'FONTE', 'FICHA']) or (
                sample_vals and all(re.search(r'^\d+(\.\d+)*$', s) for s in sample_vals[:10] if s)
            ):
                col_alignments.append('CENTER')
                col_raw_weights.append(1.1)
            elif any(k in h_up for k in ['UNIDADE', 'ORÇAMENTÁRIA', 'ORCAMENTARIA', 'DESCRIÇÃO', 'DESCRICAO', 'ESPEC', 'NOME', 'ÓRGÃO', 'ORGAO', 'SECRETARIA', 'ATIVIDADE', 'PROJETO', 'MANUTEN']):
                col_alignments.append('JUSTIFY')
                col_raw_weights.append(3.6 if avg_len > 30 else 2.8)
            else:
                col_alignments.append('JUSTIFY' if ci == 1 else 'LEFT')
                col_raw_weights.append(max(1.6, min(avg_len / 12.0, 3.2)))
        else:
            # Itens da Contratação / Objeto
            is_item_keyword = any(k in h_up for k in ['ITEM', 'Nº', 'NUM', 'ORDEM', 'POSIÇÃO'])
            is_qtd_keyword = any(k in h_up for k in ['QTD', 'QUANT', 'QUANTIDADE'])
            is_money_keyword = any(k in h_up for k in ['VALOR', 'PREÇO', 'PRECO', 'TOTAL', 'CUSTO', 'R$', 'SUBTOTAL'])
            is_unid_keyword = any(k in h_up for k in ['UNID', 'UND', 'UNIDADE', 'MEDIDA'])

            is_qtd_col = is_qtd_keyword and not is_money_keyword
            is_money_col = is_money_keyword or (
                sample_vals and any('R$' in s or re.search(r'^\d{1,3}(\.\d{3})*,\d{2}$', s) for s in sample_vals[:10])
            )
            is_unid_col = (is_unid_keyword and not is_qtd_col and not is_money_col) or (
                sample_vals and all(len(s) <= 6 for s in sample_vals[:15] if s) and any(k in h_up for k in ['UN', 'UM', 'UND', 'UNID'])
            )
            is_item_col = (
                is_item_keyword and not is_qtd_col and not is_money_col and not is_unid_col
            ) or (
                not is_qtd_keyword and not is_money_keyword and not is_unid_keyword
                and not any(k in h_up for k in ['DESC', 'ESPEC', 'LOTE', 'MARCA', 'DISP', 'SALDO'])
                and sample_vals and all(s.isdigit() and len(s) <= 4 for s in sample_vals[:15] if s)
            )
            is_code_or_brand = any(k in h_up for k in ['LOTE', 'MARCA', 'CATMAT', 'CÓDIGO', 'COD', 'REF', 'MODELO'])
            is_desc_col = any(k in h_up for k in ['DESC', 'ESPEC', 'OBJET', 'PROD', 'MATERIAL', 'SERVIÇO', 'DETALH', 'NOME']) or (
                avg_len >= 30 or max_len >= 50
            )

            if is_money_col:
                col_alignments.append("RIGHT")
                col_raw_weights.append(1.7)
            elif is_item_col:
                col_alignments.append("CENTER")
                col_raw_weights.append(0.8)
            elif is_unid_col:
                col_alignments.append("CENTER")
                col_raw_weights.append(0.85)
            elif is_qtd_col:
                col_alignments.append("CENTER")
                if any(k in h_up for k in ['DISP', 'DISPONIVEL', 'DISPONÍVEL', 'SOLICITADA', 'ESTIMADA', 'CONTRATADA', 'SALDO']) or len(h_str) >= 12:
                    col_raw_weights.append(1.65)
                elif len(h_str) > 8:
                    col_raw_weights.append(1.4)
                else:
                    col_raw_weights.append(1.15)
            elif is_code_or_brand:
                col_alignments.append("CENTER")
                col_raw_weights.append(1.3)
            elif is_desc_col:
                col_alignments.append("JUSTIFY")
                if avg_len > 90 or max_len > 140:
                    col_raw_weights.append(6.8)
                elif avg_len > 45:
                    col_raw_weights.append(5.2)
                else:
                    col_raw_weights.append(4.2)
            else:
                col_alignments.append("LEFT")
                col_raw_weights.append(max(1.4, min(avg_len / 15.0, 3.2)))

    s_weights = sum(col_raw_weights) if sum(col_raw_weights) > 0 else float(n_cols)
    col_widths = [round((w / s_weights) * total_width_pt, 2) for w in col_raw_weights]

    diff = round(total_width_pt - sum(col_widths), 2)
    if col_widths:
        col_widths[-1] = round(col_widths[-1] + diff, 2)

    return col_widths, col_alignments

def renderizar_tabela_html_formatada(headers: list, rows: list, total_geral: float = 0.0, max_height: int = 480, is_dotacao: bool = False, append_total: bool = False, has_real_headers: bool = None) -> str:
    """
    Renderiza tabela HTML responsiva e 100% fidedigna à planilha ou documento importado:
    - Na dotação orçamentária estrutural (is_dotacao=True com Exercício/Poder/Unidade), renderiza sem cabeçalho artificial,
      mesclando o Exercício horizontalmente e a 'Unidade Orçamentária /Atividade' verticalmente sobre as secretarias,
      com bordas azul-marinho oficiais (#224061) e alinhamento centrado.
    - Nas tabelas de itens/objeto, mantém cabeçalho executivo oficial (#002f6e), quebra automática de texto e destaque de resumo.
    """
    if not rows and not headers:
        return ""
    if not rows:
        rows = []
    if not headers and rows:
        headers = [f"COLUNA {i+1}" for i in range(len(rows[0]))]
    if not is_dotacao and rows:
        headers, rows = reconciliar_tabela_irregular(headers, rows, is_dotacao=False)

    # Verifica se é tabela estrutural de dotação oficial
    is_structural, start_unid, end_unid, has_exercicio = (False, -1, -1, False)
    if is_dotacao:
        # Filtra linhas vazias fantasmas onde tanto col 1 quanto col 2 são vazias (exceto linha de Exercício)
        clean_st_rows = []
        for r_i, r in enumerate(rows):
            c1_chk = str(r[1] or '').strip() if len(r) > 1 else ''
            c2_chk = str(r[2] or '').strip() if len(r) > 2 else ''
            if r_i > 0 and not c1_chk and not c2_chk:
                continue
            clean_st_rows.append(r)
        rows = clean_st_rows
        is_structural, start_unid, end_unid, has_exercicio = is_structural_budget_table(headers, rows)

    if is_structural:
        html = [
            f'<div style="max-height: {max_height}px; overflow-y: auto; overflow-x: auto; '
            f'border: 1.5px solid #224061; border-radius: 8px; box-shadow: 0 2px 8px rgba(34,64,97,0.08); margin-bottom: 12px; background: #ffffff;">',
            '<table style="width: 100%; border-collapse: collapse; border-spacing: 0; font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif; font-size: 12.5px; color: #224061; border: 1.5px solid #224061;">',
            '<tbody>'
        ]
        
        has_unid_span = (start_unid != -1 and end_unid >= start_unid)
        rowspan_unid = (end_unid - start_unid + 1) if (start_unid != -1 and end_unid > start_unid) else 1

        for ri, r in enumerate(rows):
            c0_raw = str(r[0] or "").strip()
            c1_raw = str(r[1] or "").strip() if len(r) > 1 else ""
            c2_raw = str(r[2] or "").strip() if len(r) > 2 else ""
            
            c0_html = c0_raw.replace('\n', '<br>')
            c1_html = c1_raw.replace('\n', '<br>')
            c2_html = c2_raw.replace('\n', '<br>')
            
            # Caso 1: Linha de Exercício com mesclagem horizontal nas colunas 1 e 2
            is_exerc_row = (ri == 0 and any(k in c0_raw.upper() for k in ["EXERC", "EXERCÍCIO", "EXERCICIO"]))
            if is_exerc_row:
                ano_val = c1_html if c1_html else c2_html
                html.append(
                    f'<tr style="background-color: #E2EBF4;">'
                    f'<td style="padding: 9px 14px; border: 1.2px solid #224061; text-align: left; vertical-align: middle; font-weight: 700; color: #224061; background-color: #E2EBF4; width: 28%;">'
                    f'{c0_html}</td>'
                    f'<td colspan="2" style="padding: 9px 14px; border: 1.2px solid #224061; text-align: center; vertical-align: middle; font-weight: 700; color: #224061; background-color: #E2EBF4; width: 72%;">'
                    f'{ano_val}</td>'
                    f'</tr>'
                )
            elif has_unid_span and (start_unid <= ri <= end_unid):
                html.append('<tr>')
                if ri == start_unid:
                    u_label = "Unidade Orçamentária<br>/Atividade"
                    rspan_txt = f' rowspan="{rowspan_unid}"' if rowspan_unid > 1 else ''
                    html.append(
                        f'<td{rspan_txt} style="padding: 9px 14px; border: 1.2px solid #224061; text-align: left; vertical-align: middle; font-weight: 700; color: #224061; background-color: #E2EBF4; width: 28%;">'
                        f'{u_label}</td>'
                    )
                bg_col23 = "#ebf3fb" if (ri % 2 == 1) else "#ffffff"
                html.append(
                    f'<td style="padding: 8px 12px; border: 1.2px solid #224061; text-align: justify; text-justify: inter-word; vertical-align: middle; line-height: 1.4; color: #224061; background-color: {bg_col23}; width: 48%; word-break: break-word; overflow-wrap: break-word; word-wrap: break-word; white-space: normal;">'
                    f'{c1_html}</td>'
                    f'<td style="padding: 8px 12px; border: 1.2px solid #224061; text-align: center; vertical-align: middle; line-height: 1.4; color: #224061; background-color: {bg_col23}; width: 24%; word-break: break-word;">'
                    f'{c2_html}</td>'
                    f'</tr>'
                )
            else:
                if not c0_html.strip() and c1_html.strip():
                    c0_html = "Unidade Orçamentária<br>/Atividade"
                bg_col23 = "#ebf3fb" if (ri % 2 == 1) else "#ffffff"
                html.append(
                    f'<tr>'
                    f'<td style="padding: 9px 14px; border: 1.2px solid #224061; text-align: left; vertical-align: middle; font-weight: 700; color: #224061; background-color: #E2EBF4; width: 28%;">'
                    f'{c0_html}</td>'
                    f'<td style="padding: 8px 12px; border: 1.2px solid #224061; text-align: justify; text-justify: inter-word; vertical-align: middle; line-height: 1.4; color: #224061; background-color: {bg_col23}; width: 48%; word-break: break-word; overflow-wrap: break-word; word-wrap: break-word; white-space: normal;">'
                    f'{c1_html}</td>'
                    f'<td style="padding: 8px 12px; border: 1.2px solid #224061; text-align: center; vertical-align: middle; line-height: 1.4; color: #224061; background-color: {bg_col23}; width: 24%; word-break: break-word;">'
                    f'{c2_html}</td>'
                    f'</tr>'
                )
        html.append('</tbody></table></div>')
        return "".join(html)

    num_cols = len(headers)
    col_w_pct, col_align = calcular_larguras_colunas_adaptativas(headers, rows, total_width_pt=100.0, is_dotacao=is_dotacao)
    col_widths_css = [f"{w:.1f}%" for w in col_w_pct]

    border_main = "#224061" if is_dotacao else "#002f6e"
    border_cell = "#224061" if is_dotacao else "#cbd5e1"
    hdr_bg = "#E2EBF4" if is_dotacao else "#002f6e"
    hdr_color = "#224061" if is_dotacao else "#ffffff"

    if has_real_headers is not None:
        has_real_hdrs = bool(has_real_headers)
    else:
        has_real_hdrs = headers and not any(str(h).strip().startswith("COLUNA ") for h in headers) and not all(str(h).strip().startswith("COLUNA ") for h in headers)

    html = [
        f'<div style="max-height: {max_height}px; overflow-y: auto; overflow-x: auto; '
        f'border: 1.5px solid {border_main}; border-radius: 8px; box-shadow: 0 2px 8px rgba(34,64,97,0.08); margin-bottom: 12px; background: #ffffff;">',
        f'<table style="width: 100%; border-collapse: collapse; border-spacing: 0; table-layout: fixed; font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif; font-size: 12.5px; color: #1e293b; border: 1.5px solid {border_main};">'
    ]

    if has_real_hdrs:
        html.append('<thead>')
        html.append(f'<tr style="background-color: {hdr_bg}; color: {hdr_color}; position: sticky; top: 0; z-index: 2;">')
        for ci, h in enumerate(headers):
            w_style = f"width: {col_widths_css[ci]};" if ci < len(col_widths_css) else ""
            h_disp = "" if (str(h).strip().startswith("COLUNA ") or str(h).strip().startswith("_")) else str(h).strip()
            h_disp_html = h_disp.replace('\r\n', '<br>').replace('\n', '<br>').replace('<br/>', '<br>')
            html.append(
                f'<th style="{w_style} padding: 8px 6px; border: 1.2px solid {border_main}; text-align: center; font-weight: 700; letter-spacing: 0.2px; text-transform: uppercase; font-size: 11.5px; line-height: 1.35; color: {hdr_color}; background-color: {hdr_bg}; word-break: break-word; overflow-wrap: break-word; word-wrap: break-word; white-space: normal;">'
                f'{h_disp_html}</th>'
            )
        html.append('</tr></thead>')
    html.append('<tbody>')

    has_summary_in_rows = any(is_summary_row(r) for r in rows)

    for ri, r in enumerate(rows):
        is_res = is_summary_row(r)
        if is_dotacao:
            row_bg = "#ffffff"
        else:
            row_bg = "#f1f5f9" if is_res else ("#ebf3fb" if (ri % 2 == 1) else "#ffffff")
        html.append(f'<tr style="background-color: {row_bg}; transition: background 0.15s ease;">')
        
        # Detecta se é linha de cabeçalho de seção / banner (apenas col 0 preenchida)
        is_banner = (len(r) > 1 and str(r[0] or '').strip() and not any(str(x or '').strip() for x in r[1:]))
        if is_banner:
            banner_txt = str(r[0] or '').replace('\n', '<br>')
            html.append(
                f'<td colspan="{len(headers)}" style="padding: 8px 12px; border: 1.2px solid {border_cell}; text-align: center; '
                f'background-color: #E2EBF4; font-weight: 700; color: #224061; font-size: 12px; letter-spacing: 0.3px;">'
                f'{banner_txt}</td></tr>'
            )
            continue

        has_colspan_exerc = is_dotacao and len(headers) == 3 and str(r[0]).strip().upper().startswith("EXERC") and (len(r) < 3 or not str(r[2]).strip())
        for ci, h in enumerate(headers):
            if has_colspan_exerc and ci == 2:
                continue
            raw_val = str(r[ci]) if ci < len(r) else ""
            align = col_align[ci] if ci < len(col_align) else 'left'
            if is_res and any(k in str(h).upper() for k in ['VALOR', 'PREÇO', 'PRECO', 'TOTAL', 'R$']):
                align = 'right'
            is_item_col = (ci == 0 and any(k in str(h).upper() for k in ['ITEM', 'Nº']))
            extra_bold = "font-weight: 700; color: #002f6e;" if is_res else ("font-weight: 600;" if is_item_col else "")
            cs_attr = ' colspan="2"' if (has_colspan_exerc and ci == 1) else ''
            if is_dotacao:
                if ci == 0 or has_colspan_exerc:
                    cell_bg = "#E2EBF4"
                else:
                    cell_bg = "#ebf3fb" if (ri % 2 == 1) else "#ffffff"
            else:
                cell_bg = "#f1f5f9" if is_res else ("#ebf3fb" if (ri % 2 == 1) else "#ffffff")
            cell_color = "#224061" if (is_dotacao or is_res) else "#1e293b"

            is_desc_c = (
                align.upper() == 'JUSTIFY' or 
                (ci == 1 and not is_res) or 
                (is_dotacao and any(k in str(h).upper() for k in ['DESC', 'UNID', 'ORÇAM', 'ORCAM', 'ATIVID', 'PROJET', 'NOME', 'ÓRGÃO', 'ORGAO', 'SECRETARIA', 'MANUTEN']))
            )
            if is_desc_c:
                clean_text = normalizar_quebras_texto_celula(raw_val)
                val_disp = clean_text.replace('\n', '<br>')
                align_style = "text-align: justify; text-justify: inter-word; hyphens: auto; word-break: normal; overflow-wrap: break-word; white-space: normal;"
            else:
                val_disp = raw_val.replace('\n', '<br>')
                align_style = f"text-align: {align.lower()}; word-break: break-word; overflow-wrap: anywhere; white-space: normal;"

            if is_dotacao and ci == 0 and not val_disp.strip() and (len(r) > 1 and str(r[1] or '').strip()):
                val_disp = "Unidade Orçamentária / Atividade"
                extra_bold = "font-weight: 700; color: #224061;"
                cell_bg = "#E2EBF4"

            html.append(
                f'<td{cs_attr} style="padding: 8px 12px; border: 1.2px solid {border_cell}; {align_style} '
                f'vertical-align: middle; line-height: 1.4; color: {cell_color}; background-color: {cell_bg}; {extra_bold}">'
                f'{val_disp}</td>'
            )
        html.append('</tr>')

    # Totalizador adicional APENAS se explicitamente solicitado (append_total=True), se total_geral > 0 e se a tabela NÃO possuir linha de resumo própria
    if append_total and not is_dotacao and not has_summary_in_rows and total_geral > 0:
        tot_fmt = f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        tot_col_idx = len(headers) - 1
        for i, h in enumerate(headers):
            h_up = str(h).upper()
            if any(k in h_up for k in ['TOTAL', 'VALOR TOTAL', 'PREÇO TOTAL', 'VALOR GLOBAL']):
                tot_col_idx = i
                break

        html.append('<tr style="background-color: #f1f5f9; font-weight: 700; border-top: 2px solid #002f6e; position: sticky; bottom: 0; z-index: 1;">')
        if tot_col_idx > 0:
            html.append(
                f'<td colspan="{tot_col_idx}" style="padding: 9px 12px; border: 1px solid #cbd5e1; text-align: right; color: #002f6e; font-size: 12px; font-weight: 700;">'
                f'VALOR TOTAL:</td>'
            )
        html.append(
            f'<td style="padding: 9px 12px; border: 1px solid #cbd5e1; text-align: right; color: #002f6e; font-size: 13px; font-weight: 800;">'
            f'{tot_fmt}</td>'
        )
        if tot_col_idx < len(headers) - 1:
            sobra = len(headers) - 1 - tot_col_idx
            html.append(f'<td colspan="{sobra}" style="padding: 9px 12px; border: 1px solid #cbd5e1;"></td>')
        html.append('</tr>')

    html.append('</tbody></table></div>')
    return ''.join(html)

def aplicar_estilo_tabela_oficial_docx(t_it, col_widths_pt=None):
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn, nsdecls
    from docx.shared import Pt
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    
    try:
        t_it.style = 'Table Grid'
    except Exception:
        pass
    t_it.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_it.autofit = False

    # Bordas elegantes em todas as células (Azul escuro do modelo: #224061)
    tblPr = t_it._tbl.tblPr
    for old_b in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(old_b)

    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="10" w:space="0" w:color="224061"/>\n'
        f'  <w:bottom w:val="single" w:sz="10" w:space="0" w:color="224061"/>\n'
        f'  <w:left w:val="single" w:sz="10" w:space="0" w:color="224061"/>\n'
        f'  <w:right w:val="single" w:sz="10" w:space="0" w:color="224061"/>\n'
        f'  <w:insideH w:val="single" w:sz="8" w:space="0" w:color="224061"/>\n'
        f'  <w:insideV w:val="single" w:sz="8" w:space="0" w:color="224061"/>\n'
        f'</w:tblBorders>'
    )
    tblLook = tblPr.find(qn('w:tblLook'))
    if tblLook is not None:
        tblPr.insert(tblPr.index(tblLook), tblBorders)
    else:
        tblPr.append(tblBorders)

    # Margens internas da tabela
    for old_m in tblPr.findall(qn('w:tblCellMar')):
        tblPr.remove(old_m)
    cell_mar = parse_xml(
        f'<w:tblCellMar {nsdecls("w")}>\n'
        f'  <w:top w:w="80" w:type="dxa"/>\n'
        f'  <w:bottom w:w="80" w:type="dxa"/>\n'
        f'  <w:left w:w="120" w:type="dxa"/>\n'
        f'  <w:right w:w="120" w:type="dxa"/>\n'
        f'</w:tblCellMar>'
    )
    tblLook = tblPr.find(qn('w:tblLook'))
    if tblLook is not None:
        tblPr.insert(tblPr.index(tblLook), cell_mar)
    else:
        tblPr.append(cell_mar)

    if len(t_it.rows) > 0:
        t_it.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    num_cols = len(t_it.columns)
    if not col_widths_pt or len(col_widths_pt) != num_cols:
        col_widths_pt = [485.0 / max(num_cols, 1)] * num_cols
    else:
        s_pts = sum(col_widths_pt)
        if s_pts > 0:
            col_widths_pt = [(w / s_pts) * 485.0 for w in col_widths_pt]

    # Fixa w:tblW com a largura total em dxa para garantir quebra de texto automatica e fidelidade das colunas
    for old_w in tblPr.findall(qn('w:tblW')):
        tblPr.remove(old_w)
    total_w_dxa = int(round(sum(col_widths_pt) * 20))
    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_w_dxa}" w:type="dxa"/>'))

    # Fixa w:tblGrid com as larguras em dxa (1 pt = 20 dxa) para garantir fidelidade no Word
    for old_g in tblPr.findall(qn('w:tblGrid')):
        tblPr.remove(old_g)
    tblGrid = parse_xml(f'<w:tblGrid {nsdecls("w")}/>')
    for w in col_widths_pt:
        w_dxa = int(round(w * 20))
        tblGrid.append(parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{w_dxa}"/>'))
    tblPr.append(tblGrid)

    for row in t_it.rows:
        row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for ci, cell in enumerate(row.cells):
            if ci < len(col_widths_pt):
                w_pt = col_widths_pt[ci]
                w_dxa = int(round(w_pt * 20))
                cell.width = Pt(w_pt)
                tcPr = cell._tc.get_or_add_tcPr()
                for old_w in tcPr.findall(qn('w:tcW')):
                    tcPr.remove(old_w)
                tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w_dxa}" w:type="dxa"/>'))
                for old_nw in tcPr.findall(qn('w:noWrap')):
                    tcPr.remove(old_nw)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            aplicar_bordas_celula(cell)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(2.5)
                p.paragraph_format.space_after = Pt(2.5)
                p.paragraph_format.line_spacing = 1.15


def aplicar_bordas_celula(cell, border_color="224061", border_sz="8"):
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn, nsdecls
    tcPr = cell._tc.get_or_add_tcPr()
    for old_b in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(old_b)
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>\n'
        f'  <w:left w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>\n'
        f'  <w:bottom w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>\n'
        f'  <w:right w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>\n'
        f'</w:tcBorders>'
    )
    shd = tcPr.find(qn('w:shd'))
    if shd is not None:
        tcPr.insert(tcPr.index(shd), tcBorders)
    else:
        tcPr.append(tcBorders)

def aplicar_fundo_celula_docx(cell, color_hex):
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn, nsdecls
    tcPr = cell._tc.get_or_add_tcPr()
    for old_s in tcPr.findall(qn('w:shd')):
        tcPr.remove(old_s)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{color_hex}"/>')
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
    return gerar_planilha_itens_excel_formatada(headers, rows, total_geral=21905.25, append_total_row=True)

class ResultadoPlanilhaItens(list):
    """Lista de itens para total retrocompatibilidade, enriquecida com metadados da tabela dinâmica."""
    def __init__(self, itens, tabela_dinamica=None):
        super().__init__(itens)
        self.tabela_dinamica = tabela_dinamica or {}

    @property
    def itens(self):
        return list(self)

def normalizar_quebras_texto_celula(text_val: str) -> str:
    """
    Remove quebras de linha artificiais geradas por quebra de coluna estreita (ex: PDF ou quebras manuais),
    unificando o texto em fluxo continuo com espacos para permitir a quebra automatica natural
    (word wrap) e o alinhamento justificado preenchendo toda a largura da coluna.
    Preserva listas com marcadores se houver (ex: linhas iniciadas por '-', '*', '1.', 'a)').
    """
    if not text_val:
        return ""
    s = str(text_val).replace("\r", "")
    lines = [l.strip() for l in s.split("\n")]
    lines = [l for l in lines if l]
    if not lines:
        return ""
    if len(lines) == 1:
        return lines[0]
    is_list = all(re.match(r'^([-•*]|\d+[\.\)]|[a-zA-Z][\.\)])\s+', l) for l in lines)
    if is_list:
        return "\n".join(lines)
    avg_line_len = sum(len(l) for l in lines) / len(lines)
    if avg_line_len < 55 or not any(l.endswith(('.', ';', ':')) for l in lines[:-1]):
        return " ".join(lines)
    clean_parts = []
    curr_part = []
    for l in lines:
        curr_part.append(l)
        if l.endswith(('.', ';', ':')):
            clean_parts.append(" ".join(curr_part))
            curr_part = []
    if curr_part:
        clean_parts.append(" ".join(curr_part))
    return "\n".join(clean_parts)

def parse_num_pt(val) -> float:
    if val is None or (isinstance(val, float) and str(val) == 'nan'):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'none', '-', '', 'null'):
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

def format_moeda_pt(val: float) -> str:
    try:
        return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"

def reconciliar_tabela_irregular(headers, rows, is_dotacao: bool = False):
    """
    Reconcilia e normaliza tabelas com fidelidade absoluta (1:1) a planilha original:
    1. Elimina colunas fantasmas geradas por celulas mescladas no Word/Excel/PDF (ex: 'COLUNA 2', 'COLUNA 4').
    2. Reincorpora linhas de cabecalho fragmentadas ('MATERIAL OU', 'SERVICO') nos cabecalhos autenticos.
    3. Reconcilia o desalinhamento horizontal causado por celulas mescladas na descricao.
    4. Anexa linhas orfas de quebra de texto na descricao do item precedente.
    5. Calcula o valor total quando omitido (Qtd * VU) mantendo a integridade matematica.
    6. Preserva 100% dos caracteres, totais e formatacao sem inventar colunas ou omitir informacoes.
    """
    if is_dotacao or not rows or not headers:
        return headers, rows

    hdr_txt = ' '.join([str(h) for h in headers]).upper()
    has_item_kw = any(k in hdr_txt for k in ['ITEM', 'DESCRIÇÃO', 'ESPECIFICAÇÃO', 'ESPEC', 'PRODUTO', 'SERVIÇO', 'QUANTIDADE', 'VALOR', 'PREÇO', 'UNIDADE', 'QTD'])
    numeric_item_rows = sum(1 for r in rows[:15] if [c for c in r if str(c).strip()] and re.match(r'^\d+$', [c for c in r if str(c).strip()][0]))
    
    if not (has_item_kw or numeric_item_rows >= 2):
        return headers, rows

    has_phantom_cols = any(re.match(r'^(COLUNA\s+\d+|Unnamed:\s*\d+)$', str(h).strip(), re.I) for h in headers)

    # Detecta se as primeiras linhas de rows sao continuacao de cabecalho
    header_continuation_texts = []
    data_start_idx = 0
    for idx, r in enumerate(rows[:6]):
        filled = [str(c).strip() for c in r if str(c).strip()]
        if not filled:
            continue
        if re.match(r'^\d+$', filled[0]):
            data_start_idx = idx
            break
        row_txt = ' '.join(filled).upper()
        if any(w in row_txt for w in ['MATERIAL', 'SERVIÇO', 'FORNECIMENTO', 'ESPECIFICAÇÃO', 'DESCRIÇÃO', 'UNIDADE', 'QUANTIDADE', 'PREÇO', 'VALOR', 'DISP', 'UNIT', 'TOTAL', 'MEDIDA']) or (len(filled) <= 2 and not any(re.search(r'\d+[.,]\d{2}', c) for c in filled)):
            header_continuation_texts.extend(filled)
            data_start_idx = idx + 1
        else:
            data_start_idx = idx
            break

    orphan_exists = any(not re.match(r'^\d+$', [c for c in r if str(c).strip()][0]) and not any('TOTAL' in str(c).upper() for c in r) for r in rows[data_start_idx:] if [c for c in r if str(c).strip()])
    
    if not has_phantom_cols and not header_continuation_texts and data_start_idx == 0 and not orphan_exists:
        return headers, rows

    # Reconstroi os cabecalhos autenticos
    clean_headers = [str(h).strip() for h in headers if not re.match(r'^(COLUNA\s+\d+|Unnamed:\s*\d+)$', str(h).strip(), re.I)]
    if not clean_headers:
        clean_headers = [str(h).strip() for h in headers]

    # Incorpora continuacoes de texto do cabecalho
    if header_continuation_texts:
        for ci, ch in enumerate(clean_headers):
            if any(k in ch.upper() for k in ['ESPECIFICAÇÃO', 'DESCRIÇÃO', 'ESPEC', 'OBJETO', 'MATERIAL']):
                extras = [t for t in header_continuation_texts if t.upper() not in ch.upper()]
                if extras:
                    clean_headers[ci] = (ch + ' ' + ' '.join(extras)).strip()
                    break

    num_target_cols = len(clean_headers)
    data_rows = rows[data_start_idx:]
    reconciled_rows = []
    total_row = None

    for r in data_rows:
        filled = [str(c).strip() for c in r if str(c).strip()]
        if not filled:
            continue
        row_txt = ' '.join(filled).upper()
        if any(k in row_txt for k in ['TOTAL GERAL', 'TOTAL:', 'TOTAL DO OBJETO', 'VALOR TOTAL']) or (len(filled) <= 3 and 'TOTAL' in row_txt):
            tot_val = ''
            for cell in reversed(filled):
                if any(k in cell for k in ['R$', '$', ',']) and re.search(r'\d', cell):
                    tot_val = cell
                    break
            tot_label = 'TOTAL'
            for cell in filled:
                if 'TOTAL' in cell.upper():
                    tot_label = cell
                    break
            total_row = [''] * num_target_cols
            if num_target_cols >= 2:
                total_row[1] = tot_label
                total_row[-1] = tot_val
            else:
                total_row = [tot_label, tot_val]
            continue

        if re.match(r'^\d+$', filled[0]):
            it_num = filled[0]
            if len(filled) == num_target_cols:
                row_aligned = list(filled)
            elif len(filled) == num_target_cols - 1:
                row_aligned = list(filled) + ['']
                q_val = parse_num_pt(filled[-2])
                vu_val = parse_num_pt(filled[-1])
                if q_val > 0 and vu_val > 0:
                    vt_calc = q_val * vu_val
                    row_aligned[-1] = format_moeda_pt(vt_calc)
            elif len(filled) < num_target_cols:
                row_aligned = list(filled) + [''] * (num_target_cols - len(filled))
            else:
                desc_idx = 1 if num_target_cols > 1 else 0
                extra_cells = len(filled) - num_target_cols
                combined_desc = ' '.join(filled[desc_idx : desc_idx + extra_cells + 1])
                row_aligned = [filled[0], combined_desc] + filled[desc_idx + extra_cells + 1 :]
            reconciled_rows.append(row_aligned)
        else:
            if reconciled_rows:
                desc_idx = 1 if len(reconciled_rows[-1]) > 1 else 0
                extra_text = ' '.join(filled)
                reconciled_rows[-1][desc_idx] = (reconciled_rows[-1][desc_idx] + ' ' + extra_text).strip()

    if total_row:
        reconciled_rows.append(total_row)

    if not reconciled_rows:
        return headers, rows

    return clean_headers, reconciled_rows

def converter_tabela_reconciliada_para_itens(headers: list, rows: list) -> list:
    """Converte tabela reconciliada em lista de dicionarios estruturados de itens."""
    import re
    col_item, col_desc, col_unid, col_qtd, col_vu, col_vt = 0, 1, 2, 3, 4, 5
    descartes_vt = ['QTD', 'QUANT', 'QUANTIDADE', 'DIAS', 'HORAS', 'PARCELA', 'MESES', 'PAGAMENTO']

    for ci, h in enumerate(headers):
        c_upper = str(h).strip().upper()
        if any(k in c_upper for k in ['PREÇO UNID', 'PRECO UNID', 'VALOR UNID', 'P. UNID', 'V. UNID', 'VL. UNID', 'UNIT', 'UNITARIO', 'V. UNIT', 'VL. UNIT', 'PREÇO UNIT', 'PRECO UNIT', 'VALOR UNITÁRIO', 'VALOR UNIT']):
            col_vu = ci
        elif any(k in c_upper for k in ['VALOR TOTAL', 'V. TOTAL', 'VL. TOTAL', 'PREÇO TOTAL', 'PRECO TOTAL', 'VALOR GLOBAL', 'TOTAL (R$)', 'TOTAL R$', 'SUBTOTAL']):
            if not any(d in c_upper for d in descartes_vt):
                col_vt = ci
        elif 'TOTAL' in c_upper and not any(d in c_upper for d in descartes_vt) and col_vt is None:
            col_vt = ci
        elif any(k in c_upper for k in ['QTD', 'QUANT', 'QUANTIDADE']) and not any(k in c_upper for k in ['VALOR', 'PREÇO', 'PRECO']):
            col_qtd = ci
        elif any(k in c_upper for k in ['UNIDADE', 'UNID', 'UND', 'MEDIDA']) and not any(k in c_upper for k in ['PREÇO', 'PRECO', 'VALOR', 'UNIT', 'CUSTO', 'P.', 'V.']):
            col_unid = ci
        elif any(k in c_upper for k in ['DESC', 'PROD', 'ESPEC', 'SERV', 'OBJET', 'MATERIAL', 'DISCRIMINAÇÃO']):
            col_desc = ci
        elif any(k in c_upper for k in ['ITEM', 'Nº', 'NUMERO', 'POSI']):
            col_item = ci

    def _limpar_num_local(val):
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
        s = re.sub(r'[R$\s BRL]', '', s)
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

    itens = []
    for r in rows:
        if is_summary_row(r) or not r[col_item] or not re.match(r'^\d+$', str(r[col_item]).strip()):
            continue
        try:
            it_num = int(r[col_item])
        except Exception:
            continue
        desc = str(r[col_desc]).strip() if col_desc < len(r) else ''
        unid = str(r[col_unid]).strip() if col_unid < len(r) else ''
        qtd_s = str(r[col_qtd]).strip() if col_qtd < len(r) else '1'
        vu_s = str(r[col_vu]).strip() if col_vu < len(r) else ''
        vt_s = str(r[col_vt]).strip() if col_vt < len(r) else ''
        q_f = _limpar_num_local(qtd_s)
        vu_f = _limpar_num_local(vu_s)
        vt_f = _limpar_num_local(vt_s)
        if vt_f == 0.0 and q_f > 0 and vu_f > 0:
            vt_f = q_f * vu_f
            vt_s = formatar_moeda_br(vt_f)

        itens.append({
            'item': it_num,
            'descricao': desc,
            'quantidade': str(int(q_f)) if q_f == int(q_f) else f'{q_f:.2f}'.replace('.', ','),
            'unidade': unid,
            'valor_unitario': formatar_moeda_br(vu_f) if vu_f > 0 else vu_s,
            'valor_total': formatar_moeda_br(vt_f) if vt_f > 0 else vt_s,
            '_vu_float': vu_f,
            '_vt_float': vt_f,
            '_qtd_float': q_f if q_f > 0 else 1.0
        })
    return itens

def extrair_tabela_dinamica_df(df, is_dotacao: bool = False) -> dict:
    """
    Extrai e normaliza dinamicamente a estrutura tabular a partir de um DataFrame:
    - Preserva 100% os cabeçalhos originais e o modelo de colunas subido pelo usuário.
    - Suporta qualquer quantidade de colunas (2, 3, 4, 5, 6, 7+) e qualquer modelo de tabela.
    - Remove colunas e linhas 100% vazias sem inventar colunas artificiais.
    - Se is_dotacao=True: preserva 100% os caracteres, códigos orçamentários, linhas e grades.
    - Se tabela de itens: preserva 100% das linhas, caracteres e formato da planilha importada.
    """
    import pandas as pd
    import re
    try:
        if df is None or df.empty:
            return {}

        # Remove colunas 100% vazias (onde todos os valores são vazios, espaços ou nan)
        def _col_empty(s):
            return all(x is None or str(x).strip() == '' or str(x).strip().lower() in ('nan', 'none') for x in s)
        empty_c = [c for c in df.columns if _col_empty(df[c])]
        if empty_c and len(empty_c) < len(df.columns):
            df = df.drop(columns=empty_c)

        # Remove linhas 100% vazias
        def _row_empty(r):
            return all(x is None or str(x).strip() == '' or str(x).strip().lower() in ('nan', 'none') for x in r)
        empty_r = [idx for idx, r in df.iterrows() if _row_empty(r)]
        if empty_r:
            df = df.drop(index=empty_r)

        if df.empty:
            return {}
        if df.empty:
            return {}

        has_real_hdrs_attr = df.attrs.get('has_real_headers')
        raw_cols = [str(c).strip() for c in df.columns]

        if has_real_hdrs_attr is False:
            has_real_hdrs = False
            clean_hdr = [str(c).strip() if (str(c).strip() and not str(c).startswith('Unnamed:')) else f"COLUNA {ci+1}" for ci, c in enumerate(raw_cols)]
            df.columns = clean_hdr
        elif has_real_hdrs_attr is True:
            has_real_hdrs = True
            seen = {}
            final_hdr = []
            for ci, h in enumerate(raw_cols):
                h_name = h if (h and not h.startswith('Unnamed:') and h.lower() != 'nan') else f"COLUNA {ci+1}"
                if h_name in seen:
                    seen[h_name] += 1
                    final_hdr.append(f"{h_name} ({seen[h_name]})")
                else:
                    seen[h_name] = 1
                    final_hdr.append(h_name)
            df.columns = final_hdr
        else:
            is_all_unnamed = all(not c or c.startswith('Unnamed:') or c.isdigit() for c in raw_cols)
            if is_all_unnamed:
                clean_hdr = [f"COLUNA {ci+1}" for ci in range(len(df.columns))]
                df.columns = clean_hdr
                has_real_hdrs = False
            else:
                has_real_hdrs = not all(str(h).strip().startswith("COLUNA ") for h in raw_cols)
                seen = {}
                final_hdr = []
                for ci, h in enumerate(raw_cols):
                    h_name = h if (h and not h.startswith('Unnamed:') and h.lower() != 'nan') else f"COLUNA {ci+1}"
                    if h_name in seen:
                        seen[h_name] += 1
                        final_hdr.append(f"{h_name} ({seen[h_name]})")
                    else:
                        seen[h_name] = 1
                        final_hdr.append(h_name)
                df.columns = final_hdr

        headers = [str(c).strip() for c in df.columns]
        num_cols = len(headers)

        rows = []
        for _, r in df.iterrows():
            row_vals = []
            for c in df.columns:
                val = r[c]
                if pd.isna(val) or val is None:
                    row_vals.append("")
                elif is_dotacao:
                    # Preservação 100% fidedigna de caracteres, códigos orçamentários e textos
                    s_val = re.sub(r'<[^>]+>', '', str(val)).strip()
                    if isinstance(val, float) and val.is_integer():
                        row_vals.append(str(int(val)))
                    else:
                        row_vals.append(s_val)
                elif isinstance(val, str):
                    row_vals.append(val.strip())
                elif isinstance(val, float):
                    if val.is_integer():
                        row_vals.append(str(int(val)))
                    else:
                        s_f = f"{val:.4f}".rstrip("0").rstrip(".")
                        row_vals.append(s_f.replace(".", ","))
                elif isinstance(val, int):
                    row_vals.append(str(val))
                else:
                    row_vals.append(str(val).strip())
            # Descarta linhas se todas as células forem vazias
            if any(str(x).strip() for x in row_vals):
                rows.append(row_vals)

        if is_dotacao:
            # Dotação orçamentária NÃO tem valor financeiro, preserva 100% todas as linhas
            clean_dot_rows = []
            for r_i, r in enumerate(rows):
                if len(r) >= 3:
                    c1_s = str(r[1] or '').strip()
                    c2_s = str(r[2] or '').strip()
                    if r_i > 0 and not c1_s and not c2_s:
                        continue
                clean_dot_rows.append(r)
            
            for idx_r, r in enumerate(clean_dot_rows):
                if len(r) >= 3:
                    c0_s = str(r[0] or '').strip()
                    c2_clean = re.sub(r'\s+', '', str(r[2] or ''))
                    if idx_r > 2 and re.match(r'^\d{2}\.\d{2}\.\d{2}\.\d{2}$', c2_clean):
                        r[0] = "Natureza da despesa"
                    elif not c0_s and (any(k in str(r[1] or '').upper() for k in ['MANUTEN', 'SECRETARIA', 'FUNDO', 'GABINETE', 'SEC.', 'SERVIÇO', 'AÇÕES', 'GESTÃO', 'ADMINISTRAÇÃO']) or re.search(r'\d{2}\.\d{3}', c2_clean)):
                        r[0] = "Unidade Orçamentária\n/Atividade"
                    elif idx_r == 3 and "Natureza" in c0_s:
                        r[0] = "Unidade Orçamentária\n/Atividade"

            item_rows = clean_dot_rows
            total_geral = 0.0
            total_col_idx = None
            total_formatado = ""

            col_alignments = []
            col_weights = []
            for i, h in enumerate(headers):
                h_up = h.upper()
                if any(k in h_up for k in ['EXERC', 'ANO', 'PODER', 'COD', 'CÓDIGO', 'ITEM', 'FONTE', 'FICHA']):
                    col_alignments.append('CENTER')
                else:
                    col_alignments.append('JUSTIFY')
                max_len = max(len(str(h)), max([len(str(r[i])) for r in item_rows if i < len(r)] or [0]))
                col_weights.append(max(float(max_len), 4.0))

            LARGURA_UTIL_PDF = 495.0
            soma_pesos = sum(col_weights) if sum(col_weights) > 0 else float(num_cols)
            col_widths = [round((w / soma_pesos) * LARGURA_UTIL_PDF, 2) for w in col_weights]
            diff = round(LARGURA_UTIL_PDF - sum(col_widths), 2)
            if col_widths:
                col_widths[-1] += diff

            has_real_hdrs = df.attrs.get('has_real_headers')
            if has_real_hdrs is None:
                has_real_hdrs = headers and not all(str(h).strip().startswith("COLUNA ") for h in headers)

            return {
                "headers": headers,
                "rows": item_rows,
                "col_alignments": col_alignments,
                "col_widths": col_widths,
                "total_col_idx": None,
                "total_geral": 0.0,
                "total_formatado": "",
                "qtd_colunas": num_cols,
                "qtd_linhas": len(item_rows),
                "is_dotacao": True,
                "has_real_headers": has_real_hdrs
            }

        if not is_dotacao:
            headers, rows = reconciliar_tabela_irregular(headers, rows, is_dotacao=False)
            num_cols = len(headers)

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

        has_summary_row = False
        summary_row_idx = None
        resumo_row_val = None

        for idx_r, r in enumerate(rows):
            if is_summary_row(r):
                has_summary_row = True
                summary_row_idx = idx_r
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

        # Linhas de itens para cálculo do somatório (não soma a linha de total para não duplicar o valor)
        rows_to_sum = [r for idx_r, r in enumerate(rows) if idx_r != summary_row_idx] if (has_summary_row and summary_row_idx is not None) else rows

        total_geral = 0.0
        if total_col_idx is not None:
            for r in rows_to_sum:
                if total_col_idx < len(r):
                    val_txt = r[total_col_idx]
                    total_geral += _parse_num_local(val_txt)
        elif qtd_col_idx is not None and vu_col_idx is not None:
            for r in rows_to_sum:
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
                for r in rows_to_sum:
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

        col_widths_pdf, col_alignments = calcular_larguras_colunas_adaptativas(headers, rows, total_width_pt=515.0, is_dotacao=False)
        col_widths_docx, _ = calcular_larguras_colunas_adaptativas(headers, rows, total_width_pt=485.0, is_dotacao=False)

        has_real_hdrs = df.attrs.get('has_real_headers')
        if has_real_hdrs is None:
            has_real_hdrs = headers and not all(str(h).strip().startswith("COLUNA ") for h in headers)

        return {
            "headers": headers,
            "rows": rows,  # Preserva 100% de fidelidade de todas as linhas da planilha modelo subida
            "col_alignments": col_alignments,
            "col_widths": col_widths_pdf,
            "col_widths_docx": col_widths_docx,
            "total_col_idx": total_col_idx,
            "total_geral": total_geral,
            "total_formatado": formatar_moeda_br(total_geral),
            "qtd_colunas": num_cols,
            "qtd_linhas": len(rows),
            "has_summary_row": has_summary_row,
            "summary_row_idx": summary_row_idx,
            "has_real_headers": has_real_hdrs
        }
    except Exception as e:
        print("Erro ao extrair tabela dinâmica do DataFrame:", e)
        return {}

def parse_numero_br(val) -> float:
    """Converte com segurança qualquer valor (int, float, string monetária com R$, pontos e vírgulas) para float."""
    import re
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

def conferir_valores_planilha_itens(tab_din_ou_itens, valor_contrato_atual: float = 0.0) -> dict:
    """
    Realiza a conferência de valores, reconciliação matemática e auditoria profunda dos itens da contratação:
    - Identifica automaticamente colunas de Item, Descrição, Quantidade, Valor Unitário e Valor Total.
    - Calcula o subtotal individual de cada linha (Qtd * VU) e detecta discrepâncias no cálculo de cada item.
    - Calcula a soma global de todos os itens da planilha.
    - Compara a soma da planilha com o valor cadastrado no contrato (campo 8), calculando a diferença em centavos.
    - Retorna status consolidado ('CONFERE' ou 'DIVERGENTE'), métricas de auditoria e memória de cálculo linha a linha.
    """
    if not tab_din_ou_itens:
        return {
            "tem_planilha": False,
            "confere": True,
            "status": "SEM_PLANILHA",
            "soma_itens": 0.0,
            "soma_formatada": "R$ 0,00",
            "valor_contrato": float(valor_contrato_atual or 0.0),
            "valor_contrato_formatado": formatar_moeda_br(valor_contrato_atual or 0.0),
            "diferenca": 0.0,
            "diferenca_formatada": "R$ 0,00",
            "total_itens": 0,
            "itens_com_valor": 0,
            "itens_zerados": 0,
            "linhas_divergentes_qtd": 0,
            "memoria_calculo": [],
            "mensagem": "Nenhuma planilha do objeto anexada. O valor do contrato é manual."
        }

    headers = []
    rows = []
    total_geral_pre = 0.0

    if isinstance(tab_din_ou_itens, dict):
        headers = tab_din_ou_itens.get("headers", [])
        rows = tab_din_ou_itens.get("rows", [])
        total_geral_pre = float(tab_din_ou_itens.get("total_geral", 0.0) or 0.0)
    elif hasattr(tab_din_ou_itens, "tabela_dinamica") and getattr(tab_din_ou_itens, "tabela_dinamica"):
        td = tab_din_ou_itens.tabela_dinamica
        headers = td.get("headers", [])
        rows = td.get("rows", [])
        total_geral_pre = float(td.get("total_geral", 0.0) or 0.0)
    elif isinstance(tab_din_ou_itens, list):
        if tab_din_ou_itens and isinstance(tab_din_ou_itens[0], dict):
            headers = ["Item", "Descrição", "Qtd", "Unid", "Valor Unitário", "Valor Total"]
            rows = [
                [
                    str(it.get("item", idx + 1)),
                    str(it.get("descricao", "")),
                    str(it.get("quantidade", "1")),
                    str(it.get("unidade", "Unid.")),
                    str(it.get("valor_unitario", "")),
                    str(it.get("valor_total", ""))
                ]
                for idx, it in enumerate(tab_din_ou_itens)
            ]
            total_geral_pre = sum(float(it.get("_vt_float", 0.0) or 0.0) for it in tab_din_ou_itens)
        else:
            rows = tab_din_ou_itens

    if not rows:
        return {
            "tem_planilha": False,
            "confere": True,
            "status": "SEM_PLANILHA",
            "soma_itens": 0.0,
            "soma_formatada": "R$ 0,00",
            "valor_contrato": float(valor_contrato_atual or 0.0),
            "valor_contrato_formatado": formatar_moeda_br(valor_contrato_atual or 0.0),
            "diferenca": 0.0,
            "diferenca_formatada": "R$ 0,00",
            "total_itens": 0,
            "itens_com_valor": 0,
            "itens_zerados": 0,
            "linhas_divergentes_qtd": 0,
            "memoria_calculo": [],
            "mensagem": "Planilha sem linhas de dados para conferência."
        }

    col_item = None
    col_desc = None
    col_qtd = None
    col_unid = None
    col_vu = None
    col_vt = None

    descartes_vt = ['QTD', 'QUANT', 'QUANTIDADE', 'DIAS', 'HORAS', 'PARCELA', 'MESES', 'PAGAMENTO']
    for ci, h in enumerate(headers):
        h_up = str(h).upper().strip()
        if any(k in h_up for k in ['VALOR TOTAL', 'V. TOTAL', 'VL. TOTAL', 'PREÇO TOTAL', 'PRECO TOTAL', 'VALOR GLOBAL', 'TOTAL (R$)', 'TOTAL R$', 'SUBTOTAL']):
            if not any(d in h_up for d in descartes_vt):
                col_vt = ci
        elif any(k in h_up for k in ['PREÇO UNID', 'PRECO UNID', 'VALOR UNID', 'P. UNID', 'V. UNID', 'VL. UNID', 'UNIT', 'UNITARIO', 'V. UNIT', 'VL. UNIT', 'PREÇO UNIT', 'PRECO UNIT', 'VALOR UNITÁRIO', 'VALOR UNIT']) and col_vu is None:
            col_vu = ci
        elif any(k in h_up for k in ['QTD', 'QUANT', 'QUANTIDADE']) and not any(k in h_up for k in ['VALOR', 'PREÇO', 'PRECO']) and col_qtd is None:
            col_qtd = ci
        elif any(k in h_up for k in ['UNIDADE', 'UNID', 'UND', 'MEDIDA']) and col_unid is None:
            col_unid = ci
        elif any(k in h_up for k in ['DESC', 'PROD', 'ESPEC', 'SERV', 'OBJET', 'MATERIAL', 'DISCRIMINAÇÃO']) and col_desc is None:
            col_desc = ci
        elif any(k in h_up for k in ['ITEM', 'Nº', 'NUM', 'ORDEM']) and col_item is None:
            col_item = ci

    if col_vt is None:
        for ci in reversed(range(len(headers))):
            h_up = str(headers[ci]).upper().strip()
            if any(k in h_up for k in ['VALOR', 'PREÇO', 'PRECO', 'TOTAL']) and ci != col_vu and ci != col_qtd:
                col_vt = ci
                break

    memoria_calculo = []
    soma_calculada_itens = 0.0
    linhas_divergentes = 0
    itens_com_valor = 0

    for idx, r in enumerate(rows, start=1):
        txt_row = " ".join([str(c or '').upper().strip() for c in r])
        if any(k in txt_row for k in ["TOTAL GERAL", "SUBTOTAL", "TOTAL DA PROPOSTA", "SOMA TOTAL", "VALOR GLOBAL"]):
            continue

        item_str = str(r[col_item] if col_item is not None and col_item < len(r) else idx).strip() or str(idx)
        desc_str = str(r[col_desc] if col_desc is not None and col_desc < len(r) else (r[1] if len(r) > 1 else f"Item {item_str}")).strip()
        qtd_val = parse_numero_br(r[col_qtd]) if col_qtd is not None and col_qtd < len(r) else 1.0
        if qtd_val == 0.0 and col_qtd is None:
            qtd_val = 1.0
        vu_val = parse_numero_br(r[col_vu]) if col_vu is not None and col_vu < len(r) else 0.0
        vt_inf = parse_numero_br(r[col_vt]) if col_vt is not None and col_vt < len(r) else 0.0

        vt_calc = (qtd_val * vu_val) if (qtd_val > 0 and vu_val > 0) else vt_inf
        if vt_inf == 0.0 and vt_calc > 0:
            vt_inf = vt_calc
        elif vt_calc == 0.0 and vt_inf > 0:
            vt_calc = vt_inf

        val_linha = vt_inf if vt_inf > 0 else vt_calc
        soma_calculada_itens += val_linha

        if val_linha > 0:
            itens_com_valor += 1

        tem_erro_linha = False
        dif_linha = 0.0
        if qtd_val > 0 and vu_val > 0 and vt_inf > 0:
            dif_linha = round(abs(vt_inf - (qtd_val * vu_val)), 2)
            if dif_linha >= 0.02:
                tem_erro_linha = True
                linhas_divergentes += 1

        memoria_calculo.append({
            "linha": idx,
            "item": item_str,
            "descricao": desc_str[:90] + ("..." if len(desc_str) > 90 else ""),
            "quantidade": qtd_val,
            "valor_unitario": vu_val,
            "valor_total_informado": vt_inf,
            "valor_total_calculado": vt_calc,
            "valor_adotado": val_linha,
            "tem_erro_linha": tem_erro_linha,
            "diferenca_linha": dif_linha,
            "status_linha": "⚠️ Discrepância" if tem_erro_linha else ("✅ OK" if val_linha > 0 else "⚪ Zerado")
        })

    soma_final_itens = total_geral_pre if (total_geral_pre > 0 and abs(total_geral_pre - soma_calculada_itens) < 0.05) else (soma_calculada_itens if soma_calculada_itens > 0 else total_geral_pre)
    soma_final_itens = round(soma_final_itens, 2)
    valor_contrato = round(float(valor_contrato_atual or 0.0), 2)
    diferenca = round(soma_final_itens - valor_contrato, 2)

    confere = (abs(diferenca) < 0.01) and (soma_final_itens > 0 or valor_contrato == 0.0)
    if valor_contrato == 0.0 and soma_final_itens > 0:
        confere = False
        status_geral = "PENDENTE_SINCRONIZACAO"
        mensagem = f"A planilha possui o total de {formatar_moeda_br(soma_final_itens)}, mas o Valor do Contrato ainda está zerado (R$ 0,00). Clique para sincronizar."
    elif confere:
        status_geral = "CONFERE"
        mensagem = f"Conferência Aprovada: A soma dos {len(memoria_calculo)} itens da planilha ({formatar_moeda_br(soma_final_itens)}) bate exatamente com o Valor Global do Contrato ({formatar_moeda_br(valor_contrato)})."
    else:
        status_geral = "DIVERGENTE"
        sinal = "+" if diferenca > 0 else ""
        mensagem = f"Divergência Detectada: A soma da planilha ({formatar_moeda_br(soma_final_itens)}) difere do Valor Global ({formatar_moeda_br(valor_contrato)}) em {sinal}{formatar_moeda_br(diferenca)}."

    return {
        "tem_planilha": True,
        "confere": confere,
        "status": status_geral,
        "soma_itens": soma_final_itens,
        "soma_formatada": formatar_moeda_br(soma_final_itens),
        "valor_contrato": valor_contrato,
        "valor_contrato_formatado": formatar_moeda_br(valor_contrato),
        "diferenca": diferenca,
        "diferenca_formatada": formatar_moeda_br(abs(diferenca)),
        "sinal_diferenca": "+" if diferenca > 0 else ("-" if diferenca < 0 else ""),
        "total_itens": len(memoria_calculo),
        "itens_com_valor": itens_com_valor,
        "itens_zerados": len(memoria_calculo) - itens_com_valor,
        "linhas_divergentes_qtd": linhas_divergentes,
        "memoria_calculo": memoria_calculo,
        "mensagem": mensagem
    }

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


def formatar_clausulas_inteligencia(texto_bruto: str, num_clausula_principal: int = 13) -> list:
    """Motor de formatação inteligente que utiliza o Google Gemini (LLM via REST) para readequar 
    o texto juridicamente em estrita observância à Nova Lei de Licitações (Lei nº 14.133/2021),
    com fallback para o sistema baseado em regras se a API falhar."""
    if not texto_bruto or not str(texto_bruto).strip():
        return []
    
    try:
        import requests
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


def extrair_tabela_de_linhas_brutas(raw_rows: list, is_dotacao: bool = False) -> pd.DataFrame:
    """
    Normaliza e extrai uma tabela (DataFrame) com fidelidade absoluta (1:1) a partir de linhas brutas:
    - Preserva rigorosamente 100% da quantidade original de colunas e nomes de cabeçalhos.
    - Preserva rigorosamente 100% da quantidade original de linhas e conteúdos textuais/caracteres.
    - Se a tabela tiver cabeçalhos autênticos (Item, Descrição, Qtd, Valor, etc.), mantém apenas as colunas
      pertencentes à grade real, descartando colunas fantasmas e células isoladas fora da grade.
    - Descarta blocos isolados de assinatura no rodapé (ex: 'Assinado por...', 'Prefeito', 'Secretário').
    - Preserva rigorosamente linhas de resumo / totalizador da própria planilha (TOTAL, VALOR TOTAL, etc.).
    - Se a tabela começar diretamente com dados/códigos sem cabeçalho (ou dotação chave-valor),
      mantém 100% das linhas como dados e define has_real_headers=False para não inventar cabeçalho artificial.
    - Não omite nenhuma informação da planilha modelo importada e não inventa linhas ou colunas artificiais.
    """
    import re
    import pandas as pd

    if not raw_rows:
        return pd.DataFrame()

    def _clean_str(c):
        if c is None:
            return ""
        s = str(c).strip()
        if s.lower() in ('nan', 'none', 'null'):
            return ""
        try:
            from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
            s = ILLEGAL_CHARACTERS_RE.sub('', s)
        except Exception:
            pass
        return s

    # 1. Limpa células e remove linhas completamente vazias no topo e rodapé
    clean_rows = []
    for r in raw_rows:
        row_c = [_clean_str(c) for c in r]
        clean_rows.append(row_c)

    while clean_rows and not any(clean_rows[-1]):
        clean_rows.pop()
    start_idx = 0
    while start_idx < len(clean_rows) and not any(clean_rows[start_idx]):
        start_idx += 1
    clean_rows = clean_rows[start_idx:]
    if not clean_rows:
        return pd.DataFrame()

    # Auto-detecta Dotação Orçamentária mesmo se is_dotacao=False foi passado
    if not is_dotacao:
        dota_terms = ['UNIDADE ORÇAMENTÁRIA', 'UNIDADE ORCAMENTARIA', 'NATUREZA DA DESPESA', 'PODER EXECUTIVO', 'CRÉDITO ORÇAMENTÁRIO', 'CREDITO ORCAMENTARIO', 'DOTAÇÃO ORÇAMENTÁRIA', 'DOTACAO ORCAMENTARIA']
        first_text = ' '.join(' '.join(r) for r in clean_rows[:8]).upper()
        if any(term in first_text for term in dota_terms):
            is_dotacao = True

    # Tratamento especializado para Dotação Orçamentária
    if is_dotacao:
        col_counts = {}
        for r in clean_rows:
            for ci, val in enumerate(r):
                if val:
                    col_counts[ci] = col_counts.get(ci, 0) + 1
        if not col_counts:
            return pd.DataFrame()
        min_ci = min(col_counts.keys())
        max_ci = max(col_counts.keys())
        num_cols = max_ci - min_ci + 1

        sliced_rows = []
        for r in clean_rows:
            row_slice = [r[ci] if ci < len(r) else "" for ci in range(min_ci, max_ci + 1)]
            if not any(row_slice):
                continue
            if len(row_slice) >= 3 and len(sliced_rows) > 0:
                c1_s = str(row_slice[1] or '').strip()
                c2_s = str(row_slice[2] or '').strip()
                if not c1_s and not c2_s:
                    continue
            sliced_rows.append(row_slice)

        if not sliced_rows:
            return pd.DataFrame()

        clean_headers = [f"COLUNA {ci+1}" for ci in range(num_cols)]
        df_ret = pd.DataFrame(sliced_rows, columns=clean_headers)
        df_ret.attrs['has_real_headers'] = False
        return df_ret

    # Padrões para identificar faixas isoladas de banner/título no topo ANTES da tabela
    banner_patterns = [
        r'\bPREFEITURA\b', r'\bESTADO DO\b', r'\bGOVERNO DO\b', r'\bMUNIC[IÍ]PIO\b', r'\bGABINETE\b',
        r'\bSECRETARIA\b', r'\bFUNDO MUNICIPAL\b', r'\bPREG[AÃ]O\b', r'\bPROCESSO\b', r'\bEDITAL\b',
        r'\bTERMO DE REFER[EÊ]NCIA\b', r'\bANEXO\b', r'\bCONTRATO\b', r'\bDISPENSA\b', r'\bINEXIGIBILIDADE\b'
    ]

    header_patterns = [
        r'\bITEM\b', r'\bN[º°]?\b', r'\bORDEM\b', r'\bC[OÓ]D(IGO)?\b', r'\bCATMAT\b', r'\bCATSER\b',
        r'\bDESCRI[CÇ][AÃ]O\b', r'\bDESCRICAO\b', r'\bESPECIFICA[CÇ][AÃ]O\b', r'\bESPECIFICACAO\b', r'\bDISCRIMINA[CÇ][AÃ]O\b', r'\bOBJETO\b',
        r'\bPRODUTO\b', r'\bSERVI[CÇ]O\b', r'\bSERVICO\b', r'\bMATERIAL\b',
        r'\bUNID(ADE)?\b', r'\bUND\b', r'\bMEDIDA\b',
        r'\bQTD\b', r'\bQUANT(IDADE)?\b',
        r'\bVALOR\b', r'\bPRE[CÇ]O\b', r'\bPRECO\b', r'\bUNIT([AÁ]RIO)?\b', r'\bTOTAL\b', r'\bSUBTOTAL\b', r'\bGLOBAL\b',
        r'\bMARCA\b', r'\bMODELO\b', r'\bFABRICANTE\b',
        r'\bLOTE\b', r'\bOBS(ERVA[CÇ][AÃ]O)?\b'
    ]

    sig_patterns = [
        r'\bASSINAD[OA]\b', r'\bSECRET[AÁ]RI[OA]\b', r'\bPREFEIT[OA]\b', r'\bCOORDENADOR\b',
        r'\bFISCAL\b', r'\bDIRETOR\b', r'\bDATA\s*:', r'\bP[AÁ]GINA\b', r'\bEMITIDO\s+EM\b',
        r'\bRESPONS[AÁ]VEL\b', r'\bTESTEMUNHA\b'
    ]

    summary_keywords = [
        'TOTAL', 'TOTAL GERAL', 'SUBTOTAL', 'VALOR TOTAL', 'VALOR GLOBAL',
        'SOMA', 'SOMA TOTAL', 'TOTAL DA PROPOSTA', 'TOTAL DO OBJETO', 'TOTAL ESTIMADO', 'TOTAL HOMOLOGADO'
    ]

    # Busca a linha de cabeçalho entre as 10 primeiras linhas não vazias
    header_idx = -1
    for idx, r in enumerate(clean_rows[:25]):
        filled = [c for c in r if c]
        if not filled:
            continue
        row_text = ' '.join(filled).upper()
        # Se for banner isolado de prefeitura/órgão com 1 ou 2 células, pula
        if any(re.search(pat, row_text) for pat in banner_patterns) and len(filled) <= 2:
            continue
        # Verifica se casa com termos de cabeçalho
        matches = sum(1 for pat in header_patterns if re.search(pat, row_text))
        if matches >= 1:
            header_idx = idx
            break
        # Ou se possui 2 ou mais células de texto e nenhuma é puramente numérica
        if len(filled) >= 2 and all(not re.search(r'^\d+$', c.replace('.', '').replace(',', '').strip()) for c in filled):
            header_idx = idx
            break

    if header_idx != -1:
        hdr_row = clean_rows[header_idx]
        hdr_non_empty_indices = [ci for ci, c in enumerate(hdr_row) if c]
        if not hdr_non_empty_indices:
            first_col = 0
            last_col = len(hdr_row) - 1
        else:
            first_col = min(hdr_non_empty_indices)
            last_col = max(hdr_non_empty_indices)

        candidate_data_rows = clean_rows[header_idx + 1:]

        # Verifica se alguma coluna além de last_col possui dados consistentes nas linhas de dados (>= 35% das linhas)
        extra_cols = []
        max_seen_col = max(len(r) for r in candidate_data_rows) if candidate_data_rows else last_col + 1
        for ci in range(last_col + 1, max_seen_col):
            col_filled = sum(1 for dr in candidate_data_rows if ci < len(dr) and dr[ci])
            if col_filled >= max(2, int(len(candidate_data_rows) * 0.35)):
                extra_cols.append(ci)
        if extra_cols:
            last_col = max(extra_cols)

        # Cabeçalhos exatos da linha de cabeçalho
        seen = {}
        final_headers = []
        for ci in range(first_col, last_col + 1):
            h_raw = hdr_row[ci] if ci < len(hdr_row) else ""
            h_name = h_raw if h_raw else f"COLUNA {ci - first_col + 1}"
            if h_name in seen:
                seen[h_name] += 1
                final_headers.append(f"{h_name} ({seen[h_name]})")
            else:
                seen[h_name] = 1
                final_headers.append(h_name)

        # Filtra linhas de dados: preserva todas as linhas e totais, descarta apenas blocos de assinatura isolados no rodapé
        final_data_rows = []
        for dr in candidate_data_rows:
            row_slice = [dr[ci] if ci < len(dr) else "" for ci in range(first_col, last_col + 1)]
            filled_slice = [c for c in row_slice if c]
            if not filled_slice:
                continue
            slice_text = ' '.join(filled_slice).upper()
            is_summary = any(k in slice_text for k in summary_keywords)
            # Se tiver apenas 1 célula preenchida e for assinatura/rodapé isolado no final, não é linha de dados
            if len(filled_slice) <= 1 and not is_summary:
                if any(re.search(pat, slice_text) for pat in sig_patterns):
                    continue
            final_data_rows.append(row_slice)

        if not is_dotacao:
            final_headers, final_data_rows = reconciliar_tabela_irregular(final_headers, final_data_rows, is_dotacao=False)
        df_ret = pd.DataFrame(final_data_rows, columns=final_headers)
        df_ret.attrs['has_real_headers'] = True
        return df_ret
    else:
        # Tabela sem cabeçalho: linhas são dados puros
        col_counts = {}
        for r in clean_rows:
            for ci, c in enumerate(r):
                if c:
                    col_counts[ci] = col_counts.get(ci, 0) + 1
        num_rows = len(clean_rows)
        valid_cols = [ci for ci, cnt in col_counts.items() if cnt >= max(1, int(num_rows * 0.25))]
        if not valid_cols:
            valid_cols = sorted(col_counts.keys()) if col_counts else [0]
        first_col = min(valid_cols)
        last_col = max(valid_cols)
        num_cols = last_col - first_col + 1
        headers = [f"COLUNA {i+1}" for i in range(num_cols)]

        final_data_rows = []
        for r in clean_rows:
            row_slice = [r[ci] if ci < len(r) else "" for ci in range(first_col, last_col + 1)]
            if any(row_slice):
                final_data_rows.append(row_slice)

        df_ret = pd.DataFrame(final_data_rows, columns=headers)
        df_ret.attrs['has_real_headers'] = False
        return df_ret
def extrair_dataframe_de_arquivo(file_input, file_name: str = "", is_dotacao: bool = False):
    """
    Extrai um DataFrame a partir de qualquer arquivo de tabela ou documento suportado:
    - Excel (.xlsx, .xls, .xlsm): com detecção automática do início da tabela, descartando títulos
    - Word (.docx, .doc): extrai tabelas estruturadas (preserva células e cabeçalhos)
    - CSV (.csv)
    - PDF (.pdf)
    - Texto puro / Markdown / TSV
    """
    import io
    import os
    import pandas as pd

    if file_input is None:
        return None

    raw_bytes = None
    if hasattr(file_input, 'getvalue'):
        raw_bytes = file_input.getvalue()
        buf = io.BytesIO(raw_bytes)
    elif isinstance(file_input, (bytes, bytearray)):
        raw_bytes = bytes(file_input)
        buf = io.BytesIO(raw_bytes)
    elif hasattr(file_input, 'read'):
        if hasattr(file_input, 'seek'):
            file_input.seek(0)
        raw_bytes = file_input.read()
        buf = io.BytesIO(raw_bytes)
        if hasattr(file_input, 'seek'):
            file_input.seek(0)
    elif isinstance(file_input, str) and os.path.exists(file_input):
        with open(file_input, 'rb') as f:
            raw_bytes = f.read()
        buf = io.BytesIO(raw_bytes)
        if not file_name:
            file_name = file_input
    else:
        buf = file_input

    fn_lower = (file_name or getattr(file_input, 'name', '') or "").lower()

    # Detecta se os dados brutos pertencem a arquivo binario (ZIP/DOCX/XLSX, PDF, OLE/DOC, ou contem bytes nulos)
    is_binary = False
    if raw_bytes:
        is_binary = bool(
            raw_bytes[:4] == b'PK\x03\x04' or
            raw_bytes[:4] == b'%PDF' or
            raw_bytes[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' or
            b'\x00' in raw_bytes[:4096]
        )

    def _score_candidate_table(df):
        if df is None or df.empty:
            return -1
        cols = [str(c).upper() for c in df.columns]
        num_cols = len(cols)
        num_rows = len(df)
        if num_cols < 2:
            return 0
        score = num_rows * 2
        item_kws = ['ITEM', 'DESCRIÇÃO', 'DESCRICAO', 'ESPECIFICAÇÃO', 'ESPECIFICACAO', 'PRODUTO', 'SERVIÇO', 'SERVICO', 'UNIDADE', 'QTD', 'QUANTIDADE', 'VALOR', 'PREÇO', 'PRECO', 'TOTAL', 'UNIDADE ORÇAMENTÁRIA', 'UNIDADE ORCAMENTARIA', 'NATUREZA DA DESPESA']
        for c in cols:
            if any(k in c for k in item_kws):
                score += 25
        if df.attrs.get('has_real_headers'):
            score += 30
        return score

    # 1. PLANILHAS EXCEL (.xlsx, .xls, .xlsm)
    is_excel_ext = fn_lower.endswith(('.xlsx', '.xls', '.xlsm'))
    is_docx_ext = fn_lower.endswith(('.docx', '.doc')) or (raw_bytes and (b'word/document.xml' in raw_bytes[:10000]))
    if (is_excel_ext or (raw_bytes and (raw_bytes[:4] == b'PK\x03\x04' or raw_bytes[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'))) and not is_docx_ext:
        xl_candidates = []
        # A. Tenta com openpyxl testando todas as abas
        try:
            buf.seek(0)
            import openpyxl
            wb = openpyxl.load_workbook(buf, data_only=True)
            sheet_names_to_try = [wb.active.title] + [sn for sn in wb.sheetnames if sn != wb.active.title]
            for sname in sheet_names_to_try:
                sheet = wb[sname]
                raw_rows = []
                for row in sheet.iter_rows():
                    row_vals = []
                    for cell in row:
                        val = cell.value
                        if val is None:
                            row_vals.append("")
                        elif isinstance(val, str):
                            s = val.strip()
                            row_vals.append("" if s.lower() in ('nan', 'none') else s)
                        elif isinstance(val, int):
                            row_vals.append(str(val))
                        elif isinstance(val, float):
                            num_fmt = str(getattr(cell, 'number_format', '') or '')
                            if any(k in num_fmt for k in ['R$', '$', 'BRL']):
                                row_vals.append(f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                            elif any(k in num_fmt for k in ['0.00', '#,##0.00', '0,00']):
                                row_vals.append(f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                            elif val.is_integer():
                                row_vals.append(str(int(val)))
                            else:
                                s_f = f"{val:.4f}".rstrip("0").rstrip(".").replace(".", ",")
                                row_vals.append(s_f)
                        elif hasattr(val, 'strftime'):
                            row_vals.append(val.strftime('%d/%m/%Y'))
                        else:
                            row_vals.append(str(val).strip())
                    if any(c for c in row_vals):
                        raw_rows.append(row_vals)
                if raw_rows:
                    df_res = extrair_tabela_de_linhas_brutas(raw_rows, is_dotacao=is_dotacao)
                    if not df_res.empty:
                        xl_candidates.append(df_res)
        except Exception:
            pass

        # B. Fallback com pandas read_excel para todas as abas (suporta .xlsx e .xls via xlrd)
        try:
            buf.seek(0)
            engine_xl = 'xlrd' if fn_lower.endswith('.xls') else None
            try:
                xl_dict = pd.read_excel(buf, sheet_name=None, header=None, dtype=str, engine=engine_xl)
            except Exception:
                buf.seek(0)
                xl_dict = pd.read_excel(buf, sheet_name=None, header=None, dtype=str)
            
            for sname, df_sheet in xl_dict.items():
                if not df_sheet.empty:
                    df_res = extrair_tabela_de_linhas_brutas(df_sheet.values.tolist(), is_dotacao=is_dotacao)
                    if not df_res.empty:
                        xl_candidates.append(df_res)
        except Exception:
            pass

        if xl_candidates:
            xl_candidates.sort(key=lambda x: _score_candidate_table(x), reverse=True)
            return xl_candidates[0]

    # 2. DOCUMENTO WORD (.docx, .doc)
    is_docx = fn_lower.endswith(('.docx', '.doc')) or (raw_bytes and (b'word/document.xml' in raw_bytes[:10000]))
    if is_docx:
        docx_candidates = []
        try:
            import docx
            buf.seek(0)
            doc = docx.Document(buf)
            for tbl in doc.tables:
                if not tbl.rows:
                    continue
                raw_rows = []
                for r in tbl.rows:
                    row_cells = []
                    prev_tc = None
                    for cell in r.cells:
                        if cell._tc == prev_tc:
                            row_cells.append("")
                        else:
                            txt = cell.text.strip().replace("\r", "")
                            row_cells.append(txt)
                        prev_tc = cell._tc
                    if any(c.strip() for c in row_cells):
                        raw_rows.append(row_cells)
                if raw_rows:
                    df_cand = extrair_tabela_de_linhas_brutas(raw_rows, is_dotacao=is_dotacao)
                    if not df_cand.empty:
                        docx_candidates.append(df_cand)

            if docx_candidates:
                docx_candidates.sort(key=lambda x: _score_candidate_table(x), reverse=True)
                return docx_candidates[0]

            full_txt = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
            if full_txt:
                t_tab = parse_markdown_or_tsv_table(full_txt)
                if t_tab and t_tab.get("headers") and t_tab.get("rows"):
                    df_res = extrair_tabela_de_linhas_brutas([t_tab["headers"]] + t_tab["rows"], is_dotacao=is_dotacao)
                    if not df_res.empty:
                        return df_res
        except Exception as e_docx:
            print("Erro ao ler DOCX:", e_docx)

    # 3. ARQUIVOS CSV (.csv) - Apenas se nao for arquivo binario
    if fn_lower.endswith('.csv') and not is_binary:
        for enc, sep in [('utf-8', ','), ('latin1', ';'), ('utf-8', ';'), ('latin1', ','), ('utf-8-sig', ';'), ('utf-8-sig', ',')]:
            try:
                buf.seek(0)
                df_csv = pd.read_csv(buf, header=None, dtype=str, encoding=enc, sep=sep)
                if not df_csv.empty and len(df_csv.columns) >= 2:
                    df_res = extrair_tabela_de_linhas_brutas(df_csv.values.tolist(), is_dotacao=is_dotacao)
                    if not df_res.empty:
                        return df_res
            except Exception:
                continue

    # 4. DOCUMENTOS PDF (.pdf)
    if fn_lower.endswith('.pdf') or (raw_bytes and raw_bytes[:4] == b'%PDF'):
        try:
            import pdfplumber
            buf.seek(0)
            all_tbl_rows = []
            with pdfplumber.open(buf) as pdf:
                for page in pdf.pages:
                    t_list = page.extract_tables() or []
                    for t in t_list:
                        for row in t:
                            c_clean = [str(x or '').strip().replace('\r', '') for x in row]
                            if any(c_clean):
                                all_tbl_rows.append(c_clean)
            if all_tbl_rows:
                df_res = extrair_tabela_de_linhas_brutas(all_tbl_rows, is_dotacao=is_dotacao)
                if not df_res.empty:
                    return df_res
            
            buf.seek(0)
            pdf_text = "\n".join([p.extract_text() or "" for p in pdf.pages])
            t_parsed = parse_markdown_or_tsv_table(pdf_text)
            if t_parsed and t_parsed.get("headers") and t_parsed.get("rows"):
                df_res = extrair_tabela_de_linhas_brutas([t_parsed["headers"]] + t_parsed["rows"], is_dotacao=is_dotacao)
                if not df_res.empty:
                    return df_res
        except Exception as e_pdf:
            print("Erro ao ler PDF:", e_pdf)

    # 5. TENTATIVA GENÉRICA (Apenas se NAO for arquivo binario)
    if not is_binary:
        try:
            buf.seek(0)
            df_gen = pd.read_excel(buf, header=None, dtype=str)
            if not df_gen.empty:
                df_res = extrair_tabela_de_linhas_brutas(df_gen.values.tolist(), is_dotacao=is_dotacao)
                if not df_res.empty:
                    return df_res
        except Exception:
            pass

        for enc_c, sep_c in [('utf-8', ','), ('latin1', ';'), ('utf-8', ';'), ('latin1', ',')]:
            try:
                buf.seek(0)
                df_gen = pd.read_csv(buf, header=None, dtype=str, encoding=enc_c, sep=sep_c)
                if not df_gen.empty and len(df_gen.columns) >= 2:
                    df_res = extrair_tabela_de_linhas_brutas(df_gen.values.tolist(), is_dotacao=is_dotacao)
                    if not df_res.empty:
                        return df_res
            except Exception:
                pass

        try:
            buf.seek(0)
            txt = raw_bytes.decode('utf-8', errors='ignore') if raw_bytes else ""
            t_parsed = parse_markdown_or_tsv_table(txt)
            if t_parsed and t_parsed.get("headers") and t_parsed.get("rows"):
                df_res = extrair_tabela_de_linhas_brutas([t_parsed["headers"]] + t_parsed["rows"], is_dotacao=is_dotacao)
                if not df_res.empty:
                    return df_res
        except Exception:
            pass

    return None


def processar_planilha_itens_excel(file_bytes_or_buffer, file_name: str = "", objeto_contexto: str = "") -> ResultadoPlanilhaItens:
    """
    Processa arquivo Excel (.xlsx / .xls), CSV, Word (.docx / .doc) ou PDF contendo itens da contratação.
    Auto-ajusta quantidade de colunas, linhas e cabeçalhos com total flexibilidade.
    Retorna ResultadoPlanilhaItens (compatível com lista de itens e com .tabela_dinamica).
    """
    import pandas as pd
    try:
        raw_bytes = None
        if hasattr(file_bytes_or_buffer, 'getvalue'):
            raw_bytes = file_bytes_or_buffer.getvalue()
            buf = io.BytesIO(raw_bytes)
        elif isinstance(file_bytes_or_buffer, (bytes, bytearray)):
            raw_bytes = bytes(file_bytes_or_buffer)
            buf = io.BytesIO(raw_bytes)
        elif hasattr(file_bytes_or_buffer, 'read'):
            if hasattr(file_bytes_or_buffer, 'seek'):
                file_bytes_or_buffer.seek(0)
            raw_bytes = file_bytes_or_buffer.read()
            if hasattr(file_bytes_or_buffer, 'seek'):
                file_bytes_or_buffer.seek(0)
            buf = io.BytesIO(raw_bytes)
        else:
            buf = file_bytes_or_buffer

        df = extrair_dataframe_de_arquivo(buf, file_name=file_name, is_dotacao=False)

        tab_din = extrair_tabela_dinamica_df(df)

        if tab_din and tab_din.get("headers") and tab_din.get("rows"):
            itens = converter_tabela_reconciliada_para_itens(tab_din["headers"], tab_din["rows"])
            soma_total_itens = sum(it.get("_vt_float", 0.0) for it in itens)
            if isinstance(tab_din, dict):
                if tab_din.get("total_geral", 0.0) > 0 and soma_total_itens > 0:
                    if abs(tab_din["total_geral"] - soma_total_itens) < 0.05:
                        soma_total_itens = tab_din["total_geral"]
                tab_din["total_geral"] = soma_total_itens if soma_total_itens > 0 else tab_din.get("total_geral", 0.0)
                tab_din["total_formatado"] = formatar_moeda_br(tab_din["total_geral"])
            return ResultadoPlanilhaItens(itens, tabela_dinamica=tab_din)

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
            desc = str(row[col_desc]).strip() if col_desc and not pd.isna(row[col_desc]) else (" ".join([str(v).strip() for v in row.values if v and str(v).strip()][:2]) or f"Item {item_num}")
            qtd = limpar_num(row[col_qtd]) if col_qtd else 1.0
            unid = str(row[col_unid]).strip() if (col_unid and not pd.isna(row[col_unid]) and str(row[col_unid]).strip()) else ""
            vu = limpar_num(row[col_vu]) if col_vu else 0.0
            vt = limpar_num(row[col_vt]) if col_vt else 0.0
            if vt == 0.0 and qtd > 0 and vu > 0:
                vt = qtd * vu
            elif vu == 0.0 and qtd > 0 and vt > 0:
                vu = vt / qtd

            if qtd == 0.0 and vu == 0.0 and vt == 0.0 and (not desc or desc == f"Item {item_num}"):
                continue

            qtd_str = "1" if (not col_qtd or qtd == 0.0) else (f"{int(qtd)}" if qtd == int(qtd) else f"{qtd:.2f}".replace('.', ','))

            itens.append({
                "item": item_num,
                "descricao": desc,
                "quantidade": qtd_str,
                "unidade": unid,
                "valor_unitario": formatar_moeda_br(vu) if vu > 0 else "",
                "valor_total": formatar_moeda_br(vt) if vt > 0 else "",
                "_vu_float": vu,
                "_vt_float": vt,
                "_qtd_float": qtd if qtd > 0 else 1.0
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

def parse_data_universal(val) -> Optional[date]:
    """Converte qualquer formato de data (string ISO, string BR, date, datetime, pd.Timestamp) para objeto date."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    if not s or s.lower() in ["nan", "none", "null", "nat", "-", "não informada", "nao informada"]:
        return None
    # Truncar horas caso venha '2026-10-19 14:30:00' ou '19/10/2026 00:00'
    s_clean = s.split(" ")[0].split("T")[0].strip()
    for fmt in ["%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%Y/%m/%d"]:
        try:
            return datetime.strptime(s_clean, fmt).date()
        except ValueError:
            pass
    # Regex para formatos brasileiros DD/MM/AAAA ou DD-MM-AAAA
    m = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", s_clean)
    if m:
        d, mth, y = m.groups()
        if len(y) == 2:
            y = "20" + y
        try:
            return date(int(y), int(mth), int(d))
        except Exception:
            pass
    # Regex para ISO AAAA-MM-DD
    m_iso = re.search(r"(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})", s_clean)
    if m_iso:
        y, mth, d = m_iso.groups()
        try:
            return date(int(y), int(mth), int(d))
        except Exception:
            pass
    return None

def formatar_data_br(val, padrao: str = "") -> str:
    """Formata qualquer data no padrão brasileiro oficial: DD/MM/AAAA"""
    dt = parse_data_universal(val)
    if dt is not None:
        return f"{dt.day:02d}/{dt.month:02d}/{dt.year:04d}"
    if val is not None and str(val).strip() not in ["", "None", "nan", "NaT"]:
        return str(val).strip()
    return padrao

def formatar_data_extenso(val) -> str:
    """Formata data por extenso em português do Brasil: 'DD de Mês de AAAA'"""
    dt = parse_data_universal(val)
    if dt is None:
        dt = date.today()
    meses = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
    return f"{dt.day} de {meses[dt.month - 1]} de {dt.year}"

def formatar_data_completa_extenso_br(val) -> str:
    """Formata data completa com dia da semana em português do Brasil: 'Dia-da-semana, DD de Mês de AAAA'"""
    dt = parse_data_universal(val)
    if dt is None:
        dt = date.today()
    dias_semana = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"]
    meses = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
    return f"{dias_semana[dt.weekday()]}, {dt.day} de {meses[dt.month - 1]} de {dt.year}"

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
    
    dt_ass_parsed = parse_data_universal(d.get("data_assinatura")) or date.today()
    dt_venc_raw = d.get("data_vencimento")
    dt_venc_parsed = parse_data_universal(dt_venc_raw) if dt_venc_raw else (dt_ass_parsed.replace(year=dt_ass_parsed.year + 1))
    
    d["data_assinatura"] = dt_ass_parsed
    d["data_assinatura_br"] = formatar_data_br(dt_ass_parsed)
    d["data_assinatura_extenso"] = formatar_data_extenso(dt_ass_parsed)
    d["data_vencimento"] = formatar_data_br(dt_venc_parsed)
    d["data_vencimento_br"] = formatar_data_br(dt_venc_parsed)
    d["data_vencimento_extenso"] = formatar_data_extenso(dt_venc_parsed)
    d["vigencia_descricao"] = d.get("vigencia_descricao") or "12 (doze) meses a contar da data de assinatura"
    d["dotacao_orcamentaria"] = d.get("dotacao_orcamentaria") or "Unidade Orçamentária: 02 - Recursos Ordinários / Próprios do Município"
    d["modelo_agu"] = (d.get("modelo_agu") or "COMPRAS").upper()
    d["foro"] = d.get("foro") or "Comarca de Imperatriz - MA"

    # Metadados de Dotação e Itens para Tabelas Visuais
    exercicio_padrao = str(dt_ass_parsed.year)
    d["dotacao_exercicio"] = d.get("dotacao_exercicio") or exercicio_padrao
    d["dotacao_poder"] = d.get("dotacao_poder") or "Poder Executivo 02.00"
    raw_dot_orc = str(d.get("dotacao_orcamentaria") or "").strip()
    if not d.get("dotacao_unidade"):
        if not (raw_dot_orc.startswith("{") or "---|" in raw_dot_orc or "\t" in raw_dot_orc):
            d["dotacao_unidade"] = raw_dot_orc or "MANUTENÇÃO E FUNCIONAMENTO 12.361.0402.2022.0000"
        else:
            d["dotacao_unidade"] = "MANUTENÇÃO E FUNCIONAMENTO 12.361.0402.2022.0000"
    d["dotacao_natureza"] = d.get("dotacao_natureza") or ("Material de consumo 33.90.30.00" if "COMPRA" in str(d.get("tipo_contratacao", "")).upper() else "Outros Serviços de Terceiros 33.90.39.00")
    
    # Auto-desempacota dotacao_tabela se dotacao_orcamentaria estiver em formato JSON ou Markdown/TSV/Grade
    if not d.get("dotacao_tabela") and d.get("dotacao_orcamentaria"):
        val_d_raw = str(d["dotacao_orcamentaria"]).strip()
        if val_d_raw.startswith("{") and "headers" in val_d_raw:
            try:
                import json
                d_js = json.loads(val_d_raw)
                if isinstance(d_js, dict) and d_js.get("headers"):
                    d["dotacao_tabela"] = d_js
            except Exception:
                pass
        if not d.get("dotacao_tabela"):
            t_p = parse_markdown_or_tsv_table(val_d_raw)
            if t_p and t_p.get("headers") and t_p.get("rows"):
                d["dotacao_tabela"] = t_p
    
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
            if isinstance(it, (list, tuple)):
                num_it = str(it[0]) if len(it) > 0 else str(idx)
                desc_it = str(it[1]) if len(it) > 1 else d["objeto"]
                qtd_it = str(it[2]) if len(it) > 2 else "1"
                unid_it = str(it[3]) if len(it) > 3 else "Unid."
                vu = str(it[4]) if len(it) > 4 else d["valor_total_formatado"]
                vt = str(it[5]) if len(it) > 5 else (str(it[-1]) if len(it) > 1 else d["valor_total_formatado"])
                vt_float = 0.0
                s_vt = str(vt).replace("R$", "").replace(" ", "").strip()
                if "," in s_vt and "." in s_vt:
                    s_vt = s_vt.replace(".", "").replace(",", ".")
                elif "," in s_vt:
                    s_vt = s_vt.replace(",", ".")
                try:
                    vt_float = float(s_vt)
                except Exception:
                    vt_float = 0.0
            elif isinstance(it, dict):
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
            else:
                num_it = idx
                desc_it = str(it)
                qtd_it = "1"
                unid_it = "Unid."
                vu = d["valor_total_formatado"]
                vt = d["valor_total_formatado"]
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


def parse_markdown_or_tsv_table(text, is_dotacao: bool = False):
    if not text: return None
    import re
    lines = [line.strip() for line in str(text).strip().split('\n') if line.strip()]
    if not lines: return None
    
    # 1. Tabela Markdown com separador de cabeçalho (|---|)
    start_idx = -1
    for i in range(len(lines) - 1):
        if '|' in lines[i] and '|' in lines[i+1] and '---' in lines[i+1]:
            start_idx = i
            break
            
    if start_idx != -1 and len(lines) >= start_idx + 2:
        headers = [re.sub(r'<[^>]+>', '', h).strip() for h in lines[start_idx].strip('|').split('|')]
        rows = []
        for line in lines[start_idx + 2:]:
            if '---' in line: continue
            if '|' not in line: break
            parts = [re.sub(r'<[^>]+>', '', p).strip() for p in line.strip('|').split('|')]
            if any(parts):
                rows.append(parts)
        if headers and rows:
            return {"headers": headers, "rows": rows}

    # 2. Tabela TSV (com tabs \t), Pipe (|) sem divisória, ou linhas de Atividade + Código
    raw_rows = []
    max_cols = 0
    for l in lines:
        if '\t' in l:
            parts = [re.sub(r'<[^>]+>', '', p).strip() for p in l.split('\t')]
        elif '|' in l:
            parts = [re.sub(r'<[^>]+>', '', p).strip() for p in l.strip('|').split('|')]
        else:
            # Tenta dividir descrição e código orçamentário
            m_c = re.search(r'^(.*?)\s+((?:\d{2}\.|\d{4})[\d\.\s\-]{4,}.*)$', l)
            if m_c:
                parts = [m_c.group(1).strip(), m_c.group(2).strip()]
            else:
                parts = [l]
        parts = [p for p in parts if p]
        if any(parts):
            raw_rows.append(parts)
            max_cols = max(max_cols, len(parts))

    if not raw_rows or max_cols < 2:
        return None

    # Normaliza número de colunas
    norm_rows = []
    for r in raw_rows:
        if len(r) < max_cols:
            r = r + [''] * (max_cols - len(r))
        norm_rows.append(r[:max_cols])

    first_row = norm_rows[0]
    has_text = any(first_row) and not all(str(x).replace('.', '').replace(',', '').isdigit() for x in first_row if str(x).strip())
    if has_text and len(norm_rows) > 1:
        seen = {}
        clean_cols = []
        for col_name in first_row:
            c_str = str(col_name).strip() or "COLUNA"
            if c_str in seen:
                seen[c_str] += 1
                clean_cols.append(f"{c_str} ({seen[c_str]})")
            else:
                seen[c_str] = 1
                clean_cols.append(c_str)
        return {"headers": clean_cols, "rows": norm_rows[1:]}
    else:
        headers = [f"COLUNA {i+1}" for i in range(max_cols)]
        return {"headers": headers, "rows": norm_rows}

def parse_mixed_markdown(text):
    if not text: return []
    lines = str(text).strip().split('\n')
    
    blocks = []
    current_text = []
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Check if table starts here
        if i < len(lines) - 1 and '|' in line and '|' in lines[i+1] and '---' in lines[i+1]:
            # Flush text
            if current_text:
                blocks.append(("text", "\n".join(current_text)))
                current_text = []
                
            headers = [h.strip() for h in line.strip('|').split('|')]
            rows = []
            i += 2
            while i < len(lines):
                tline = lines[i].strip()
                if not tline or '|' not in tline:
                    break
                if '---' not in tline:
                    parts = [p.strip() for p in tline.strip('|').split('|')]
                    if any(parts): rows.append(parts)
                i += 1
            if headers and rows:
                blocks.append(("table", {"headers": headers, "rows": rows}))
            continue
            
        if line:
            current_text.append(line)
        i += 1
        
    if current_text:
        blocks.append(("text", "\n".join(current_text)))
        
    # Se não foi detectada nenhuma tabela por '---|', mas o texto é uma tabela válida (TSV, pipe ou lista orçamentária)
    has_table = any(b[0] == 'table' for b in blocks)
    if not has_table:
        tab_parsed = parse_markdown_or_tsv_table(text)
        if tab_parsed and tab_parsed.get("headers") and tab_parsed.get("rows"):
            return [("table", tab_parsed)]
        
    return blocks

def parse_dotacao_item_multi(k, v):
    import re
    rows = []
    lines = [line.strip() for line in str(v).strip().split('\n') if line.strip()]
    if not lines:
        return []

    # 1. Se for uma tabela com múltiplas linhas (markdown, TSV ou Atividade + Código)
    if len(lines) > 1 or '\t' in str(v) or '|' in str(v):
        t_tab = parse_markdown_or_tsv_table(v)
        if t_tab and t_tab.get("rows"):
            for i, r_vals in enumerate(t_tab["rows"]):
                if len(r_vals) >= 2:
                    rows.append((k if i == 0 else "", str(r_vals[0]).strip(), " ".join([str(x).strip() for x in r_vals[1:] if str(x).strip()])))
                elif len(r_vals) == 1:
                    rows.append((k if i == 0 else "", str(r_vals[0]).strip(), ""))
            if rows:
                return rows

    # 2. Processamento linha a linha tradicional ou caso tabela não tenha sido detectada
    for i, line in enumerate(lines):
        if '\t' in line:
            parts = [p.strip() for p in line.split('\t') if p.strip()]
            if len(parts) >= 2:
                rows.append((k if len(rows) == 0 else "", parts[0], " ".join(parts[1:])))
                continue
            elif len(parts) == 1:
                line = parts[0]
        if '|' in line:
            parts = [p.strip() for p in line.strip('|').split('|') if p.strip()]
            if len(parts) >= 2:
                rows.append((k if len(rows) == 0 else "", parts[0], " ".join(parts[1:])))
                continue
            elif len(parts) == 1:
                line = parts[0]
        
        m = re.search(r'^(.*?)\s+((?:\d{2}\.|\d{4})[\d\.\s\-]{4,}.*)$', line)
        if m:
            rows.append((k if len(rows) == 0 else "", m.group(1).strip(), m.group(2).strip()))
        else:
            if k == "Exercício":
                rows.append((k, line, ""))
            else:
                rows.append((k if len(rows) == 0 else "", line, ""))
    return rows

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
                f"2.1. O prazo de vigência deste Contrato é de {dados['vigencia_descricao']}, com início na data de sua assinatura e término previsto em {formatar_data_br(dados['data_vencimento'])}.",
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
                    tab_din = dados.get("tabela_dinamica") or dados.get("tabela_itens_dinamica")
                    if not tab_din and hasattr(dados.get("itens"), "tabela_dinamica"):
                        tab_din = dados["itens"].tabela_dinamica
                    if not tab_din:
                        tab_din = {}

                    if tab_din and tab_din.get("rows"):
                        dyn_rows = tab_din["rows"]
                        dyn_headers = tab_din.get("headers") or [f"Coluna {i+1}" for i in range(len(dyn_rows[0]))]
                        dyn_headers, dyn_rows = reconciliar_tabela_irregular(dyn_headers, dyn_rows, is_dotacao=False)
                        has_real_hdrs = tab_din.get("has_real_headers")
                        if has_real_hdrs is None:
                            has_real_hdrs = bool(dyn_headers and not all(str(h).strip().startswith("COLUNA ") for h in dyn_headers))
                        col_w_docx, dyn_alignments = calcular_larguras_colunas_adaptativas(dyn_headers, dyn_rows, total_width_pt=485.0, is_dotacao=False)
                        tot_col = tab_din.get("total_col_idx")
                        tot_val_txt = tab_din.get("total_formatado") or tot_display
                    elif dados.get("itens") and len(dados["itens"]) > 0:
                        itens_list = dados["itens"]
                        first_it = itens_list[0] if isinstance(itens_list[0], dict) else {}
                        known_agu_keys = ["item", "descricao", "quantidade", "unidade", "valor_unitario", "valor_total"]
                        internal_keys = ["_vu_float", "_vt_float", "_qtd_float"]
                        custom_keys = [k for k in first_it.keys() if k not in known_agu_keys and k not in internal_keys]
                        
                        if custom_keys:
                            dyn_headers = [k for k in first_it.keys() if k not in internal_keys]
                            dyn_rows = [[str(it.get(k, "")) for k in dyn_headers] for it in itens_list]
                        else:
                            cand_cols = [
                                ("item", "Item"),
                                ("descricao", "Descrição do Objeto / Especificação"),
                                ("quantidade", "Qtd"),
                                ("unidade", "Unid."),
                                ("valor_unitario", "Valor Unitário"),
                                ("valor_total", "Valor Total")
                            ]
                            active_cols = []
                            for k_field, lbl in cand_cols:
                                has_val = any(bool(str(it.get(k_field, "")).strip()) for it in itens_list)
                                if has_val or k_field in ("item", "descricao", "valor_total"):
                                    active_cols.append((k_field, lbl))
                            dyn_headers = [lbl for _, lbl in active_cols]
                            dyn_rows = [[str(it.get(kf, "")) for kf, _ in active_cols] for it in itens_list]
                        has_real_hdrs = True
                        col_w_docx, dyn_alignments = calcular_larguras_colunas_adaptativas(dyn_headers, dyn_rows, total_width_pt=485.0, is_dotacao=False)
                        tot_col = len(dyn_headers) - 1
                        tot_val_txt = tot_display
                    else:
                        dyn_rows = []
                        dyn_headers = []
                        has_real_hdrs = False
                        col_w_docx = []
                        dyn_alignments = []
                        tot_col = None
                        tot_val_txt = tot_display

                    n_cols = len(dyn_headers) if dyn_headers else (len(dyn_rows[0]) if dyn_rows else 0)

                    has_summary_row = tab_din.get("has_summary_row") or any(is_summary_row(r) for r in dyn_rows)
                    has_tot_row = (not has_summary_row) and (tot_col is not None and tot_col < n_cols and bool(tab_din.get("append_total_row", False)))
                    n_rows = len(dyn_rows) + (1 if has_real_hdrs else 0) + (1 if has_tot_row else 0)
                    t_it = doc.add_table(rows=n_rows, cols=n_cols)
                    aplicar_estilo_tabela_oficial_docx(t_it, col_widths_pt=col_w_docx)

                    row_offset = 0
                    if has_real_hdrs:
                        # Cabeçalho dinâmico adaptativo (#002F6E fundo com texto branco em negrito)
                        hdr_cells = t_it.rows[0].cells
                        for c_idx, h_text in enumerate(dyn_headers):
                            h_str = "" if str(h_text).startswith("_") else str(h_text)
                            h_str_docx = h_str.replace("<br>", "\n").replace("<br/>", "\n").replace("<BR>", "\n")
                            hdr_cells[c_idx].text = h_str_docx
                            aplicar_fundo_celula_docx(hdr_cells[c_idx], '002F6E')
                            aplicar_bordas_celula(hdr_cells[c_idx])
                            for p in hdr_cells[c_idx].paragraphs:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                p.paragraph_format.space_before = Pt(2.5)
                                p.paragraph_format.space_after = Pt(2.5)
                                p.paragraph_format.line_spacing = 1.15
                                for r in p.runs:
                                    r.font.name = "Arial"
                                    r.font.size = Pt(8.5) if n_cols <= 6 else Pt(7.5)
                                    r.bold = True
                                    r.font.color.rgb = RGBColor(255, 255, 255)
                        row_offset = 1

                    # Linhas de dados (fundo alternado #F8FAFC / #FFFFFF, bordas finas e quebra de texto)
                    for r_idx, row_vals in enumerate(dyn_rows):
                        row_cells = t_it.rows[r_idx + row_offset].cells
                        is_res = is_summary_row(row_vals)
                        bg_color = 'F1F5F9' if is_res else ('EBF3FB' if (r_idx + 1) % 2 == 0 else 'FFFFFF')
                        for c in row_cells:
                            aplicar_fundo_celula_docx(c, bg_color)
                        for c_idx, cell_val in enumerate(row_vals):
                            if c_idx < len(row_cells):
                                al = dyn_alignments[c_idx] if c_idx < len(dyn_alignments) else "LEFT"
                                if is_res and any(k in str(dyn_headers[c_idx]).upper() for k in ['VALOR', 'PREÇO', 'PRECO', 'TOTAL', 'R$']):
                                    al = "RIGHT"
                                elif al == "JUSTIFY" or (c_idx == 1 and not is_res):
                                    al = "JUSTIFY"

                                if al == "JUSTIFY":
                                    val_docx = normalizar_quebras_texto_celula(str(cell_val))
                                else:
                                    val_docx = str(cell_val).replace("<br>", "\n").replace("<br/>", "\n")

                                row_cells[c_idx].text = val_docx
                                aplicar_bordas_celula(row_cells[c_idx])
                                for p in row_cells[c_idx].paragraphs:
                                    if al == "RIGHT":
                                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                                    elif al == "CENTER":
                                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    elif al == "JUSTIFY":
                                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                                    else:
                                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                    p.paragraph_format.space_before = Pt(2.5)
                                    p.paragraph_format.space_after = Pt(2.5)
                                    p.paragraph_format.line_spacing = 1.05
                                    for r in p.runs:
                                        r.font.name = "Arial"
                                        r.font.size = Pt(8.0) if n_cols <= 6 else Pt(7.5)
                                        if is_res:
                                            r.bold = True
                                            r.font.color.rgb = RGBColor(0, 47, 110)
                                        else:
                                            r.font.color.rgb = RGBColor(30, 41, 59)

                    # Linha de total adicional (se explicitamente configurada e ausente na planilha original)
                    if has_tot_row:
                        tot_row_cells = t_it.rows[-1].cells
                        for c in tot_row_cells:
                            aplicar_fundo_celula_docx(c, 'F1F5F9')
                            aplicar_bordas_celula(c)
                        merge_end = tot_col - 1 if tot_col > 0 else 0
                        a = tot_row_cells[0]
                        b = tot_row_cells[merge_end]
                        if merge_end > 0:
                            a.merge(b)
                        a.text = "TOTAL GERAL:"
                        for p_lbl in a.paragraphs:
                            p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            p_lbl.paragraph_format.space_before = Pt(3.0)
                            p_lbl.paragraph_format.space_after = Pt(3.0)
                            for r in p_lbl.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(8.5)
                                r.bold = True
                                r.font.color.rgb = RGBColor(0, 47, 110)

                        tot_row_cells[tot_col].text = tot_val_txt
                        for p_val in tot_row_cells[tot_col].paragraphs:
                            p_val.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            p_val.paragraph_format.space_before = Pt(3.0)
                            p_val.paragraph_format.space_after = Pt(3.0)
                            for r in p_val.runs:
                                r.font.name = "Arial"
                                r.font.size = Pt(8.5)
                                r.bold = True
                                r.font.color.rgb = RGBColor(0, 47, 110)
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
                    val_unidade = str(dados.get("dotacao_unidade") or "")
                    dot_tab = dados.get("dotacao_tabela")
                    has_table = False
                    mixed_blocks = []
                    if dot_tab and isinstance(dot_tab, dict) and dot_tab.get("headers") and dot_tab.get("rows"):
                        mixed_blocks = [("table", dot_tab)]
                        has_table = True
                    else:
                        val_orc = str(dados.get("dotacao_orcamentaria") or "")
                        if val_orc and val_orc.strip().startswith("{") and "headers" in val_orc:
                            try:
                                import json
                                d_json = json.loads(val_orc)
                                if isinstance(d_json, dict) and d_json.get("headers") and d_json.get("rows"):
                                    mixed_blocks = [("table", d_json)]
                                    has_table = True
                            except Exception:
                                pass
                        if not has_table:
                            mixed_blocks = parse_mixed_markdown(val_orc)
                            has_table = any(b[0] == 'table' for b in mixed_blocks)
                    
                    if has_table:
                        for btype, bdata in mixed_blocks:
                            if btype == "text":
                                for txt_line in bdata.split('\n'):
                                    if not txt_line.strip():
                                        continue
                                    p_c = doc.add_paragraph()
                                    p_c.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                                    p_c.paragraph_format.first_line_indent = Inches(0.4)
                                    p_c.paragraph_format.line_spacing = 1.15
                                    r_c = p_c.add_run(txt_line.strip())
                                    r_c.font.name = "Times New Roman"
                                    r_c.font.size = Pt(10)
                            elif btype == "table":
                                headers = [str(h).strip() for h in bdata["headers"]]
                                rows = bdata["rows"]
                                n_cols = len(headers)
                                
                                # Verifica se é tabela estrutural de dotação oficial
                                clean_rows = []
                                for r_i, r in enumerate(rows):
                                    c1_chk = str(r[1] or '').strip() if len(r) > 1 else ''
                                    c2_chk = str(r[2] or '').strip() if len(r) > 2 else ''
                                    if r_i > 0 and not c1_chk and not c2_chk:
                                        continue
                                    clean_rows.append(r)
                                rows = clean_rows
                                is_st, start_unid, end_unid, has_ex = is_structural_budget_table(headers, rows)
                                
                                if is_st:
                                    t_dot = doc.add_table(rows=len(rows), cols=3)
                                    t_dot.alignment = WD_TABLE_ALIGNMENT.CENTER
                                    aplicar_estilo_tabela_oficial_docx(t_dot, col_widths_pt=[130.0, 240.0, 120.0])
                                    
                                    for r_idx, row_vals in enumerate(rows):
                                        row_cells = t_dot.rows[r_idx].cells
                                        c0_val = str(row_vals[0] if len(row_vals) > 0 else "")
                                        c1_val = str(row_vals[1] if len(row_vals) > 1 else "")
                                        c2_val = str(row_vals[2] if len(row_vals) > 2 else "")
                                        
                                        if not c0_val.strip() and c1_val.strip():
                                            c0_val = "Unidade Orçamentária\n/Atividade"

                                        for c_idx, val in enumerate([c0_val, c1_val, c2_val]):
                                            cell = row_cells[c_idx]
                                            cell.text = str(val).replace("<br>", "\n").replace("<br/>", "\n")
                                            if c_idx == 0 or r_idx == 0:
                                                c_bg = 'E2EBF4'
                                            else:
                                                c_bg = 'EBF3FB' if (r_idx % 2 == 1) else 'FFFFFF'
                                            aplicar_fundo_celula_docx(cell, c_bg)
                                            aplicar_bordas_celula(cell, border_color="224061", border_sz="8")
                                            p = cell.paragraphs[0]
                                            p.paragraph_format.line_spacing = 1.15
                                            p.paragraph_format.space_before = Pt(2.5)
                                            p.paragraph_format.space_after = Pt(2.5)
                                            if c_idx == 0:
                                                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                            elif c_idx == 1 and r_idx == 0:
                                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                            elif c_idx == 1:
                                                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                                            else:
                                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                            for r in p.runs:
                                                r.font.name = "Arial"
                                                r.font.size = Pt(8.5)
                                                r.font.color.rgb = RGBColor(34, 64, 97)
                                                if c_idx == 0 or r_idx == 0:
                                                    r.bold = True
                                    
                                    # Mesclagem horizontal do Exercício (Col 1 e Col 2)
                                    t_dot.rows[0].cells[1].merge(t_dot.rows[0].cells[2])
                                    ano_txt = str(rows[0][1] if len(rows[0]) > 1 and str(rows[0][1]).strip() else (rows[0][2] if len(rows[0]) > 2 else ""))
                                    t_dot.rows[0].cells[1].text = ano_txt
                                    aplicar_fundo_celula_docx(t_dot.rows[0].cells[1], 'E2EBF4')
                                    aplicar_bordas_celula(t_dot.rows[0].cells[1], border_color="224061", border_sz="8")
                                    p_ex = t_dot.rows[0].cells[1].paragraphs[0]
                                    p_ex.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    for r in p_ex.runs:
                                        r.font.name = "Arial"
                                        r.font.size = Pt(8.5)
                                        r.font.color.rgb = RGBColor(34, 64, 97)
                                        r.bold = True

                                    # Mesclagem vertical da Unidade Orçamentária na Coluna 0 (ZERO LINHAS NA COLUNA 0!)
                                    if start_unid != -1 and end_unid > start_unid:
                                        c_top = t_dot.cell(start_unid, 0)
                                        c_bot = t_dot.cell(end_unid, 0)
                                        c_top.merge(c_bot)
                                        c_top.text = "Unidade Orçamentária\n/Atividade"
                                        c_top.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                                        aplicar_fundo_celula_docx(c_top, 'E2EBF4')
                                        aplicar_bordas_celula(c_top, border_color="224061", border_sz="8")
                                        p_top = c_top.paragraphs[0]
                                        p_top.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                        p_top.paragraph_format.line_spacing = 1.15
                                        for r in p_top.runs:
                                            r.font.name = "Arial"
                                            r.font.size = Pt(8.5)
                                            r.font.color.rgb = RGBColor(34, 64, 97)
                                            r.bold = True
                                    elif start_unid != -1 and end_unid == start_unid:
                                        c_top = t_dot.cell(start_unid, 0)
                                        c_top.text = "Unidade Orçamentária\n/Atividade"
                                        c_top.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                                        aplicar_fundo_celula_docx(c_top, 'E2EBF4')
                                        aplicar_bordas_celula(c_top, border_color="224061", border_sz="8")
                                        p_top = c_top.paragraphs[0]
                                        p_top.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                        p_top.paragraph_format.line_spacing = 1.15
                                        for r in p_top.runs:
                                            r.font.name = "Arial"
                                            r.font.size = Pt(8.5)
                                            r.font.color.rgb = RGBColor(34, 64, 97)
                                            r.bold = True
                                else:
                                    has_hdr = bdata.get("has_real_headers")
                                    if has_hdr is None:
                                        has_hdr = headers and not all(str(h).strip().startswith("COLUNA ") for h in headers)
                                    n_rows = len(rows) + (1 if has_hdr else 0)
                                    t_dot = doc.add_table(rows=n_rows, cols=n_cols)
                                    t_dot.alignment = WD_TABLE_ALIGNMENT.CENTER
                                    
                                    col_w_pts = bdata.get("col_widths")
                                    if not col_w_pts or len(col_w_pts) != n_cols:
                                        col_weights = []
                                        for ci, h in enumerate(headers):
                                            m_len = max(len(str(h)), max([len(str(r[ci])) for r in rows if ci < len(r)] or [0]))
                                            col_weights.append(max(float(m_len), 4.0))
                                        s_w = sum(col_weights) if sum(col_weights) > 0 else (n_cols * 4.0)
                                        col_w_pts = [(w / s_w) * 480.0 for w in col_weights]
                                    aplicar_estilo_tabela_oficial_docx(t_dot, col_widths_pt=col_w_pts)
                                    
                                    dot_offset = 0
                                    if has_hdr:
                                        hdr_cells = t_dot.rows[0].cells
                                        for c_idx in range(n_cols):
                                            h_text = headers[c_idx] if c_idx < len(headers) else ""
                                            hdr_cells[c_idx].text = str(h_text)
                                            aplicar_fundo_celula_docx(hdr_cells[c_idx], 'E2EBF4')
                                            aplicar_bordas_celula(hdr_cells[c_idx], border_color="224061", border_sz="8")
                                            for p in hdr_cells[c_idx].paragraphs:
                                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                                for r in p.runs:
                                                    r.font.name = "Arial"
                                                    r.font.size = Pt(8.5)
                                                    r.bold = True
                                                    r.font.color.rgb = RGBColor(34, 64, 97)
                                        dot_offset = 1
                                            
                                    col_alignments = bdata.get("col_alignments", [])
                                    for r_idx, row_vals in enumerate(rows):
                                        row_cells = t_dot.rows[r_idx + dot_offset].cells
                                        is_banner = (len(row_vals) > 1 and str(row_vals[0] or '').strip() and not any(str(x or '').strip() for x in row_vals[1:]))
                                        if is_banner:
                                            row_cells[0].merge(row_cells[-1])
                                            row_cells[0].text = str(row_vals[0] or '').replace("<br>", "\n").replace("<br/>", "\n")
                                            aplicar_fundo_celula_docx(row_cells[0], 'E2EBF4')
                                            aplicar_bordas_celula(row_cells[0], border_color="224061", border_sz="8")
                                            p = row_cells[0].paragraphs[0]
                                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                            for r in p.runs:
                                                r.font.name = "Arial"
                                                r.font.size = Pt(8.5)
                                                r.bold = True
                                                r.font.color.rgb = RGBColor(34, 64, 97)
                                            continue

                                        for c_idx in range(n_cols):
                                            cell_val = row_vals[c_idx] if c_idx < len(row_vals) else ""
                                            cell = row_cells[c_idx]
                                            cell.text = str(cell_val).replace("<br>", "\n").replace("<br/>", "\n")
                                            if c_idx == 0:
                                                c_bg = 'E2EBF4'
                                            else:
                                                c_bg = 'EBF3FB' if (r_idx % 2 == 1) else 'FFFFFF'
                                            aplicar_fundo_celula_docx(cell, c_bg)
                                            aplicar_bordas_celula(cell, border_color="224061", border_sz="8")
                                            al_code = col_alignments[c_idx] if c_idx < len(col_alignments) else "LEFT"
                                            for p in cell.paragraphs:
                                                if al_code == "CENTER":
                                                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                                elif al_code == "RIGHT":
                                                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                                                elif al_code == "JUSTIFY":
                                                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                                                else:
                                                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                                for r in p.runs:
                                                    r.font.name = "Arial"
                                                    r.font.size = Pt(8)
                                                    r.font.color.rgb = RGBColor(34, 64, 97)

                                    # Se for dotação com 3 colunas e a linha for Exercício sem código, mescla colunas 1 e 2
                                    if is_clausula_dotacao and n_cols == 3:
                                        for r_i, r_vals in enumerate(rows):
                                            if r_vals and str(r_vals[0]).strip().upper().startswith("EXERC") and (len(r_vals) < 3 or not str(r_vals[2]).strip()):
                                                r_c = t_dot.rows[r_i + 1].cells
                                                r_c[1].merge(r_c[2])
                                                aplicar_fundo_celula_docx(r_c[1], 'E2EBF4')
                                                aplicar_bordas_celula(r_c[1], border_color="224061", border_sz="8")
                    else:
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

                        expanded_items = []
                        for k, v in dot_items:
                            expanded_items.extend(parse_dotacao_item_multi(k, v))

                        t_dot = doc.add_table(rows=len(expanded_items), cols=3)
                        t_dot.alignment = WD_TABLE_ALIGNMENT.CENTER
                        aplicar_estilo_tabela_oficial_docx(t_dot, col_widths_pt=[130.0, 240.0, 120.0])
                        for r_i, (k, v1, v2) in enumerate(expanded_items):
                            c0 = t_dot.cell(r_i, 0)
                            c1 = t_dot.cell(r_i, 1)
                            c2 = t_dot.cell(r_i, 2)
                            
                            aplicar_fundo_celula_docx(c0, 'E2EBF4')
                            bg_23 = 'E2EBF4' if r_i == 0 else 'FFFFFF'
                            aplicar_fundo_celula_docx(c1, bg_23)
                            aplicar_fundo_celula_docx(c2, bg_23)
                            
                            if r_i == 0:
                                c1.merge(c2)
                            
                            aplicar_bordas_celula(c0, border_color="224061", border_sz="8")
                            aplicar_bordas_celula(c1, border_color="224061", border_sz="8")
                            if r_i != 0:
                                aplicar_bordas_celula(c2, border_color="224061", border_sz="8")
                            
                            c0.paragraphs[0].text = ""
                            r0 = c0.paragraphs[0].add_run(k)
                            r0.bold = True
                            r0.font.name = "Arial"
                            r0.font.size = Pt(8.5)
                            r0.font.color.rgb = RGBColor(34, 64, 97)
                            c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                            
                            c1.paragraphs[0].text = ""
                            p1 = c1.paragraphs[0]
                            r1 = p1.add_run(v1)
                            r1.font.name = "Arial"
                            r1.font.size = Pt(8.5)
                            r1.font.color.rgb = RGBColor(34, 64, 97)
                            
                            if r_i == 0:
                                r1.bold = True
                                p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            else:
                                p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                c2.paragraphs[0].text = ""
                                p2 = c2.paragraphs[0]
                                r2 = p2.add_run(v2)
                                r2.font.name = "Arial"
                                r2.font.size = Pt(8.5)
                                r2.font.color.rgb = RGBColor(34, 64, 97)
                                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
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
        leading=10.0,
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
    style_tbl_td_j = ParagraphStyle(
        'TblTDJ',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#1e293b"),
        alignment=4
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
                    
                    tab_din = dados.get("tabela_dinamica") or dados.get("tabela_itens_dinamica")
                    if not tab_din and hasattr(dados.get("itens"), "tabela_dinamica"):
                        tab_din = dados["itens"].tabela_dinamica
                    if not tab_din:
                        tab_din = {}

                    if tab_din and tab_din.get("rows"):
                        dyn_rows = tab_din["rows"]
                        dyn_headers = tab_din.get("headers") or [f"Coluna {i+1}" for i in range(len(dyn_rows[0]))]
                        has_real_hdrs = tab_din.get("has_real_headers")
                        if has_real_hdrs is None:
                            has_real_hdrs = bool(dyn_headers and not all(str(h).strip().startswith("COLUNA ") for h in dyn_headers))
                        col_widths, dyn_alignments = calcular_larguras_colunas_adaptativas(dyn_headers, dyn_rows, total_width_pt=515.0, is_dotacao=False)
                        tot_col = tab_din.get("total_col_idx")
                        tot_val_txt = tab_din.get("total_formatado") or tot_display
                    elif dados.get("itens") and len(dados["itens"]) > 0:
                        itens_list = dados["itens"]
                        first_it = itens_list[0] if isinstance(itens_list[0], dict) else {}
                        known_agu_keys = ["item", "descricao", "quantidade", "unidade", "valor_unitario", "valor_total"]
                        internal_keys = ["_vu_float", "_vt_float", "_qtd_float"]
                        custom_keys = [k for k in first_it.keys() if k not in known_agu_keys and k not in internal_keys]
                        
                        if custom_keys:
                            dyn_headers = [k for k in first_it.keys() if k not in internal_keys]
                            dyn_rows = [[str(it.get(k, "")) for k in dyn_headers] for it in itens_list]
                        else:
                            cand_cols = [
                                ("item", "Item"),
                                ("descricao", "Descrição do Objeto / Especificação"),
                                ("quantidade", "Qtd"),
                                ("unidade", "Unid."),
                                ("valor_unitario", "Valor Unitário"),
                                ("valor_total", "Valor Total")
                            ]
                            active_cols = []
                            for k_field, lbl in cand_cols:
                                has_val = any(bool(str(it.get(k_field, "")).strip()) for it in itens_list)
                                if has_val or k_field in ("item", "descricao", "valor_total"):
                                    active_cols.append((k_field, lbl))
                            dyn_headers = [lbl for _, lbl in active_cols]
                            dyn_rows = [[str(it.get(kf, "")) for kf, _ in active_cols] for it in itens_list]
                        has_real_hdrs = True
                        col_widths, dyn_alignments = calcular_larguras_colunas_adaptativas(dyn_headers, dyn_rows, total_width_pt=515.0, is_dotacao=False)
                        tot_col = len(dyn_headers) - 1
                        tot_val_txt = tot_display
                    else:
                        dyn_rows = []
                        dyn_headers = []
                        has_real_hdrs = False
                        col_widths = []
                        dyn_alignments = []
                        tot_col = None
                        tot_val_txt = tot_display

                    n_cols = len(dyn_headers) if dyn_headers else (len(dyn_rows[0]) if dyn_rows else 0)

                    has_summary_row = tab_din.get("has_summary_row") or any(is_summary_row(r) for r in dyn_rows)
                    has_tot = (not has_summary_row) and (tot_col is not None and tot_col < n_cols and bool(tab_din.get("append_total_row", False)))

                    table_itens_data = []
                    if has_real_hdrs:
                        th_cells = []
                        for h in dyn_headers:
                            h_clean = "" if str(h).startswith("_") else str(h)
                            h_pdf = h_clean.replace('\r\n', '<br/>').replace('\n', '<br/>').replace('<br>', '<br/>')
                            th_cells.append(Paragraph(h_pdf, style_tbl_th))
                        table_itens_data.append(th_cells)
                    for r_vals in dyn_rows:
                        is_res = is_summary_row(r_vals)
                        row_cells = []
                        for ci in range(n_cols):
                            val_str = str(r_vals[ci]) if ci < len(r_vals) else ""
                            align_code = dyn_alignments[ci] if ci < len(dyn_alignments) else "LEFT"
                            if is_res and any(k in str(dyn_headers[ci]).upper() for k in ['VALOR', 'PREÇO', 'PRECO', 'TOTAL', 'R$']):
                                align_code = "RIGHT"
                            elif align_code == "JUSTIFY" or (ci == 1 and not is_res):
                                align_code = "JUSTIFY"
                                val_str = normalizar_quebras_texto_celula(val_str)

                            style_cell = style_tbl_td_c if align_code == "CENTER" else (
                                style_tbl_td_r if align_code == "RIGHT" else (style_tbl_td_j if align_code == "JUSTIFY" else style_tbl_td)
                            )
                            if is_res:
                                cell_p = Paragraph(f"<b>{val_str}</b>", style_cell)
                            else:
                                cell_p = Paragraph(val_str.replace("\n", "<br/>").replace("<br>", "<br/>"), style_cell)
                            row_cells.append(cell_p)
                        table_itens_data.append(row_cells)

                    if has_tot:
                        footer_row = ["" for _ in range(n_cols)]
                        footer_row[0] = Paragraph("<b>TOTAL GERAL:</b>", ParagraphStyle('TblTotL', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=2, textColor=colors.HexColor("#002F6E")))
                        footer_row[tot_col] = Paragraph(f"<b>{tot_val_txt}</b>", ParagraphStyle('TblTotV', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=2, textColor=colors.HexColor("#002F6E")))
                        table_itens_data.append(footer_row)

                    t_itens = Table(table_itens_data, colWidths=col_widths, repeatRows=1 if has_real_hdrs else 0)
                    t_styles = [
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                        ('TOPPADDING', (0, 0), (-1, -1), 3),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                        ('LEFTPADDING', (0, 0), (-1, -1), 4),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                    ]
                    if has_real_hdrs:
                        t_styles.append(('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#002F6E")))
                        t_styles.append(('ALIGN', (0, 0), (-1, 0), 'CENTER'))

                    start_row_idx = 1 if has_real_hdrs else 0
                    for r_offset, r_vals in enumerate(dyn_rows):
                        r_i = start_row_idx + r_offset
                        if is_summary_row(r_vals):
                            t_styles.append(('BACKGROUND', (0, r_i), (-1, r_i), colors.HexColor("#F1F5F9")))
                            t_styles.append(('LINEABOVE', (0, r_i), (-1, r_i), 1.0, colors.HexColor("#002F6E")))
                            t_styles.append(('LINEBELOW', (0, r_i), (-1, r_i), 1.0, colors.HexColor("#002F6E")))
                        else:
                            bg = colors.HexColor("#EBF3FB") if (r_offset + 1) % 2 == 0 else colors.HexColor("#FFFFFF")
                            t_styles.append(('BACKGROUND', (0, r_i), (-1, r_i), bg))

                    if has_tot:
                        if tot_col > 0:
                            t_styles.append(('SPAN', (0, -1), (tot_col - 1, -1)))
                        t_styles.extend([
                            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#F1F5F9")),
                            ('LINEABOVE', (0, -1), (-1, -1), 1.2, colors.HexColor("#002F6E")),
                            ('LINEBELOW', (0, -1), (-1, -1), 1.2, colors.HexColor("#002F6E")),
                            ('GRID', (0, -1), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                        ])
                    t_itens.setStyle(TableStyle(t_styles))
                    story.append(t_itens)
                    story.append(Spacer(1, 4))
        # Inserir tabela de dotação orçamentária na Cláusula Quarta
        elif is_clausula_dotacao:
            for idx_d, d_txt in enumerate(cl["conteudo"]):
                story.append(Paragraph(d_txt, style_clause_body))
                if idx_d == 0:
                    story.append(Spacer(1, 3))
                    val_unidade = str(dados.get("dotacao_unidade") or "")
                    dot_tab = dados.get("dotacao_tabela")
                    has_table = False
                    mixed_blocks = []
                    if dot_tab and isinstance(dot_tab, dict) and dot_tab.get("headers") and dot_tab.get("rows"):
                        mixed_blocks = [("table", dot_tab)]
                        has_table = True
                    else:
                        val_orc = str(dados.get("dotacao_orcamentaria") or "")
                        if val_orc and val_orc.strip().startswith("{") and "headers" in val_orc:
                            try:
                                import json
                                d_json = json.loads(val_orc)
                                if isinstance(d_json, dict) and d_json.get("headers") and d_json.get("rows"):
                                    mixed_blocks = [("table", d_json)]
                                    has_table = True
                            except Exception:
                                pass
                        if not has_table:
                            mixed_blocks = parse_mixed_markdown(val_orc)
                            has_table = any(b[0] == 'table' for b in mixed_blocks)
                    
                    if has_table:
                        for btype, bdata in mixed_blocks:
                            if btype == "text":
                                for txt_line in bdata.split('\n'):
                                    if txt_line.strip():
                                        story.append(Paragraph(txt_line.strip(), style_clause_body))
                            elif btype == "table":
                                headers = [str(h).strip() for h in bdata["headers"]]
                                rows = bdata["rows"]
                                n_cols = len(headers)
                                col_alignments = bdata.get("col_alignments", [])
                                
                                # Verifica se é tabela estrutural de dotação oficial
                                clean_rows = []
                                for r_i, r in enumerate(rows):
                                    c1_chk = str(r[1] or '').strip() if len(r) > 1 else ''
                                    c2_chk = str(r[2] or '').strip() if len(r) > 2 else ''
                                    if r_i > 0 and not c1_chk and not c2_chk:
                                        continue
                                    clean_rows.append(r)
                                rows = clean_rows
                                is_st, start_unid, end_unid, has_ex = is_structural_budget_table(headers, rows)
                                has_unid_span = (start_unid != -1 and end_unid >= start_unid)
                                
                                if is_st:
                                    t_dot_data = []
                                    for r_idx, row_vals in enumerate(rows):
                                        c0_val = str(row_vals[0] if len(row_vals) > 0 else "")
                                        c1_val = str(row_vals[1] if len(row_vals) > 1 else "")
                                        c2_val = str(row_vals[2] if len(row_vals) > 2 else "")
                                        
                                        d_row = []
                                        # Col 0: Título/rótulo em negrito com alinhamento à esquerda
                                        if has_unid_span and (start_unid <= r_idx <= end_unid):
                                            if r_idx == start_unid:
                                                d_row.append(Paragraph(f"<b><font color='#224061'>Unidade Orçamentária<br/>/Atividade</font></b>", style_tbl_td))
                                            else:
                                                d_row.append(Paragraph("", style_tbl_td))
                                        else:
                                            c0_clean = c0_val.replace('\n', '<br/>').replace('<br>', '<br/>')
                                            if not c0_clean.strip() and c1_val.strip():
                                                c0_clean = "Unidade Orçamentária<br/>/Atividade"
                                            d_row.append(Paragraph(f"<b><font color='#224061'>{c0_clean}</font></b>", style_tbl_td))
                                        
                                        # Col 1 e Col 2
                                        c1_clean = c1_val.replace('\n', '<br/>').replace('<br>', '<br/>')
                                        if r_idx == 0:
                                            d_row.append(Paragraph(f"<b><font color='#224061'>{c1_clean}</font></b>", style_tbl_td_c))
                                            d_row.append(Paragraph("", style_tbl_td_c))
                                        else:
                                            d_row.append(Paragraph(f"<font color='#224061'>{c1_clean}</font>", style_tbl_td_j))
                                            c2_clean = c2_val.replace('\n', '<br/>').replace('<br>', '<br/>')
                                            d_row.append(Paragraph(f"<font color='#224061'>{c2_clean}</font>", style_tbl_td_c))
                                        
                                        t_dot_data.append(d_row)
                                    
                                    col_widths = [125.0, 265.0, 110.0]
                                    t_dot = Table(t_dot_data, colWidths=col_widths, repeatRows=0)
                                    t_dot_style = [
                                        ('SPAN', (1, 0), (2, 0)),
                                        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#FFFFFF")),
                                        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor("#E2EBF4")),
                                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E2EBF4")),
                                        ('GRID', (0, 0), (-1, -1), 0.8, colors.HexColor("#224061")),
                                        ('BOX', (0, 0), (-1, -1), 1.25, colors.HexColor("#224061")),
                                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                                        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                                        ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                                        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
                                        ('TOPPADDING', (0, 0), (-1, -1), 3),
                                        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                                        ('LEFTPADDING', (0, 0), (-1, -1), 5),
                                        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                                    ]
                                    if start_unid != -1 and end_unid > start_unid:
                                        t_dot_style.append(('SPAN', (0, start_unid), (0, end_unid)))
                                        t_dot_style.append(('VALIGN', (0, start_unid), (0, end_unid), 'MIDDLE'))
                                        t_dot_style.append(('BACKGROUND', (0, start_unid), (0, end_unid), colors.HexColor("#E2EBF4")))
                                    elif start_unid != -1 and end_unid == start_unid:
                                        t_dot_style.append(('BACKGROUND', (0, start_unid), (0, start_unid), colors.HexColor("#E2EBF4")))
                                        t_dot_style.append(('VALIGN', (0, start_unid), (0, start_unid), 'MIDDLE'))
                                
                                    for r_idx in range(1, len(rows)):
                                        if r_idx % 2 == 1:
                                            t_dot_style.append(('BACKGROUND', (1, r_idx), (2, r_idx), colors.HexColor("#EBF3FB")))
                                    t_dot.setStyle(TableStyle(t_dot_style))
                                    story.append(t_dot)
                                    story.append(Spacer(1, 4))
                                else:
                                    has_hdr = bdata.get("has_real_headers")
                                    if has_hdr is None:
                                        has_hdr = headers and not all(str(h).strip().startswith("COLUNA ") for h in headers)
                                    t_dot_data = []
                                    t_dot_style = []
                                    if has_hdr:
                                        h_row = []
                                        for h in headers:
                                            h_row.append(Paragraph(f"<b><font color='#224061'>{h}</font></b>", style_tbl_td_c))
                                        t_dot_data.append(h_row)
                                        t_dot_style.append(('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E2EBF4")))
                                    for r_idx, r_vals in enumerate(rows):
                                        is_banner = (len(r_vals) > 1 and str(r_vals[0] or '').strip() and not any(str(x or '').strip() for x in r_vals[1:]))
                                        if is_banner:
                                            b_clean = str(r_vals[0] or '').replace('\n', '<br/>').replace('<br>', '<br/>')
                                            d_row = [Paragraph(f"<b><font color='#224061'>{b_clean}</font></b>", style_tbl_td_c)]
                                            d_row.extend([Paragraph("", style_tbl_td)] * (n_cols - 1))
                                            t_dot_data.append(d_row)
                                            t_dot_style.append(('SPAN', (0, r_idx + 1), (-1, r_idx + 1)))
                                            t_dot_style.append(('BACKGROUND', (0, r_idx + 1), (-1, r_idx + 1), colors.HexColor("#E2EBF4")))
                                            t_dot_style.append(('ALIGN', (0, r_idx + 1), (-1, r_idx + 1), 'CENTER'))
                                            continue

                                        d_row = []
                                        for ci in range(n_cols):
                                            v = r_vals[ci] if ci < len(r_vals) else ""
                                            al_code = col_alignments[ci] if ci < len(col_alignments) else "LEFT"
                                            st_cell = style_tbl_td_c if al_code == "CENTER" else (
                                                style_tbl_td_r if al_code == "RIGHT" else (style_tbl_td_j if al_code == "JUSTIFY" else style_tbl_td)
                                            )
                                            v_clean = str(v).replace('\n', '<br/>').replace('<br>', '<br/>')
                                            d_row.append(Paragraph(f"<font color='#224061'>{v_clean}</font>", st_cell))
                                        t_dot_data.append(d_row)
                                        row_in_table = r_idx + (1 if has_hdr else 0)
                                        t_dot_style.append(('BACKGROUND', (0, row_in_table), (0, row_in_table), colors.HexColor("#E2EBF4")))
                                        bg_c = colors.HexColor("#EBF3FB") if (r_idx % 2 == 1) else colors.HexColor("#FFFFFF")
                                        t_dot_style.append(('BACKGROUND', (1, row_in_table), (-1, row_in_table), bg_c))
                                    
                                    col_widths = bdata.get("col_widths")
                                    if not col_widths or len(col_widths) != n_cols:
                                        col_weights = []
                                        for ci, h in enumerate(headers):
                                            m_len = max(len(str(h)), max([len(str(r[ci])) for r in rows if ci < len(r)] or [0]))
                                            col_weights.append(max(float(m_len), 4.0))
                                        s_w = sum(col_weights) if sum(col_weights) > 0 else (n_cols * 4.0)
                                        col_widths = [(w / s_w) * 490.0 for w in col_weights]
                                    else:
                                        s_w = sum(col_widths)
                                        if s_w > 0:
                                            col_widths = [(w / s_w) * 490.0 for w in col_widths]

                                    t_dot = Table(t_dot_data, colWidths=col_widths, repeatRows=1)

                                    t_dot_style.extend([
                                        ('GRID', (0, 0), (-1, -1), 0.8, colors.HexColor("#224061")),
                                        ('INNERGRID', (0, 0), (-1, -1), 0.8, colors.HexColor("#224061")),
                                        ('BOX', (0, 0), (-1, -1), 1.25, colors.HexColor("#224061")),
                                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                                        ('TOPPADDING', (0, 0), (-1, -1), 4),
                                        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                                        ('LEFTPADDING', (0, 0), (-1, -1), 6),
                                        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                                    ])

                                    # Se for dotação com 3 colunas e a linha for Exercício sem código, aplica SPAN entre colunas 1 e 2
                                    if is_clausula_dotacao and n_cols == 3:
                                        for r_i, r_vals in enumerate(rows):
                                            if r_vals and str(r_vals[0]).strip().upper().startswith("EXERC") and (len(r_vals) < 3 or not str(r_vals[2]).strip()):
                                                t_dot_style.append(('SPAN', (1, r_i + 1), (2, r_i + 1)))
                                                t_dot_style.append(('BACKGROUND', (1, r_i + 1), (2, r_i + 1), colors.HexColor("#E2EBF4")))

                                    t_dot.setStyle(TableStyle(t_dot_style))
                                    story.append(t_dot)
                                    story.append(Spacer(1, 4))
                    else:
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

                        expanded_items = []
                        for k, v in dot_items:
                            expanded_items.extend(parse_dotacao_item_multi(k, v))

                        t_dot_data = []
                        for r_i, (k, v1, v2) in enumerate(expanded_items):
                            if r_i == 0:
                                t_dot_data.append([
                                    Paragraph(f"<b><font color='#224061'>{k}</font></b>", style_tbl_td), 
                                    Paragraph(f"<b><font color='#224061'>{v1}</font></b>", style_tbl_td_c), 
                                    Paragraph("", style_tbl_td_c)
                                ])
                            else:
                                t_dot_data.append([
                                    Paragraph(f"<b><font color='#224061'>{k}</font></b>", style_tbl_td), 
                                    Paragraph(f"<font color='#224061'>{v1}</font>", style_tbl_td), 
                                    Paragraph(f"<font color='#224061'>{v2}</font>", style_tbl_td_c)
                                ])
                        t_dot = Table(t_dot_data, colWidths=[130.0, 240.0, 120.0])
                        t_dot.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#FFFFFF")),
                            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor("#E2EBF4")),
                            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E2EBF4")),
                            ('GRID', (0, 0), (-1, -1), 0.8, colors.HexColor("#224061")),
                            ('BOX', (0, 0), (-1, -1), 1.25, colors.HexColor("#224061")),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                            ('TOPPADDING', (0, 0), (-1, -1), 4),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                            ('SPAN', (1, 0), (2, 0)),
                            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                            ('ALIGN', (1, 0), (2, 0), 'CENTER'),
                            ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                            ('ALIGN', (2, 1), (2, -1), 'CENTER'),
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
                Paragraph(formatar_data_br(prot.get("data_rec", "—")), style_tbl_val),
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
    dt_inicio = formatar_data_br(dados.get("data_inicio_aditivo", ""))
    dt_fim = formatar_data_br(dados.get("nova_data_vencimento", ""))
    
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
