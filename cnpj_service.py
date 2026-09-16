import urllib.request
import urllib.parse
import json
import re
import logging
from typing import Dict, Any, Optional, Tuple

logger = logging.getLogger("cnpj_service")

def limpar_cnpj(cnpj: Any) -> str:
    """Remove caracteres não numéricos do CNPJ"""
    return re.sub(r'\D', '', str(cnpj or '')).strip()

def formatar_cnpj(cnpj: Any) -> str:
    """Formata CNPJ no padrão XX.XXX.XXX/XXXX-XX"""
    c = limpar_cnpj(cnpj)
    if len(c) == 14:
        return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    elif len(c) == 11:
        return f"{c[:3]}.{c[3:6]}.{c[6:9]}-{c[9:]}"
    return str(cnpj or "")

def formatar_cep(cep: Any) -> str:
    """Formata CEP no padrão XXXXX-XXX"""
    c = re.sub(r'\D', '', str(cep or '')).strip()
    if len(c) == 8:
        return f"{c[:5]}-{c[5:]}"
    return str(cep or "")

def validar_cnpj(cnpj: Any) -> bool:
    """Valida se o CNPJ tem 14 dígitos e não é uma sequência inválida"""
    c = limpar_cnpj(cnpj)
    if len(c) != 14:
        return False
    if len(set(c)) == 1:
        return False
    return True

def consultar_cnpj_api(cnpj: str, timeout: int = 8) -> Dict[str, Any]:
    """
    Consulta dados cadastrais da empresa na Receita Federal via API pública.
    1ª Tentativa: BrasilAPI (rápida, gratuita e completa)
    2ª Tentativa: ReceitaWS (fallback de alta disponibilidade)
    """
    c = limpar_cnpj(cnpj)
    if not validar_cnpj(c):
        return {
            "sucesso": False,
            "erro": f"CNPJ '{cnpj}' inválido. Certifique-se de informar os 14 dígitos numéricos."
        }

    # 1. Tentativa via BrasilAPI
    try:
        url_brasil_api = f"https://brasilapi.com.br/api/cnpj/v1/{c}"
        req = urllib.request.Request(
            url_brasil_api,
            headers={
                "User-Agent": "GestaoContratosGEL/2026 (Prefeitura Municipal de Ribeiraozinho do Maranhao)",
                "Accept": "application/json"
            }
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status == 200:
                data = json.loads(resp.read().decode("utf-8"))
                
                # Identificar representante legal a partir do QSA
                qsa = data.get("qsa", [])
                rep_nome = "Representante Legal"
                rep_cargo = "Sócio Administrador"
                for socio in qsa:
                    qual = str(socio.get("qualificacao_socio") or socio.get("qualificacao_representante_legal") or "").upper()
                    if any(k in qual for k in ["ADMINISTRADOR", "DIRETOR", "PRESIDENTE", "SÓCIO-ADMINISTRADOR", "TITULAR"]):
                        rep_nome = socio.get("nome_socio") or socio.get("nome") or rep_nome
                        rep_cargo = qual.title() if qual else "Sócio Administrador"
                        break
                if rep_nome == "Representante Legal" and qsa:
                    rep_nome = qsa[0].get("nome_socio") or qsa[0].get("nome") or rep_nome

                logr = str(data.get("logradouro") or "").strip()
                num = str(data.get("numero") or "").strip()
                end_completo = f"{logr}, {num}" if num and num.upper() != "SN" else logr

                tel = str(data.get("ddd_telefone_1") or "").strip()
                if tel:
                    tel = re.sub(r'\D', '', tel)
                    if len(tel) == 10:
                        tel = f"({tel[:2]}) {tel[2:6]}-{tel[6:]}"
                    elif len(tel) == 11:
                        tel = f"({tel[:2]}) {tel[2:7]}-{tel[7:]}"

                return {
                    "sucesso": True,
                    "cnpj": c,
                    "cnpj_formatado": formatar_cnpj(c),
                    "razao_social": str(data.get("razao_social") or "").strip(),
                    "nome_fantasia": str(data.get("nome_fantasia") or data.get("razao_social") or "").strip(),
                    "endereco": end_completo if end_completo else "Endereço Comercial",
                    "numero": num if num else "s/n",
                    "complemento": str(data.get("complemento") or "").strip(),
                    "bairro": str(data.get("bairro") or "Centro").strip(),
                    "cidade": str(data.get("municipio") or "").strip(),
                    "uf": str(data.get("uf") or "").strip(),
                    "cep": formatar_cep(data.get("cep")),
                    "telefone": tel,
                    "email": str(data.get("email") or "").strip().lower(),
                    "situacao_cadastral": str(data.get("descricao_situacao_cadastral") or "ATIVA").strip(),
                    "representante_nome": rep_nome,
                    "representante_cargo": rep_cargo,
                    "fonte": "BrasilAPI / Receita Federal"
                }
    except Exception as e_br:
        logger.warning(f"Consulta BrasilAPI falhou ({e_br}), acionando fallback ReceitaWS...")

    # 2. Tentativa via ReceitaWS (Fallback)
    try:
        url_receita_ws = f"https://receitaws.com.br/v1/cnpj/{c}"
        req_ws = urllib.request.Request(
            url_receita_ws,
            headers={
                "User-Agent": "GestaoContratosGEL/2026",
                "Accept": "application/json"
            }
        )
        with urllib.request.urlopen(req_ws, timeout=timeout) as resp_ws:
            data_ws = json.loads(resp_ws.read().decode("utf-8"))
            if data_ws.get("status") == "ERROR":
                return {
                    "sucesso": False,
                    "erro": data_ws.get("message", "CNPJ não localizado na base da Receita Federal.")
                }

            qsa_ws = data_ws.get("qsa", [])
            rep_nome = qsa_ws[0].get("nome") if qsa_ws else "Representante Legal"
            rep_cargo = qsa_ws[0].get("qual") if qsa_ws else "Sócio Administrador"
            for socio in qsa_ws:
                qual = str(socio.get("qual") or "").upper()
                if any(k in qual for k in ["ADMINISTRADOR", "DIRETOR", "PRESIDENTE"]):
                    rep_nome = socio.get("nome", rep_nome)
                    rep_cargo = qual.title()
                    break

            logr = str(data_ws.get("logradouro") or "").strip()
            num = str(data_ws.get("numero") or "").strip()
            end_completo = f"{logr}, {num}" if num and num.upper() != "SN" else logr

            return {
                "sucesso": True,
                "cnpj": c,
                "cnpj_formatado": formatar_cnpj(c),
                "razao_social": str(data_ws.get("nome") or "").strip(),
                "nome_fantasia": str(data_ws.get("fantasia") or data_ws.get("nome") or "").strip(),
                "endereco": end_completo if end_completo else "Endereço Comercial",
                "numero": num if num else "s/n",
                "complemento": str(data_ws.get("complemento") or "").strip(),
                "bairro": str(data_ws.get("bairro") or "Centro").strip(),
                "cidade": str(data_ws.get("municipio") or "").strip(),
                "uf": str(data_ws.get("uf") or "").strip(),
                "cep": formatar_cep(data_ws.get("cep")),
                "telefone": str(data_ws.get("telefone") or "").strip(),
                "email": str(data_ws.get("email") or "").strip().lower(),
                "situacao_cadastral": str(data_ws.get("situacao") or "ATIVA").strip(),
                "representante_nome": rep_nome,
                "representante_cargo": rep_cargo,
                "fonte": "ReceitaWS / Receita Federal"
            }
    except Exception as e_ws:
        logger.error(f"Consulta ReceitaWS também falhou: {e_ws}")
        return {
            "sucesso": False,
            "erro": f"Não foi possível conectar aos servidores da Receita Federal ({e_ws}). Verifique o CNPJ ou a conexão."
        }


def buscar_ou_cadastrar_fornecedor(cnpj_input: str, engine) -> Dict[str, Any]:
    """
    Busca o fornecedor pelo CNPJ no banco de dados local.
    Caso não exista, busca na API oficial da Receita Federal e cadastra
    automaticamente no SQLite para uso imediato e futuro.
    """
    from sqlalchemy import text

    c_limpo = limpar_cnpj(cnpj_input)
    c_fmt = formatar_cnpj(c_limpo)

    if not c_limpo:
        return {"sucesso": False, "erro": "Informe um CNPJ válido para consulta."}

    # 1. Consulta no banco de dados local
    try:
        with engine.connect() as conn:
            query = text("""
                SELECT id, razao_social, nome_fantasia, cnpj_cpf, endereco,
                       numero, complemento, bairro, cidade, uf, cep,
                       telefone, email, nome_representante, cargo_representante,
                       situacao_cadastral
                FROM fornecedores
                WHERE REPLACE(REPLACE(REPLACE(REPLACE(cnpj_cpf, '.', ''), '-', ''), '/', ''), ' ', '') = :c_limpo
                   OR cnpj_cpf = :c_fmt
                LIMIT 1
            """)
            row = conn.execute(query, {"c_limpo": c_limpo, "c_fmt": c_fmt}).mappings().first()
            if row:
                d = dict(row)
                return {
                    "sucesso": True,
                    "encontrado_no_banco": True,
                    "recem_cadastrado": False,
                    "id": d.get("id"),
                    "cnpj": c_limpo,
                    "cnpj_formatado": d.get("cnpj_cpf") or c_fmt,
                    "razao_social": d.get("razao_social") or d.get("nome_fantasia"),
                    "nome_fantasia": d.get("nome_fantasia") or d.get("razao_social"),
                    "endereco": d.get("endereco") or "Endereço Comercial",
                    "numero": d.get("numero") or "s/n",
                    "complemento": d.get("complemento") or "",
                    "bairro": d.get("bairro") or "Centro",
                    "cidade": d.get("cidade") or "Imperatriz",
                    "uf": d.get("uf") or "MA",
                    "cep": d.get("cep") or "65900-000",
                    "telefone": d.get("telefone") or "",
                    "email": d.get("email") or "",
                    "representante_nome": d.get("nome_representante") or "Representante Legal",
                    "representante_cargo": d.get("cargo_representante") or "Sócio Administrador",
                    "situacao_cadastral": d.get("situacao_cadastral") or "ATIVA",
                    "fonte": "Banco de Dados Local"
                }
    except Exception as e_db:
        logger.debug(f"Erro ao buscar fornecedor no banco local: {e_db}")

    # 2. Se não encontrou no banco, puxa via API
    res_api = consultar_cnpj_api(c_limpo)
    if not res_api.get("sucesso"):
        return res_api

    # 3. Cadastra automaticamente no banco de dados local
    try:
        with engine.connect() as conn:
            sql_insert = text("""
                INSERT INTO fornecedores (
                    razao_social, nome_fantasia, cnpj_cpf, endereco,
                    numero, complemento, bairro, cidade, uf, cep,
                    telefone, email, nome_representante, cargo_representante,
                    situacao_cadastral, ativo
                ) VALUES (
                    :rz, :fant, :cnpj, :end,
                    :num, :comp, :bairro, :cid, :uf, :cep,
                    :tel, :email, :rep, :cargo,
                    :sit, 1
                )
            """)
            res_insert = conn.execute(sql_insert, {
                "rz": res_api["razao_social"],
                "fant": res_api["nome_fantasia"],
                "cnpj": res_api["cnpj_formatado"],
                "end": res_api["endereco"],
                "num": res_api["numero"],
                "comp": res_api["complemento"],
                "bairro": res_api["bairro"],
                "cid": res_api["cidade"],
                "uf": res_api["uf"],
                "cep": res_api["cep"],
                "tel": res_api["telefone"],
                "email": res_api["email"],
                "rep": res_api["representante_nome"],
                "cargo": res_api["representante_cargo"],
                "sit": res_api["situacao_cadastral"]
            })
            conn.commit()
            new_id = res_insert.lastrowid
            res_api["id"] = new_id
            res_api["encontrado_no_banco"] = False
            res_api["recem_cadastrado"] = True
    except Exception as e_ins:
        logger.warning(f"Não foi possível persistir fornecedor no banco ({e_ins}), retornando dados da API.")
        res_api["encontrado_no_banco"] = False
        res_api["recem_cadastrado"] = False

    return res_api
