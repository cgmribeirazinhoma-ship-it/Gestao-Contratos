import urllib.request
import urllib.parse
import json
import re
import ssl
import logging
from typing import Dict, Any, Optional, Tuple

logger = logging.getLogger("cnpj_service")

def _http_get_json(url: str, timeout: int = 7) -> Optional[Dict[str, Any]]:
    """Executa requisição HTTP GET com User-Agent de navegador e fallback automático de SSL."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json"
    }
    req = urllib.request.Request(url, headers=headers)
    # 1. Tentativa com SSL padrão
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status == 200:
                return json.loads(resp.read().decode("utf-8"))
    except Exception as e1:
        # 2. Tentativa com contexto desabilitado para evitar falha de certificado no Windows
        try:
            ctx = ssl._create_unverified_context()
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                if resp.status == 200:
                    return json.loads(resp.read().decode("utf-8"))
        except Exception as e2:
            logger.debug(f"HTTP GET falhou para {url}: {e1} / {e2}")
    return None

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

def _parse_minhareceita_brasilapi(data: dict, cnpj_limpo: str, fonte_nome: str) -> Dict[str, Any]:
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

    tel = str(data.get("ddd_telefone_1") or data.get("telefone") or "").strip()
    if tel:
        tel_num = re.sub(r'\D', '', tel)
        if len(tel_num) == 10:
            tel = f"({tel_num[:2]}) {tel_num[2:6]}-{tel_num[6:]}"
        elif len(tel_num) == 11:
            tel = f"({tel_num[:2]}) {tel_num[2:7]}-{tel_num[7:]}"

    return {
        "sucesso": True,
        "cnpj": cnpj_limpo,
        "cnpj_formatado": formatar_cnpj(cnpj_limpo),
        "razao_social": str(data.get("razao_social") or data.get("nome") or "").strip(),
        "nome_fantasia": str(data.get("nome_fantasia") or data.get("razao_social") or data.get("nome") or "").strip(),
        "endereco": end_completo if end_completo else "Endereço Comercial",
        "numero": num if num else "s/n",
        "complemento": str(data.get("complemento") or "").strip(),
        "bairro": str(data.get("bairro") or "Centro").strip(),
        "cidade": str(data.get("municipio") or data.get("cidade") or "").strip(),
        "uf": str(data.get("uf") or "").strip(),
        "cep": formatar_cep(data.get("cep")),
        "telefone": tel,
        "email": str(data.get("email") or "").strip().lower(),
        "situacao_cadastral": str(data.get("descricao_situacao_cadastral") or data.get("situacao") or "ATIVA").strip(),
        "representante_nome": rep_nome,
        "representante_cargo": rep_cargo,
        "fonte": fonte_nome
    }

def consultar_cnpj_api(cnpj: str, timeout: int = 8) -> Dict[str, Any]:
    """
    Consulta dados cadastrais oficiais da empresa via API pública com alta disponibilidade.
    Ordem de consulta com failover transparente:
    1º: Minha Receita (espelho oficial da Receita Federal de altíssima velocidade e sem quota)
    2º: BrasilAPI (API comunitária consolidada)
    3º: ReceitaWS (fallback de segurança)
    """
    c = limpar_cnpj(cnpj)
    if not validar_cnpj(c):
        return {
            "sucesso": False,
            "erro": f"CNPJ '{cnpj}' inválido. Certifique-se de informar os 14 dígitos numéricos."
        }

    # 1. Tentativa via Minha Receita (Altíssima velocidade, dados oficiais completos)
    try:
        url_mr = f"https://minhareceita.org/{c}"
        data_mr = _http_get_json(url_mr, timeout=timeout)
        if data_mr and (data_mr.get("razao_social") or data_mr.get("nome")):
            return _parse_minhareceita_brasilapi(data_mr, c, "Receita Federal (Minha Receita)")
    except Exception as e_mr:
        logger.debug(f"MinhaReceita falhou: {e_mr}")

    # 2. Tentativa via BrasilAPI
    try:
        url_br = f"https://brasilapi.com.br/api/cnpj/v1/{c}"
        data_br = _http_get_json(url_br, timeout=timeout)
        if data_br and (data_br.get("razao_social") or data_br.get("nome_fantasia")):
            return _parse_minhareceita_brasilapi(data_br, c, "Receita Federal (BrasilAPI)")
    except Exception as e_br:
        logger.debug(f"BrasilAPI falhou: {e_br}")

    # 3. Tentativa via ReceitaWS (Fallback)
    try:
        url_ws = f"https://receitaws.com.br/v1/cnpj/{c}"
        data_ws = _http_get_json(url_ws, timeout=timeout)
        if data_ws:
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
                "fonte": "Receita Federal (ReceitaWS)"
            }
    except Exception as e_ws:
        logger.error(f"Consulta ReceitaWS também falhou: {e_ws}")

    return {
        "sucesso": False,
        "erro": "Não foi possível obter os dados da empresa nos servidores da Receita Federal. Verifique o número informado ou sua conexão à internet."
    }


def buscar_fornecedor_no_banco(cnpj_input: str, engine) -> Optional[Dict[str, Any]]:
    """
    Busca o fornecedor exclusivamente no banco de dados local (tabela fornecedores ou contratos).
    Retorna o dicionário de dados caso conste no banco, ou None caso não conste.
    """
    from sqlalchemy import text

    c_limpo = limpar_cnpj(cnpj_input)
    c_fmt = formatar_cnpj(c_limpo)

    if not c_limpo:
        return None

    # 1. Consulta prioritária na tabela de fornecedores
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
                    "id": d.get("id"),
                    "cnpj": c_limpo,
                    "cnpj_formatado": d.get("cnpj_cpf") or c_fmt,
                    "razao_social": d.get("razao_social") or d.get("nome_fantasia") or "",
                    "nome_fantasia": d.get("nome_fantasia") or d.get("razao_social") or "",
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

            # 2. Consulta fallback em contratos anteriores
            query_ct = text("""
                SELECT c.fornecedor, c.cnpj_fornecedor
                FROM contratos c
                WHERE REPLACE(REPLACE(REPLACE(REPLACE(c.cnpj_fornecedor, '.', ''), '-', ''), '/', ''), ' ', '') = :c_limpo
                   OR c.cnpj_fornecedor = :c_fmt
                LIMIT 1
            """)
            row_ct = conn.execute(query_ct, {"c_limpo": c_limpo, "c_fmt": c_fmt}).mappings().first()
            if row_ct and row_ct.get("fornecedor"):
                return {
                    "sucesso": True,
                    "encontrado_no_banco": True,
                    "cnpj": c_limpo,
                    "cnpj_formatado": row_ct.get("cnpj_fornecedor") or c_fmt,
                    "razao_social": row_ct.get("fornecedor"),
                    "nome_fantasia": row_ct.get("fornecedor"),
                    "endereco": "Endereço Comercial",
                    "numero": "s/n",
                    "complemento": "",
                    "bairro": "Centro",
                    "cidade": "Imperatriz",
                    "uf": "MA",
                    "cep": "65900-000",
                    "telefone": "",
                    "email": "",
                    "representante_nome": "Representante Legal",
                    "representante_cargo": "Sócio Administrador",
                    "situacao_cadastral": "ATIVA",
                    "fonte": "Histórico de Contratos (Banco de Dados)"
                }
    except Exception as e_db:
        logger.debug(f"Erro ao buscar fornecedor no banco local: {e_db}")

    return None


def puxar_fornecedor_api_e_cadastrar(cnpj_input: str, engine) -> Dict[str, Any]:
    """
    Puxa os dados da empresa via API oficial da Receita Federal / BrasilAPI e
    cadastra automaticamente no banco de dados local para uso imediato e futuro.
    """
    from sqlalchemy import text

    c_limpo = limpar_cnpj(cnpj_input)
    if not c_limpo:
        return {"sucesso": False, "erro": "Informe um CNPJ válido com 14 dígitos."}

    # 1. Consulta na API da Receita Federal
    res_api = consultar_cnpj_api(c_limpo)
    if not res_api.get("sucesso"):
        return res_api

    # 2. Cadastra automaticamente no banco de dados local
    try:
        with engine.connect() as conn:
            # Verifica se já existe para evitar erro de integridade
            existe_id = conn.execute(
                text("SELECT id FROM fornecedores WHERE REPLACE(REPLACE(REPLACE(REPLACE(cnpj_cpf, '.', ''), '-', ''), '/', ''), ' ', '') = :c"),
                {"c": c_limpo}
            ).scalar()

            if not existe_id:
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
                    ) RETURNING id
                """)
                row_ins = conn.execute(sql_insert, {
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
                }).fetchone()
                conn.commit()
                res_api["id"] = row_ins[0] if row_ins else None
                res_api["recem_cadastrado"] = True
            else:
                sql_update = text("""
                    UPDATE fornecedores SET
                        razao_social = :rz,
                        nome_fantasia = :fant,
                        endereco = :end,
                        numero = :num,
                        complemento = :comp,
                        bairro = :bairro,
                        cidade = :cid,
                        uf = :uf,
                        cep = :cep,
                        telefone = :tel,
                        email = :email,
                        nome_representante = :rep,
                        cargo_representante = :cargo,
                        situacao_cadastral = :sit
                    WHERE id = :id
                """)
                conn.execute(sql_update, {
                    "id": existe_id,
                    "rz": res_api["razao_social"],
                    "fant": res_api["nome_fantasia"],
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
                res_api["id"] = existe_id
                res_api["recem_cadastrado"] = False

            res_api["encontrado_no_banco"] = True
    except Exception as e_ins:
        logger.warning(f"Não foi possível persistir fornecedor no banco ({e_ins}), retornando dados da API.")
        res_api["encontrado_no_banco"] = False
        res_api["recem_cadastrado"] = False

    return res_api


def buscar_ou_cadastrar_fornecedor(cnpj_input: str, engine) -> Dict[str, Any]:
    """
    Busca o fornecedor pelo CNPJ no banco de dados local.
    Caso não exista, busca na API oficial da Receita Federal e cadastra
    automaticamente no banco para uso imediato e futuro.
    """
    dados_banco = buscar_fornecedor_no_banco(cnpj_input, engine)
    if dados_banco:
        return dados_banco

    return puxar_fornecedor_api_e_cadastrar(cnpj_input, engine)
