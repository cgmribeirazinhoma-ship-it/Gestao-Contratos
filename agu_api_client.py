import urllib.request
import urllib.parse
import http.cookiejar
import ssl
import re
import json
import logging
from typing import Dict, Any, Optional

logger = logging.getLogger("agu_api_client")

class AGUApiClient:
    """
    Cliente para integração com o Sistema Ger@ AGU / CGU Contratos
    (https://cgu.agu.gov.br/contrato/)
    """
    BASE_URL = "https://cgu.agu.gov.br/contrato/"
    MONTAGEM_URL = "https://cgu.agu.gov.br/contrato/montagem/index.php"
    API_COLETA_URL = "https://cgu.agu.gov.br/cgi-bin/sapiens_com/relsapiens/coleta1.py"

    def __init__(self, timeout: int = 15):
        self.timeout = timeout
        self.cookie_jar = http.cookiejar.CookieJar()
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        handler = urllib.request.HTTPSHandler(context=ctx)
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cookie_jar), handler)
        self.csrf_token = ""
        self.api_key = ""

    def obter_sessao(self) -> bool:
        """Acessa a página inicial para capturar cookies, CSRF e API_KEY"""
        try:
            req = urllib.request.Request(
                self.BASE_URL,
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                }
            )
            res = self.opener.open(req, timeout=self.timeout)
            html = res.read().decode("utf-8", errors="ignore")

            cookies = {c.name: c.value for c in self.cookie_jar}
            self.api_key = cookies.get("API_KEY", "")
            self.csrf_token = cookies.get("CSRF_TOKEN", "")

            # Se não estiver no cookie, busca no HTML
            if not self.csrf_token:
                m_csrf = re.search(r'name=["\']csrf_token["\']\s+value=["\']([^"\']+)["\']', html)
                if m_csrf:
                    self.csrf_token = m_csrf.group(1)

            return bool(self.csrf_token)
        except Exception as e:
            logger.warning(f"Não foi possível obter sessão com portal AGU: {e}")
            return False

    def consultar_codigo_agu(self, codigo: str) -> Optional[Dict[str, Any]]:
        """
        Consulta um contrato ou edital salvo no Ger@ AGU pelo código de 19 dígitos
        via API coleta1.py
        """
        if not self.csrf_token:
            if not self.obter_sessao():
                return None

        codigo_limpo = str(codigo).strip()
        for script_id in [81, 84, 101]:
            try:
                data = urllib.parse.urlencode({
                    "script": script_id,
                    "base": 1,
                    "sqlpronto": "VAZIO",
                    "csrf_token": self.csrf_token,
                    "param1": codigo_limpo
                }).encode("utf-8")

                req = urllib.request.Request(
                    self.API_COLETA_URL,
                    data=data,
                    headers={
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                        "Referer": self.BASE_URL,
                        "Origin": "https://cgu.agu.gov.br",
                        "Content-Type": "application/x-www-form-urlencoded",
                        "X-API-KEY": self.api_key
                    }
                )

                res = self.opener.open(req, timeout=self.timeout)
                resp_txt = res.read().decode("utf-8", errors="ignore")
                resp_json = json.loads(resp_txt)

                if resp_json.get("erro") == "OK" and resp_json.get("info"):
                    return resp_json
            except Exception as e:
                logger.debug(f"Tentativa script {script_id} falhou: {e}")
                continue

        return None

    def gerar_minuta_agu(self, dados: Dict[str, Any]) -> Dict[str, Any]:
        """
        Envia dados cadastrais ao endpoint oficial da AGU (montagem/index.php)
        e retorna a minuta jurídica processada.
        """
        if not self.csrf_token:
            self.obter_sessao()

        # Mapeamento do tipo de contrato AGU
        # aq: 0 = Aquisição em geral, 1 = Aquisição contínua
        # serv: 0 = Não contínuo/escopo, 1 = Contínuo sem dedicação, 2 = Contínuo com dedicação, 3 = Obra/Engenharia
        modelo = dados.get("modelo_agu", "COMPRAS").upper()
        tipo_contrat = str(dados.get("tipo_contratacao", "")).upper()
        
        if "AQUIS" in tipo_contrat or "COMPRA" in modelo or "BEM" in modelo or "AQUIS" in modelo:
            sel_tipo = "0"  # 0 = Aquisições
            aq = "0"
            serv = ""
            continuo = ""
            dedicacao = ""
        elif "SERV" in tipo_contrat or "CONTINUO" in modelo:
            sel_tipo = "1"  # 1 = Serviços
            aq = ""
            serv = "1"
            continuo = "1"
            dedicacao = "2"  # Sem dedicação exclusiva por padrão
        elif "OBRA" in modelo or "ENGENHARIA" in modelo:
            sel_tipo = "1"
            aq = ""
            serv = "2"
            continuo = ""
            dedicacao = ""
        elif "DISPENSA" in modelo or "INEXIGIBILIDADE" in modelo:
            sel_tipo = "0"
            aq = "0"
            serv = ""
            continuo = ""
            dedicacao = ""
        else:
            # Serviços gerais / escopo
            sel_tipo = "1"
            aq = ""
            serv = "1"
            continuo = "2"
            dedicacao = "2"

        # Mapeamento da modalidade / Demais Informações (CGUTEC Versão 1.0.2)
        mod_raw = str(dados.get("modalidade", "")).upper()
        demais_info = str(dados.get("demais_informacoes", "")).upper()
        if "PREG" in mod_raw or "PREG" in demais_info:
            modalidade = "1"  # 1 = Pregão
            tp_mod = "1"
        elif "CONCORR" in mod_raw or "CONCORR" in demais_info:
            modalidade = "2"  # 2 = Concorrência
            tp_mod = "2"
        elif "DIRET" in mod_raw or "DIRET" in demais_info or "DISP" in mod_raw or "INEX" in mod_raw:
            modalidade = "3"  # 3 = Contratação Direta (Dispensa / Inexigibilidade)
            tp_mod = "3"
        else:
            modalidade = "1"
            tp_mod = "1"

        payload = {
            "csrf_token": self.csrf_token,
            "opinicial": "0",
            "op_gera": "0",
            "tp_mod": tp_mod,
            "modalidade": modalidade,
            "sel_tipo": sel_tipo,
            "aq": aq,
            "serv": serv,
            "continuo": continuo,
            "dedicacao": dedicacao,
            # Dados do Contratante (Órgão / Município - CGUTEC 1.0.2)
            "razao_social_contrato": dados.get("orgao_nome", "PREFEITURA MUNICIPAL DE RIBEIRÃOZINHO DO MARANHÃO - MA"),
            "cnpj_o": dados.get("orgao_cnpj", "01.612.834/0001-86"),
            "representado_o": dados.get("secretario", "SECRETÁRIO MUNICIPAL TITULAR"),
            "cargo_o": dados.get("cargo_secretario", "Secretário(a) Municipal Titular"),
            "setor": dados.get("setor_licitacoes") or dados.get("setor") or dados.get("orgao_nome", "Setor de Licitações e Contratos"),
            "cep_contrato": dados.get("orgao_cep", "65928-000"),
            "logradouro_contrato": dados.get("orgao_endereco", "Rua Principal"),
            "bairro_contrato": dados.get("orgao_bairro", "Centro"),
            "numero_contrato": dados.get("numero_imovel") or dados.get("orgao_numero", "s/n"),
            "complemento_contrato": dados.get("complemento_imovel") or dados.get("orgao_complemento", ""),
            "municipio_contrato": dados.get("orgao_cidade", "Ribeirãozinho do Maranhão"),
            "uf_contrato": dados.get("orgao_uf", "MA"),
            # Dados do Contratado (Fornecedor)
            "empresa": dados.get("fornecedor_nome", "FORNECEDOR CONTRATADO LTDA"),
            "cnpj_c": dados.get("fornecedor_cnpj", "00.000.000/0001-00"),
            "representado_c": dados.get("fornecedor_representante", "Representante Legal"),
            "cargo_c": dados.get("fornecedor_cargo", "Sócio Administrador"),
            "cepct": dados.get("fornecedor_cep", "65900-000"),
            "logradouroct": dados.get("fornecedor_endereco", "Endereço da Empresa"),
            "bairroct": dados.get("fornecedor_bairro", "Centro"),
            "numeroct": dados.get("fornecedor_numero", "100"),
            "complementoct": dados.get("fornecedor_complemento", ""),
            "localidadect": dados.get("fornecedor_cidade", "Imperatriz"),
            "ufct": dados.get("fornecedor_uf", "MA"),
            # Dados do Contrato
            "processoadm": dados.get("processo_adm", "001/2026"),
            "pregao": dados.get("numero_licitacao", dados.get("numero_completo", "001/2026")),
            "objeto": dados.get("objeto", "Contratação administrativa conforme especificações"),
            "form_loaded_at": "1700000000"
        }

        try:
            encoded_data = urllib.parse.urlencode(payload).encode("utf-8")
            req = urllib.request.Request(
                self.MONTAGEM_URL,
                data=encoded_data,
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                    "Referer": self.BASE_URL,
                    "Origin": "https://cgu.agu.gov.br",
                    "Content-Type": "application/x-www-form-urlencoded"
                }
            )

            res = self.opener.open(req, timeout=self.timeout)
            html_retorno = res.read().decode("utf-8", errors="ignore")

            # Extrai o documento montado do painel direito
            m_right = re.search(r'<div class="right" id="right">(.*?)</div>\s*</body>', html_retorno, re.DOTALL)
            if m_right:
                conteudo_html = m_right.group(1).strip()
                # Limpa tags desnecessárias
                texto_puro = re.sub(r'<[^>]+>', ' ', conteudo_html)
                texto_puro = re.sub(r'\s+', ' ', texto_puro).strip()

                return {
                    "sucesso": True,
                    "origem": "API_AGU_ONLINE",
                    "html": conteudo_html,
                    "texto_completo": texto_puro,
                    "payload_enviado": payload
                }
            else:
                return {
                    "sucesso": False,
                    "origem": "API_AGU_OFFLINE",
                    "erro": "Painel de montagem não retornado pela AGU",
                    "payload_enviado": payload
                }
        except Exception as e:
            logger.warning(f"Erro ao chamar API de montagem da AGU: {e}")
            return {
                "sucesso": False,
                "origem": "API_AGU_ERRO",
                "erro": str(e),
                "payload_enviado": payload
            }
