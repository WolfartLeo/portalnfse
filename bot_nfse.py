# bot_nfse.py
import os
import re
import time
import shutil
import datetime
from typing import List, Dict, Optional, Tuple

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

try:
    import pyautogui as pg
except Exception:
    pg = None
import xml.etree.ElementTree as ET

from cert_image_selector import selecionar_certificado_por_imagem

try:
    from webdriver_manager.chrome import ChromeDriverManager
    USE_WEBDRIVER_MANAGER = True
except ImportError:
    USE_WEBDRIVER_MANAGER = False

if pg is not None:
    pg.FAILSAFE = True
    pg.PAUSE = 0.2

# ==============================
# CONFIGURAÇÕES GERAIS
# ==============================

CAMINHO_PLANILHA = r"C:\Python\INNOVE\PortalNFSe\planilhas\ACESSO_PORTAL_NACIONAL.xlsx"
PASTA_BASE_SAIDA = r"Z:\COMUM\LEONARDO\Portal Nacional"
PASTA_DOWNLOAD_TEMP = r"C:\Python\INNOVE\PortalNFSe\downloads_temp"

PASTA_IMAGENS_CERT = r"C:\Python\INNOVE\PortalNFSe\imagens"

DELAY_ACAO = 3.5

URL_PORTAL = "https://www.nfse.gov.br/EmissorNacional/Login?ReturnUrl=%2fEmissorNacional"

ID_INPUT_LOGIN = "Inscricao"
ID_INPUT_SENHA = "Senha"
ID_BTN_ACESSAR = ""  # se descobrir um id fixo, coloca aqui

XPATH_MENU_NFSE_EMITIDAS = '//*[@id="navbar"]/ul/li[3]/a'

# XPaths da tela de Visualizar NF:
XPATH_BTN_XML = '//*[@id="searchbar"]/ul/li[3]/a'
XPATH_BTN_PDF = '//*[@id="searchbar"]/ul/li[4]/a'

# XPATH do botão "próxima página" na listagem de NFS-e emitidas
XPATH_BTN_PROXIMA_PAGINA = '/html/body/div[1]/div[3]/div[1]/ul/li[8]/a'


# ==============================
# UTILITÁRIOS
# ==============================

def calcular_competencia_anterior(
    ano: Optional[int] = None,
    mes: Optional[int] = None,
) -> Tuple[int, int]:
    """
    Se ano/mes forem informados, usa diretamente.
    Se não, calcula a competência ANTERIOR ao mês atual.
    Ex: hoje 15/12/2025 -> competência anterior = 11/2025.
    """
    if ano is not None and mes is not None:
        return ano, mes

    hoje = datetime.date.today()
    primeiro_dia_mes_atual = hoje.replace(day=1)
    competencia_anterior = primeiro_dia_mes_atual - datetime.timedelta(days=1)
    return competencia_anterior.year, competencia_anterior.month


def montar_nome_pasta_competencia(ano: int, mes: int) -> str:
    return f"{ano:04d}-{mes:02d}"


def garantir_pasta(caminho: str) -> None:
    os.makedirs(caminho, exist_ok=True)


def limpar_nome_arquivo(nome: str) -> str:
    proibidos = r'\/:*?"<>|'
    for ch in proibidos:
        nome = nome.replace(ch, " ")
    return " ".join(nome.split())


def carregar_clientes_da_planilha() -> List[Dict]:
    df = pd.read_excel(CAMINHO_PLANILHA)
    df = df.fillna("")

    colunas_obrigatorias = [
        "EMPRESA",
        "CNPJ",
        "TIPO_ACESSO",
        "LOGIN",
        "SENHA",
        "ATIVO",
        "PREFEITURA",
        "IDENT_CERT",
        "IMG_CERT",
    ]
    for col in colunas_obrigatorias:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente na planilha: {col}")

    df_ativos = df[df["ATIVO"].astype(str).str.upper().str.strip() == "S"]
    return df_ativos.to_dict(orient="records")


def aguardar_novo_arquivo(extensao: str, timeout: int = 30) -> Optional[str]:
    extensao = extensao.lower()
    inicio = time.time()
    arquivos_iniciais = set(os.listdir(PASTA_DOWNLOAD_TEMP))

    while time.time() - inicio < timeout:
        atuais = set(os.listdir(PASTA_DOWNLOAD_TEMP))
        novos = atuais - arquivos_iniciais
        for nome in novos:
            if nome.lower().endswith(extensao):
                return os.path.join(PASTA_DOWNLOAD_TEMP, nome)
        time.sleep(1)

    return None


def mover_com_nome_base(caminho_origem: str, pasta_destino: str, nome_base: str) -> str:
    garantir_pasta(pasta_destino)
    base_limpo = limpar_nome_arquivo(nome_base)
    ext = os.path.splitext(caminho_origem)[1].lower()
    destino = os.path.join(pasta_destino, base_limpo + ext)

    contador = 2
    while os.path.exists(destino):
        destino = os.path.join(pasta_destino, f"{base_limpo} ({contador}){ext}")
        contador += 1

    shutil.move(caminho_origem, destino)
    return destino


def _parse_valor_monetario(texto: Optional[str]) -> Optional[float]:
    """
    Parser de valor robusto para:
    - "281.31"
    - "281,31"
    - "1.234,56"
    - "R$ 1.234,56"
    """
    if not texto:
        return None
    s = texto.strip()
    if not s:
        return None
    s = s.replace("R$", "").strip()

    # Se tiver . e , assume padrão brasileiro (1.234,56)
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    # Se tiver só , assume que é decimal
    elif "," in s and "." not in s:
        s = s.replace(",", ".")

    s = re.sub(r"[^\d\.\-]", "", s)
    if not s or s in (".", "-", ".-"):
        return None
    try:
        return float(s)
    except Exception:
        return None


def _formatar_data_iso_para_br(data_str: Optional[str]) -> Optional[str]:
    """
    Converte 'YYYY-MM-DD' ou 'YYYY-MM-DDTHH:MM...' para 'DD/MM/YYYY'.
    """
    if not data_str:
        return None
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", data_str)
    if not m:
        return None
    ano, mes, dia = m.groups()
    return f"{dia}/{mes}/{ano}"


def _formatar_id_para_excel(id_str: Optional[str]) -> Optional[str]:
    """
    Formata CNPJ/CPF como string para não virar notação científica no Excel.
    Retorna algo como: '01234567000189 (com apóstrofo).
    """
    if not id_str:
        return None
    digits = re.sub(r"\D", "", id_str)
    if not digits:
        return None
    return f"'{digits}"


def extrair_dados_nfse_do_xml(caminho_xml: str) -> Dict[str, Optional[str]]:
    """
    Extrai dados relevantes da NFS-e (layout NFSe Nacional) para o relatório:

    - numero_nf
    - data_emissao (DD/MM/AAAA)
    - data_competencia (DD/MM/AAAA quando disponível)
    - cnpj_prestador, razao_prestador
    - cnpj_tomador, razao_tomador
    - optante_sn (S/N)
    - codigo_trib_nacional
    - valor_servico
    - ir, iss, iss_retido, csll, deducoes, pis, cofins, inss
    - desc_incond, desc_cond
    - outras_retencoes
    - aliquota
    - base_calculo
    - valor_liquido
    - situacao (NORMAL / CANCELADA / COD_xxx)
    """
    dados = {k: None for k in [
        "numero_nf",
        "data_emissao",
        "data_competencia",
        "cnpj_prestador",
        "razao_prestador",
        "cnpj_tomador",
        "razao_tomador",
        "optante_sn",
        "codigo_trib_nacional",
        "valor_servico",
        "ir",
        "iss",
        "iss_retido",
        "csll",
        "deducoes",
        "pis",
        "cofins",
        "inss",
        "desc_incond",
        "desc_cond",
        "outras_retencoes",
        "aliquota",
        "base_calculo",
        "valor_liquido",
        "situacao",
    ]}

    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
    except Exception as e:
        print(f"[AVISO] Não consegui ler o XML '{caminho_xml}': {e}")
        return dados

    if "}" in root.tag:
        ns_uri = root.tag.split("}")[0].strip("{")
        ns = {"n": ns_uri}
    else:
        ns = {"n": ""}

    def get_text(xpath: str) -> Optional[str]:
        try:
            el = root.find(xpath, ns)
        except Exception:
            el = None
        if el is not None and el.text:
            t = el.text.strip()
            return t or None
        return None

    # Número da NF
    numero_raw = None
    for xp in [".//n:nNFSe", ".//n:nDFSe", ".//n:nDPS"]:
        t = get_text(xp)
        if t:
            numero_raw = t
            break

    if numero_raw:
        dig = re.sub(r"\D", "", numero_raw)
        dados["numero_nf"] = dig.lstrip("0") or dig

    # Datas
    dh_emi = get_text(".//n:DPS/n:infDPS/n:dhEmi") or get_text(".//n:dhProc")
    d_comp = get_text(".//n:DPS/n:infDPS/n:dCompet")

    dados["data_emissao"] = _formatar_data_iso_para_br(dh_emi)
    dados["data_competencia"] = _formatar_data_iso_para_br(d_comp)

    # Prestador
    emit = root.find(".//n:emit", ns)
    if emit is not None:
        for tagname in ("CNPJ", "CPF", "NIF"):
            el = emit.find(f"n:{tagname}", ns)
            if el is not None and el.text:
                dados["cnpj_prestador"] = re.sub(r"\D", "", el.text)
                break
        xN = emit.find("n:xNome", ns)
        if xN is not None and xN.text:
            dados["razao_prestador"] = xN.text.strip().upper()

    # Tomador
    toma = root.find(".//n:DPS/n:infDPS/n:toma", ns)
    if toma is not None:
        for tagname in ("CNPJ", "CPF", "NIF"):
            el = toma.find(f"n:{tagname}", ns)
            if el is not None and el.text:
                dados["cnpj_tomador"] = re.sub(r"\D", "", el.text)
                break
        xN = toma.find("n:xNome", ns)
        if xN is not None and xN.text:
            dados["razao_tomador"] = xN.text.strip().upper()

    # Optante Simples Nacional
    op_simp = get_text(".//n:DPS/n:infDPS/n:prest/n:regTrib/n:opSimpNac")
    if op_simp:
        # Heurística: 1 = Não optante / 2 ou 3 = Optante
        dados["optante_sn"] = "S" if op_simp in ("2", "3") else "N"

    # Código de Tributação Nacional
    codigo_trib = get_text(".//n:DPS/n:infDPS/n:serv/n:cServ/n:cTribNac") or get_text(".//n:cTribNac")
    if codigo_trib:
        dados["codigo_trib_nacional"] = codigo_trib.strip()

    # Valores principais
    v_serv = get_text(".//n:DPS/n:infDPS/n:valores/n:vServPrest/n:vServ")
    dados["valor_servico"] = _parse_valor_monetario(v_serv) if v_serv else None

    v_bc = get_text(".//n:infNFSe/n:valores/n:vBC")
    dados["base_calculo"] = _parse_valor_monetario(v_bc) if v_bc else None

    v_liq = get_text(".//n:infNFSe/n:valores/n:vLiq")
    dados["valor_liquido"] = _parse_valor_monetario(v_liq) if v_liq else None

    v_total_ret_txt = get_text(".//n:infNFSe/n:valores/n:vTotalRet")
    v_total_ret = _parse_valor_monetario(v_total_ret_txt) if v_total_ret_txt else None

    # Alíquota ISS (quando vier)
    aliq_txt = get_text(".//n:DPS/n:infDPS/n:valores/n:trib/n:tribMun/n:pAliq")
    dados["aliquota"] = _parse_valor_monetario(aliq_txt) if aliq_txt else None

    # ISS e ISS Retido
    v_iss_txt = get_text(".//n:infNFSe/n:valores/n:vISSQN") or get_text(".//n:valores/n:vISSQN")
    dados["iss"] = _parse_valor_monetario(v_iss_txt) if v_iss_txt else None

    v_iss_ret_txt = None
    for elem in root.iter():
        tag = elem.tag.split('}')[-1].lower()
        txt = (elem.text or "").strip()
        if not txt:
            continue
        if tag in ("vissqnret", "vretissqn"):
            v_iss_ret_txt = txt
            break
    dados["iss_retido"] = _parse_valor_monetario(v_iss_ret_txt) if v_iss_ret_txt else None

    # Tributos federais
    tribFed = root.find(".//n:DPS/n:infDPS/n:valores/n:trib/n:tribFed", ns)
    if tribFed is not None:
        piscofins = tribFed.find("n:piscofins", ns)
        if piscofins is not None:
            vpis = piscofins.find("n:vPis", ns)
            vcof = piscofins.find("n:vCofins", ns)
            if vpis is not None and vpis.text:
                dados["pis"] = _parse_valor_monetario(vpis.text)
            if vcof is not None and vcof.text:
                dados["cofins"] = _parse_valor_monetario(vcof.text)
        vRetCSLL = tribFed.find("n:vRetCSLL", ns)
        if vRetCSLL is not None and vRetCSLL.text:
            dados["csll"] = _parse_valor_monetario(vRetCSLL.text)

        # INSS / IRRF se existirem
        for elem in tribFed.iter():
            tag = elem.tag.split('}')[-1].lower()
            txt = (elem.text or "").strip()
            if not txt:
                continue
            if tag in ("vretinss", "vinss"):
                dados["inss"] = _parse_valor_monetario(txt)
            if tag in ("vretir", "vretirrf", "virrf"):
                dados["ir"] = _parse_valor_monetario(txt)

    # Deduções / Descontos (somente tags vDesc*)
    for elem in root.iter():
        tag = elem.tag.split('}')[-1]
        txt = (elem.text or "").strip()
        if not txt or txt == "-":
            continue
        tag_low = tag.lower()
        # Descontos incondicional/condicional
        if tag_low.startswith("vdesc"):
            if "cond" in tag_low:
                if dados["desc_cond"] is None:
                    dados["desc_cond"] = _parse_valor_monetario(txt)
            else:
                if dados["desc_incond"] is None:
                    dados["desc_incond"] = _parse_valor_monetario(txt)
        # Deduções
        if "deduc" in tag_low or "dedu" in tag_low:
            if dados["deducoes"] is None:
                dados["deducoes"] = _parse_valor_monetario(txt)

    # Outras retenções = vTotalRet - (IR + ISS_RET + CSLL + PIS + COFINS + INSS)
    soma_explicita = 0.0
    for k in ("ir", "iss_retido", "csll", "pis", "cofins", "inss"):
        v = dados[k]
        if isinstance(v, (int, float)):
            soma_explicita += v

    if v_total_ret is not None:
        outras = v_total_ret - soma_explicita
        if abs(outras) > 0.009:
            dados["outras_retencoes"] = outras
        else:
            dados["outras_retencoes"] = 0.0
    else:
        dados["outras_retencoes"] = None

    # Situação via cStat
    cstat = get_text(".//n:infNFSe/n:cStat")
    if cstat == "100":
        dados["situacao"] = "NORMAL"
    elif cstat in ("135", "136", "151"):
        dados["situacao"] = "CANCELADA"
    elif cstat:
        dados["situacao"] = f"COD_{cstat}"

    return dados


# ==============================
# CLASSE DO ROBÔ
# ==============================

class NFSePortalBot:
    def __init__(self, ano_competencia: Optional[int] = None, mes_competencia: Optional[int] = None):
        # sempre competência ANTERIOR, se não informado
        self.ano, self.mes = calcular_competencia_anterior(ano_competencia, mes_competencia)
        self.competencia_str = montar_nome_pasta_competencia(self.ano, self.mes)  # AAAA-MM
        self.competencia_label = f"{self.mes:02d}/{self.ano:04d}"  # MM/AAAA (igual tela Portal)

        self.pasta_competencia = os.path.join(PASTA_BASE_SAIDA, self.competencia_str)
        garantir_pasta(self.pasta_competencia)
        garantir_pasta(PASTA_DOWNLOAD_TEMP)
        garantir_pasta(PASTA_IMAGENS_CERT)

        self.driver: Optional[webdriver.Chrome] = None
        self.registros_log: List[Dict] = []

    # ---------- Navegador ----------

    def _inicializar_navegador(self) -> None:
        chrome_options = Options()

        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": PASTA_DOWNLOAD_TEMP,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        })

        # Modo cloud/headless (ex.: Render/Docker)
        headless = os.getenv("HEADLESS", "").strip().lower() in {"1", "true", "yes", "y"}
        if headless:
            chrome_options.add_argument("--headless=new")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--window-size=1920,1080")
        else:
            chrome_options.add_argument("--start-maximized")

        # Permite informar o binário do Chrome/Chromium via env (útil em Docker)
        chrome_bin = os.getenv("CHROME_BIN") or os.getenv("GOOGLE_CHROME_BIN")
        if chrome_bin:
            chrome_options.binary_location = chrome_bin

        # Permite informar o caminho do chromedriver via env (útil em Docker)
        chromedriver_path = os.getenv("CHROMEDRIVER_PATH")
        if chromedriver_path and os.path.exists(chromedriver_path):
            service = Service(executable_path=chromedriver_path)
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            return

        if USE_WEBDRIVER_MANAGER:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
        else:
            self.driver = webdriver.Chrome(options=chrome_options)

    def _finalizar_navegador(self) -> None:
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    # ---------- Aguardar tela logada ----------

    def _aguardar_tela_logada(self, timeout: int = 60) -> bool:
        assert self.driver is not None
        driver = self.driver

        inicio = time.time()
        while time.time() - inicio < timeout:
            try:
                driver.find_element(By.XPATH, XPATH_MENU_NFSE_EMITIDAS)
                return True
            except Exception:
                time.sleep(2)

        return False

    # ---------- Ir para NFS-e Emitidas ----------

    def _ir_para_nfse_emitidas(self) -> bool:
        assert self.driver is not None
        driver = self.driver

        try:
            link_emitidas = driver.find_element(By.XPATH, XPATH_MENU_NFSE_EMITIDAS)
        except Exception:
            print("[ERRO] Não encontrei o menu 'NFS-e Emitidas' no topo.")
            return False

        try:
            link_emitidas.click()
        except Exception as e:
            print(f"[ERRO] Falha ao clicar no menu 'NFS-e Emitidas': {e}")
            return False

        time.sleep(5)
        print("[INFO] Naveguei para a tela 'NFS-e Emitidas'.")
        return True


    def _linha_esta_cancelada(self, linha) -> bool:
        """
        Detecta se a NFS-e está CANCELADA olhando o ícone da coluna 'Situação' na listagem.
        Exemplos:
          - /EmissorNacional/img/tb-cancelada.svg  (cancelada)
          - /EmissorNacional/img/tb-gerada.svg     (emitida)
        Estratégia:
          - busca <img> dentro do td[6] (coluna Situação)
          - valida por src ou tooltip (data-original-title / title)
        """
        try:
            # coluna 6 = Situação (pela tela padrão do Portal)
            imgs = linha.find_elements(By.XPATH, "./td[6]//img")
            for img in imgs:
                src = (img.get_attribute("src") or "").lower()
                tip = (img.get_attribute("data-original-title") or img.get_attribute("title") or "").lower()
                if "tb-cancelada.svg" in src or "cancelada" in tip:
                    return True
        except Exception:
            return False
        return False


    # ---------- Download PDF/XML na Visualização \+ LOG ----------

    def _baixar_pdf_xml_da_visualizacao(
        self,
        cliente: Dict,
        emissao_tabela: str,
        competencia_tabela: str,
        is_cancelada: bool = False,
    ) -> None:
        assert self.driver is not None
        driver = self.driver

        # XML
        print("[INFO] Aguardando botão 'Download XML' na tela de Visualizar...")

        btn_xml = None
        inicio = time.time()
        while time.time() - inicio < 25:
            try:
                btn_xml = driver.find_element(By.XPATH, XPATH_BTN_XML)
            except Exception:
                try:
                    btn_xml = driver.find_element(
                        By.XPATH,
                        "//a[contains(@class,'btn') and (contains(., 'XML') or contains(@title,'XML'))]"
                    )
                except Exception:
                    btn_xml = None

            if btn_xml is not None:
                break
            time.sleep(1)

        if btn_xml is None:
            print("[ERRO] Botão 'Download XML' não encontrado na tela de Visualizar.")
            return

        try:
            btn_xml.click()
        except Exception as e:
            print(f"[ERRO] Falha ao clicar em 'Download XML': {e}")
            return

        caminho_xml = aguardar_novo_arquivo(".xml", timeout=40)
        if not caminho_xml:
            print("[ERRO] Nenhum XML novo encontrado após o clique em Download XML.")
            return

        print(f"[INFO] XML baixado: {caminho_xml}")

        # Extrair dados do XML
        dados_xml = extrair_dados_nfse_do_xml(caminho_xml)
        # ===== Situação (CANCELADA) =====
        # Prioridade: ícone da listagem (td Situação) e, como fallback, cStat no XML.
        situacao_xml = str(dados_xml.get("situacao") or "").strip().upper()
        flag_cancelada = bool(is_cancelada) or ("CANCEL" in situacao_xml)

        if flag_cancelada:
            # Regra de negócio: nota cancelada NÃO gera faturamento.
            # Então, no LOG final, todos os valores monetários devem sair como 0.
            for k in (
                "valor_servico",
                "ir",
                "iss",
                "iss_retido",
                "csll",
                "deducoes",
                "pis",
                "cofins",
                "inss",
                "desc_incond",
                "desc_cond",
                "outras_retencoes",
                "aliquota",
                "base_calculo",
                "valor_liquido",
            ):
                dados_xml[k] = 0.0
            dados_xml["situacao"] = "CANCELADA"


        numero_nf = str(dados_xml.get("numero_nf") or "").strip()
        if not numero_nf:
            base_nome_nf = os.path.splitext(os.path.basename(caminho_xml))[0]
            digitos = re.sub(r"\D", "", base_nome_nf)
            numero_nf = digitos if digitos else "SEM_NUMERO"

        razao_tomador_raw = dados_xml.get("razao_tomador") or cliente["EMPRESA"]
        razao_tomador = (razao_tomador_raw or "").strip().upper() or cliente["EMPRESA"].strip().upper()
        nome_base = f"{razao_tomador} - NF {numero_nf}"

        # Mover XML para pasta por competência
        caminho_xml_final = mover_com_nome_base(caminho_xml, self.pasta_competencia, nome_base)
        print(f"[INFO] XML movido para: {caminho_xml_final}")

        # PDF
        print("[INFO] Aguardando botão 'Download DANFS-e/PDF' na tela de Visualizar...")

        btn_pdf = None
        inicio = time.time()
        while time.time() - inicio < 25:
            try:
                btn_pdf = driver.find_element(By.XPATH, XPATH_BTN_PDF)
            except Exception:
                try:
                    btn_pdf = driver.find_element(
                        By.XPATH,
                        "//a[contains(@class,'btn') and "
                        "(contains(., 'DANFS') or contains(., 'DANF') "
                        "or contains(., 'PDF') or contains(@title,'DANFSe'))]"
                    )
                except Exception:
                    btn_pdf = None

            if btn_pdf is not None:
                break
            time.sleep(1)

        if btn_pdf is not None:
            try:
                btn_pdf.click()
                caminho_pdf = aguardar_novo_arquivo(".pdf", timeout=40)
                if caminho_pdf:
                    caminho_pdf_final = mover_com_nome_base(caminho_pdf, self.pasta_competencia, nome_base)
                    print(f"[INFO] PDF movido para: {caminho_pdf_final}")
                else:
                    caminho_pdf_final = None
                    print("[AVISO] Nenhum PDF novo encontrado após o clique em Download DANFS-e/PDF.")
            except Exception as e:
                caminho_pdf_final = None
                print(f"[ERRO] Falha ao clicar em Download DANFS-e/PDF: {e}")
        else:
            caminho_pdf_final = None
            print("[AVISO] Botão de Download PDF/DANFS-e não encontrado. Vou seguir só com o XML.")

        # ===== Montagem do LOG conforme layout solicitado =====

        def fmt2(v: Optional[float]) -> Optional[float]:
            if isinstance(v, (int, float)):
                return round(float(v), 2)
            return None

        # Datas: prioriza XML; se não tiver, usa o que veio da tabela
        data_emissao_log = dados_xml.get("data_emissao") or emissao_tabela
        data_comp_log = dados_xml.get("data_competencia") or competencia_tabela

        # IDs formatados para não virar notação científica
        cnpj_prest = _formatar_id_para_excel(dados_xml.get("cnpj_prestador"))
        cnpj_tom = _formatar_id_para_excel(dados_xml.get("cnpj_tomador"))

        registro = {
            # A – Nº NF
            "NUMERO_NF": numero_nf,
            # B – Data Emissão
            "DATA_EMISSAO": data_emissao_log,
            # C – Data Competência
            "DATA_COMPETENCIA": data_comp_log,
            # D – CNPJ Prestador
            "CNPJ_PRESTADOR": cnpj_prest,
            # E – Razão Prestador
            "RAZAO_PRESTADOR": (dados_xml.get("razao_prestador") or "").strip() or cliente["EMPRESA"],
            # F – CNPJ Tomador (pode ser CPF)
            "CNPJ_TOMADOR": cnpj_tom,
            # G – Razão Tomador
            "RAZAO_TOMADOR": razao_tomador,
            # H – Optante SN
            "OPTANTE_SN": dados_xml.get("optante_sn"),
            # I – Código de Tributação
            "CODIGO_TRIBUTACAO_NACIONAL": dados_xml.get("codigo_trib_nacional"),
            # J – Valor Serviço
            "VALOR_SERVICO": fmt2(dados_xml.get("valor_servico")),
            # K – IR
            "IR": fmt2(dados_xml.get("ir")),
            # L – ISS
            "ISS": fmt2(dados_xml.get("iss")),
            # M – ISS Retido
            "ISS_RETIDO": fmt2(dados_xml.get("iss_retido")),
            # N – CSLL
            "CSLL": fmt2(dados_xml.get("csll")),
            # O – Deduções
            "DEDUCOES": fmt2(dados_xml.get("deducoes")),
            # P – PIS
            "PIS": fmt2(dados_xml.get("pis")),
            # Q – COFINS
            "COFINS": fmt2(dados_xml.get("cofins")),
            # R – INSS
            "INSS": fmt2(dados_xml.get("inss")),
            # S – Desc. Incond.
            "DESC_INCOND": fmt2(dados_xml.get("desc_incond")),
            # T – Desc. Cond.
            "DESC_COND": fmt2(dados_xml.get("desc_cond")),
            # U – Outras Ret.
            "OUTRAS_RET": fmt2(dados_xml.get("outras_retencoes")),
            # V – Alíquota
            "ALIQUOTA": fmt2(dados_xml.get("aliquota")),
            # X – Base de Cálculo
            "BASE_CALCULO": fmt2(dados_xml.get("base_calculo")),
            # W – Valor Líquido
            "VALOR_LIQUIDO": fmt2(dados_xml.get("valor_liquido")),
            # Y – Situação
            "SITUACAO": dados_xml.get("situacao"),
        }

        self.registros_log.append(registro)
        print(f"[INFO] Registro de log incluído para NF {numero_nf}.")

    # ---------- Processar todas as páginas de Notas Emitidas ----------

    def _processar_notas_emitidas(self, cliente: Dict) -> None:
        """
        Percorre TODAS as páginas de NFS-e Emitidas, mas só baixa
        as notas da competência alvo (self.competencia_label, ex: 11/2025).
        """
        assert self.driver is not None
        driver = self.driver

        alvo_ano = self.ano
        alvo_mes = self.mes
        alvo_label = self.competencia_label

        print(f"[INFO] Competência ALVO para esse cliente: {alvo_label}")

        pagina = 1
        chave_primeira_anterior = None

        while True:
            linhas = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            if not linhas:
                print(f"[INFO] Nenhuma linha encontrada na página {pagina}. Encerrando paginação.")
                break

            try:
                chave_primeira = linhas[0].text.strip()
            except Exception:
                chave_primeira = f"pag_{pagina}_linha_0"

            if chave_primeira == chave_primeira_anterior:
                print("[INFO] Primeira linha repetida em relação à página anterior. Parece ser a última página. Encerrando paginação.")
                break

            chave_primeira_anterior = chave_primeira
            print(f"[INFO] Processando página {pagina}. Total de linhas: {len(linhas)}")

            # percorre linhas da página
            for idx in range(len(linhas)):
                linhas = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                if idx >= len(linhas):
                    break

                linha = linhas[idx]
                is_cancelada = self._linha_esta_cancelada(linha)

                try:
                    emissao = linha.find_element(By.XPATH, "./td[1]").text.strip()
                except Exception:
                    emissao = "?"

                try:
                    competencia_texto = linha.find_element(By.XPATH, "./td[3]").text.strip()
                except Exception:
                    competencia_texto = "?"

                # converte "MM/AAAA" para (ano, mes)
                comp_parsed = None
                m = re.search(r"(\d{2})/(\d{4})", competencia_texto)
                if m:
                    mes_linha = int(m.group(1))
                    ano_linha = int(m.group(2))
                    comp_parsed = (ano_linha, mes_linha)

                if not comp_parsed:
                    print(f"[INFO] Linha {idx+1}: competência '{competencia_texto}' não reconhecida, pulando.")
                    continue

                if comp_parsed != (alvo_ano, alvo_mes):
                    print(f"[INFO] Linha {idx+1}: competência {competencia_texto} != alvo {alvo_label}, pulando.")
                    continue

                print(
                    f"[INFO] Linha {idx+1}: Emissão={emissao} | Competência={competencia_texto} (ALVO) | Cancelada={is_cancelada} "
                    "-> iniciando fluxo Visualizar"
                )

                # 3 pontinhos
                try:
                    menu_3_pontos = linha.find_element(
                        By.XPATH,
                        "./td[7]//a[contains(@class,'icone-trigger')]"
                    )
                except Exception as e:
                    print(f"[ERRO] Não achei o menu de ações (3 pontinhos) na linha {idx+1}: {e}")
                    continue

                try:
                    driver.execute_script("arguments[0].click();", menu_3_pontos)
                except Exception as e:
                    print(f"[ERRO] Falha ao clicar nos 3 pontinhos da linha {idx+1}: {e}")
                    continue

                time.sleep(1.5)

                # Visualizar
                try:
                    link_visualizar = driver.find_element(
                        By.XPATH,
                        "//div[contains(@class,'popover') and contains(@style,'display: block')]"
                        "//a[contains(@class,'list-group-item') and contains(., 'Visualizar')]"
                    )
                except Exception as e:
                    print(f"[ERRO] Não encontrei a opção 'Visualizar' para a linha {idx+1}: {e}")
                    continue

                janela_atual = driver.current_window_handle
                handles_antes = set(driver.window_handles)

                try:
                    link_visualizar.click()
                except Exception as e:
                    print(f"[ERRO] Falha ao clicar em 'Visualizar' na linha {idx+1}: {e}")
                    continue

                time.sleep(3)

                handles_depois = set(driver.window_handles)

                if len(handles_depois) > len(handles_antes):
                    nova_janela = list(handles_depois - handles_antes)[0]
                    try:
                        driver.switch_to.window(nova_janela)
                        print(f"[INFO] Visualização da linha {idx+1} aberta em nova aba.")
                        self._baixar_pdf_xml_da_visualizacao(cliente, emissao, competencia_texto, is_cancelada=is_cancelada)
                        time.sleep(2)
                        driver.close()
                    finally:
                        driver.switch_to.window(janela_atual)
                        time.sleep(3)
                else:
                    print(f"[INFO] Visualização da linha {idx+1} aberta na mesma aba.")
                    self._baixar_pdf_xml_da_visualizacao(cliente, emissao, competencia_texto, is_cancelada=is_cancelada)
                    time.sleep(2)
                    driver.back()
                    time.sleep(3)

            # tenta ir para a próxima página de notas
            try:
                btn_prox = driver.find_element(By.XPATH, XPATH_BTN_PROXIMA_PAGINA)
            except Exception:
                # fallback genérico
                try:
                    btn_prox = driver.find_element(
                        By.XPATH,
                        "//ul/li/a[contains(., 'Próxima') or contains(., '>') or contains(., '>>')]"
                    )
                except Exception:
                    btn_prox = None

            if not btn_prox:
                print("[INFO] Não encontrei botão de próxima página. Encerrando paginação.")
                break

            try:
                btn_prox.click()
                pagina += 1
                time.sleep(4)
            except Exception as e:
                print(f"[INFO] Falha ao clicar na próxima página ({e}). Encerrando paginação.")
                break

        print("[INFO] Ciclo de páginas (Visualizar + Download) concluído para todas as notas da competência alvo.")

    # ---------- Login: LOGIN/SENHA ----------

    def _login_por_login_senha(self, cliente: Dict) -> bool:
        assert self.driver is not None
        driver = self.driver

        driver.get(URL_PORTAL)
        time.sleep(DELAY_ACAO)

        try:
            input_login = driver.find_element(By.ID, ID_INPUT_LOGIN)
            input_senha = driver.find_element(By.ID, ID_INPUT_SENHA)
        except Exception:
            print(f"[ERRO] Não encontrei campos de login/senha para o cliente: {cliente['EMPRESA']}")
            return False

        input_login.clear()
        input_login.send_keys(str(cliente.get("LOGIN", "")).strip())

        input_senha.clear()
        input_senha.send_keys(str(cliente.get("SENHA", "")).strip())

        time.sleep(1.0)

        try:
            if ID_BTN_ACESSAR:
                btn_entrar = driver.find_element(By.ID, ID_BTN_ACESSAR)
            else:
                btn_entrar = driver.find_element(
                    By.XPATH,
                    "//*[ (self::a or self::button) "
                    " and (contains(., 'Acessar') or contains(., 'Entrar')) "
                    " and not(contains(., 'certificado')) ]"
                )
        except Exception:
            print(f"[ERRO] Não encontrei botão de login para o cliente: {cliente['EMPRESA']}")
            return False

        btn_entrar.click()
        time.sleep(DELAY_ACAO)

        if not self._aguardar_tela_logada(timeout=60):
            print(f"[ERRO] Não identifiquei a tela logada após login/senha de {cliente['EMPRESA']}.")
            return False

        print(f"[INFO] Login (usuário/senha) OK para {cliente['EMPRESA']}")
        return True

    # ---------- Login: CERTIFICADO (imagem) ----------

    def _login_por_certificado(self, cliente: Dict) -> bool:
        if pg is None:
            raise RuntimeError("pyautogui não está instalado. Para acesso por CERTIFICADO, instale: pip install pyautogui")
        assert self.driver is not None
        driver = self.driver

        driver.get(URL_PORTAL)
        time.sleep(DELAY_ACAO)

        try:
            driver.maximize_window()
        except Exception:
            pass

        img_btn_cert = os.path.join(PASTA_IMAGENS_CERT, "btn_acesso_cert.png")
        if not os.path.exists(img_btn_cert):
            print(f"[ERRO] Imagem do botão de certificado não encontrada: {img_btn_cert}")
            return False

        print(f"[INFO] Vou procurar o botão 'Acesso via certificado digital' na tela: {img_btn_cert}")
        inicio = time.time()
        pos_btn = None

        time.sleep(3)

        while time.time() - inicio < 30:
            try:
                pos_btn = pg.locateOnScreen(img_btn_cert, confidence=0.8)
            except Exception as e:
                print(f"[ERRO] locateOnScreen (botão certificado) falhou: {e}")
                return False

            if pos_btn:
                x, y = pg.center(pos_btn)
                print(f"[INFO] Botão de certificado localizado em ({x}, {y}). Vou clicar.")
                try:
                    pg.moveTo(x, y, duration=0.5)
                    pg.click()
                except Exception as e:
                    print(f"[ERRO] Falha ao clicar no botão de certificado: {e}")
                    return False
                break

            time.sleep(1.0)

        if not pos_btn:
            print("[ERRO] Não localizei o botão 'Acesso via certificado digital' na tela dentro do timeout.")
            return False

        print(f"[INFO] Clique no botão de certificado disparado para {cliente['EMPRESA']}. Aguardando popup...")

        nome_img_cert = str(cliente.get("IMG_CERT", "")).strip()
        if not nome_img_cert:
            print(f"[ERRO] Cliente {cliente['EMPRESA']} com TIPO_ACESSO=CERTIFICADO, mas IMG_CERT vazio na planilha.")
            return False

        time.sleep(3)

        print(f"[INFO] Vou selecionar o certificado por imagem: {nome_img_cert}")
        ok = selecionar_certificado_por_imagem(
            base_dir_imagens=PASTA_IMAGENS_CERT,
            nome_img_cert=nome_img_cert,
            timeout_cert=40,
            confidence=0.8,
            debug=True,
        )
        if not ok:
            print(f"[ERRO] Falha ao selecionar certificado por imagem para {cliente['EMPRESA']}.")
            return False

        print("[INFO] Certificado selecionado. Aguardando tela logada do Portal...")
        if not self._aguardar_tela_logada(timeout=60):
            print(f"[ERRO] Não identifiquei a tela logada após seleção de certificado para {cliente['EMPRESA']}.")
            return False

        print(f"[INFO] Login (certificado) OK para {cliente['EMPRESA']}")
        return True

    # ---------- Processar cliente ----------

    def _processar_cliente(self, cliente: Dict) -> None:
        print(f"\n=== Processando cliente: {cliente['EMPRESA']} | TIPO_ACESSO={cliente['TIPO_ACESSO']} ===")

        self._inicializar_navegador()

        try:
            tipo_acesso = str(cliente["TIPO_ACESSO"]).strip().upper()
            if tipo_acesso == "LOGIN_SENHA":
                autenticado = self._login_por_login_senha(cliente)
            elif tipo_acesso == "CERTIFICADO":
                autenticado = self._login_por_certificado(cliente)
            else:
                print(f"[ERRO] TIPO_ACESSO inválido para {cliente['EMPRESA']}: {tipo_acesso}")
                autenticado = False

            if not autenticado:
                print(f"[ERRO] Falha no login para o cliente: {cliente['EMPRESA']}")
                return

            print(f"[INFO] Login bem-sucedido para {cliente['EMPRESA']} | Competência alvo: {self.competencia_str}")

            time.sleep(3)

            if not self._ir_para_nfse_emitidas():
                print(f"[ERRO] Não consegui navegar para 'NFS-e Emitidas' para {cliente['EMPRESA']}.")
                return

            self._processar_notas_emitidas(cliente)

        finally:
            self._finalizar_navegador()

    # ---------- Execução geral ----------

    def rodar(self) -> None:
        clientes = carregar_clientes_da_planilha()
        if not clientes:
            print("[AVISO] Nenhum cliente ATIVO na planilha.")
            return

        print(f"=== Rodando PortalNFSe para competência {self.competencia_str} ({self.competencia_label}) ===")
        print(f"Total de clientes ativos: {len(clientes)}")

        for cliente in clientes:
            self._processar_cliente(cliente)

        if self.registros_log:
            df_log = pd.DataFrame(self.registros_log)

            # Garante a ordem das colunas no Excel
            colunas_ordem = [
                "NUMERO_NF",
                "DATA_EMISSAO",
                "DATA_COMPETENCIA",
                "CNPJ_PRESTADOR",
                "RAZAO_PRESTADOR",
                "CNPJ_TOMADOR",
                "RAZAO_TOMADOR",
                "OPTANTE_SN",
                "CODIGO_TRIBUTACAO_NACIONAL",
                "VALOR_SERVICO",
                "IR",
                "ISS",
                "ISS_RETIDO",
                "CSLL",
                "DEDUCOES",
                "PIS",
                "COFINS",
                "INSS",
                "DESC_INCOND",
                "DESC_COND",
                "OUTRAS_RET",
                "ALIQUOTA",
                "BASE_CALCULO",
                "VALOR_LIQUIDO",
                "SITUACAO",
            ]
            df_log = df_log.reindex(columns=colunas_ordem)

            caminho_log = os.path.join(
                self.pasta_competencia,
                f"LOG_NFSE_{self.competencia_str}.xlsx"
            )
            df_log.to_excel(caminho_log, index=False)
            print(f"[INFO] Log salvo em: {caminho_log}")
        else:
            print("[INFO] Nenhum registro para log.")

        print("\n=== Fim da execução geral ===")


if __name__ == "__main__":
    bot = NFSePortalBot()
    bot.rodar()
