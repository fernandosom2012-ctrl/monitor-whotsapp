"""WhatsApp Web monitor that saves support tickets into an Excel file."""

from __future__ import annotations

import logging
import re
import shutil
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from config import CONFIG
from excel_handler import ExcelHandler


def _configure_stdout() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass


_configure_stdout()


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("monitor.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


def normalizar_texto(texto: str) -> str:
    """Return a normalized version of text for tolerant comparisons."""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = texto.replace("#", " ")
    texto = texto.replace("*", " ")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip().lower()


def extrair_campo(texto: str, *aliases: str) -> str:
    """Extract the first matching field, accepting aliases and next-line values."""
    linhas = texto.splitlines()
    aliases_norm = [normalizar_texto(alias) for alias in aliases]

    for indice, linha in enumerate(linhas):
        linha_limpa = linha.strip()
        if not linha_limpa:
            continue

        linha_norm = normalizar_texto(linha_limpa)
        for alias_norm in aliases_norm:
            if not linha_norm.startswith(alias_norm):
                continue

            partes = re.split(r"[:\-]", linha_limpa, maxsplit=1)
            if len(partes) == 2 and partes[1].strip():
                return limpar_valor_extraido(partes[1].strip())

            for proxima_linha in linhas[indice + 1 :]:
                valor = proxima_linha.strip()
                if valor:
                    return limpar_valor_extraido(valor)

    return ""


def limpar_valor_extraido(valor: str) -> str:
    """Remove ruido comum do WhatsApp, como horario grudado no fim da linha."""
    valor = re.sub(r"\s+", " ", valor).strip()
    valor = re.sub(r"(?<!\d)\d{2}:\d{2}$", "", valor).strip()
    valor = valor.replace("*", "").strip()
    return valor


def detectar_prioridade(texto: str) -> str:
    """Infer ticket priority from explicit fields or keywords."""
    campo = extrair_campo(texto, "prioridade", "urgencia", "urgência")
    if campo:
        campo_norm = normalizar_texto(campo)
        if any(token in campo_norm for token in ("alta", "urgente", "critico", "critica", "emergencia")):
            return "Alta"
        if any(token in campo_norm for token in ("media", "mediana", "importante")):
            return "Média"
        if "baixa" in campo_norm:
            return "Baixa"
        return campo.strip().capitalize()

    texto_norm = normalizar_texto(texto)
    if any(token in texto_norm for token in ("urgente", "critico", "critica", "emergencia", "parado", "fora do ar")):
        return "Alta"
    if any(token in texto_norm for token in ("importante", "precisa", "necessario")):
        return "Média"
    return "Baixa"


def gerar_id_chamado(sequencial: int) -> str:
    prefixo = CONFIG.get("prefixo_id", "CH")
    ano = datetime.now().strftime("%Y")
    return f"{prefixo}-{ano}-{sequencial:04d}"


def extrair_descricao_fallback(texto: str, trigger: str) -> str:
    """Fallback description when the structured field was not found."""
    trigger_norm = normalizar_texto(trigger)
    linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]
    coletar = False
    restantes: list[str] = []

    for indice, linha in enumerate(linhas):
        linha_norm = normalizar_texto(linha)
        if not coletar and trigger_norm in linha_norm:
            coletar = True
            continue
        if coletar:
            if ("problema" in linha_norm or "descri" in linha_norm) and linha.endswith(":"):
                for proxima_linha in linhas[indice + 1 :]:
                    valor = proxima_linha.strip()
                    if valor:
                        return valor[:200]
            restantes.append(linha)

    if restantes:
        nao_estruturadas = [
            linha
            for linha in restantes
            if ":" not in linha and "-" not in linha[:15]
        ]
        if nao_estruturadas:
            return limpar_valor_extraido(" ".join(nao_estruturadas))[:200]
        return limpar_valor_extraido(" ".join(restantes))[:200]
    return limpar_valor_extraido(texto.strip())[:200] or "(sem descricao)"


def parse_mensagem(texto: str, remetente: str, timestamp: str, sequencial: int) -> dict | None:
    """Parse a visible WhatsApp message into a ticket payload."""
    trigger = CONFIG.get("trigger_chamado", "#chamado")
    if normalizar_texto(trigger) not in normalizar_texto(texto):
        return None

    solicitante = extrair_campo(texto, "nome", "solicitante") or remetente
    setor = extrair_campo(texto, "setor", "departamento", "area", "área")
    problema = extrair_campo(
        texto,
        "problema",
        "descricao",
        "descrição",
        "assunto",
        "descricao do problema",
        "descrição do problema",
    )

    if not problema:
        problema = extrair_descricao_fallback(texto, trigger)

    return {
        "id": gerar_id_chamado(sequencial),
        "data_hora": timestamp,
        "solicitante": solicitante,
        "setor": setor,
        "problema": problema,
        "prioridade": detectar_prioridade(texto),
        "status": "Aberto",
        "mensagem_original": texto.strip(),
    }


def localizar_chromedriver() -> Path | None:
    """Try to find a local ChromeDriver before using webdriver-manager."""
    caminho_config = str(CONFIG.get("chromedriver_path", "")).strip()
    if caminho_config:
        caminho = Path(caminho_config).expanduser().resolve()
        if caminho.exists():
            return caminho
        log.warning("ChromeDriver configurado nao encontrado: %s", caminho)

    caminho_path = shutil.which("chromedriver")
    if caminho_path:
        return Path(caminho_path).resolve()

    cache_dir = Path.home() / ".wdm" / "drivers" / "chromedriver"
    if not cache_dir.exists():
        return None

    candidatos = sorted(
        [*cache_dir.rglob("chromedriver.exe"), *cache_dir.rglob("chromedriver")],
        key=lambda caminho: caminho.stat().st_mtime,
        reverse=True,
    )
    return candidatos[0] if candidatos else None


def extrair_timestamp(pre_plain_text: str) -> str:
    match = re.search(r"\[(\d{2}:\d{2}),\s*(\d{2}/\d{2}/\d{4})\]", pre_plain_text)
    if match:
        hora, data = match.groups()
        return f"{data} {hora}"
    return datetime.now().strftime("%d/%m/%Y %H:%M")


class WhatsAppMonitor:
    WHATSAPP_URL = "https://web.whatsapp.com"
    CHAT_LIST_SELECTOR = 'div[data-testid="chat-list"]'
    MESSAGE_SELECTOR = "div[data-id]"
    GROUP_XPATH = '//span[@title="{nome_grupo}"]'
    SEARCH_SELECTORS = (
        (By.CSS_SELECTOR, '[data-testid="chat-list-search-container"] input[role="textbox"]'),
        (By.CSS_SELECTOR, 'input[aria-label="Pesquisar ou começar uma nova conversa"]'),
        (By.CSS_SELECTOR, 'input[placeholder*="Pesquisar"]'),
    )

    def __init__(self, config: dict | None = None):
        self.config = config or CONFIG
        self.driver: webdriver.Chrome | None = None
        self.excel = ExcelHandler(self.config["arquivo_excel"])
        self.msgs_vistas: set[str] = set()
        self.seq_chamado = self.excel.proximo_sequencial()

    def _chrome_options(self) -> Options:
        opts = Options()
        perfil = Path("whatsapp_profile").resolve()
        perfil.mkdir(exist_ok=True)

        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--no-first-run")
        opts.add_argument(f"--user-data-dir={perfil}")
        return opts

    def iniciar_browser(self) -> None:
        log.info("Iniciando Chrome...")
        opts = self._chrome_options()
        driver_local = localizar_chromedriver()

        if driver_local:
            log.info("Usando ChromeDriver local: %s", driver_local)
            try:
                self.driver = webdriver.Chrome(service=Service(str(driver_local)), options=opts)
                return
            except Exception as exc:
                log.warning("Falha ao iniciar com ChromeDriver local: %s", exc)

        log.info("Buscando ChromeDriver automaticamente...")
        try:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=opts)
        except Exception as exc:
            raise RuntimeError(
                "Nao foi possivel iniciar o ChromeDriver. "
                "Verifique a conexao com a internet ou configure chromedriver_path."
            ) from exc

    def aguardar_login(self, timeout: int | None = None) -> None:
        if not self.driver:
            raise RuntimeError("Browser ainda nao foi iniciado.")

        tempo_limite = timeout or int(self.config.get("timeout_login", 120))
        log.info("Aguardando login no WhatsApp Web...")
        self.driver.get(self.WHATSAPP_URL)

        try:
            WebDriverWait(self.driver, tempo_limite).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, self.CHAT_LIST_SELECTOR))
            )
            log.info("Login realizado com sucesso.")
        except TimeoutException as exc:
            raise RuntimeError("Timeout ao aguardar login no WhatsApp Web.") from exc

    def _encontrar_grupo_visivel(self, nome_grupo: str, timeout: int) -> webdriver.remote.webelement.WebElement | None:
        if not self.driver:
            return None
        try:
            return WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, self.GROUP_XPATH.format(nome_grupo=nome_grupo)))
            )
        except TimeoutException:
            return None

    def _encontrar_campo_busca(self) -> webdriver.remote.webelement.WebElement:
        if not self.driver:
            raise RuntimeError("Browser ainda nao foi iniciado.")

        for by, selector in self.SEARCH_SELECTORS:
            try:
                return WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((by, selector)))
            except TimeoutException:
                continue
        raise RuntimeError("Campo de busca do WhatsApp nao encontrado.")

    def abrir_grupo(self, nome_grupo: str) -> None:
        if not self.driver:
            raise RuntimeError("Browser ainda nao foi iniciado.")

        log.info("Abrindo grupo: %s", nome_grupo)
        timeout_busca = int(self.config.get("timeout_busca_grupo", 20))
        grupo = self._encontrar_grupo_visivel(nome_grupo, 5)

        if not grupo:
            campo_busca = self._encontrar_campo_busca()
            campo_busca.click()
            campo_busca.send_keys(Keys.CONTROL, "a")
            campo_busca.send_keys(Keys.BACKSPACE)
            campo_busca.send_keys(nome_grupo)
            time.sleep(1.5)
            grupo = self._encontrar_grupo_visivel(nome_grupo, timeout_busca)

        if not grupo:
            raise RuntimeError(f"Grupo '{nome_grupo}' nao encontrado. Verifique o nome em config.py.")

        grupo.click()
        time.sleep(2)
        log.info("Grupo '%s' aberto.", nome_grupo)

    def extrair_texto_mensagem(self, elemento) -> str:
        """Extract the visible message body without quote blocks or metadata."""
        if not self.driver:
            return ""

        try:
            texto = self.driver.execute_script(
                """
                const original = arguments[0];
                const clone = original.cloneNode(true);
                clone.querySelectorAll(
                    '[data-testid="quoted-message"], ' +
                    '[data-testid="author"], ' +
                    '[data-testid="msg-meta"], ' +
                    '[data-icon="tail-in"], ' +
                    '[data-icon="tail-out"]'
                ).forEach(node => node.remove());

                const spans = Array.from(clone.querySelectorAll('span.selectable-text'))
                    .map(node => (node.innerText || '').trim())
                    .filter(Boolean);
                if (spans.length) {
                    return spans.join('\\n').trim();
                }

                const copyable = clone.querySelector('div.copyable-text');
                if (copyable) {
                    return (copyable.innerText || '').trim();
                }

                return (clone.innerText || '').trim();
                """,
                elemento,
            )
            return str(texto).strip()
        except Exception:
            return elemento.text.strip()

    def sincronizar_mensagens_visiveis(self) -> None:
        if not self.driver:
            return

        try:
            for elemento in self.driver.find_elements(By.CSS_SELECTOR, self.MESSAGE_SELECTOR):
                message_id = elemento.get_attribute("data-id")
                if message_id:
                    self.msgs_vistas.add(message_id)
            log.info("Mensagens iniciais sincronizadas: %s", len(self.msgs_vistas))
        except Exception as exc:
            log.warning("Nao foi possivel sincronizar mensagens iniciais: %s", exc)

    def ler_mensagens(self) -> list[dict]:
        if not self.driver:
            return []

        novas: list[dict] = []

        try:
            conteineres = self.driver.find_elements(By.CSS_SELECTOR, self.MESSAGE_SELECTOR)
            for elemento in conteineres:
                message_id = elemento.get_attribute("data-id")
                if not message_id or message_id in self.msgs_vistas:
                    continue

                try:
                    texto = self.extrair_texto_mensagem(elemento)
                    if not texto:
                        continue

                    try:
                        remetente_el = elemento.find_element(By.CSS_SELECTOR, "span[data-testid='author'], span._ahxt")
                        remetente = remetente_el.text.strip() or "Desconhecido"
                    except NoSuchElementException:
                        remetente = "Desconhecido"

                    try:
                        pre_plain = elemento.find_element(By.CSS_SELECTOR, "div[data-pre-plain-text]")
                        timestamp = extrair_timestamp(pre_plain.get_attribute("data-pre-plain-text") or "")
                    except NoSuchElementException:
                        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")

                    novas.append(
                        {
                            "id_msg": message_id,
                            "texto": texto,
                            "remetente": remetente,
                            "timestamp": timestamp,
                        }
                    )
                    self.msgs_vistas.add(message_id)
                except StaleElementReferenceException:
                    continue
        except Exception as exc:
            log.warning("Erro ao ler mensagens: %s", exc)

        return novas

    def processar_mensagens(self, mensagens: list[dict]) -> int:
        registrados = 0

        for mensagem in mensagens:
            chamado = parse_mensagem(
                mensagem["texto"],
                mensagem["remetente"],
                mensagem["timestamp"],
                self.seq_chamado,
            )
            if not chamado:
                continue

            self.excel.salvar_chamado(chamado)
            self.seq_chamado += 1
            registrados += 1
            log.info("Chamado registrado: %s | %s", chamado["id"], chamado["problema"][:60])

        return registrados

    def monitorar(self) -> None:
        intervalo = int(self.config.get("intervalo_segundos", 10))
        grupo = self.config["nome_grupo"]
        log.info("Monitorando grupo '%s' a cada %ss. Use Ctrl+C para parar.", grupo, intervalo)

        while True:
            try:
                mensagens = self.ler_mensagens()
                if mensagens:
                    self.processar_mensagens(mensagens)
                time.sleep(intervalo)
            except KeyboardInterrupt:
                raise
            except Exception as exc:
                log.error("Erro no loop principal: %s", exc)
                time.sleep(5)

    def fechar(self) -> None:
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    def executar(self) -> int:
        try:
            self.iniciar_browser()
            self.aguardar_login()
            self.abrir_grupo(self.config["nome_grupo"])
            self.sincronizar_mensagens_visiveis()
            self.monitorar()
            return 0
        except KeyboardInterrupt:
            log.info("Monitoramento encerrado pelo usuario.")
            return 0
        except Exception as exc:
            log.exception("Falha na execucao do monitor: %s", exc)
            return 1
        finally:
            self.fechar()
            self.excel.fechar()
            log.info("Script finalizado.")


if __name__ == "__main__":
    raise SystemExit(WhatsAppMonitor().executar())
