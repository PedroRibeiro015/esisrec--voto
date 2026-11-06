"""
Pauta os processos no e-SISREC: abre a lista 'Relatar', pergunta quais processos serão pautados e
qual será a pauta, e em cada processo digitado que estiver na lista abre o 'Relatório e Voto',
preenche Relatório, Voto e Ementa com "xxx", faz as marcações, preenche a pauta e SALVA.
Não baixa PDF, não anexa documento e nunca assina nem finaliza.

Este arquivo funciona sozinho. Só precisa do Google Chrome instalado e de:
  pip install playwright openpyxl

Uso:
  python pauta_processos.py

Marcações: matéria de alçada 'Não', decisão 'Conhecer do Recurso e negar-lhe provimento', tipo de
decisão 'Sem concessão' e tipo de objeto 'Sim'.

Salvaguardas:
  - Pede confirmação antes de gravar. Para testar, digite um processo só.
  - Não apaga texto de verdade: se o Relatório, o Voto ou a Ementa já tiver outro texto, não salva
    esse processo e avisa no final.
  - Sem a pauta preenchida, o processo não é salvo.

Login: o script abre o Chrome com um perfil próprio, você faz o login e ele continua sozinho.
Os números podem ser digitados ou colados, um por linha ou separados por vírgula, espaço ou ';',
com ou sem pontuação. Linha vazia termina a digitação. Controle em Downloads/processos/pauta_log.xlsx.

Pauta: é uma lista suspensa na tela do Relatório e Voto. Digite o nome da pauta; acentos, maiúsculas
e zeros à esquerda não importam ('sessao 615' escolhe 'Sessão 0615'). Se a pauta não estiver na lista
(ou corresponder a mais de uma opção), o processo não é salvo e o robô mostra as opções que encontrou.
"""
import os
import re
import shutil
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

# ---------------------------------------------------------------------------
# Robô do e-SISREC (mesmo código do pipeline_processos.py; correção lá não chega aqui sozinha)
# ---------------------------------------------------------------------------

URL_SISTEMA = "https://esisrec.inss.gov.br/esisrec/"
UNIDADE = "GOV:CRPS.10ª JR"
TEXTO_LINK_LISTA = 'Processos com Localizador "Relatar"'

PASTA_BASE = Path.home() / "Downloads" / "processos"

# Chrome em que você faz o login (modo padrão da coleta). Perfil próprio, separado do seu
# Chrome pessoal: guarda o login do e-SISREC entre uma execução e outra.
PASTA_PERFIL_CHROME = PASTA_BASE / "perfil_chrome_robo"
PORTA_CHROME = 9222

# Tempos de espera do robô, em milissegundos.
ESPERA_MS = 120_000        # elementos da página (2 min). Use 0 para esperar sem limite.
ESPERA_CURTA_MS = 15_000   # tentativa principal, antes de usar a alternativa
ESPERA_PDF_MS = 300_000    # geração e carregamento do PDF (5 min)

PADRAO_PROCESSO = re.compile(r"\d{5}\.\d{6}/\d{4}-\d{2}")

TABELA_PROCESSOS = "#tblProcessosMovimentacao"
PAGINADORES = (f"{TABELA_PROCESSOS}_paginator_top", f"{TABELA_PROCESSOS}_paginator_bottom",
               "[id$='_paginator_top']", "[id$='_paginator_bottom']")


def agora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class Log:
    """Planilha de controle: uma linha por processo, atualizada a cada passo.
    Se a planilha já existe, ela é reaproveitada (o histórico não é apagado)."""

    def __init__(self, caminho: Path):
        from openpyxl import Workbook, load_workbook
        self.caminho = caminho
        self.linhas: dict[str, int] = {}
        if caminho.exists():
            try:
                self.wb = load_workbook(caminho)
                # Pela aba "Processos", não pela aba ativa (que muda se alguém salvar em outra aba).
                self.ws = self.wb["Processos"] if "Processos" in self.wb.sheetnames else self.wb.active
                for linha in range(2, self.ws.max_row + 1):
                    numero = self.ws.cell(row=linha, column=1).value
                    if numero:
                        self.linhas[str(numero)] = linha
                return
            except Exception as e:
                # Não cria uma planilha nova: ela seria salva por cima desta, e o robô esqueceria
                # quais processos já têm o documento anexado (e anexaria de novo).
                raise SystemExit(f"Não consegui ler {caminho} ({e}). Feche-a no Excel e rode de novo. "
                                 f"Se ela estiver estragada, restaure uma cópia; não apague.")
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "Processos"
        self.ws.append(["Processo", "Página", "Status", "Data/Hora"])

    def adicionar(self, numero: str, pagina: int):
        if numero in self.linhas:
            self.ws.cell(row=self.linhas[numero], column=2, value=pagina)
            return
        self.ws.append([numero, pagina, "Pendente", agora()])
        self.linhas[numero] = self.ws.max_row

    def status(self, numero: str) -> str:
        if numero not in self.linhas:
            return ""
        return str(self.ws.cell(row=self.linhas[numero], column=3).value or "")

    def atualizar(self, numero: str, status: str):
        linha = self.linhas[numero]
        self.ws.cell(row=linha, column=3, value=status)
        self.ws.cell(row=linha, column=4, value=agora())
        self.salvar()

    def salvar(self):
        try:
            self.wb.save(self.caminho)
        except PermissionError:
            print("AVISO: não consegui salvar a planilha (ela está aberta no Excel?). Feche-a.")


def esperar_visivel(locator, timeout: int = ESPERA_MS):
    locator.wait_for(state="visible", timeout=timeout)
    return locator


def clicar(locator, descricao: str, tag: str, timeout: int = ESPERA_MS):
    esperar_visivel(locator, timeout).click()
    print(f"OK [{tag}] {descricao}")


# Onde ficam os números de processo da lista, em ordem de preferência: a lista "Relatar" (id
# conhecido) e, em outras telas (como o Relatório de Sessão), a tabela de dados da página, pelos
# links ou, se os números não forem links, pelas linhas.
SELETORES_LISTA = (f"{TABELA_PROCESSOS} a", ".ui-datatable-data a", ".ui-datatable-data > tr")


def extrair_numeros_pagina(page, seletor: str) -> list[str]:
    """Extrai os números de processo dos elementos da lista (um número por elemento)."""
    numeros = []
    for texto in page.locator(seletor).all_inner_texts():
        achado = PADRAO_PROCESSO.search(texto)
        if achado and achado.group() not in numeros:
            numeros.append(achado.group())
    return numeros


def localizar_lista(page, timeout_ms: int = ESPERA_CURTA_MS) -> str | None:
    """Espera a lista de processos aparecer e devolve o seletor dela (de SELETORES_LISTA), ou None.
    Esperar é preciso: com a tabela ainda carregando, o robô lia uma lista vazia ou incompleta."""
    limite = time.monotonic() + timeout_ms / 1000
    while True:
        for seletor in SELETORES_LISTA:
            if extrair_numeros_pagina(page, seletor):
                return seletor
        if time.monotonic() > limite:
            return None
        page.wait_for_timeout(500)


def achar_no_paginador(page, classe: str):
    """Primeiro elemento visível com essa classe no paginador da lista, ou None.
    Procura no paginador de cima e no de baixo; se a tabela vier com outro id, em qualquer paginador."""
    for paginador in PAGINADORES:
        alvo = page.locator(f"{paginador} {classe}").filter(visible=True)
        if alvo.count():
            return alvo.first
    return None


def texto_contador(page) -> str:
    """Texto do contador do paginador, ex.: '(1 de 9)'. Vazio se não achar."""
    try:
        contador = achar_no_paginador(page, ".ui-paginator-current")
        return contador.inner_text().strip() if contador else ""
    except Exception:
        return ""


def obter_total_paginas(page) -> int:
    """Lê o total de páginas a partir do texto '(1 de 9)'."""
    # Espera a lista carregar (lendo na hora, com a lista carregando, dava "1 página").
    limite = time.monotonic() + ESPERA_CURTA_MS / 1000
    while True:
        texto = texto_contador(page)
        achado = re.search(r"de\s+(\d+)", texto)
        if achado:
            total = int(achado.group(1))
            print(f"Total de páginas detectado: {total} (contador: '{texto}')")
            return total
        if time.monotonic() > limite:
            print(f"AVISO: não achei o contador de páginas ('1 de N'). Texto lido: '{texto}'")
            return 1
        page.wait_for_timeout(500)


def ir_para_proxima_pagina(page, seletor: str, numeros_atuais: list[str]) -> bool:
    """Clica em 'próxima' e espera a tabela trocar de conteúdo.
    (Os botões numerados não servem: o paginador só mostra ~10 deles por vez.)"""
    return trocar_pagina(page, seletor, numeros_atuais, ".ui-paginator-next", "próxima")


def trocar_pagina(page, seletor: str, numeros_atuais: list[str], classe: str, nome: str) -> bool:
    """Clica no botão do paginador (classe) e espera a tabela trocar de conteúdo."""
    try:
        botao = achar_no_paginador(page, classe)
        if botao is None:
            print(f"AVISO: não achei o botão '{nome} página' na lista.")
            return False
        if "ui-state-disabled" in (botao.get_attribute("class") or ""):
            return False  # já está na última (ou na primeira) página
        contador_antes = texto_contador(page)
        botao.click()
        # Aqui o limite é sempre finito (mesmo com ESPERA_MS = 0): se a tabela nunca mudar,
        # o robô ficaria preso para sempre.
        limite = time.monotonic() + (ESPERA_MS / 1000 if ESPERA_MS else ESPERA_PDF_MS / 1000)
        reclicar = time.monotonic() + ESPERA_CURTA_MS / 1000
        while time.monotonic() < limite:
            page.wait_for_timeout(500)
            novos = extrair_numeros_pagina(page, seletor)
            if novos and novos != numeros_atuais:
                page.wait_for_timeout(1000)  # deixa a tabela terminar de atualizar
                return True
            # O clique às vezes se perde (lista ainda carregando). Clica de novo uma vez, mas só se
            # o contador não mudou: se a página já trocou, outro clique pularia uma página.
            if reclicar and time.monotonic() > reclicar:
                reclicar = None
                if texto_contador(page) == contador_antes:
                    print(f"AVISO: a lista não mudou; clicando em '{nome}' de novo.")
                    botao.dispatch_event("click")
        print(f"AVISO: a tabela não mudou depois de clicar em '{nome}' "
              f"(contador antes: '{contador_antes}', agora: '{texto_contador(page)}').")
    except Exception as e:
        print(f"AVISO: erro ao ir para a {nome} página: {e}")
    return False


def so_digitos(texto: str) -> str:
    return re.sub(r"[./-]", "", texto)


def processo_na_tela(page, numero: str, timeout_ms: int = ESPERA_CURTA_MS) -> bool:
    """A tela aberta é mesmo a do processo 'numero'? (Evita baixar ou gravar no processo errado.)
    Compara só os dígitos, porque a tela pode escrever o número com ou sem pontuação."""
    alvo = so_digitos(numero)
    limite = time.monotonic() + timeout_ms / 1000
    while True:
        try:
            if alvo in so_digitos(page.locator("body").inner_text()):
                return True
        except Exception:
            pass  # a página estava carregando
        if time.monotonic() > limite:
            return False
        page.wait_for_timeout(500)


def coletar_todos_processos(page) -> list[tuple[str, int]]:
    """Percorre todas as páginas da lista. Retorna [(numero, pagina), ...] sem repetidos."""
    processos, vistos = [], set()
    seletor = localizar_lista(page)
    if seletor is None:
        print("AVISO: não achei números de processo na tela aberta.")
        return processos
    total = obter_total_paginas(page)
    print(f"Coletando processos de {total} página(s)...")
    pagina = 1
    while True:
        numeros = extrair_numeros_pagina(page, seletor)
        novos = 0
        for numero in numeros:
            if numero not in vistos:
                vistos.add(numero)
                processos.append((numero, pagina))
                novos += 1
        print(f"Página {pagina}: {novos} processos. Total parcial: {len(processos)}")
        if not ir_para_proxima_pagina(page, seletor, numeros):
            if pagina < total:
                print(f"AVISO: não foi possível ir para a página {pagina + 1} de {total}. "
                      f"Parando a coleta.")
            break
        pagina += 1
    print(f"Total coletado: {len(processos)} processos.")
    return processos


def abrir_lista_relatar(page):
    """Com o login já feito (esperar_login), escolhe a unidade e abre a lista 'Relatar'."""
    selecionar_unidade_se_preciso(page)

    print("Aguardando o link da lista de processos...")
    clicar(page.get_by_text(TEXTO_LINK_LISTA, exact=True), "Abriu a lista 'Relatar'", "lista",
           timeout=0)
    page.wait_for_timeout(5000)


def selecionar_unidade_se_preciso(page):
    unidade = page.locator("select#domains")
    if unidade.count() > 0 and unidade.first.is_visible():
        print(f"Tela de unidade detectada. Selecionando '{UNIDADE}'...")
        page.select_option("select#domains", label=UNIDADE)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)


def voltar_pagina_inicial(page):
    try:
        # Espera por qualquer um dos dois botões (antes o robô olhava uma vez só, sem esperar).
        botao = page.locator("[id='formProcesso:btPaginaInicialFim'], "
                             "[id='formProcesso:btPaginaInicial']").first
        clicar(botao, "Voltou à Página Inicial", "navegação", timeout=10_000)
        page.wait_for_timeout(3000)
    except Exception:
        print("AVISO: botão 'Página Inicial' não encontrado; recarregando o sistema.")
        recarregar_sistema(page)


def fechar_dialogos(page):
    """Fecha as janelas abertas por cima da tela (o visualizador de PDF e, se um erro a deixou
    aberta, a de anexar). Com elas abertas, a busca rápida e os botões da tela não recebem o clique."""
    for _ in range(3):
        try:
            abertas = page.locator(".ui-dialog").filter(visible=True)
            if not abertas.count():
                return
            fechar = abertas.last.locator(".ui-dialog-titlebar-close")
            if fechar.count():
                fechar.first.click(timeout=ESPERA_CURTA_MS)
            else:
                page.keyboard.press("Escape")
        except Exception:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
        page.wait_for_timeout(1000)
    print("AVISO: não consegui fechar a janela aberta por cima da tela (visualizador de PDF?).")


def busca_rapida_na_tela(page) -> bool:
    """O campo de busca rápida está visível na tela atual (sem precisar ir à Página Inicial)?"""
    try:
        return page.locator("#formBuscaRapida\\:buscaRapida").filter(visible=True).count() > 0
    except Exception:
        return False


def recarregar_sistema(page):
    """Abre o sistema de novo (tela limpa) e escolhe a unidade, se ele pedir."""
    page.goto(URL_SISTEMA)
    try:
        page.locator("select#domains").or_(page.locator("[id='formBuscaRapida:buscaRapida']")) \
            .filter(visible=True).first.wait_for(state="visible", timeout=ESPERA_MS)
    except Exception:
        pass
    selecionar_unidade_se_preciso(page)
    page.wait_for_timeout(3000)


def fazer_busca_processo(page, numero: str, tentar_de_novo: bool = True) -> bool:
    """Preenche a busca rápida com o número do processo e espera a tela do processo.
    Às vezes a tela fica com DUAS buscas rápidas (mesmo id, depois de várias telas abertas por
    cima umas das outras). Usa a última visível e o botão do mesmo formulário; se a busca falhar,
    recarrega o sistema e tenta mais uma vez."""
    try:
        campos = page.locator("[id='formBuscaRapida:buscaRapida']").filter(visible=True)
        campos.first.wait_for(state="visible", timeout=ESPERA_MS)
        if campos.count() > 1:
            print(f"AVISO: a tela tem {campos.count()} campos de busca rápida; usando o último.")
        campo = campos.last
        campo.fill(numero)
        botao = campo.locator("xpath=ancestor::form[1]").locator("[id='formBuscaRapida:btBuscaRapida']")
        if not botao.count():
            botao = page.locator("[id='formBuscaRapida:btBuscaRapida']").filter(visible=True)
        botao.last.click()
        page.wait_for_selector("[id='formProcesso:incluirDocumentoRequerimento']",
                               state="visible", timeout=ESPERA_MS)
        if not processo_na_tela(page, numero):
            print(f"ERRO A busca abriu um processo, mas o número {numero} não aparece na tela.")
            return False
        print(f"OK Carregou processo: {numero}")
        return True
    except Exception as e:
        print(f"ERRO Falha ao buscar processo {numero}: {str(e)[:300]}")
        if tentar_de_novo:
            print("Recarregando o sistema e tentando a busca de novo...")
            try:
                fechar_dialogos(page)
                recarregar_sistema(page)
            except Exception as e2:
                print(f"AVISO: não consegui recarregar o sistema: {str(e2)[:200]}")
                return False
            return fazer_busca_processo(page, numero, tentar_de_novo=False)
        return False


def encontrar_chrome() -> str | None:
    """Caminho do Chrome (ou do Edge, se não houver Chrome). CHROME_PATH tem prioridade."""
    candidatos = [os.getenv("CHROME_PATH"), shutil.which("chrome"), shutil.which("msedge")]
    for base in (os.getenv("PROGRAMFILES"), os.getenv("PROGRAMFILES(X86)"), os.getenv("LOCALAPPDATA")):
        if base:
            candidatos.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))
    for base in (os.getenv("PROGRAMFILES(X86)"), os.getenv("PROGRAMFILES")):
        if base:
            candidatos.append(os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe"))
    return next((c for c in candidatos if c and os.path.isfile(c)), None)


def endereco_chrome() -> str | None:
    """Endereço da porta de conexão do Chrome do robô, ou None se ele não estiver aberto.
    O Chrome às vezes escuta só no IPv6 ([::1]) e às vezes só no IPv4 (127.0.0.1): testa os dois."""
    from urllib.request import urlopen
    for host in ("127.0.0.1", "[::1]"):
        endereco = f"http://{host}:{PORTA_CHROME}"
        try:
            with urlopen(f"{endereco}/json/version", timeout=2):
                return endereco
        except Exception:
            pass
    return None


def chrome_aberto() -> bool:
    """O Chrome do robô (com a porta de conexão) já está aberto?"""
    return endereco_chrome() is not None


def comando_chrome(chrome: str) -> list[str]:
    return [chrome,
            f"--remote-debugging-port={PORTA_CHROME}",  # só aceita conexões deste computador
            f"--user-data-dir={PASTA_PERFIL_CHROME}",   # perfil separado do seu Chrome pessoal
            "--no-first-run", "--no-default-browser-check"]


def abrir_chrome_para_login():
    """Abre um Chrome NORMAL (não controlado pelo robô), sem abrir o site: você entra no
    e-SISREC e faz o login nele. Assim o login (e o reCAPTCHA, se aparecer) é feito por você
    num navegador comum. O robô percebe sozinho quando o login termina (esperar_login)."""
    import subprocess

    if chrome_aberto():
        print(f"Chrome do robô já está aberto (porta {PORTA_CHROME}); usando ele.")
    else:
        chrome = encontrar_chrome()
        if not chrome:
            raise SystemExit("Chrome não encontrado. Instale o Google Chrome, defina CHROME_PATH "
                             "com o caminho do chrome.exe, ou use --navegador playwright.")
        PASTA_PERFIL_CHROME.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(comando_chrome(chrome))
        limite = time.monotonic() + 30
        while not chrome_aberto():
            if time.monotonic() > limite:
                raise SystemExit(f"O Chrome abriu, mas a porta {PORTA_CHROME} não respondeu. Se já "
                                 f"havia um Chrome com este perfil aberto, feche-o e rode de novo.")
            time.sleep(0.5)


def tela_pos_login(page):
    """Elementos que só aparecem depois do login: a escolha de unidade ou o link da lista."""
    return page.locator("select#domains").or_(page.get_by_text(TEXTO_LINK_LISTA, exact=True))


def esperar_login(context):
    """Espera, sem limite de tempo, alguma aba chegar à tela de unidade ou à tela principal do
    e-SISREC (ou seja, o login foi feito) e devolve essa aba. Não abre nem recarrega o site."""
    print("=" * 60)
    print(f"No navegador do robô, entre no e-SISREC ({URL_SISTEMA}) e faça o login.")
    print("O robô começa sozinho quando aparecer a tela de unidade ou a tela principal.")
    print("(Se o sistema já estiver aberto em outra tela, volte à página inicial dele.)")
    print("=" * 60)
    host = urlparse(URL_SISTEMA).netloc
    while True:
        for page in list(context.pages):
            if urlparse(page.url).netloc != host:
                continue
            try:
                if tela_pos_login(page).filter(visible=True).count() > 0:
                    print("Login detectado. Começando.")
                    return page
            except Exception:
                pass  # a aba estava trocando de página ou foi fechada
        # Uma chamada ao navegador (e não time.sleep), para a lista de abas se atualizar.
        try:
            (context.pages[0] if context.pages else context.new_page()).wait_for_timeout(1000)
        except Exception:
            time.sleep(1)


@contextmanager
def sistema_logado(navegador: str = "chrome"):
    """Abre o navegador (ou conecta ao Chrome do robô), espera você fazer o login e entrega a aba
    do e-SISREC, já com a unidade escolhida. Usado pelas etapas 'coletar' e 'enviar'."""
    from playwright.sync_api import sync_playwright

    if navegador == "chrome":
        abrir_chrome_para_login()

    with sync_playwright() as p:
        if navegador == "chrome":
            # Conecta ao Chrome em que você já fez o login.
            browser = p.chromium.connect_over_cdp(endereco_chrome())
            context = browser.contexts[0]
        else:
            # Modo antigo: o próprio Playwright abre o navegador (o reCAPTCHA pode bloquear o login).
            browser = p.chromium.launch(headless=False)
            # locale pt-BR: o contador do visualizador de PDF precisa mostrar "de N" (e não "of N").
            context = browser.new_context(accept_downloads=True, locale="pt-BR",
                                          viewport={"width": 1920, "height": 1080})
            context.new_page()
        try:
            page = esperar_login(context)
            page.bring_to_front()
            selecionar_unidade_se_preciso(page)
            yield page
        finally:
            # Se você fechou o navegador na mão, fechar de novo dá erro e esconderia o erro real.
            try:
                if navegador == "chrome":
                    browser.close()  # só desconecta; o Chrome continua aberto
                    print("Robô desconectado. Pode fechar a janela do Chrome do robô.")
                else:
                    context.close()
                    browser.close()
            except Exception:
                pass


FORM_VOTO = "formRelatorioVoto"
# Parte do .docx (título sem acento) -> editor da tela.
CAMPOS_TEXTO = {"RELATORIO": f"{FORM_VOTO}:relator:txtRelatorio",
                "VOTO": f"{FORM_VOTO}:relator:txtVoto",
                "EMENTA": f"{FORM_VOTO}:relator:txtEmenta"}
TABELA_TIPO_OBJETO = f"{FORM_VOTO}:relator:tabParametrosDecisoesMultiplas"
BOTAO_SALVAR = f"{FORM_VOTO}:btnSalvar"

# Estrutura das telas (vista no DevTools):
#  - Tipo de decisão: tabela; em cada linha, <td role="gridcell" class="ui-selection-column"> com
#    div.ui-radiobutton > div.ui-radiobutton-box, e o texto no <td role="gridcell"> seguinte.
#  - Tipo de objeto: tabela; em cada linha, <td role="gridcell" class="ui-selection-column"> com
#    div.ui-chkbox > div.ui-chkbox-box (o quadrado de marcar).
#  - Sim/Não do tipo de objeto: <div id="...:tabParametrosDecisoesMultiplas:0:sltParametroDecisaoSimNaoMultiplo"
#    class="ui-selectoneradio ui-grid ui-grid-responsive">, com as opções em div.ui-g > div.ui-g-12.
#  - Pauta: <label id="formRelatorioVoto:relator:sltPauta_label" class="ui-selectonemenu-label">
#    (mostra &nbsp; enquanto nada foi escolhido).
# Marcado = a caixa (…-box) ganha a classe ui-state-active; desmarcado = ui-state-default.
CAIXA_RADIO_LINHA = "td.ui-selection-column .ui-radiobutton-box"
CAIXA_CHECK_LINHA = "td.ui-selection-column .ui-chkbox-box"
SUFIXO_SIM_NAO = ":sltParametroDecisaoSimNaoMultiplo"

# Marcações feitas em todo processo enviado (o resultado é sempre negar provimento; a revisão
# fica com o advogado depois).
OPCAO_ALCADA = "Não"
OPCAO_DECISAO = "Conhecer do Recurso e negar-lhe provimento"
OPCAO_TIPO_DECISAO = "Sem concessão/Reativação/Revisão de benefício"
OPCAO_TIPO_OBJETO = "Sim"
ESPERA_SALVAR_MS = 60_000

STATUS_SALVO = "Salvo"
STATUS_SALVO_INCERTO = "Salvo? O sistema não confirmou: conferir"


class ErroAoSalvar(Exception):
    """O sistema recusou o salvamento (campo obrigatório, por exemplo)."""


def abrir_relatorio_voto(page):
    clicar(page.locator("[id='formProcesso:relatorioVoto']"), "Abriu 'Relatório e Voto'", "enviar")
    page.locator(f"[id='{CAMPOS_TEXTO['RELATORIO']}']").wait_for(state="attached", timeout=ESPERA_MS)


JS_EDITOR = """id => Object.values(window.CKEDITOR ? CKEDITOR.instances : {}).find(
    e => e.name === id || (e.element && e.element.getId() === id))"""


def esperar_editor(page, id_campo: str):
    page.wait_for_function(f"id => {{ const e = ({JS_EDITOR})(id); return e && e.status === 'ready'; }}",
                           arg=id_campo, timeout=ESPERA_MS)


def texto_do_editor(page, id_campo: str) -> str:
    esperar_editor(page, id_campo)
    return page.evaluate(f"id => ({JS_EDITOR})(id).getData().replace(/<[^>]+>|&nbsp;/g, '').trim()",
                         id_campo)


def preencher_editor(page, id_campo: str, html: str, descricao: str):
    """Põe o texto no editor (CKEditor) e copia para o campo que o formulário envia."""
    esperar_editor(page, id_campo)
    tamanho = page.evaluate(f"""async ([id, html]) => {{
        const e = ({JS_EDITOR})(id);
        await new Promise(fim => e.setData(html, {{callback: fim}}));
        e.updateElement();
        e.fire('change');
        return e.getData().replace(/<[^>]+>/g, '').trim().length;
    }}""", [id_campo, html])
    if not tamanho:
        raise RuntimeError(f"{descricao}: o texto não entrou no editor")
    print(f"OK [enviar] {descricao} preenchido ({tamanho} caracteres)")


def caixa_ativa(caixa) -> bool:
    return "ui-state-active" in (caixa.get_attribute("class") or "")


def esperar_condicao_caixa(caixa, condicao) -> bool:
    """Espera a caixa atender à condição: o sistema atualiza a tela depois de cada clique, às vezes devagar."""
    limite = time.monotonic() + ESPERA_CURTA_MS / 1000
    while not condicao(caixa):
        if time.monotonic() > limite:
            return False
        caixa.page.wait_for_timeout(500)
    return True


def caixa_liberada(caixa) -> bool:
    return "ui-state-disabled" not in (caixa.get_attribute("class") or "")


def marcar_opcao(escopo, rotulo: str, descricao: str):
    """Marca o botão de opção (PrimeFaces) que tem o texto 'rotulo', dentro de 'escopo'."""
    exato = re.compile(rf"^\s*{re.escape(rotulo)}\s*$", re.IGNORECASE)
    rotulos = escopo.locator("label").filter(visible=True).filter(has_text=exato)
    # Algumas listas (ex.: Tipo de Decisão) são tabelas: a bolinha fica numa coluna e o texto em outra,
    # sem <label>.
    linhas = escopo.locator("tr:has(> td.ui-selection-column)").filter(visible=True).filter(
        has=escopo.page.locator("td[role='gridcell']", has_text=exato))
    # Algumas opções só aparecem depois de outra escolha (o sistema atualiza a tela, às vezes devagar).
    limite = time.monotonic() + ESPERA_MS / 1000
    if not (rotulos.count() or linhas.count()):
        print(f"[enviar] {descricao}: esperando a opção '{rotulo}' aparecer (até {ESPERA_MS // 1000} s)...")
    while not (rotulos.count() or linhas.count()) and time.monotonic() < limite:
        escopo.page.wait_for_timeout(500)
    if rotulos.count() + linhas.count() != 1:
        # Mostra as opções que estão na tela, para saber se foi lentidão ou se o texto é outro.
        textos = escopo.locator("label, td[role='gridcell']").filter(visible=True).all_inner_texts()
        na_tela = [t.strip() for t in textos if t.strip()]
        raise RuntimeError(f"{descricao}: esperava 1 opção '{rotulo}' na tela, "
                           f"achei {rotulos.count() + linhas.count()}. Opções na tela: {na_tela[:30]}")
    alvo = rotulos.first.get_attribute("for") if rotulos.count() else None
    if linhas.count():
        caixa = linhas.first.locator(CAIXA_RADIO_LINHA)
    elif alvo:
        caixa = escopo.page.locator(f"[id='{alvo}']").locator(
            "xpath=ancestor::div[contains(@class,'ui-radiobutton')][1]/div[contains(@class,'ui-radiobutton-box')]")
    else:
        caixa = rotulos.first.locator(
            "xpath=ancestor::*[.//div[contains(@class,'ui-radiobutton-box')]][1]"
            "//div[contains(@class,'ui-radiobutton-box')]")
    if caixa.count() != 1:
        raise RuntimeError(f"{descricao}: não achei o botão da opção '{rotulo}'")
    if not caixa_ativa(caixa):
        # O 'Sim' do tipo de objeto só é liberado depois que o quadrado da linha é marcado.
        if not esperar_condicao_caixa(caixa, caixa_liberada):
            raise RuntimeError(f"{descricao}: a opção '{rotulo}' continuou desativada")
        caixa.click()
    if not esperar_condicao_caixa(caixa, caixa_ativa):
        raise RuntimeError(f"{descricao}: a opção '{rotulo}' não ficou marcada")
    print(f"OK [enviar] {descricao}: {rotulo}")


def marcar_tipo_objeto(page):
    """Marca o quadrado da linha do tipo de objeto e 'Sim' nela. Só se houver uma linha: com
    mais de uma, não dá para saber qual marcar."""
    # data-ri: só as linhas da tabela (cada linha tem dentro o grupo Sim/Não).
    linhas = page.locator(f"[id='{TABELA_TIPO_OBJETO}'] tr[data-ri]")
    linhas.first.wait_for(state="visible", timeout=ESPERA_CURTA_MS)
    if linhas.count() != 1:
        raise RuntimeError(f"Tipo de objeto: esperava 1 linha, achei {linhas.count()}")
    linha = linhas.first
    quadrado = linha.locator(CAIXA_CHECK_LINHA)
    if quadrado.count() != 1:
        raise RuntimeError("Tipo de objeto: não achei o quadrado de marcar")
    if not caixa_ativa(quadrado):
        quadrado.click()
    if not esperar_condicao_caixa(quadrado, caixa_ativa):
        raise RuntimeError("Tipo de objeto: o quadrado não ficou marcado")
    # O Sim/Não é o div ...:N:sltParametroDecisaoSimNaoMultiplo dentro da linha (N = número da linha).
    grupo_sim_nao = linha.locator(f"[id$='{SUFIXO_SIM_NAO}']")
    try:
        grupo_sim_nao.first.wait_for(state="visible", timeout=ESPERA_CURTA_MS)
        escopo = grupo_sim_nao.first
    except Exception:
        print("AVISO [enviar] Tipo de objeto: não achei o grupo Sim/Não pelo id; procurando na linha.")
        escopo = linha
    marcar_opcao(escopo, OPCAO_TIPO_OBJETO, "Tipo de objeto")


def salvar_relatorio_voto(page) -> str:
    """Clica em Salvar e lê a mensagem do sistema."""
    mensagens = page.locator(".ui-growl-item, .ui-messages-info, .ui-messages-warn, "
                             ".ui-messages-error, .ui-messages-fatal")
    # Some com mensagens antigas, para não confundir com a resposta deste salvamento.
    page.evaluate("""() => document.querySelectorAll('.ui-growl-item, .ui-messages-info, .ui-messages-warn,'
        + ' .ui-messages-error, .ui-messages-fatal').forEach(m => m.remove())""")
    page.locator(f"[id='{BOTAO_SALVAR}']").click()
    try:
        mensagens.first.wait_for(state="visible", timeout=ESPERA_SALVAR_MS)
    except Exception:
        return STATUS_SALVO_INCERTO
    page.wait_for_timeout(1000)
    texto = " | ".join(t.strip() for t in mensagens.filter(visible=True).all_inner_texts())[:300]
    erro = page.locator(".ui-messages-error, .ui-messages-fatal, .ui-growl-image-error, .ui-growl-image-fatal")
    if erro.filter(visible=True).count():
        raise ErroAoSalvar(texto or "o sistema mostrou um erro")
    print(f"OK [enviar] Salvo. Mensagem do sistema: {texto}")
    return STATUS_SALVO


# ---------------------------------------------------------------------------
# Pauta
# ---------------------------------------------------------------------------

ARQ_LOG_PAUTA = PASTA_BASE / "pauta_log.xlsx"
TEXTO_PROVISORIO = "xxx"
ID_PAUTA = f"{FORM_VOTO}:relator:sltPauta"  # lista suspensa: {ID_PAUTA}_label e {ID_PAUTA}_panel

# Aceita 12345.678901/2026-12 e também 12345.678901_2026-12 ou só os 17 dígitos.
PADRAO_DIGITADO = re.compile(r"(?<!\d)(\d{5})\.?(\d{6})[/_]?(\d{4})-?(\d{2})(?!\d)")


class TextoExistente(Exception):
    """O processo já tem texto de verdade, que o "xxx" apagaria."""


def ler_numeros(texto: str) -> tuple[list[str], list[str]]:
    """Separa o texto em (números de processo no formato do sistema, trechos que não são números)."""
    numeros = [f"{a}.{b}/{c}-{d}" for a, b, c, d in PADRAO_DIGITADO.findall(texto)]
    nao_entendidos = re.findall(r"[^\s,;]+", PADRAO_DIGITADO.sub(" ", texto))
    return numeros, nao_entendidos


def perguntar_processos() -> list[str]:
    """Pergunta os processos a pautar. Devolve os números na ordem digitada, sem repetir."""
    print("\n" + "=" * 60)
    print("Quais processos serão pautados?")
    print("Digite ou cole os números (ex.: 12345.678901/2026-12). Linha vazia termina.")
    print("=" * 60)
    pedidos = []
    while True:
        try:
            linha = input("> ")
        except EOFError:
            break
        if not linha.strip():
            break
        numeros, nao_entendidos = ler_numeros(linha)
        for numero in numeros:
            if numero not in pedidos:
                pedidos.append(numero)
        if nao_entendidos:
            print(f"AVISO: não entendi {', '.join(nao_entendidos)}. Digite esse(s) de novo.")
    return pedidos


def perguntar_pauta() -> str:
    """Pergunta a pauta. O texto é preenchido exatamente como foi digitado."""
    try:
        return input("\nQual será a pauta? ").strip()
    except EOFError:
        return ""


def confirmar(encontrados: list[str], pauta: str) -> bool:
    print(f"\nProcessos: {', '.join(encontrados)}")
    print(f"Pauta: {pauta}")
    try:
        resposta = input(f"O robô vai preencher '{TEXTO_PROVISORIO}' no Relatório, Voto e Ementa, fazer as "
                         f"marcações, preencher a pauta e SALVAR no e-SISREC. Continuar? (s/n) ")
    except EOFError:
        resposta = ""
    return resposta.strip().lower() in ("s", "sim")


def lista_lida_ate_o_fim(page) -> bool:
    """Depois de ler a lista: ela está na última página? (O botão 'próxima' fica desativado nela.)"""
    proxima = achar_no_paginador(page, ".ui-paginator-next")
    return proxima is None or "ui-state-disabled" in (proxima.get_attribute("class") or "")


def so_palavras(texto: str) -> str:
    return " ".join(texto.split()).lower()


def normalizar_pauta(texto: str) -> list[str]:
    """Palavras da pauta para comparar: sem acento, minúsculas e números sem zeros à esquerda.
    Assim 'sessao 615' fica igual a 'Sessão 0615'."""
    import unicodedata
    sem_acento = "".join(c for c in unicodedata.normalize("NFD", texto)
                         if unicodedata.category(c) != "Mn")
    palavras = re.findall(r"\d+|[^\W\d_]+", sem_acento.lower())
    # A lista mostra só '0763/2026 E - 11/12/2026 09:45', sem a palavra 'Sessão': essas palavras
    # digitadas não contam.
    return [str(int(p)) if p.isdigit() else p for p in palavras if p not in PALAVRAS_IGNORADAS]


PALAVRAS_IGNORADAS = {"sessao", "sessoes", "pauta", "n", "no", "nº", "numero", "de", "da", "do"}


def escolher_opcao_pauta(pauta: str, textos: list[str]) -> int:
    """Índice da opção que corresponde à pauta digitada. Primeiro procura uma opção igual
    (ex.: 'sessao 615' = 'Sessão 0615'); se não houver, uma que comece pelo que foi digitado
    (ex.: 'Sessão 0615 - 10/11/2026'). Gera erro se não houver nenhuma ou houver mais de uma."""
    digitado = normalizar_pauta(pauta)
    normalizados = [normalizar_pauta(t) for t in textos]
    iguais = [i for i, n in enumerate(normalizados) if n == digitado]
    def combina(n: list[str]) -> bool:
        # O número da sessão (1º da opção) tem que ser o digitado; o resto do que foi digitado
        # tem que aparecer na opção, na mesma ordem ('764 E' combina com '0764/2026 E - ...').
        if not digitado or not n or n[0] != digitado[0]:
            return False
        resto = iter(n[1:])
        return all(p in resto for p in digitado[1:])
    comecam = [i for i, n in enumerate(normalizados) if combina(n)]
    na_tela = [t.strip() for t in textos if t.strip()]
    for achados in (iguais, comecam):
        if len(achados) == 1:
            return achados[0]
        if len(achados) > 1:
            raise RuntimeError(f"Pauta: '{pauta}' corresponde a mais de uma opção: "
                               f"{[textos[i].strip() for i in achados]}. Digite de forma mais completa.")
    raise RuntimeError(f"Pauta: '{pauta}' não está entre as opções: {na_tela[:30]}")


def preencher_pauta(page, pauta: str) -> str:
    """Escolhe na lista suspensa a opção que corresponde à pauta digitada (sem contar acentos,
    maiúsculas, espaços e zeros à esquerda: 'sessao 615' escolhe 'Sessão 0615') e confere.
    Devolve o texto escolhido. Se não conseguir, gera erro (nada é salvo)."""
    campo = page.locator(f"[id='{ID_PAUTA}_label']")
    try:
        campo.wait_for(state="visible", timeout=ESPERA_CURTA_MS)
    except Exception:
        raise RuntimeError("Pauta: o campo não apareceu na tela") from None
    campo.click()
    opcoes = page.locator(f"[id='{ID_PAUTA}_panel'] li")
    opcoes.first.wait_for(state="visible", timeout=ESPERA_CURTA_MS)
    textos = opcoes.all_inner_texts()
    indice = escolher_opcao_pauta(pauta, textos)
    opcao = textos[indice].strip()
    opcoes.nth(indice).click()
    page.wait_for_timeout(1500)
    escolhido = campo.inner_text()
    if so_palavras(escolhido) != so_palavras(opcao):
        raise RuntimeError(f"Pauta: escolhi '{opcao}', mas o campo ficou com '{escolhido.strip()}'")
    print(f"OK [pauta] Pauta preenchida: {escolhido.strip()} (digitado: '{pauta}')")
    return escolhido.strip()


def pautar_processo(page, numero: str, pauta: str) -> str:
    """Abre o Relatório e Voto do processo, preenche tudo e salva. Devolve o status."""
    fechar_dialogos(page)
    buscas = page.locator("[id='formBuscaRapida:buscaRapida']").filter(visible=True).count()
    if buscas > 1:
        # A tela acumulou cópias (mesmo id) de tanto abrir processos: começa de uma tela limpa.
        print(f"AVISO: a tela tem {buscas} campos de busca rápida; recarregando o sistema.")
        recarregar_sistema(page)
    elif not buscas:
        voltar_pagina_inicial(page)
    if not fazer_busca_processo(page, numero):
        return "Erro na busca"
    abrir_relatorio_voto(page)

    for parte, id_campo in CAMPOS_TEXTO.items():
        antigo = texto_do_editor(page, id_campo)
        if antigo and antigo.lower() != TEXTO_PROVISORIO:
            raise TextoExistente(f"{parte.capitalize()} já tem {len(antigo)} caracteres de texto")
    for parte, id_campo in CAMPOS_TEXTO.items():
        preencher_editor(page, id_campo, f"<p>{TEXTO_PROVISORIO}</p>", parte.capitalize())

    grupo_alcada = page.locator("xpath=//*[text()[contains(., 'Alçada')]]"
                                "/following::*[contains(@class,'ui-selectoneradio')][1]")
    formulario = page.locator(f"[id='{FORM_VOTO}']")
    marcacoes = [
        ("Matéria de alçada", lambda: marcar_opcao(grupo_alcada, OPCAO_ALCADA, "Matéria de alçada")),
        ("Decisão do relator", lambda: marcar_opcao(formulario, OPCAO_DECISAO, "Decisão do relator")),
        ("Tipo de decisão", lambda: marcar_opcao(formulario, OPCAO_TIPO_DECISAO, "Tipo de decisão")),
        ("Tipo de objeto", lambda: marcar_tipo_objeto(page)),
    ]
    # Como no pipeline: uma marcação que falha não impede o salvamento; o status diz o que conferir.
    falhas = []
    for nome, marcar in marcacoes:
        try:
            marcar()
        except Exception as e:
            print(f"AVISO [{numero}] {str(e)[:300]}")
            falhas.append(nome)

    pauta_escolhida = preencher_pauta(page, pauta)  # por último; se falhar, nada é salvo

    status = salvar_relatorio_voto(page)
    status += f" (pauta: {pauta_escolhida})"
    if falhas:
        status += f" - conferir: {', '.join(falhas)}"
    return status


def main():
    ARQ_LOG_PAUTA.parent.mkdir(parents=True, exist_ok=True)  # no 1º uso, a pasta ainda não existe
    log = Log(ARQ_LOG_PAUTA)
    with sistema_logado() as page:
        abrir_lista_relatar(page)

        pedidos = perguntar_processos()
        if not pedidos:
            print("Nenhum processo informado. Encerrando.")
            return
        pauta = perguntar_pauta()
        if not pauta:
            print("Nenhuma pauta informada. Encerrando.")
            return

        paginas = dict(coletar_todos_processos(page))
        if not paginas:
            print("Não consegui ler a lista 'Relatar'. Nada foi alterado.")
            return
        leitura_completa = lista_lida_ate_o_fim(page)
        if not leitura_completa:
            print("AVISO: a leitura parou antes da última página da lista.")

        encontrados = [n for n in pedidos if n in paginas]
        fora_da_lista = [n for n in pedidos if n not in paginas]
        if fora_da_lista:
            print(f"Fora da lista (não vou mexer): {', '.join(fora_da_lista)}")
        if not encontrados:
            print("Nenhum dos processos digitados está na lista. Nada foi alterado.")
            return
        if not confirmar(encontrados, pauta):
            print("Cancelado. Nada foi alterado.")
            return

        problemas = []
        for idx, numero in enumerate(encontrados, 1):
            print(f"\n--- Pautando {idx}/{len(encontrados)}: {numero} ---")
            log.adicionar(numero, paginas[numero])
            try:
                status = pautar_processo(page, numero, pauta)
            except TextoExistente as e:
                status = f"Conferir: {e}; não substituí nem salvei"
            except ErroAoSalvar as e:
                status = f"Erro ao salvar: {e}"
            except Exception as e:
                status = f"Erro antes de salvar (nada foi salvo): {str(e)[:200]}"
            print(f"[{numero}] {status}")
            log.atualizar(numero, status)
            if not status.startswith("Salvo (") or " - conferir" in status:
                problemas.append((numero, status))
            page.wait_for_timeout(2000)

        print("\n" + "=" * 60)
        print(f"Terminado. {len(encontrados) - len(problemas)} de {len(pedidos)} processo(s) "
              f"salvo(s) sem pendência na pauta '{pauta}'. Controle em: {ARQ_LOG_PAUTA}")
        if fora_da_lista:
            motivo = ("não encontrados nas páginas lidas (a leitura ficou incompleta)"
                      if not leitura_completa else "não estão na lista 'Relatar'")
            print(f"Não mexi, {motivo}:")
            for numero in fora_da_lista:
                print(f"  {numero}")
        if problemas:
            print("Para conferir:")
            for numero, status in problemas:
                print(f"  {numero}: {status}")


if __name__ == "__main__":
    main()
