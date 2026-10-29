from playwright.sync_api import sync_playwright
import os
import re
import json
import urllib.request
import time
import socket
import subprocess
import sys
import traceback
from openpyxl import Workbook
from datetime import datetime
from urllib.parse import urlparse


# ======================================================================
# CONFIGURAÇÃO
# ======================================================================
PASTA_LOG = r"C:\Users\Pedro\Downloads\processos"
UNIDADE = "GOV:CRPS.10ª JR"
URL_SISTEMA = "https://esisrec.inss.gov.br/esisrec/"

# True  = NÃO valida (só simula e registra). Teste assim primeiro!
# False = registra o voto decisório de verdade.
DRY_RUN = False

MIN_CARACTERES = 10

SEL = {
    # Paginador da lista de processos (AJUSTAR para a tela da sessão, se diferente)
    "paginador": "[id$='_paginator_top']",

    # Tela do processo: item "Voto Decisório" da lista de Ações
    "link_voto": "[id='formProcesso:votoDecisorio']",

    # Tela do voto: título "Relatório". O texto é lido do primeiro campo
    # (textarea / input / editor) ou texto que vier depois dele.
    "titulo_relatorio": "xpath=(//*[normalize-space(text())='Relatório'])[1]",

    # Textos que NÃO são o relatório (se aparecerem logo após o título,
    # significa que o relatório está vazio)
    "textos_nao_relatorio": ["Voto", "Concordo com o voto do(a) Relator(a)"],

    # Radio "Concordo" do voto do presidente (opção 1). Clica na caixinha visível;
    # o <input> real fica escondido dentro do ui-helper-hidden-accessible.
    "radio_concordo": "[id='formRelatorioVoto:presidente:opt1']",
    "radio_concordo_caixa": "[id='formRelatorioVoto:presidente:opt1'] .ui-radiobutton-box",

    # Botão de validar o voto (AJUSTAR com o F12)
    "btn_validar_voto": "[id='formRelatorioVoto:btnSalvar']",
    "btn_confirmar": "",   # AJUSTAR se houver diálogo de confirmação ("" = não tem)

    # Mensagens de erro do PrimeFaces exibidas após validar
    "msg_erro": ".ui-messages-error, .ui-message-error, .ui-growl-item-container.ui-state-error",
}

PADRAO_PROCESSO = re.compile(r'\d{5}\.\d{6}/\d{4}-\d{2}')
# ======================================================================


# ---------------------- coleta da lista ----------------------

def area_lista(page):
    """Tabela que contém o paginador; se não achar, a página inteira."""
    tabela = page.locator(SEL["paginador"]).first.locator(
        "xpath=ancestor::*[contains(@class,'ui-datatable')][1]"
    )
    try:
        if tabela.count() > 0:
            return tabela
    except Exception:
        pass
    return page.locator("body")


def extrair_numeros_processos_pagina(page):
    textos = area_lista(page).locator("a").all_inner_texts()
    return [t.strip() for t in textos if PADRAO_PROCESSO.fullmatch(t.strip())]


def obter_total_paginas(page):
    try:
        paginator = page.locator(f"{SEL['paginador']} .ui-paginator-current").first
        texto = paginator.inner_text(timeout=5000)
        match = re.search(r'de\s+(\d+)', texto)
        if match:
            total = int(match.group(1))
            print(f" Total de páginas detectado: {total}")
            return total
    except Exception as e:
        print(f" Paginador não encontrado (assumindo 1 página): {str(e).splitlines()[0]}")
    return 1


def navegar_para_pagina(page, numero_pagina, timeout=30):
    try:
        btn = page.locator(f"{SEL['paginador']} .ui-paginator-page").filter(
            has_text=re.compile(rf"^\s*{numero_pagina}\s*$")
        ).first
        if btn.count() == 0:
            return False
        antes = extrair_numeros_processos_pagina(page)
        btn.click()
        # O PrimeFaces troca a tabela via AJAX: espera o conteúdo mudar de fato
        limite = time.time() + timeout
        while time.time() < limite:
            agora = extrair_numeros_processos_pagina(page)
            if agora and agora != antes:
                page.wait_for_timeout(500)
                return True
            page.wait_for_timeout(500)
        print(f" A página {numero_pagina} não carregou em {timeout}s.")
    except Exception as e:
        print(f" Erro ao navegar para página {numero_pagina}: {e}")
    return False


def coletar_todos_processos(page):
    todos = []
    total_paginas = obter_total_paginas(page)
    print(f" Coletando processos de {total_paginas} página(s)...")
    for pagina in range(1, total_paginas + 1):
        if pagina > 1 and not navegar_para_pagina(page, pagina):
            print(f" Não foi possível ir à página {pagina}. Parando coleta.")
            break
        nums = extrair_numeros_processos_pagina(page)
        todos.extend(nums)
        print(f"Página {pagina}: {len(nums)} processos. Total parcial: {len(todos)}")
    # remove duplicados mantendo a ordem
    todos = list(dict.fromkeys(todos))
    print(f"\n Total coletado: {len(todos)} processos.")
    return todos


# ---------------------- navegação / voto ----------------------

def voltar_inicio(page):
    try:
        for sel in ("[id='formProcesso:btPaginaInicialFim']", "[id='formProcesso:btPaginaInicial']"):
            btn = page.locator(sel)
            if btn.count() > 0 and btn.first.is_visible():
                btn.first.click()
                page.wait_for_timeout(3000)
                return
    except Exception:
        pass
    page.goto(URL_SISTEMA)
    page.wait_for_timeout(3000)


def fazer_busca_processo(page, numero):
    try:
        campo = page.locator("#formBuscaRapida\\:buscaRapida")
        campo.wait_for(state="visible", timeout=30000)
        campo.fill(numero)
        page.locator("#formBuscaRapida\\:btBuscaRapida").click()
        page.wait_for_selector(SEL["link_voto"], state="visible", timeout=30000)
        # Garante que abriu o processo certo antes de votar
        if numero not in page.locator("body").inner_text():
            print(f" O processo exibido não é o {numero}")
            return False
        print(f" Carregou processo: {numero}")
        return True
    except Exception as e:
        print(f" Falha ao buscar processo {numero}: {e}")
        return False


JS_LER_APOS_TITULO = """
(titulo) => {
  const w = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
  w.currentNode = titulo;
  let n;
  while ((n = w.nextNode())) {
    if (titulo.contains(n)) continue;
    const tag = n.tagName;
    if (tag === 'SCRIPT' || tag === 'STYLE') continue;
    if (tag === 'TEXTAREA') return n.value;
    if (tag === 'INPUT' && ['text', ''].includes(n.type)) return n.value;
    if (n.isContentEditable) return n.innerText;
    if (tag === 'IFRAME') {
      try { return n.contentDocument.body.innerText; } catch (e) { return ''; }
    }
    const proprio = [...n.childNodes].filter(c => c.nodeType === 3)
                     .map(c => c.textContent).join('').trim();
    if (proprio) return n.innerText;
  }
  return '';
}
"""


def ler_relatorio(page):
    titulo = page.locator(SEL["titulo_relatorio"]).first
    titulo.wait_for(state="visible", timeout=30000)
    texto = (titulo.evaluate(JS_LER_APOS_TITULO) or "").strip()
    # Se o que veio depois do título é outro rótulo da tela, o relatório está vazio
    if any(t in texto for t in SEL["textos_nao_relatorio"]):
        return ""
    return texto


def marcar_concordo(page):
    caixa = page.locator(SEL["radio_concordo_caixa"]).first
    caixa.wait_for(state="visible", timeout=30000)
    caixa.scroll_into_view_if_needed()
    caixa.click()
    # O PrimeFaces põe ui-state-active na caixa quando a opção fica marcada
    limite = time.time() + 10
    while time.time() < limite:
        if "ui-state-active" in (caixa.get_attribute("class") or ""):
            return
        page.wait_for_timeout(300)
    raise RuntimeError("Não consegui marcar a opção 'Concordo' do voto decisório")


def concordar_e_validar(page):
    marcar_concordo(page)
    page.wait_for_timeout(500)
    btn = page.locator(SEL["btn_validar_voto"]).first
    btn.wait_for(state="visible", timeout=30000)
    btn.scroll_into_view_if_needed()
    btn.click()
    if SEL["btn_confirmar"]:
        conf = page.locator(SEL["btn_confirmar"]).first
        conf.wait_for(state="visible", timeout=15000)
        conf.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    erro = page.locator(SEL["msg_erro"])
    if erro.count() > 0 and erro.first.is_visible():
        raise RuntimeError(f"Sistema recusou o voto decisório: {erro.first.inner_text().strip()[:150]}")


def processar_voto(page, numero):
    """Retorna (status, detalhe)."""
    link = page.locator(SEL["link_voto"]).first
    link.wait_for(state="visible", timeout=30000)
    link.scroll_into_view_if_needed()
    link.click()
    # Espera a TELA DO VOTO DECISÓRIO aparecer (evita ler o "Relatório" da tela anterior)
    page.locator(SEL["radio_concordo"]).first.wait_for(state="visible", timeout=30000)
    page.wait_for_timeout(1000)

    relatorio = ler_relatorio(page)
    tam = len(relatorio)
    print(f"   Relatório lido ({tam} caracteres): {relatorio[:60]!r}")

    if tam >= MIN_CARACTERES:
        if DRY_RUN:
            print(" DRY_RUN: marcaria 'Concordo' no voto decisório e salvaria")
            return "Simulado", f"{tam} caracteres"
        concordar_e_validar(page)
        print(" Voto decisório (concordo) salvo")
        return "Validado", f"{tam} caracteres"

    print(f" Relatório com {tam} caracteres: só volta")
    return "Ignorado", f"{tam} caracteres"


PORTA_CDP = 9222
PERFIL_CHROME = r"C:\chrome_esisrec"
CAMINHOS_CHROME = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
]


def porta_aberta():
    try:
        with socket.create_connection(("127.0.0.1", PORTA_CDP), timeout=1):
            return True
    except OSError:
        return False


def garantir_chrome():
    """Abre o Chrome com depuração remota (sem abrir o site), se ainda não estiver aberto assim."""
    if porta_aberta():
        return
    caminho = next((c for c in CAMINHOS_CHROME if os.path.exists(c)), None)
    if not caminho:
        raise RuntimeError("Chrome não encontrado. Ajuste CAMINHOS_CHROME no topo do script.")
    print(" Abrindo o Chrome...")
    subprocess.Popen([
        caminho,
        f"--remote-debugging-port={PORTA_CDP}",
        f"--user-data-dir={PERFIL_CHROME}",
    ])
    for _ in range(40):
        if porta_aberta():
            return
        time.sleep(0.5)
    raise RuntimeError("O Chrome abriu, mas a porta 9222 não respondeu.")


def abas_chrome():
    """URLs das abas abertas, lidas pela porta de depuração SEM conectar o Playwright."""
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{PORTA_CDP}/json/list", timeout=3) as r:
            return [t.get("url", "") for t in json.load(r) if t.get("type") == "page"]
    except Exception:
        return []


def esperar_login():
    """Espera o login terminar sem tocar na página.

    A tela de login do GERID usa reCAPTCHA invisível: o botão "Entrar" só envia
    o formulário depois que o reCAPTCHA gera o token. Com o Playwright conectado
    durante o login, o reCAPTCHA pode não responder e o botão "não faz nada".

    Como agora é você quem digita o endereço, a aba passa pelo e-SISREC por um instante
    antes de ir para a tela de login. Por isso o script só considera o login feito quando
    a aba fica no e-SISREC em 3 checagens seguidas (cerca de 6 s).
    """
    print(f"No Chrome, entre no e-SISREC ({URL_SISTEMA}) e faça o login (CPF e senha).")
    print("O script só se conecta depois do login...")
    host = urlparse(URL_SISTEMA).netloc
    seguidas = 0
    while seguidas < 3:
        no_sistema = any(urlparse(u).netloc == host for u in abas_chrome())
        seguidas = seguidas + 1 if no_sistema else 0
        time.sleep(2)


# Marca no sessionStorage da aba quando VOCÊ clica em "Consultar". O sessionStorage sobrevive
# à troca de página no mesmo site, então funciona se a consulta for por AJAX ou recarregando.
JS_OUVIR_CONSULTAR = """
() => {
  if (window.__roboOuvindo) return;
  window.__roboOuvindo = true;
  document.addEventListener('click', (e) => {
    for (let n = e.target; n; n = n.parentElement) {
      const t = ((n.tagName === 'INPUT' ? n.value : n.innerText) || '').trim();
      if (/^consultar$/i.test(t)) { sessionStorage.setItem('roboConsultou', '1'); return; }
    }
  }, true);
}
"""


def ouvir_consultar(context, page):
    """Começa a observar o clique em "Consultar" (na página atual e nas próximas)."""
    context.add_init_script(f"({JS_OUVIR_CONSULTAR})()")
    page.evaluate("() => sessionStorage.removeItem('roboConsultou')")
    page.evaluate(JS_OUVIR_CONSULTAR)


def esperar_lista(page):
    """Espera você clicar em "Consultar" e a lista de processos terminar de carregar.
    A tela inicial também tem links de processos, por isso o clique é obrigatório."""
    def lista_pronta():
        if not page.evaluate("() => sessionStorage.getItem('roboConsultou') === '1'"):
            return False
        antes = extrair_numeros_processos_pagina(page)
        if not antes:
            return False
        page.wait_for_timeout(2000)  # a lista tem que ficar igual por 2 s (terminou de carregar)
        return extrair_numeros_processos_pagina(page) == antes

    esperar_ate(lista_pronta)


def esperar_ate(condicao, intervalo=1.5):
    """Repete a checagem até condicao() ser verdadeira, ignorando erros de navegação."""
    while True:
        try:
            if condicao():
                return
        except Exception:
            pass
        time.sleep(intervalo)


def salvar_log(wb, caminho):
    """Salva sem derrubar o lote se o arquivo estiver aberto no Excel."""
    try:
        wb.save(caminho)
    except PermissionError:
        print(f" AVISO: Não consegui salvar {caminho} (está aberto no Excel?). "
              "Tento de novo no próximo processo.")


# ---------------------- principal ----------------------

def main():
    # Alguns terminais do Windows não mostram certos caracteres: troca por '?' em vez de parar.
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
    os.makedirs(PASTA_LOG, exist_ok=True)
    caminho_log = os.path.join(
        PASTA_LOG, f"voto_decisorio_log_{datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx"
    )
    wb = Workbook()
    ws = wb.active
    ws.title = "Voto Decisorio"
    ws.append(["Processo", "Status", "Detalhe", "Data/Hora"])

    def log(numero, status, detalhe=""):
        ws.append([numero, status, detalhe, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        salvar_log(wb, caminho_log)

    garantir_chrome()
    esperar_login()

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{PORTA_CDP}")
        context = browser.contexts[0]
        page = next((pg for pg in context.pages if "esisrec.inss.gov.br" in pg.url), None)
        if page is None:
            page = context.pages[0] if context.pages else context.new_page()

        try:
            ouvir_consultar(context, page)
            print("Escolha o perfil no navegador, se for pedido. Aguardando...")

            def logado():
                if page.locator("select#domains").is_visible():
                    try:
                        page.select_option("select#domains", label=UNIDADE)
                    except Exception as e:
                        print(f" Não consegui selecionar a unidade '{UNIDADE}': {e}")
                    page.wait_for_timeout(2000)
                    return False
                return page.locator("#formBuscaRapida\\:buscaRapida").is_visible()

            esperar_ate(logado)
            print(" Login detectado.")

            print("Selecione a sessão e clique em CONSULTAR. "
                  "O script começa sozinho quando a lista aparecer...")
            esperar_lista(page)
            print(" Lista detectada.")

            todos = coletar_todos_processos(page)
            if not todos:
                print(" Nenhum processo encontrado. Encerrando.")
                return
            print(f" Primeiros processos: {', '.join(todos[:3])}")
            if not DRY_RUN:
                print(" ATENÇÃO: DRY_RUN = False, os votos decisórios serão registrados de verdade. "
                      "Começando em 10 s (Ctrl+C para cancelar)...")
                time.sleep(10)

            for idx, numero in enumerate(todos):
                print(f"\n--- {idx + 1}/{len(todos)}: {numero} ---")
                try:
                    if idx > 0:
                        voltar_inicio(page)
                    if not fazer_busca_processo(page, numero):
                        log(numero, "Erro na busca")
                        continue
                    status, detalhe = processar_voto(page, numero)
                    log(numero, status, detalhe)
                except Exception as e:
                    print(f" Erro em {numero}: {e}")
                    log(numero, "Erro", str(e)[:200])
                page.wait_for_timeout(1500)

            print("\n Fim. Log em:", caminho_log)

        except BaseException:
            traceback.print_exc()
        finally:
            input("\nPressione ENTER para encerrar o script (o Chrome continua aberto)...")


if __name__ == "__main__":
    main()
