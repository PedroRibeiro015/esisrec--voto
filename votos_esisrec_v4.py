from playwright.sync_api import sync_playwright
import os
import re
import time
import traceback
from openpyxl import Workbook
from datetime import datetime


# ======================================================================
# CONFIGURAÇÃO
# ======================================================================
PASTA_LOG = r"C:\Users\Pedro\Downloads\processos"
UNIDADE = "GOV:CRPS.10ª JR"

# True  = NÃO valida (só simula e registra). Teste assim primeiro!
# False = registra o voto de verdade.
DRY_RUN = True

MIN_CARACTERES = 10

SEL = {
    # Paginador da lista de processos (AJUSTAR para a tela da sessão, se diferente)
    "paginador": "[id$='_paginator_top']",

    # Tela do processo: item "Voto" da lista de Ações
    "link_voto": "[id='formProcesso:cadastrarRelatorioVoto']",

    # Tela do voto: texto logo abaixo do título "Relatório"
    "campo_relatorio": "xpath=(//*[normalize-space(text())='Relatório'])[1]"
                       "/following::*[normalize-space(.)!=''][1]",   # AJUSTAR se ler texto errado

    "texto_concordo": "Concordo com o voto do(a) Relator(a)",

    # Botão de validar o voto (AJUSTAR com o F12)
    "btn_validar_voto": "[id$='btValidarVoto']",
    "btn_confirmar": "",   # AJUSTAR se houver diálogo de confirmação ("" = não tem)
}

PADRAO_PROCESSO = re.compile(r'\d{5}\.\d{6}/\d{4}-\d{2}')
# ======================================================================


# ---------------------- coleta da lista ----------------------

def extrair_numeros_processos_pagina(page):
    numeros = []
    for link in page.locator("a").all():
        try:
            texto = link.inner_text().strip()
        except Exception:
            continue
        if PADRAO_PROCESSO.fullmatch(texto):
            numeros.append(texto)
    return numeros


def obter_total_paginas(page):
    try:
        paginator = page.locator(f"{SEL['paginador']} .ui-paginator-current").first
        texto = paginator.inner_text()
        match = re.search(r'de\s+(\d+)', texto)
        if match:
            total = int(match.group(1))
            print(f"📄 Total de páginas detectado: {total}")
            return total
    except Exception as e:
        print(f"ℹ️ Paginador não encontrado (assumindo 1 página): {e}")
    return 1


def navegar_para_pagina(page, numero_pagina):
    try:
        btn = page.locator(
            f"{SEL['paginador']} .ui-paginator-page:has-text('{numero_pagina}')"
        ).first
        if btn.count() > 0:
            btn.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(3000)
            return True
    except Exception as e:
        print(f"⚠️ Erro ao navegar para página {numero_pagina}: {e}")
    return False


def coletar_todos_processos(page):
    todos = []
    total_paginas = obter_total_paginas(page)
    print(f"🔍 Coletando processos de {total_paginas} página(s)...")
    for pagina in range(1, total_paginas + 1):
        if pagina > 1 and not navegar_para_pagina(page, pagina):
            print(f"⚠️ Não foi possível ir à página {pagina}. Parando coleta.")
            break
        nums = extrair_numeros_processos_pagina(page)
        todos.extend(nums)
        print(f"Página {pagina}: {len(nums)} processos. Total parcial: {len(todos)}")
        page.wait_for_timeout(1000)
    # remove duplicados mantendo a ordem
    todos = list(dict.fromkeys(todos))
    print(f"\n📊 Total coletado: {len(todos)} processos.")
    return todos


# ---------------------- navegação / voto ----------------------

def voltar_inicio(page):
    try:
        btn = page.locator("[id='formProcesso:btPaginaInicialFim']")
        if btn.count() == 0:
            btn = page.locator("[id='formProcesso:btPaginaInicial']")
        btn.wait_for(state="visible", timeout=10000)
        btn.click()
        page.wait_for_timeout(3000)
    except Exception:
        page.goto("https://esisrec.inss.gov.br/esisrec/")
        page.wait_for_timeout(3000)


def fazer_busca_processo(page, numero):
    try:
        campo = page.locator("#formBuscaRapida\\:buscaRapida")
        campo.wait_for(state="visible", timeout=30000)
        campo.fill(numero)
        page.locator("#formBuscaRapida\\:btBuscaRapida").click()
        page.wait_for_selector(SEL["link_voto"], state="visible", timeout=30000)
        print(f"✅ Carregou processo: {numero}")
        return True
    except Exception as e:
        print(f"❌ Falha ao buscar processo {numero}: {e}")
        return False


def ler_relatorio(page):
    campo = page.locator(SEL["campo_relatorio"]).first
    campo.wait_for(state="visible", timeout=30000)
    try:
        texto = campo.input_value()
    except Exception:
        texto = campo.inner_text()
    return (texto or "").strip()


def concordar_e_validar(page):
    page.get_by_text(SEL["texto_concordo"], exact=False).first.click()
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


def processar_voto(page, numero):
    """Retorna (status, detalhe)."""
    link = page.locator(SEL["link_voto"]).first
    link.wait_for(state="visible", timeout=30000)
    link.scroll_into_view_if_needed()
    link.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)

    relatorio = ler_relatorio(page)
    tam = len(relatorio)
    print(f"   Relatório lido ({tam} caracteres): {relatorio[:60]!r}")

    if tam >= MIN_CARACTERES:
        if DRY_RUN:
            print("🧪 DRY_RUN: concordaria com o relator e validaria")
            return "Simulado", f"{tam} caracteres"
        concordar_e_validar(page)
        print("✅ Voto (concordo) validado")
        return "Validado", f"{tam} caracteres"

    print(f"⚪ Relatório com {tam} caracteres → só volta")
    return "Ignorado", f"{tam} caracteres"


def esperar_ate(condicao, intervalo=1.5):
    """Repete a checagem até condicao() ser verdadeira, ignorando erros de navegação."""
    while True:
        try:
            if condicao():
                return
        except Exception:
            pass
        time.sleep(intervalo)


# ---------------------- principal ----------------------

def main():
    os.makedirs(PASTA_LOG, exist_ok=True)
    caminho_log = os.path.join(PASTA_LOG, "votos_log.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Votos"
    ws.append(["Processo", "Status", "Detalhe", "Data/Hora"])

    def log(numero, status, detalhe=""):
        ws.append([numero, status, detalhe, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        wb.save(caminho_log)

    with sync_playwright() as p:
        # Chrome aberto manualmente com --remote-debugging-port=9222
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        context = browser.contexts[0]
        page = context.pages[0] if context.pages else context.new_page()

        try:
            if "esisrec.inss.gov.br" not in page.url:
                page.goto("https://esisrec.inss.gov.br/esisrec/")

            print("Faça o login e escolha o perfil no navegador. Aguardando...")

            def logado():
                if page.locator("select#domains").is_visible():
                    page.select_option("select#domains", label=UNIDADE)
                    page.wait_for_timeout(2000)
                    return False
                return page.locator("#formBuscaRapida\\:buscaRapida").is_visible()

            esperar_ate(logado)
            print("✅ Login detectado.")

            print("Selecione a sessão e clique em CONSULTAR. Aguardando a lista de processos...")
            esperar_ate(lambda: len(extrair_numeros_processos_pagina(page)) > 0)
            page.wait_for_timeout(2000)
            print("✅ Lista detectada.")

            todos = coletar_todos_processos(page)
            if not todos:
                print("⚠️ Nenhum processo encontrado. Encerrando.")
                return

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
                    print(f"❌ Erro em {numero}: {e}")
                    log(numero, "Erro", str(e)[:200])
                page.wait_for_timeout(1500)

            print("\n🎉 Fim. Log em:", caminho_log)

        except BaseException:
            traceback.print_exc()
        finally:
            input("\nPressione ENTER para encerrar o script (o Chrome continua aberto)...")


if __name__ == "__main__":
    main()
