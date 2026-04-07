"""
Bling - Automação de exportação + tratamento de CSVs (versão otimizada)

Funcionalidades:
- Login no Bling (Selenium + Chrome)
- Abre tela de exportação de produtos e baixa os 7 'Download'
- Intervalo entre downloads: 12s (configurável)
- Detecta arquivos CSV novos na pasta de downloads
- Faz tratamento: quando NCM == "7113.19.00" e Observações != 0 (numérico),
  Preço = Observações * 1400. Formata preço com vírgula.
- Gera arquivo processed_<orig>.csv com delimitador ';' e encoding latin-1 (Windows)
"""

import os
import time
import glob
from pathlib import Path
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    NoSuchElementException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    TimeoutException
)
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ================== CONFIGURAÇÃO ==================
EMAIL = "09839356100"         # substitua
SENHA = "@%Gabriel99"         # substitua
DOWNLOAD_DIR = str(Path.home() / "Downloads" / "bling_exports")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# Timeout / delays (ajuste conforme sua máquina/ internet)
DOWNLOAD_WAIT_TIMEOUT = 240      # segundos para aguardar todos downloads concluírem
DOWNLOAD_POLL_INTERVAL = 1.0
DELAY_BETWEEN_DOWNLOADS = 10.0   # **requisitado**: 10-12s entre cliques

# PROCESSAMENTO
MULTIPLICADOR = 1400.0
NCM_ALVO = "7113.19.00"

# Selenium retry / waits
MAX_CLICK_RETRIES = 3
PAGE_LOAD_TIMEOUT = 20

# Chrome preferences para permitir downloads automáticos e impedir prompts
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "profile.default_content_setting_values.notifications": 2,
    # se quiser headless (não recomendado para debugging)
    # "headless": True,
}
chrome_options.add_experimental_option("prefs", prefs)
# chrome_options.add_argument("--start-maximized")

# inicia driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, PAGE_LOAD_TIMEOUT)

def safe_sleep(t):
    time.sleep(t)

def remove_overlays_js():
    js = """
    try{
      document.querySelectorAll('.modal, .modal-backdrop, .overlay, .ui-widget-overlay, .fancybox-overlay, .lightbox-overlay').forEach(e => e.remove());
      document.querySelectorAll('[style*="z-index"]').forEach(el => { try{ el.style.zIndex = '0'; } catch(e){} });
    }catch(e){}
    """
    try:
        driver.execute_script(js)
    except Exception:
        pass

def try_close_popups():
    xpaths = [
        "//button[contains(@class,'close')]",
        "//button[contains(text(),'×') or contains(text(),'X')]",
        "//div[contains(@class,'modal')]//a[contains(@class,'close')]",
        "//button[contains(@aria-label,'close') or contains(@aria-label,'fechar')]",
        "//span[contains(@class,'close') or contains(text(),'×')]",
        "//a[contains(@class,'close') or contains(text(),'×')]",
        "//i[contains(@class,'close')]",
    ]
    for xp in xpaths:
        try:
            elems = driver.find_elements(By.XPATH, xp)
            for e in elems:
                try:
                    if e.is_displayed():
                        try:
                            e.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", e)
                        safe_sleep(0.6)
                except Exception:
                    pass
        except Exception:
            pass

def safe_click(element):
    for _ in range(MAX_CLICK_RETRIES):
        try:
            element.click()
            return True
        except (ElementClickInterceptedException, ElementNotInteractableException):
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", element)
                driver.execute_script("arguments[0].click();", element)
                return True
            except Exception:
                remove_overlays_js()
                safe_sleep(0.8)
    return False

def wait_for_all_downloads_to_finish(expected_count, timeout=DOWNLOAD_WAIT_TIMEOUT):
    start = time.time()
    while True:
        cr = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
        csvs = glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv"))
        no_cr = len(cr) == 0
        if no_cr and len(csvs) >= expected_count:
            return True
        if time.time() - start > timeout:
            print("⚠️ Timeout esperando downloads. crdownloads:", cr, "csv:", csvs)
            return False
        time.sleep(DOWNLOAD_POLL_INTERVAL)

# ---------- fluxo ----------
try:
    print("🔐 Abrindo Bling (login)...")
    driver.get("https://www.bling.com.br/login")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    safe_sleep(1.0)

    # USERNAME
    try:
        campo_username = driver.find_element(By.ID, "username")
    except NoSuchElementException:
        # alternativa por placeholder
        campo_username = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='usuário' i], input[placeholder*='e-mail' i], input[name='username']")
    campo_username.clear()
    campo_username.send_keys(EMAIL)
    safe_sleep(0.4)

    # SENHA
    try:
        campo_senha = driver.find_element(By.CSS_SELECTOR, "input[type='password'], input[placeholder*='senha' i]")
    except NoSuchElementException:
        campo_senha = driver.find_element(By.CSS_SELECTOR, "input[type='text'][placeholder*='senha' i]")
    campo_senha.clear()
    campo_senha.send_keys(SENHA)
    safe_sleep(0.4)

    # BOTÃO LOGIN (tentativas)
    botao = None
    try:
        botao = driver.find_element(By.ID, "login-submit")
    except Exception:
        try:
            botao = driver.find_element(By.XPATH, "//button[contains(., 'Entrar') or contains(., 'Login') or contains(., 'entrar')]")
        except Exception:
            botao = None

    if not botao:
        raise Exception("Botão de login não encontrado. Ajuste seletores.")

    safe_click(botao)
    print("⏳ Fazendo login, aguardando dashboard...")
    # aguardar mudança de URL / dashboard carregado
    try:
        wait.until(EC.url_contains("bling.com.br"))
    except TimeoutException:
        pass
    safe_sleep(4)
    try_close_popups()
    remove_overlays_js()
    safe_sleep(1.0)

    # ================= ir para exportação =================
    print("📦 Acessando página de exportação de produtos...")
    # URL direta (ajustada) — a que você confirmou funcionar:
    driver.get("https://www.bling.com.br/exportacao.produtos.php?criterio=opc-ultimos&pesquisa=&idTag=0&dataValidadeIni=&dataValidadeFim=&fornecedor=0&marca=&ncm=&dataAlteracaoIni=&dataAlteracaoFim=&codigoProdutoNoFabricante=&tipoProduto=T&idComponente=0&filtroTipoItem=&situacaoVinculoLoja=&idLojaVinculo=&filtroEstoque=&categoria=T")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    safe_sleep(1.0)
    try_close_popups()
    remove_overlays_js()
    safe_sleep(1.0)

    # ========== localizar elemento 'Exportar dados para planilha' e ativar ==========
    print("🔎 Procurando 'Exportar dados para planilha' ...")
    try:
        export_el = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'li[data-function="exportar"]')))
    except TimeoutException:
        # busca alternativa por texto
        els = driver.find_elements(By.XPATH, "//li[contains(., 'Exportar dados para planilha') or contains(., 'Exportar dados') or contains(., 'Exportar')]")
        export_el = els[0] if els else None

    if not export_el:
        raise Exception("Elemento 'Exportar dados para planilha' não encontrado.")

    safe_click(export_el)
    safe_sleep(2.0)
    try_close_popups()
    remove_overlays_js()
    safe_sleep(1.2)

    # ========== estamos na página de exportação: achar links 'Download' ==========
    print("📄 Página de exportação aberta — procurando links 'Download' ...")
    # aguarda que os links existam
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.action-link")))
    anchors = driver.find_elements(By.CSS_SELECTOR, "a.action-link")
    download_links = [a for a in anchors if "download" in a.text.strip().lower()]

    # fallback
    if not download_links:
        download_links = driver.find_elements(By.XPATH, "//a[contains(translate(., 'DOWNLOAD', 'download'), 'download')]")

    print(f"🔘 {len(download_links)} links 'Download' encontrados.")

    # limpar csv antigos
    for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")):
        try:
            os.remove(f)
        except Exception:
            pass

    # clicar nos downloads com delays
    started = 0
    for idx, link in enumerate(download_links, start=1):
        print(f"➡️ Iniciando download {idx}/{len(download_links)} ...")
        # estratégia robusta:
        # 1) focar (JS focus), 2) enviar ENTER via active element (simula teclado)
        try:
            driver.execute_script("arguments[0].scrollIntoView(true);", link)
        except Exception:
            pass
        safe_sleep(0.3)
        try:
            # foco e Enter (teclado)
            driver.execute_script("arguments[0].focus();", link)
            active = driver.switch_to.active_element
            try:
                active.send_keys(Keys.ENTER)
            except Exception:
                # fallback JS click
                driver.execute_script("arguments[0].click();", link)
            started += 1
        except Exception:
            # último recurso: click direto
            try:
                link.click()
                started += 1
            except Exception as e:
                print(f"❌ Falha ao iniciar download {idx}: {e}")

        # aguarda o início do download (detecta novo csv no diretório)
        # diferenciar: iremos aguardar que pelo menos um novo arquivo apareça com timestamp recente
        # implementação simples: pause fixa e depois próximo
        print(f"    ⏳ Aguardando {DELAY_BETWEEN_DOWNLOADS:.0f}s antes do próximo download...")
        safe_sleep(DELAY_BETWEEN_DOWNLOADS)

    print(f"🔁 Cliques finalizados. Downloads iniciados (detectados attempts): {started}")

    # aguardar todos os downloads terminarem
    print("⏳ Aguardando término de quaisquer .crdownload remanescentes...")
    # esperamos pelo menos 'started' arquivos .csv
    downloads_ok = wait_for_all_downloads_to_finish(expected_count=started, timeout=DOWNLOAD_WAIT_TIMEOUT)
    csv_files = sorted(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
    print(f"🔘 Encontrados {len(csv_files)} novos arquivos CSV (detectados).")

    # ========== TRATAMENTO AUTOMÁTICO ==========
    print("🛠️ Iniciando tratamento automático das planilhas baixadas...")
    if not csv_files:
        print("⚠️ Nenhum CSV detectado para processar.")
    else:
        for csv_path in csv_files:
            try:
                print(f"🔎 Processando: {os.path.basename(csv_path)}")
                # tentativas de leitura: geralmente Bling usa ';' e encoding latin-1
                df = None
                read_success = False
                encs = ["latin-1", "utf-8", "cp1252"]
                for enc in encs:
                    try:
                        df = pd.read_csv(csv_path, sep=';', encoding=enc, dtype=str, low_memory=False)
                        read_success = True
                        used_enc = enc
                        break
                    except Exception:
                        continue
                if not read_success:
                    # tentativa fallback sem separador
                    df = pd.read_csv(csv_path, dtype=str, low_memory=False, encoding="latin-1", sep=None, engine="python")

                # limpar BOM no nome da primeira coluna se existir
                cols = [c.strip() for c in df.columns]
                # remover aspas duplas soltas nos nomes: por vezes Bling traz "ID";"Código"...
                cols = [c.replace('\ufeff','').replace('"','').strip() for c in cols]
                df.columns = cols

                # checar colunas essenciais
                if not all(x in df.columns for x in ["NCM", "Observações", "Preço"]):
                    print("    ⚠️ Colunas esperadas ausentes. Salvando cópia sem alteração.")
                    out_path = os.path.join(DOWNLOAD_DIR, "processed_" + os.path.basename(csv_path))
                    df.to_csv(out_path, index=False, sep=';', encoding='latin-1')
                    continue

                # função para converter Observações para número (float)
                def to_number(x):
                    if pd.isna(x):
                        return None
                    s = str(x).strip()
                    # remover espaços, pontos de milhar e trocar vírgula por ponto
                    s = s.replace(" ", "").replace(".", "").replace(",", ".")
                    try:
                        return float(s)
                    except:
                        return None

                updated = 0
                # iterar linhas onde NCM == NCM_ALVO
                mask = df["NCM"].astype(str).str.strip() == NCM_ALVO
                for idx in df.index[mask]:
                    obs_raw = df.at[idx, "Observações"]
                    num = to_number(obs_raw)
                    if num is None:
                        # se não for número, pular
                        continue
                    # quando Observações == 0 -> não multiplicar
                    if num == 0:
                        continue
                    new_price = num * MULTIPLICADOR
                    # formatar com vírgula e 2 casas (R$ BR): 1234.5 -> "1.234,50" ou "1234,50" (usar sem milhar)
                    # O usuário pediu vírgula como decimal; não exigiu milhares. Então:
                    price_str = f"{new_price:,.2f}"  # usa vírgula de milhar por padrão en_US? -> depende do locale
                    # Normalizar: garantir ponto decimal para substituir por vírgula e remoção de milhares (usar replace)
                    # Vamos formatar sem milhares:
                    price_str = f"{new_price:.2f}".replace(".", ",")
                    df.at[idx, "Preço"] = price_str
                    updated += 1

                out_path = os.path.join(DOWNLOAD_DIR, "processed_" + os.path.basename(csv_path))
                df.to_csv(out_path, index=False, sep=';', encoding='latin-1')
                print(f"    ✅ Salvo {os.path.basename(out_path)} — registros atualizados: {updated}")
            except Exception as e:
                print(f"    ❌ Erro ao processar {os.path.basename(csv_path)} : {e}")

    print("🏁 Todos os processos finalizados. Verifique a pasta:", DOWNLOAD_DIR)

except Exception as e:
    print("❌ Erro durante o processo:\n", e)

finally:
    print("🔚 Finalizando driver em 6s...")
    safe_sleep(6)
    try:
        driver.quit()
    except Exception:
        pass
