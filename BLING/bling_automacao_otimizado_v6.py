"""
Bling Automation v6
- Selenium automation to click all "Download" links on Bling export page.
- Processes downloaded CSVs:
    * For rows with NCM == NCM_ALVO and Observações > 0 => Preço = Observações * MULTIPLICADOR
    * Observações parsed from formats like "5,5" or "5.5", ignores non-numeric
    * If Observações == 0 -> skip multiplication
    * Preço final formatted as BR (thousands '.' and decimal ',')
- Saves processed files into DOWNLOAD_DIR/processed and builds processed_all.csv
- GUI: Tkinter prompt to input multiplicador before running.
Requirements:
  pip install selenium webdriver-manager pandas
Notes:
  - Chrome must allow automatic downloads; prefs in code try to set it.
  - chromedriver is obtained via webdriver-manager at runtime.
  - To create EXE use pyinstaller (instructions below).
"""

import os
import time
import glob
import shutil
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import simpledialog, messagebox

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (
    NoSuchElementException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    TimeoutException,
    WebDriverException
)
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ===========================
# CONFIGURAÇÃO - Ajuste se desejar
# ===========================
EMAIL = "09839356100"
SENHA = "@%Gabriel99"

# local para salvar downloads
DOWNLOAD_DIR = str(Path.home() / "Downloads" / "bling_exports")
PROCESSED_DIR = os.path.join(DOWNLOAD_DIR, "processed")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(PROCESSED_DIR, exist_ok=True)

# NCM alvo e multiplicador padrão (será solicitado via Tkinter)
NCM_ALVO = "7113.19.00"
MULTIPLICADOR_PADRAO = 1400.0

# intervalos
CLICK_INTERVAL_SECONDS = 12  # 10-12s recomendados
DOWNLOAD_WAIT_TIMEOUT = 300  # tempo total máximo para downloads (segundos)
DOWNLOAD_POLL_INTERVAL = 1

# Chrome options
def build_chrome_options(download_dir):
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    # chrome_options.add_argument("--start-maximized")
    # chrome_options.add_argument("--headless=new")  # NÃO recomendo headless durante depuração
    return chrome_options

# ===========================
# UTILITÁRIOS
# ===========================
def safe_sleep(t):
    print(f"⏳ Aguardar {t}s...")
    time.sleep(t)

def remove_overlays_js(driver):
    js = """
    try{
      document.querySelectorAll('.modal, .modal-backdrop, .overlay, .ui-widget-overlay, .fancybox-overlay, .lightbox-overlay').forEach(el => el.remove());
      document.querySelectorAll('[style*="z-index"]').forEach(el => { try{ el.style.zIndex = '0'; } catch(e){} });
    } catch(e) {}
    """
    try:
        driver.execute_script(js)
    except Exception:
        pass

def try_close_popups(driver):
    xpaths = [
        "//button[contains(@class,'close')]", 
        "//button[contains(text(),'×') or contains(text(),'X') or contains(text(),'Fechar') or contains(text(),'fechar')]",
        "//a[contains(@class,'close') or contains(text(),'×') or contains(text(),'Fechar') or contains(text(),'fechar')]",
        "//div[contains(@class,'modal')]//button",
        "//div[contains(@class,'popup')]//button[contains(.,'Fechar') or contains(.,'fechar') or contains(.,'×')]"
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
                            try:
                                driver.execute_script("arguments[0].click();", e)
                            except Exception:
                                pass
                        safe_sleep(0.4)
                except Exception:
                    pass
        except Exception:
            pass

def wait_for_downloads_to_finish(download_dir, expected_new_count, timeout=DOWNLOAD_WAIT_TIMEOUT):
    start = time.time()
    while True:
        crdownloads = glob.glob(os.path.join(download_dir, "*.crdownload"))
        csv_files = glob.glob(os.path.join(download_dir, "*.csv"))
        # cond: no crdownload and minimum new csv count achieved
        if len(crdownloads) == 0 and len(csv_files) >= expected_new_count:
            return True, csv_files
        if time.time() - start > timeout:
            return False, csv_files
        time.sleep(DOWNLOAD_POLL_INTERVAL)

def detect_new_csvs(before_set, download_dir):
    all_csvs = set(glob.glob(os.path.join(download_dir, "*.csv")))
    new = sorted(list(all_csvs - before_set))
    return new

def parse_number_from_str(s):
    """Converte strings como '5,5', '1.234,56', '5.5' em float; retorna None se não for numérico."""
    if s is None: return None
    st = str(s).strip()
    if st == "": return None
    # remover aspas e espaços
    st = st.replace('"', '').replace("'", "").strip()
    # eliminar espaços no meio
    st = st.replace(" ", "")
    # se for algo como '-' ou '0' tratamos
    if st in ['-', 'null', 'None', 'nan', 'NaN']:
        return None
    # remover pontos de milhar e trocar vírgula por ponto
    # cuidado: '5.5' -> talvez seja 5.5 already, do generic approach:
    # se contém ',' assume formato BR -> remove pontos e troca ',' por '.'
    try:
        if ',' in st and '.' in st:
            # exemplo "1.234,56" -> remove ponto, replace comma
            st2 = st.replace('.', '').replace(',', '.')
        elif ',' in st:
            st2 = st.replace(',', '.')
        else:
            st2 = st
        return float(st2)
    except Exception:
        # último recurso: remove non digits except dot and minus
        import re
        filtered = re.sub(r'[^0-9\.\-]', '', st)
        try:
            return float(filtered) if filtered not in ['', '-', '.'] else None
        except:
            return None

def format_price_br(value):
    """Formata número float para string BR com ponto de milhar e vírgula decimal (2 casas)."""
    try:
        v = float(value)
    except Exception:
        return value
    # separador de milhar com '.' e decimal com ','
    # usar pandas for convenience:
    s = f"{v:,.2f}"  # usa ',' como milhar e '.' como decimal na locale C (ex: 1,234.56)
    # agora inverter: trocar ','-> temporary, '.'->',', temporary->'.'
    s = s.replace(',', 'X').replace('.', ',').replace('X', '.')
    return s

# ===========================
# PROCESSAMENTO DOS CSVs
# ===========================
def process_csv_file(path, multiplicador, ncm_alvo):
    base = os.path.basename(path)
    print(f"🔎 Processando: {base}")
    # tentar leituras com possíveis encodings e separadores
    df = None
    tried = []
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            # Bling costuma usar ';' como separador, com campos entre aspas
            df = pd.read_csv(path, sep=';', quotechar='"', encoding=enc, dtype=str)
            tried.append((enc, ';'))
            break
        except Exception:
            try:
                df = pd.read_csv(path, sep=',', quotechar='"', encoding=enc, dtype=str)
                tried.append((enc, ','))
                break
            except Exception:
                continue
    if df is None:
        print("❌ Falha ao ler CSV com tentativas de encodings; pulando:", base)
        return None, 0

    # limpar colunas
    df.columns = [c.strip() for c in df.columns]

    # checar colunas necessárias
    if "NCM" not in df.columns or "Observações" not in df.columns or "Preço" not in df.columns:
        print("⚠️ Colunas esperadas ausentes. Salvando cópia sem alteração:", base)
        out_path = os.path.join(PROCESSED_DIR, "processed_" + base)
        df.to_csv(out_path, index=False, encoding='utf-8-sig', sep=';')
        return out_path, 0

    updated_count = 0
    # máscara NCM alvo
    mask = df["NCM"].astype(str).str.strip() == ncm_alvo

    for idx in df.index[mask]:
        obs_raw = df.at[idx, "Observações"]
        val = parse_number_from_str(obs_raw)
        if val is None:
            # não numérico -> pular
            continue
        if val == 0:
            # requisito: se Observações == 0 não multiplicar
            continue
        new_price = val * multiplicador
        # salvar com formatação BR
        df.at[idx, "Preço"] = format_price_br(new_price)
        updated_count += 1

    out_path = os.path.join(PROCESSED_DIR, "processed_" + base)
    # salvar em ; e utf-8-sig
    df.to_csv(out_path, index=False, encoding='utf-8-sig', sep=';')
    print(f"✅ Salvo {os.path.basename(out_path)} — registros atualizados: {updated_count}")
    return out_path, updated_count

# ===========================
# RODA AUTOMATION
# ===========================
def run_automation(multiplicador):
    chrome_options = build_chrome_options(DOWNLOAD_DIR)
    driver = None
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    except WebDriverException as e:
        print("❌ Erro iniciando WebDriver:", e)
        return

    wait = WebDriverWait(driver, 25)

    try:
        print("🔐 Abrindo Bling (login)...")
        driver.get("https://www.bling.com.br/login")
        # aguardar página e campos
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        safe_sleep(1)

        # preencher username
        try:
            campo_username = wait.until(EC.presence_of_element_located((By.ID, "username")))
        except TimeoutException:
            campo_username = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='usuário' i], input[placeholder*='e-mail' i]")
        campo_username.clear()
        campo_username.send_keys(EMAIL)
        safe_sleep(0.3)

        # preencher senha
        try:
            campo_senha = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        except NoSuchElementException:
            campo_senha = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='senha' i], input[placeholder='Insira sua senha']")
        campo_senha.clear()
        campo_senha.send_keys(SENHA)
        safe_sleep(0.3)

        # clicar entrar
        botao_entrar = None
        try:
            botao_entrar = driver.find_element(By.ID, "login-submit")
        except NoSuchElementException:
            try:
                botao_entrar = driver.find_element(By.XPATH, "//button[contains(., 'Entrar') or contains(., 'Login') or contains(., 'ACESSAR') or contains(., 'Acessar')]")
            except NoSuchElementException:
                pass
        if not botao_entrar:
            raise Exception("Botão de login não encontrado automaticamente.")
        # clicar via JS para maior estabilidade
        driver.execute_script("arguments[0].click();", botao_entrar)

        print("⏳ Fazendo login, aguardando dashboard...")
        safe_sleep(4)
        try_close_popups(driver)
        remove_overlays_js(driver)
        safe_sleep(1)

        # abrir a página de exportação diretamente (você indicou que funciona)
        print("📦 Acessando página de exportação de produtos...")
        export_page_url = "https://www.bling.com.br/exportacao.produtos.php?criterio=opc-ultimos&pesquisa=&idTag=0&dataValidadeIni=&dataValidadeFim=&fornecedor=0&marca=&ncm=&dataAlteracaoIni=&dataAlteracaoFim=&codigoProdutoNoFabricante=&tipoProduto=T&idComponente=0&filtroTipoItem=&situacaoVinculoLoja=&idLojaVinculo=&filtroEstoque=&categoria=T"
        driver.get(export_page_url)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        safe_sleep(1)
        try_close_popups(driver)
        remove_overlays_js(driver)
        safe_sleep(1)

        # localizar links "Download" na página (ancoras com texto "Download")
        print("🔎 Procurando links 'Download' na página...")
        # esperar anchors aparecerem
        wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))
        safe_sleep(0.5)

        anchors = driver.find_elements(By.TAG_NAME, "a")
        download_links = []
        for a in anchors:
            try:
                txt = a.text.strip().lower()
                if "download" in txt:
                    download_links.append(a)
            except Exception:
                pass

        if not download_links:
            # fallback por selector conhecido
            anchors_alt = driver.find_elements(By.CSS_SELECTOR, "a.action-link")
            for a in anchors_alt:
                if "download" in a.text.strip().lower():
                    download_links.append(a)

        total_to_click = len(download_links)
        print(f"🔘 {total_to_click} links 'Download' encontrados.")

        if total_to_click == 0:
            raise Exception("Nenhum link de Download encontrado na página.")

        # registrar arquivos CSV existentes para detectar novos
        before_csvs = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
        # remover .crdownload antigos
        for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload")):
            try: os.remove(f)
            except: pass

        # clicar em todos os links via JS, com intervalo
        successful_clicks = 0
        for idx, link in enumerate(download_links, start=1):
            print(f"➡️ Iniciando download {idx}/{total_to_click} (JS click)...")
            try:
                # scroll and JS click
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
                safe_sleep(0.3)
                driver.execute_script("arguments[0].click();", link)
                successful_clicks += 1
            except Exception as e:
                print("⚠️ Falha no click JS:", e)
                # tentar fallback com click normal
                try:
                    link.click()
                    successful_clicks += 1
                except Exception:
                    print("❌ Não foi possível iniciar o download para o link", idx)
            # esperar o intervalo antes do próximo clique para evitar problemas de throttle
            print(f"    ⏳ Aguardando {CLICK_INTERVAL_SECONDS}s antes do próximo download...")
            time.sleep(CLICK_INTERVAL_SECONDS)

        print(f"🔁 Cliques finalizados. Downloads iniciados (tentativas): {successful_clicks}")

        # aguardar downloads finalizarem
        expected_new = successful_clicks
        print("⏳ Aguardando conclusão dos downloads (timeout: {}s)...".format(DOWNLOAD_WAIT_TIMEOUT))
        ok, all_csvs = wait_for_downloads_to_finish(DOWNLOAD_DIR, expected_new, timeout=DOWNLOAD_WAIT_TIMEOUT)
        if not ok:
            print("⚠️ Timeout esperando downloads. Verifique manualmente em:", DOWNLOAD_DIR)

        # identificar novos arquivos (baixados nesta execução)
        new_files = detect_new_csvs(before_csvs, DOWNLOAD_DIR)
        print(f"🔘 Encontrados {len(new_files)} novos arquivos CSV:")
        for f in new_files:
            print("   ", os.path.basename(f))

        # ================= TRATAMENTO
        if not new_files:
            print("⚠️ Nenhum novo CSV para processar. Encerrando execução.")
            return

        processed_paths = []
        total_updates = 0
        for csv in new_files:
            out_path, updated = process_csv_file(csv, multiplicador, NCM_ALVO)
            if out_path:
                processed_paths.append(out_path)
                total_updates += updated

        # gerar unificado se houver processed files
        if processed_paths:
            print("🔗 Unindo processed CSVs em único arquivo...")
            # ler todos e concatenar (mesmo separador e encoding)
            dfs = []
            for p in processed_paths:
                try:
                    dfp = pd.read_csv(p, sep=';', encoding='utf-8', dtype=str)
                except Exception:
                    dfp = pd.read_csv(p, sep=';', encoding='latin-1', dtype=str)
                dfs.append(dfp)
            if dfs:
                df_all = pd.concat(dfs, ignore_index=True, sort=False)
                out_all = os.path.join(PROCESSED_DIR, "processed_all.csv")
                df_all.to_csv(out_all, index=False, encoding='utf-8-sig', sep=';')
                print(f"✅ Arquivo unificado salvo em: {out_all}")

        print("🏁 Processo concluído. Downloads/processed em:", DOWNLOAD_DIR)

    except Exception as e:
        print("❌ Erro durante o processo:", str(e))
    finally:
        print("🔚 Fechando navegador em 5s...")
        safe_sleep(5)
        try:
            driver.quit()
        except Exception:
            pass

# ===========================
# UI Tkinter para multiplicador
# ===========================
def ask_multiplier_and_run():
    root = tk.Tk()
    root.withdraw()
    prompt = f"Informe o multiplicador (padrão {MULTIPLICADOR_PADRAO:.2f}). Use vírgula ou ponto (ex: 1.400, 1400, 5,5):"
    answer = simpledialog.askstring("Multiplicador", prompt)
    if answer is None:
        # usuário cancelou
        if messagebox.askyesno("Sair", "Deseja encerrar a execução?"):
            return
        else:
            multiplier = MULTIPLICADOR_PADRAO
    else:
        # normalizar string e parse
        ans = str(answer).strip().replace(".", "").replace(",", ".")  # '1.400' -> '1400', '5,5'->'5.5'
        try:
            multiplier = float(ans)
        except:
            messagebox.showwarning("Valor inválido", f"Valor informado inválido. Será usado multiplicador padrão {MULTIPLICADOR_PADRAO:.2f}")
            multiplier = MULTIPLICADOR_PADRAO

    # confirmar e executar
    if messagebox.askyesno("Confirmar", f"Executar automação com multiplicador = {multiplier:.2f}?"):
        run_automation(multiplier)

if __name__ == "__main__":
    ask_multiplier_and_run()
