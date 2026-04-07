"""
bling_automacao_otimizado_v3.py
- Selenium automation to download all "Download" CSVs from Bling export page using
  TAB navigation after the first click (12s wait between clicks).
- Processes each CSV: if NCM == "7113.19.00" and Observações != 0, Preço = Observações * 1400.0
  The output Preço will use comma as decimal separator (e.g. 7700,00).
"""

import os
import time
import glob
import shutil
from pathlib import Path
import pandas as pd
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    NoSuchElementException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    TimeoutException
)
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ========== CONFIG ==========
EMAIL = "09839356100"     
SENHA = "@%Gabriel99"                

# pasta onde os arquivos serão baixados
DOWNLOAD_DIR = str(Path.home() / "Downloads" / "bling_exports")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# tempos / ajustes
DOWNLOAD_WAIT_TIMEOUT = 240          # tempo máximo para aguardar que downloads terminem
DOWNLOAD_POLL_INTERVAL = 1          # polling interval
BETWEEN_DOWNLOADS_SEC = 12          # tempo entre cada download (você pediu 10-12s; deixei 12s)
LOGIN_WAIT = 6                      # tempo de espera pós-login
CLICK_RETRY = 3

# regra de negócio
MULTIPLICADOR = 1400.0
NCM_ALVO = "7113.19.00"

# Chrome options: permite downloads automáticos
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "profile.default_content_setting_values.notifications": 2,
}
chrome_options.add_experimental_option("prefs", prefs)
# Recomendo manter janela visível durante testes
# chrome_options.add_argument("--start-maximized")

# inicializa driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 20)
actions = ActionChains(driver)

def safe_sleep(t):
    time.sleep(t)

def remove_overlays_js():
    js = """
    try {
      document.querySelectorAll('.modal, .modal-backdrop, .overlay, .ui-widget-overlay, .fancybox-overlay, .lightbox-overlay').forEach(e=>e.remove());
      document.querySelectorAll('[style*=\"z-index\"]').forEach(el=>{
         try{ el.style.zIndex='0'; }catch(e){}
      });
    } catch(e){}
    """
    try:
        driver.execute_script(js)
    except Exception:
        pass

def try_close_popups():
    # sequencia de seletores para botões fechar
    xpaths = [
        "//button[contains(@class,'close')]",
        "//button[contains(text(),'×') or contains(text(),'X')]",
        "//a[contains(@class,'close') or contains(text(),'×') or contains(text(),'Fechar') or contains(text(),'fechar')]",
        "//div[contains(@class,'modal')]//button[contains(., 'Fechar') or contains(., 'OK')]",
        "//button[@aria-label='Close' or @aria-label='Fechar']",
    ]
    for xp in xpaths:
        try:
            elems = driver.find_elements(By.XPATH, xp)
            for e in elems:
                try:
                    if e.is_displayed():
                        try:
                            e.click()
                        except:
                            driver.execute_script("arguments[0].click();", e)
                        safe_sleep(0.6)
                except Exception:
                    pass
        except Exception:
            pass
    remove_overlays_js()

def list_csvs(folder):
    return sorted(glob.glob(os.path.join(folder, "*.csv")))

def wait_for_new_file(before_files, timeout=60):
    start = time.time()
    while time.time() - start < timeout:
        current = list_csvs(DOWNLOAD_DIR)
        new = [f for f in current if f not in before_files]
        # ignore processed_ files
        new = [f for f in new if "processed_" not in os.path.basename(f)]
        if new:
            # prefer the most recent
            newest = sorted(new, key=os.path.getctime)[-1]
            # ensure .crdownload not present for that file
            if not os.path.exists(newest + ".crdownload"):
                return newest
        time.sleep(1)
    return None

def wait_for_all_crdownloads_to_finish(timeout=DOWNLOAD_WAIT_TIMEOUT):
    start = time.time()
    while True:
        crs = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
        if len(crs) == 0:
            return True
        if time.time() - start > timeout:
            return False
        time.sleep(DOWNLOAD_POLL_INTERVAL)

def safe_click(elem, scroll=True):
    for _ in range(CLICK_RETRY):
        try:
            if scroll:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
            elem.click()
            return True
        except (ElementClickInterceptedException, ElementNotInteractableException):
            try:
                driver.execute_script("arguments[0].click();", elem)
                return True
            except Exception:
                remove_overlays_js()
                time.sleep(0.5)
    return False

def clean_dataframe_and_fix_columns(df):
    # strip spaces in columns and remove BOM if present
    new_cols = []
    for c in df.columns:
        c2 = str(c).strip()
        # remove BOM characters and stray quotes
        c2 = c2.lstrip('\ufeff').strip()
        c2 = c2.strip('"').strip("'")
        new_cols.append(c2)
    df.columns = new_cols

    # Remove surrounding quotes in string fields and trim spaces
    def clean_cell(x):
        if pd.isna(x):
            return x
        s = str(x).strip()
        # remove leading/trailing quotes
        if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
            s = s[1:-1]
        return s.strip()

    # apply cleaning for object columns
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].apply(clean_cell)

    return df

def parse_number_from_brazil(s):
    """Converte strings tipo '5,5' ou '5.5' ou '1.234,56' para float.
       Retorna None se não for conversível ou for vazio."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    s = str(s).strip()
    if s == "":
        return None
    # remover espaços e apóstrofos
    s = s.replace(" ", "").replace("'", "")
    # se já é um número puro
    if re.fullmatch(r'^\d+(\.\d+)?$', s):
        return float(s)
    # trocar pontos de milhar e vírgula decimal -> padrão en
    # Ex: '1.234,56' => '1234.56'
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace(',', '.')
    try:
        return float(s)
    except:
        return None

def process_csv_file(path):
    fname = os.path.basename(path)
    try:
        # tentativa 1: sep=';' encoding latin-1 com engine python
        try:
            df = pd.read_csv(path, sep=';', dtype=str, encoding='latin-1', engine='python')
        except Exception:
            df = pd.read_csv(path, sep=';', dtype=str, encoding='utf-8', engine='python')

        df = clean_dataframe_and_fix_columns(df)

        # check required columns
        expected = {"NCM", "Observações", "Preço"}
        if not expected.issubset(set(df.columns)):
            # tenta detectar variações (Observacoes sem acento, Preco sem acento)
            cols_map = {}
            for c in df.columns:
                low = c.lower().replace("ç", "c").replace("ã", "a").replace("ó", "o").replace("õ","o").replace("é","e").replace("í","i")
                if "ncm" == low:
                    cols_map["NCM"] = c
                if "observ" in low:
                    cols_map["Observações"] = c
                if "pre" in low and "preco" in low or low.startswith("pre"):
                    # best match
                    cols_map.setdefault("Preço", c)
            # update df if mapped
            for k,v in cols_map.items():
                df.rename(columns={v:k}, inplace=True)
        if not expected.issubset(set(df.columns)):
            # não encontrou as colunas - salva cópia sem alteração e sai
            out = os.path.join(DOWNLOAD_DIR, "processed_" + fname)
            df.to_csv(out, sep=';', index=False, encoding='latin-1')
            return (fname, False, "Colunas esperadas ausentes. Salvo sem alteração.")
        # agora temos as colunas necessárias
        # converter Observações para número (float)
        updated_count = 0
        for i in df.index:
            try:
                ncm = str(df.at[i, "NCM"]).strip()
            except Exception:
                ncm = ""
            obs_raw = df.at[i, "Observações"]
            obs_num = parse_number_from_brazil(obs_raw)
            # tratar zero explicitamente (não multiplicar)
            if ncm == NCM_ALVO and obs_num is not None and obs_num != 0:
                new_price = obs_num * MULTIPLICADOR
                # format with comma decimal and two places, no thousands separator
                price_str = f"{new_price:.2f}".replace(".", ",")
                df.at[i, "Preço"] = price_str
                updated_count += 1
        out = os.path.join(DOWNLOAD_DIR, "processed_" + fname)
        df.to_csv(out, sep=';', index=False, encoding='latin-1')
        return (fname, True, f"atualizados: {updated_count}")
    except Exception as e:
        return (fname, False, f"erro: {e}")

# ========== FLUXO PRINCIPAL ==========
try:
    print("🔐 Abrindo Bling (login)...")
    driver.get("https://www.bling.com.br/login")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    safe_sleep(1)

    # Preencher username (diversos seletores)
    try:
        username_el = driver.find_element(By.ID, "username")
    except NoSuchElementException:
        try:
            username_el = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='usuário' i]")
        except NoSuchElementException:
            username_el = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='e-mail' i]")

    username_el.clear()
    username_el.send_keys(EMAIL)
    safe_sleep(0.3)

    # preencher senha (diversos seletores)
    try:
        senha_el = driver.find_element(By.ID, "password")
    except NoSuchElementException:
        try:
            senha_el = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='senha' i]")
        except NoSuchElementException:
            # fallback to password type
            senha_el = driver.find_element(By.CSS_SELECTOR, "input[type='password'], input[type='text'][placeholder*='senha' i]")
    senha_el.clear()
    senha_el.send_keys(SENHA)
    safe_sleep(0.3)

    # clicar no botão de login
    try:
        btn = driver.find_element(By.ID, "login-submit")
    except NoSuchElementException:
        try:
            btn = driver.find_element(By.XPATH, "//button[contains(., 'Entrar') or contains(., 'Login') or contains(., 'Entrar')]")
        except NoSuchElementException:
            btn = None
    if btn is None:
        raise Exception("Botão de login não identificado automaticamente. Ajuste o seletor.")
    safe_click(btn)
    print("⏳ Fazendo login, aguardando dashboard...")
    safe_sleep(LOGIN_WAIT)
    try_close_popups()

    # abrir página de exportação (URL direta — mais estável)
    print("📦 Acessando página de exportação de produtos...")
    # usei a URL que funcionou anteriormente; se necessário altere
    driver.get("https://www.bling.com.br/exportacao.produtos.php?criterio=opc-ultimos")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    safe_sleep(2)
    try_close_popups()
    remove_overlays_js()

    # localizar o elemento 'Exportar dados para planilha' e clicar
    print("🔎 Procurando 'Exportar dados para planilha' ...")
    export_selector = None
    try:
        export_selector = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'li[data-function="exportar"]')))
    except TimeoutException:
        # fallback: localizar li via texto
        try:
            elems = driver.find_elements(By.XPATH, "//li[contains(., 'Exportar dados para planilha') or contains(., 'Exportar dados')]")
            export_selector = elems[0] if elems else None
        except Exception:
            export_selector = None

    if not export_selector:
        raise Exception("Elemento 'Exportar dados para planilha' não encontrado.")

    safe_click(export_selector)
    safe_sleep(2)
    try_close_popups()
    remove_overlays_js()
    safe_sleep(1)

    # abrir export page complete; aguardar os links 'Download'
    print("📄 Página de exportação aberta — procurando links 'Download' ...")
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.action-link")), timeout=10)
    anchors = driver.find_elements(By.CSS_SELECTOR, "a.action-link")
    download_links = [a for a in anchors if "download" in a.text.strip().lower()]

    if not download_links:
        # fallback by xpath
        download_links = driver.find_elements(By.XPATH, "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'download')]")

    print(f"🔘 {len(download_links)} links 'Download' encontrados.")
    if len(download_links) == 0:
        raise Exception("Nenhum link 'Download' localizado na página de exportação.")

    # limpa CSVs antigos da pasta
    for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")):
        try:
            os.remove(f)
        except:
            pass

    # snapshot before
    before_files = list_csvs(DOWNLOAD_DIR)

    # 1) clicar no primeiro link diretamente
    first = download_links[0]
    print("➡️ Iniciando download 1/{} ...".format(len(download_links)))
    if not safe_click(first):
        # try JS click as last resort
        driver.execute_script("arguments[0].click();", first)
    # aguardar detecção de novo arquivo
    new = wait_for_new_file(before_files, timeout=30)
    if new:
        print("    ✅ Novo arquivo detectado:", os.path.basename(new))
        before_files = list_csvs(DOWNLOAD_DIR)  # atualizar snapshot
    else:
        print("    ⚠️ Não detectado novo arquivo do primeiro clique (verificar).")

    # 2) Para os demais: usar TAB + ENTER e esperar entre downloads
    remaining = len(download_links) - 1
    if remaining > 0:
        print("➡️ Iniciando downloads subsequentes via TAB + ENTER ...")
        # garantir que o foco está no primeiro link - clicamos novamente para garantir foco
        try:
            first.send_keys(Keys.NULL)
            first.click()
        except:
            pass
        # deixar o foco nativo
        safe_sleep(0.5)

        for idx in range(remaining):
            print(f"    ➡️ Acionando download {idx+2}/{len(download_links)} via TAB+ENTER ...")
            # envia TAB para avançar foco
            actions.send_keys(Keys.TAB).perform()
            safe_sleep(0.4)
            # press ENTER to trigger download
            actions.send_keys(Keys.ENTER).perform()
            # aguardar até novo arquivo aparecer
            newf = wait_for_new_file(before_files, timeout=40)
            if newf:
                print("        ✅ Novo arquivo detectado:", os.path.basename(newf))
                before_files = list_csvs(DOWNLOAD_DIR)
            else:
                print("        ⚠️ Novo arquivo NÃO detectado para esse acionamento.")
            print(f"    ⏳ Aguardando {BETWEEN_DOWNLOADS_SEC}s antes do próximo download...")
            safe_sleep(BETWEEN_DOWNLOADS_SEC)

    # esperar por quaisquer .crdownload restantes
    print("🔁 Cliques finalizados. Aguardando término de quaisquer .crdownload remanescentes...")
    wait_for_all_crdownloads_to_finish(timeout=DOWNLOAD_WAIT_TIMEOUT)
    csvs_found = list_csvs(DOWNLOAD_DIR)
    print(f"🔘 Encontrados {len(csvs_found)} novos arquivos CSV (detectados).")

    # ========== TRATAMENTO AUTOMÁTICO ==========
    print("🛠️ Iniciando tratamento automático das planilhas baixadas...")
    if len(csvs_found) == 0:
        print("⚠️ Nenhum CSV encontrado para processar.")
    else:
        for f in csvs_found:
            fname = os.path.basename(f)
            print("🔎 Processando:", fname)
            res_name, ok, message = process_csv_file(f)
            if ok:
                print("    ✅ Salvo processed_" + res_name, "-", message)
            else:
                print("    ⚠️", message)

    print("🏁 Todos os processos finalizados. Verifique a pasta:", DOWNLOAD_DIR)

except Exception as exc:
    print("❌ Erro durante o processo:\n", exc)

finally:
    print("🔚 Finalizando driver em 6s...")
    safe_sleep(6)
    try:
        driver.quit()
    except:
        pass
