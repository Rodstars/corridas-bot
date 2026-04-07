# =========================================================
#            BLING AUTOMATION - VERSION 5 (FINAL)
#       Download via TAB → ENTER + CSV Processing BR
# =========================================================

import os
import time
import glob
import shutil
import re
from pathlib import Path
import pandas as pd
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

# ============================
# === CONFIGURAÇÃO ==========
# ============================
EMAIL = "09839356100"   # <<-- ajuste
SENHA = "@%Gabriel99"   # <<-- ajuste
# pasta onde os arquivos serão baixados
DOWNLOAD_DIR = str(Path.home() / "Downloads" / "bling_exports")
PROCESSED_DIR = os.path.join(DOWNLOAD_DIR, "processed")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(PROCESSED_DIR, exist_ok=True)

# Espera / timeouts
DOWNLOAD_WAIT_TIMEOUT = 300        # segundos para aguardar downloads finalizarem
DOWNLOAD_POLL_INTERVAL = 1         # polling interval (segundos)
CLICK_INTERVAL_SECONDS = 12        # tempo entre cliques (recomendado 10-12s)
MAX_WAIT_ELEMENT = 25              # tempo máximo para esperar elementos em tela

# Multiplicador e NCM
MULTIPLICADOR = 1400.0
NCM_ALVO = "7113.19.00"

# ============================
# === CHROME OPTIONS ========
# ============================
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "profile.default_content_setting_values.notifications": 2,
    "profile.default_content_settings.popups": 0,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True,
}
chrome_options.add_experimental_option("prefs", prefs)
# Recomendo ver a janela enquanto desenvolve:
# chrome_options.add_argument("--start-maximized")
# Para reduzir logs
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Inicializa driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, MAX_WAIT_ELEMENT)


# ============================
# === UTILITÁRIOS ===========
# ============================
def safe_sleep(t):
    print(f"⏳ aguardando {t}s...")
    time.sleep(t)

def remove_overlays_js():
    """Remove blocos comuns de overlay/modal via JS."""
    js = """
    try {
      document.querySelectorAll('.modal, .modal-backdrop, .overlay, .ui-widget-overlay, .fancybox-overlay, .lightbox-overlay').forEach(el => el.remove());
      document.querySelectorAll('[style*="z-index"]').forEach(el => {
        try { el.style.zIndex = '0'; } catch(e) {}
      });
    } catch(e) {}
    """
    try:
        driver.execute_script(js)
    except Exception:
        pass

def try_close_popups():
    """Várias tentativas de cliques em botões X/Fechar comuns."""
    xpaths = [
        "//button[contains(@class,'close')]", 
        "//button[contains(text(),'×') or contains(text(),'X') or contains(text(),'Fechar') or contains(text(),'fechar')]",
        "//a[contains(@class,'close') or contains(text(),'×') or contains(text(),'Fechar') or contains(text(),'fechar')]",
        "//div[contains(@class,'modal')]//button",
        "//div[contains(@class,'popup')]//button[contains(.,'Fechar') or contains(.,'fechar') or contains(.,'×')]",
        "//button[contains(@aria-label,'close') or contains(@aria-label,'fechar')]"
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
                        time.sleep(0.5)
                except Exception:
                    pass
        except Exception:
            pass

def wait_for_new_csvs(existing_set, expected_count, timeout=DOWNLOAD_WAIT_TIMEOUT):
    """Aguarda até que pelo menos expected_count novos CSVs apareçam e .crdownload desapareça."""
    start = time.time()
    while True:
        cr = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
        all_csvs = sorted(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
        new_csvs = [p for p in all_csvs if p not in existing_set]
        if len(cr) == 0 and len(new_csvs) >= expected_count:
            return True, new_csvs
        if time.time() - start > timeout:
            return False, new_csvs
        time.sleep(DOWNLOAD_POLL_INTERVAL)

def js_click(element):
    """Clica via JavaScript com scrollIntoView e tratamento."""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
        safe_sleep(0.4)
        driver.execute_script("arguments[0].click();", element)
        return True
    except WebDriverException as e:
        print("⚠️ js_click falhou:", e)
        return False

def normalize_col_name(s):
    """Normaliza nome de coluna: remove BOM, aspas e espaços extras."""
    if s is None:
        return s
    # remover BOM
    s = str(s)
    s = s.lstrip("\ufeff").strip()
    # remover aspas extras
    s = s.strip('"').strip("'")
    # remover caracteres invisíveis
    s = re.sub(r'[\x00-\x1f]+', '', s)
    # strip final
    return s.strip()

def to_number_obs(x):
    """Converte string Observações para float. Retorna None se não for número válido."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip()
    if s == "":
        return None
    # limpar aspas
    s = s.replace('"', '').replace("'", "").strip()
    # remover espaçamento
    s = s.replace(" ", "")
    # Se for apenas '0' ou '0,0' etc -> 0.0
    # remover pontos de milhar e trocar vírgula para ponto
    # abordagem: remover todos '.' (pontos) e trocar ',' por '.'
    s_clean = s.replace(".", "").replace(",", ".")
    # permitir números como '5.5' ou '5'
    try:
        val = float(s_clean)
        return val
    except:
        # tenta inverter (caso tenham sido usados pontos decimais no lugar de milhar)
        try:
            s_alt = s.replace(",", "")
            return float(s_alt)
        except:
            return None

def format_brazil_number(value):
    """Formata float para padrão BR: ponto milhares e vírgula decimal (2 casas)."""
    try:
        v = float(value)
    except:
        return value
    # usa formatação com separador de milhares US (',' thousands '.' decimal), depois troca
    s = f"{v:,.2f}"   # ex: '7,700.00'
    # troca: ',' -> 'X', '.' -> ',', 'X' -> '.'
    s = s.replace(',', 'X').replace('.', ',').replace('X', '.')
    return s

# ============================
# === FLUXO PRINCIPAL =======
# ============================
try:
    print("🔐 Abrindo Bling (login)...")
    driver.get("https://www.bling.com.br/login")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    safe_sleep(1)

    # === preencher username ===
    try:
        el_user = wait.until(EC.presence_of_element_located((By.ID, "username")))
    except TimeoutException:
        # fallback por placeholder
        el_user = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='usuário' i], input[placeholder*='e-mail' i]")
    el_user.clear()
    el_user.send_keys(EMAIL)
    safe_sleep(0.3)

    # === preencher senha ===
    try:
        el_pass = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
    except NoSuchElementException:
        el_pass = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='senha' i], input[placeholder='Insira sua senha']")
    el_pass.clear()
    el_pass.send_keys(SENHA)
    safe_sleep(0.3)

    # === botão login (vários fallbacks) ===
    botao_entrar = None
    try:
        botao_entrar = driver.find_element(By.ID, "login-submit")
    except NoSuchElementException:
        try:
            botao_entrar = driver.find_element(By.XPATH, "//button[contains(., 'Entrar') or contains(., 'Login') or contains(., 'ACESSAR') or contains(., 'Acessar')]")
        except NoSuchElementException:
            botao_entrar = None

    if not botao_entrar:
        raise Exception("Botão de login não encontrado automaticamente. Ajuste o seletor no código.")
    # click via JS
    js_click(botao_entrar)

    print("⏳ Fazendo login, aguardando dashboard...")
    safe_sleep(4)
    try_close_popups()
    remove_overlays_js()
    safe_sleep(1)

    # === acessar página de exportação diretamente (melhor) ===
    print("📦 Acessando página de exportação de produtos...")
    driver.get("https://www.bling.com.br/exportacao.produtos.php?criterio=opc-ultimos&pesquisa=&idTag=0&dataValidadeIni=&dataValidadeFim=&fornecedor=0&marca=&ncm=&dataAlteracaoIni=&dataAlteracaoFim=&codigoProdutoNoFabricante=&tipoProduto=T&idComponente=0&filtroTipoItem=&situacaoVinculoLoja=&idLojaVinculo=&filtroEstoque=&categoria=T")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    safe_sleep(2)
    try_close_popups()
    remove_overlays_js()
    safe_sleep(1)

    # Se houver um passo extra (clicar no menu exportar), tenta: (fallback)
    try:
        export_li = driver.find_element(By.CSS_SELECTOR, 'li[data-function="exportar"]')
        # se necessário, ativar
        try:
            js_click(export_li)
            safe_sleep(1)
        except Exception:
            pass
    except Exception:
        pass

    # === localizar os links "Download" dinamicamente ===
    print("🔎 Procurando 'Download' links na página...")
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.action-link, a")))
    safe_sleep(0.8)
    anchors = driver.find_elements(By.CSS_SELECTOR, "a.action-link, a")
    download_links = []
    for a in anchors:
        try:
            txt = a.text.strip().lower()
            if "download" in txt:
                download_links.append(a)
        except Exception:
            pass

    if not download_links:
        # fallback xpath
        download_links = driver.find_elements(By.XPATH, "//a[contains(translate(., 'DOWNLOAD', 'download'), 'download')]")

    print(f"🔘 {len(download_links)} links 'Download' encontrados.")

    if len(download_links) == 0:
        raise Exception("Nenhum link de Download encontrado na página. Verifique manualmente.")

    # marcar arquivos CSV existentes para detectar novos
    existing_csvs = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
    # remover .crdownload remanescentes antigos
    for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload")):
        try:
            os.remove(f)
        except Exception:
            pass

    # === clicar em cada link via JS com intervalo ===
    successful_clicks = 0
    for idx, link in enumerate(download_links, start=1):
        print(f"➡️ Iniciando download {idx}/{len(download_links)} ...")
        try:
            # remover overlays a cada iteração
            try_close_popups()
            remove_overlays_js()
            safe_sleep(0.5)
            ok = js_click(link)
            if not ok:
                # tentativa extra com scroll e click
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'})", link)
                    driver.execute_script("arguments[0].click();", link)
                    ok = True
                except Exception:
                    ok = False
            if ok:
                successful_clicks += 1
                # aguarda intervalo configurado antes do próximo clique
                print(f"    ✅ Novo download solicitado. Aguardando {CLICK_INTERVAL_SECONDS}s antes do próximo...")
                time.sleep(CLICK_INTERVAL_SECONDS)
            else:
                print(f"    ⚠️ Não conseguiu acionar download {idx}.")
        except Exception as e:
            print("    ⚠️ Erro ao tentar iniciar download:", e)

    print(f"🔁 Cliques finalizados. Downloads solicitados (detectados): {successful_clicks}")

    # === aguardar novos CSVs serem criados e .crdownload sumir ===
    print("⏳ Aguardando término dos downloads...")
    ok, new_csvs = wait_for_new_csvs(existing_csvs, successful_clicks, timeout=DOWNLOAD_WAIT_TIMEOUT)
    if not ok:
        print("⚠️ Timeout aguardando downloads. Arquivos parciais detectados ou menos arquivos que o esperado.")
    new_csvs = sorted(new_csvs)
    print(f"🔘 Encontrados {len(new_csvs)} novos arquivos CSV:")
    for p in new_csvs:
        print("   ", os.path.basename(p))

    # ============================
    # TRATAMENTO DOS CSVs
    # ============================
    print("🛠️ Iniciando tratamento automático das planilhas baixadas...")
    processed_paths = []
    for csv_path in new_csvs:
        base = os.path.basename(csv_path)
        try:
            print("🔎 Processando:", base)
            # tentar leitura padrão (Bling costuma usar ; como separador)
            df = None
            tried = []
            for enc in ("utf-8", "latin-1", "cp1252"):
                try:
                    df = pd.read_csv(csv_path, sep=';', quotechar='"', encoding=enc, dtype=str)
                    tried.append(enc)
                    break
                except Exception as e:
                    tried.append(f"fail({enc})")
                    continue
            if df is None:
                raise Exception("Falha ao ler CSV com encodings testados: " + ",".join(tried))

            # normalizar nomes de colunas
            df.columns = [normalize_col_name(c) for c in df.columns]

            # checar colunas necessárias
            if not all(col in df.columns for col in ["NCM", "Observações", "Preço"]):
                print("    ⚠️ Colunas esperadas ausentes. Salvando cópia sem alteração.")
                out_path = os.path.join(PROCESSED_DIR, "processed_" + base)
                df.to_csv(out_path, index=False, encoding='utf-8-sig', sep=';')
                processed_paths.append(out_path)
                continue

            # converte Observações para número com to_number_obs
            obs_series = df["Observações"].apply(to_number_obs)

            # aplicar máscara NCM e Observações > 0
            mask = df["NCM"].astype(str).str.strip() == NCM_ALVO
            updated = 0
            for i in df.index[mask]:
                obs_val = obs_series.iloc[i] if i < len(obs_series) else to_number_obs(df.at[i, "Observações"])
                if obs_val is None:
                    continue
                # se é zero ou muito próximo de zero -> pular
                if abs(obs_val) < 1e-9:
                    continue
                # somente multiplicar quando Observações > 0
                if obs_val > 0:
                    new_price = obs_val * MULTIPLICADOR
                    # formatar para BR: 7.700,00
                    df.at[i, "Preço"] = format_brazil_number(new_price)
                    updated += 1

            out_path = os.path.join(PROCESSED_DIR, "processed_" + base)
            df.to_csv(out_path, index=False, encoding='utf-8-sig', sep=';')
            processed_paths.append(out_path)
            print(f"    ✅ Salvo {os.path.basename(out_path)} — registros atualizados: {updated}")

        except Exception as e:
            print("    ❌ Erro ao processar", base, ":", e)

    # criar arquivo unificado processed_all.csv
    if processed_paths:
        print("🔗 Unindo processed CSVs em único arquivo...")
        try:
            dfs = []
            for p in processed_paths:
                try:
                    d = pd.read_csv(p, sep=';', encoding='utf-8', dtype=str)
                except:
                    d = pd.read_csv(p, sep=';', encoding='latin-1', dtype=str)
                dfs.append(d)
            all_df = pd.concat(dfs, ignore_index=True, sort=False)
            all_out = os.path.join(PROCESSED_DIR, "processed_all.csv")
            all_df.to_csv(all_out, index=False, encoding='utf-8-sig', sep=';')
            print("    ✅ Arquivo unificado salvo em:", all_out)
        except Exception as e:
            print("    ⚠️ Falha ao criar arquivo unificado:", e)
    else:
        print("⚠️ Nenhum processed CSV gerado para unir.")

    print("🏁 Processo concluído. Verifique as pastas:")
    print(" - downloads:", DOWNLOAD_DIR)
    print(" - processed:", PROCESSED_DIR)

except Exception as exc:
    print("❌ Erro durante o processo:\n", exc)

finally:
    print("🔚 Finalizando driver em 6s...")
    time.sleep(6)
    try:
        driver.quit()
    except Exception:
        pass
