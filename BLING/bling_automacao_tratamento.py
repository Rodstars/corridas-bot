import os
import time
import glob
from pathlib import Path
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (
    NoSuchElementException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
    TimeoutException
)
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ============================
# CONFIGURAÇÕES (ajuste)
# ============================
EMAIL = "09839356100"
SENHA = "@%Gabriel99"

# pasta onde os arquivos serão baixados
DOWNLOAD_DIR = str(Path.home() / "Downloads" / "bling_exports")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# tempo máximo para aguardar downloads (em segundos)
DOWNLOAD_WAIT_TIMEOUT = 240
DOWNLOAD_POLL_INTERVAL = 1

# multiplcador e NCM alvo
MULTIPLICADOR = 1400.0
NCM_ALVO = "7113.19.00"

# quantos downloads você espera iniciar nessa tela (normalmente 7)
TOTAL_EXPECTED_DOWNLOADS = 7

# ============================
# OPÇÕES DO CHROME
# ============================
chrome_options = webdriver.ChromeOptions()

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "profile.default_content_setting_values.automatic_downloads": 1,

    # ESSENCIAL PARA PERMITIR MÚLTIPLOS DOWNLOADS SEGUIDOS:
    "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,

    # bloquear notificações
    "profile.default_content_setting_values.notifications": 2,
}

chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--disable-popup-blocking")


# Opcional (comente se quiser ver janela)
# chrome_options.add_argument("--headless=new")  # só use se tiver certeza que tudo funciona em headless
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 30)

# ============================
# UTILITÁRIOS
# ============================
def safe_sleep(t):
    time.sleep(t)

def remove_overlays_js():
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

def try_close_popups():
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
                            driver.execute_script("arguments[0].click();", e)
                        safe_sleep(0.6)
                except Exception:
                    pass
        except Exception:
            pass

def wait_for_downloads_to_finish(expected_new_count, timeout=DOWNLOAD_WAIT_TIMEOUT):
    """
    Espera até que não haja .crdownload e existam pelo menos expected_new_count
    novos CSVs criados após o início da operação.
    """
    start = time.time()
    while True:
        crdownloads = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
        csv_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv"))
        if len(crdownloads) == 0 and len(csv_files) >= expected_new_count:
            return True, csv_files
        if time.time() - start > timeout:
            return False, csv_files
        time.sleep(DOWNLOAD_POLL_INTERVAL)

def safe_click(element, scroll=True, timeout=5):
    """Tenta várias formas de click"""
    try:
        if scroll:
            driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", element)
            safe_sleep(0.4)
        element.click()
        return True
    except (ElementClickInterceptedException, ElementNotInteractableException):
        try:
            driver.execute_script("arguments[0].click();", element)
            return True
        except Exception:
            return False
    except Exception:
        return False

# ============================
# FLUXO PRINCIPAL
# ============================
try:
    print("🔐 Abrindo Bling (login)...")
    driver.get("https://www.bling.com.br/login")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    safe_sleep(1)

    # preencher username (id="username" conforme HTML que você enviou)
    try:
        campo_username = wait.until(EC.presence_of_element_located((By.ID, "username")))
    except TimeoutException:
        # fallback por placeholder (case-insensitive)
        campo_username = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='usuário' i], input[placeholder*='e-mail' i]")
    campo_username.clear()
    campo_username.send_keys(EMAIL)
    safe_sleep(0.3)

    # preencher senha (muitos sites usam input type=text para senha em JS, por isso fallback)
    try:
        campo_senha = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
    except NoSuchElementException:
        campo_senha = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='senha' i], input[placeholder='Insira sua senha']")
    campo_senha.clear()
    campo_senha.send_keys(SENHA)
    safe_sleep(0.3)

    # clicar botão entrar (vários fallbacks)
    botao_entrar = None
    try:
        botao_entrar = driver.find_element(By.ID, "login-submit")
    except NoSuchElementException:
        try:
            botao_entrar = driver.find_element(By.XPATH, "//button[contains(., 'Entrar') or contains(., 'Login') or contains(., 'ACESSAR') or contains(., 'Acessar')]")
        except NoSuchElementException:
            pass

    if not botao_entrar:
        raise Exception("Botão de login não localizado automaticamente. Ajuste seletor.")
    safe_click(botao_entrar)
    print("⏳ Fazendo login, aguardando dashboard...")
    # espera URL ou elemento do dashboard
    safe_sleep(4)
    try_close_popups()
    remove_overlays_js()
    safe_sleep(1)

    # acessar diretamente a página de exportação (você indicou que funcionou)
    print("📦 Acessando página de exportação de produtos...")
    driver.get("https://www.bling.com.br/exportacao.produtos.php")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    safe_sleep(2)
    try_close_popups()
    remove_overlays_js()
    safe_sleep(1)

    # procurar e clicar no menu "Exportar dados para planilha" se necessário
    try:
        export_btn = None
        # primeiro tenta o li com data-function exportar
        try:
            export_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'li[data-function="exportar"]')))
        except TimeoutException:
            # alternativa por texto
            elems = driver.find_elements(By.XPATH, "//li[contains(., 'Exportar dados para planilha') or contains(., 'Exportar dados') or contains(., 'Exportar')]")
            export_btn = elems[0] if elems else None

        if export_btn:
            print("🔎 Clicando em 'Exportar dados para planilha'...")
            safe_click(export_btn)
            safe_sleep(2)
            try_close_popups()
            remove_overlays_js()
            safe_sleep(1)
        else:
            print("ℹ️ Já estamos provavelmente na tela de exportação.")
    except Exception as e:
        print("⚠️ Falha ao acionar exportar (seguimos tentando).", e)

    # ============================
    # coletar links de Download (visíveis)
    # ============================
    print("📄 Página de exportação aberta — procurando links 'Download' ...")
    # aguarda que existam anchors
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.action-link, a[href]")))
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

    # fallback: xpath procurar 'Download' explicitamente
    if not download_links:
        download_links = driver.find_elements(By.XPATH, "//a[contains(., 'Download') or contains(., 'download')]")

    print(f"🔘 {len(download_links)} links 'Download' encontrados.")

    if len(download_links) == 0:
        raise Exception("Nenhum link de Download encontrado — verifique a página manualmente.")

    # antes de clicar: registrar CSVs existentes para distinguir novos
    existing_csvs = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))

    # limpar arquivos temporários antigos .crdownload (opcional)
    for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload")):
        try:
            os.remove(f)
        except Exception:
            pass

    # clicar nos links (com intervalos)
    successful_clicks = 0
    for idx, link in enumerate(download_links[:TOTAL_EXPECTED_DOWNLOADS], start=1):
        print(f"➡️ Iniciando download {idx}/{min(len(download_links), TOTAL_EXPECTED_DOWNLOADS)} ...")
        try:
            ok = safe_click(link)
            if not ok:
                # tentar via JS
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
                    driver.execute_script("arguments[0].click();", link)
                    ok = True
                except Exception:
                    ok = False
            if ok:
                successful_clicks += 1
                safe_sleep(1.2)  # dar tempo para começar o download
            else:
                print(f"⚠️ Falha ao iniciar download {idx}.")
        except Exception as e:
            print("⚠️ Erro durante clique:", e)

    print(f"🔁 Cliques finalizados. Downloads iniciados (tentativas bem sucedidas): {successful_clicks}")

    # ============================
    # esperar downloads e obter novos csvs
    # ============================
    print("⏳ Aguardando conclusão dos downloads...")
    ok, csv_files_all = wait_for_downloads_to_finish(expected_new_count=successful_clicks, timeout=DOWNLOAD_WAIT_TIMEOUT)
    if not ok:
        print("⚠️ Timeout esperando downloads. Verifique manualmente em:", DOWNLOAD_DIR)

    # identificar quais CSVs são novos (gerados nesta execução)
    csv_files = sorted([f for f in csv_files_all if f not in existing_csvs])
    print(f"🔘 Encontrados {len(csv_files)} novos arquivos CSV: ")
    for f in csv_files:
        print("   ", os.path.basename(f))

    # ============================
    # TRATAMENTO DAS PLANILHAS
    # ============================
    if not csv_files:
        print("⚠️ Nenhum CSV novo para processar. Encerrando etapa de tratamento.")
    else:
        print("🛠️ Iniciando tratamento das planilhas baixadas...")
        for csv_path in csv_files:
            base = os.path.basename(csv_path)
            try:
                print("🔎 Processando:", base)
                # leitura robusta: Bling usa ; como separador e campos entre aspas
                try:
                    df = pd.read_csv(csv_path, sep=';', quotechar='"', encoding='utf-8', dtype=str)
                except Exception:
                    # fallback para latin-1
                    df = pd.read_csv(csv_path, sep=';', quotechar='"', encoding='latin-1', dtype=str)

                # strip nas colunas
                df.columns = [c.strip() for c in df.columns]

                # checar colunas necessárias
                if "NCM" not in df.columns or "Observações" not in df.columns or "Preço" not in df.columns:
                    print("⚠️ Arquivo não contém todas as colunas esperadas (NCM, Observações, Preço). Salvando cópia sem alteração.")
                    out_path = os.path.join(DOWNLOAD_DIR, "processed_" + base)
                    df.to_csv(out_path, index=False, encoding='utf-8-sig', sep=';')
                    continue

                # função para transformar 'Observações' em número (tratando 5,5 ou 5.5, removendo pontos de milhar)
                def to_number(x):
                    if pd.isna(x):
                        return None
                    s = str(x).strip()
                    if s == "":
                        return None
                    # remover aspas internas e espaços
                    s = s.replace('"', '').replace("'", "").strip()
                    # se contém letras, rejeita (retorna None)
                    # mas pode ter "10" ou "5,5" ou "1.234,56"
                    # primeiro remover pontos de milhar (quando houver) - detecta padrão com '.' antes de 3 dígitos
                    # uma forma simples: remover todos os pontos e depois transformar vírgula para ponto
                    s2 = s.replace('.', '').replace(',', '.')
                    try:
                        return float(s2)
                    except:
                        # último recurso: tentar substituir vírgula por ponto sem remover pontos
                        s3 = s.replace(',', '.')
                        try:
                            return float(s3)
                        except:
                            return None

                # máscara NCM alvo (limpando espaços)
                mask = df["NCM"].astype(str).str.strip() == NCM_ALVO

                updated_count = 0
                for idx in df.index[mask]:
                    obs_raw = df.at[idx, "Observações"]
                    val = to_number(obs_raw)
                    if val is not None:
                        new_price = val * MULTIPLICADOR
                        # opcional: formatar com 2 decimais e vírgula como decimal (se preferir)
                        # df.at[idx, "Preço"] = f"{new_price:.2f}".replace('.', ',')
                        # armazenamos como número (string) com ponto decimal para evitar confusão
                        df.at[idx, "Preço"] = f"{new_price:.2f}"
                        updated_count += 1
                    else:
                        # não numérico -> pula
                        pass

                out_path = os.path.join(DOWNLOAD_DIR, "processed_" + base)
                df.to_csv(out_path, index=False, encoding='utf-8-sig', sep=';')
                print(f"✅ Salvo {os.path.basename(out_path)} — registros atualizados: {updated_count}")

            except Exception as e:
                print("❌ Erro ao processar", base, ":", e)

    print("🏁 Todos os processos finalizados. Verifique a pasta:", DOWNLOAD_DIR)

except Exception as e:
    print("❌ Erro durante o processo:\n", e)

finally:
    print("🔚 Fechando driver em 5s...")
    time.sleep(5)
    try:
        driver.quit()
    except Exception:
        pass
