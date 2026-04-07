# bling_automacao_otimizado.py
import os
import time
import glob
from pathlib import Path
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
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

# ========== CONFIGURAÇÃO (EDITE AQUI) ==========
EMAIL = "09839356100"     # substitua
SENHA = "@%Gabriel99"                # substitua
DOWNLOAD_DIR = str(Path.home() / "Downloads" / "bling_exports")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# Timeout / delays
DOWNLOAD_WAIT_TIMEOUT = 180        # tempo máximo (segundos) para aguardar cada novo arquivo aparecer
POLL_INTERVAL = 1                  # polling para checagem de arquivos
CLICK_DELAY = 12                   # segundos entre cada clique de download (recomendado 10-12s)
MAX_CLICK_RETRIES = 3

# Regras de tratamento
MULTIPLICADOR = 1400.0
NCM_ALVO = "7113.19.00"

# ========== OPÇÕES DO CHROME ==========
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "profile.default_content_setting_values.notifications": 2,
}
chrome_options.add_experimental_option("prefs", prefs)
# chrome_options.add_argument("--start-maximized")  # opcional
# chrome_options.add_argument("--headless=new")     # opcional, desative se quiser ver o navegador

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 20)

# ========== UTILITÁRIOS ==========
def safe_sleep(sec):
    time.sleep(sec)

def remove_overlays_js():
    try:
        js = """
        document.querySelectorAll('.modal, .modal-backdrop, .overlay, .ui-widget-overlay, .fancybox-overlay, .lightbox-overlay').forEach(el => el.remove());
        document.querySelectorAll('[style*="z-index"]').forEach(el => {
            try { el.style.zIndex = '0'; } catch(e){}
        });
        """
        driver.execute_script(js)
    except Exception:
        pass

def try_close_popups():
    xpaths = [
        "//button[contains(@class,'close')]",
        "//button[contains(text(),'×') or contains(text(),'X')]",
        "//a[contains(@class,'close') or contains(text(),'×') or contains(text(),'X')]",
        "//button[contains(@aria-label,'fechar') or contains(@aria-label,'close')]",
        "//div[contains(@class,'modal')]//button[contains(., 'Fechar') or contains(., 'OK')]"
    ]
    for xp in xpaths:
        try:
            for e in driver.find_elements(By.XPATH, xp):
                try:
                    if e.is_displayed():
                        e.click()
                        time.sleep(0.5)
                except Exception:
                    try:
                        driver.execute_script("arguments[0].click();", e)
                        time.sleep(0.5)
                    except Exception:
                        pass
        except Exception:
            pass

def safe_click(element, tries=MAX_CLICK_RETRIES):
    for i in range(tries):
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
                time.sleep(0.8)
    return False

def wait_new_file(old_files, timeout=DOWNLOAD_WAIT_TIMEOUT):
    """Espera até que surja um novo .csv na pasta (que não esteja em .crdownload)."""
    start = time.time()
    while True:
        csvs = [f for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv"))]
        new = [f for f in csvs if f not in old_files]
        # garantir que não há .crdownload relacionado ao novo arquivo
        crs = [f for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))]
        if new and not crs:
            return new[0]  # retorna o primeiro novo arquivo detectado
        if time.time() - start > timeout:
            return None
        time.sleep(POLL_INTERVAL)

def to_number_from_brazil(s):
    """Converte texto BR (ex: '1.234,56' ou '5,5' ou '10') em float. Retorna None se inválido."""
    if s is None: 
        return None
    st = str(s).strip()
    if st == "":
        return None
    # retirar espaços e possíveis textos não numéricos
    # remover pontos de milhar e trocar vírgula decimal por ponto
    try:
        # remover tudo que não for dígito, ponto ou vírgula ou sinal
        allowed = "0123456789,.-"
        st = ''.join(ch for ch in st if ch in allowed)
        # se houver mais de one '-' ou in wrong pos, let float conversion handle
        st = st.replace(".", "").replace(",", ".")
        val = float(st)
        return val
    except Exception:
        return None

def format_brazilian(value):
    """Formata float para moeda BR: 7700 -> '7.700,00'"""
    try:
        v = float(value)
    except Exception:
        return value
    s = f"{v:,.2f}"          # ex: '7,700.00'
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s

# ========== FLUXO PRINCIPAL ==========
try:
    print("🔐 Abrindo Bling (login)...")
    driver.get("https://www.bling.com.br/login")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(1)

    # username
    try:
        campo_user = driver.find_element(By.ID, "username")
    except NoSuchElementException:
        try:
            campo_user = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='usuário' i], input[placeholder*='e-mail' i]")
        except Exception:
            campo_user = None

    if campo_user:
        campo_user.clear()
        campo_user.send_keys(EMAIL)
        time.sleep(0.2)
    else:
        print("⚠️ Campo usuário não encontrado automaticamente. Prossiga manualmente e pressione Enter no terminal para continuar.")
        input("Pressione Enter após logar manualmente...")

    # senha
    try:
        # preferir campo password se tiver
        campo_senha = None
        try:
            campo_senha = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        except NoSuchElementException:
            campo_senha = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='senha' i], input[placeholder*='Senha' i], input[type='text'][placeholder*='senha' i]")
        if campo_senha:
            campo_senha.clear()
            campo_senha.send_keys(SENHA)
            time.sleep(0.2)
    except Exception:
        pass

    # botão entrar (flexível)
    try:
        botao = driver.find_element(By.ID, "login-submit")
    except NoSuchElementException:
        try:
            botao = driver.find_element(By.XPATH, "//button[contains(., 'Entrar') or contains(., 'Login') or contains(., 'Log in') or contains(., 'ACESSAR')]")
        except Exception:
            botao = None

    if botao:
        safe_click(botao)
    else:
        print("⚠️ Botão de login não encontrado automaticamente. Prossiga manualmente e pressione Enter no terminal para continuar.")
        input("Pressione Enter após logar manualmente...")

    print("⏳ Fazendo login, aguardando dashboard...")
    # aguardar URL ou elemento que comprove login
    try:
        wait.until(EC.url_contains("bling.com.br"))
    except TimeoutException:
        pass
    time.sleep(4)
    try_close_popups()
    remove_overlays_js()
    time.sleep(1)

    # Ir para a página de exportação direta (você indicou que exportacao.produtos funciona)
    print("📦 Acessando página de exportação de produtos...")
    driver.get("https://www.bling.com.br/exportacao.produtos.php?criterio=opc-ultimos")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(2)
    try_close_popups()
    remove_overlays_js()
    time.sleep(1)

    print("📄 Página de exportação aberta — procurando links 'Download' ...")
    # localizar links com texto 'Download' (action-link)
    anchors = driver.find_elements(By.CSS_SELECTOR, "a.action-link")
    download_links = []
    for a in anchors:
        try:
            if "download" in a.text.strip().lower():
                download_links.append(a)
        except Exception:
            pass

    # fallback: procurar por ancoras que contenham 'Download' no xpath
    if not download_links:
        anchors_x = driver.find_elements(By.XPATH, "//a[contains(., 'Download') or contains(., 'download')]")
        download_links = anchors_x

    print(f"🔘 {len(download_links)} links 'Download' encontrados.")

    if not download_links:
        raise Exception("Nenhum link de Download encontrado na página de exportação.")

    # Limpar arquivos antigos
    old_csvs = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))

    # Estratégia de clique:
    # - Clicar no primeiro link diretamente.
    # - Para cada link subsequente: esperar CLICK_DELAY segundos e clicar também (com scroll+JS)
    # - Após cada clique, aguardar aparecimento de novo arquivo na pasta (wait_new_file).
    new_files = []
    for idx, link in enumerate(download_links, start=1):
        print(f"➡️ Iniciando download {idx}/{len(download_links)} ...")
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
        except Exception:
            pass
        ok = safe_click(link)
        if not ok:
            # tentativa via JS click direta
            try:
                driver.execute_script("arguments[0].click();", link)
                ok = True
            except Exception:
                ok = False
        if not ok:
            print(f"⚠️ Falha ao iniciar clique no link {idx}. Pulando.")
            continue

        # após iniciar o clique, aguardar que um novo arquivo CSV surja (até DOWNLOAD_WAIT_TIMEOUT)
        new = wait_new_file(old_csvs, timeout=DOWNLOAD_WAIT_TIMEOUT)
        if new:
            print("    ✅ Novo arquivo detectado:", os.path.basename(new))
            new_files.append(new)
            old_csvs.add(new)
        else:
            print("    ⚠️ Novo arquivo NÃO detectado dentro do tempo limite. Continue verificando manualmente.")
        # esperar um tempo extra entre cliques (evita prompt/permissão do Chrome)
        print(f"    ⏳ Aguardando {CLICK_DELAY}s antes do próximo download...")
        time.sleep(CLICK_DELAY)

    print(f"🔁 Cliques finalizados. Downloads iniciados (detectados): {len(new_files)}")

    # Se nenhum arquivo novo detectado, avisar e continuar
    if not new_files:
        print("⚠️ Nenhum arquivo foi detectado como baixado. Verifique permissões e teste manualmente.")
    else:
        print("⏳ Aguardando término de quaisquer .crdownload remanescentes...")
        # aguardar até que não haja .crdownload
        start = time.time()
        while True:
            crs = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
            if not crs:
                break
            if time.time() - start > DOWNLOAD_WAIT_TIMEOUT:
                break
            time.sleep(1)
        print("🔘 Encontrados", len(new_files), "novos arquivos CSV (detectados).")

    # ========== TRATAMENTO DOS CSVs ==========
    print("🛠️ Iniciando tratamento automático das planilhas baixadas...")
    csv_files = sorted(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
    if not csv_files:
        print("⚠️ Nenhum CSV encontrado para processar.")
    else:
        for csv_path in csv_files:
            try:
                print("🔎 Processando:", os.path.basename(csv_path))
                # leitura tolerante (pulando linhas corrompidas)
                try:
                    df = pd.read_csv(csv_path, dtype=str, on_bad_lines='skip', sep=',', encoding='utf-8')
                except Exception:
                    df = pd.read_csv(csv_path, dtype=str, on_bad_lines='skip', sep=',', encoding='latin-1')

                # saneamento de colunas
                df.columns = [c.strip() for c in df.columns]

                required = {"NCM", "Observações", "Preço"}
                if not required.issubset(set(df.columns)):
                    print("    ⚠️ Colunas esperadas ausentes. Salvando cópia sem alteração.")
                    out_path = os.path.join(DOWNLOAD_DIR, "processed_" + os.path.basename(csv_path))
                    df.to_csv(out_path, index=False, encoding="utf-8-sig")
                    continue

                updated = 0
                # iterar linhas, aplicar regra
                for idx in df.index:
                    try:
                        ncm_val = str(df.at[idx, "NCM"]).strip()
                    except Exception:
                        ncm_val = ""
                    if ncm_val != NCM_ALVO:
                        continue

                    obs_raw = df.at[idx, "Observações"]
                    num = to_number_from_brazil(obs_raw)
                    # se não é número ou é zero -> pular
                    if num is None:
                        continue
                    if abs(num) < 1e-9:
                        continue

                    # calcular novo preço
                    novo = num * MULTIPLICADOR
                    # formatar para BR como string
                    novo_fmt = format_brazilian(novo)
                    df.at[idx, "Preço"] = novo_fmt
                    updated += 1

                out_path = os.path.join(DOWNLOAD_DIR, "processed_" + os.path.basename(csv_path))
                df.to_csv(out_path, index=False, encoding="utf-8-sig")
                print(f"    ✅ Salvo {os.path.basename(out_path)} — registros atualizados: {updated}")
            except Exception as e:
                print("    ❌ Erro ao processar", os.path.basename(csv_path), ":", e)

    print("🏁 Todos os processos finalizados. Verifique a pasta:", DOWNLOAD_DIR)

except Exception as e:
    print("❌ Erro durante o processo:", e)

finally:
    print("🔚 Finalizando driver em 6s...")
    time.sleep(6)
    try:
        driver.quit()
    except Exception:
        pass
