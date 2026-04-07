import os
import time
import glob
import pandas as pd
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# ============================================================
# CONFIGURAÇÕES
# ============================================================
EMAIL = "09839356100"
SENHA = "@%Gabriel99"

DOWNLOAD_DIR = str(Path.home() / "Downloads" / "bling_exports")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

WAIT_BETWEEN_DOWNLOADS = 10          # segundos entre TAB + ENTER
DOWNLOAD_WAIT_TIMEOUT = 200

NCM_ALVO = "7113.19.00"
MULTIPLICADOR = 1400.0


# ============================================================
# CONFIGURAÇÃO DO CHROME
# ============================================================
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 20)


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
def safe_sleep(t):
    time.sleep(t)


def wait_for_downloads(target_count):
    """Aguarda downloads terminarem."""
    start = time.time()
    while True:
        cr = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
        ok = glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv"))
        if len(cr) == 0 and len(ok) >= target_count:
            return True
        if time.time() - start > DOWNLOAD_WAIT_TIMEOUT:
            print("⚠️ Timeout aguardando downloads.")
            return False
        time.sleep(1)


def limpar_csv_antigos():
    for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")):
        try:
            os.remove(f)
        except:
            pass


# ============================================================
# LOGIN
# ============================================================
print("🔐 Abrindo Bling (login)...")
driver.get("https://www.bling.com.br/login")

wait.until(EC.presence_of_element_located((By.ID, "username")))
driver.find_element(By.ID, "username").send_keys(EMAIL)

senha_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='senha' i]")
senha_input.send_keys(SENHA)

button = driver.find_element(By.ID, "login-submit")
button.click()

print("⏳ Fazendo login, aguardando dashboard...")
safe_sleep(5)


# ============================================================
# NAVEGAR PARA A EXPORTAÇÃO
# ============================================================
print("📦 Acessando página de exportação de produtos...")
driver.get("https://www.bling.com.br/exportacao.produtos.php")
safe_sleep(3)


# ============================================================
# LOCALIZAR LINKS "DOWNLOAD"
# ============================================================
print("📄 Página de exportação aberta — procurando links 'Download' ...")
wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))

links = driver.find_elements(By.XPATH, "//a[contains(text(),'Download')]")
total_links = len(links)

print(f"🔘 {total_links} links 'Download' encontrados.")

if total_links == 0:
    raise Exception("Nenhum link de download encontrado.")


# ============================================================
# PASSO 1 — CLICAR NO PRIMEIRO DOWNLOAD
# ============================================================
limpar_csv_antigos()

print("➡️ Iniciando download 1/7 (clique direto)...")
links[0].click()
safe_sleep(WAIT_BETWEEN_DOWNLOADS)


# ============================================================
# PASSO 2 — ACIONAR TAB + ENTER PARA OS DOWNLOADS RESTANTES
# ============================================================
for i in range(2, total_links + 1):
    print(f"➡️ Iniciando download {i}/{total_links} via TAB + ENTER ...")
    driver.switch_to.active_element.send_keys(Keys.TAB)
    safe_sleep(0.5)
    driver.switch_to.active_element.send_keys(Keys.ENTER)
    safe_sleep(WAIT_BETWEEN_DOWNLOADS)


# ============================================================
# AGUARDAR DOWNLOADS
# ============================================================
print("⏳ Aguardando conclusão dos downloads...")
wait_for_downloads(total_links)

csv_files = sorted(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
print("🔘 Arquivos baixados:", [os.path.basename(f) for f in csv_files])


# ============================================================
# TRATAMENTO DAS PLANILHAS
# ============================================================
def to_number(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None


print("🛠️ Iniciando tratamento das planilhas...")
for csv_path in csv_files:

    print("🔎 Processando:", os.path.basename(csv_path))

    try:
        df = pd.read_csv(csv_path, dtype=str, encoding="utf-8")
    except:
        df = pd.read_csv(csv_path, dtype=str, encoding="latin-1")

    df.columns = [c.strip() for c in df.columns]

    if not {"NCM", "Observações", "Preço"}.issubset(df.columns):
        print("⚠️ Colunas obrigatórias não encontradas:", csv_path)
        continue

    updated = 0

    for idx in df.index:
        if df.at[idx, "NCM"] == NCM_ALVO:

            valor_obs = to_number(df.at[idx, "Observações"])

            # Ignorar ZERO
            if valor_obs is None or valor_obs == 0:
                continue

            novo_preco = valor_obs * MULTIPLICADOR

            # Formato brasileiro com vírgula
            df.at[idx, "Preço"] = f"{novo_preco:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

            updated += 1

    out = os.path.join(DOWNLOAD_DIR, "processed_" + os.path.basename(csv_path))
    df.to_csv(out, index=False, encoding="utf-8-sig")

    print(f"✅ Salvo {os.path.basename(out)} — registros atualizados: {updated}")


print("🏁 Processo finalizado com sucesso!")
driver.quit()
