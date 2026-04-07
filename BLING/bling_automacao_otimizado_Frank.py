# =========================================================
#            BLING AUTOMATION - VERSION 5 (FINAL)
#       Download via TAB → ENTER + CSV Processing BR
# =========================================================

import os
import time
import glob
from pathlib import Path
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException
)

from webdriver_manager.chrome import ChromeDriverManager


# =========================================================
# CONFIGURAÇÕES GERAIS
# =========================================================

EMAIL = "09839356100"
SENHA = "@%Gabriel99"

DOWNLOAD_DIR = str(Path.home() / "Downloads" / "bling_exports")
PROCESSED_DIR = os.path.join(DOWNLOAD_DIR, "processed")

os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(PROCESSED_DIR, exist_ok=True)

TOTAL_DOWNLOADS = 7
INTERVALO_DOWNLOAD = 12  # segundos

MULTIPLICADOR = 1400.0
NCM_ALVO = "7113.19.00"


# =========================================================
# CONFIG CHROME
# =========================================================

chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "profile.default_content_setting_values.automatic_downloads": 1,
}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--disable-popup-blocking")
chrome_options.add_argument("--start-maximized")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

wait = WebDriverWait(driver, 20)


# =========================================================
# UTILITÁRIOS
# =========================================================

def safe_sleep(t):
    time.sleep(t)

def pagina_estavel(url):
    """Evita que o Bling redirecione para o dashboard."""
    return url in driver.current_url


# =========================================================
# LOGIN
# =========================================================

print("🔐 Abrindo Bling...")
driver.get("https://www.bling.com.br/login")

wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
safe_sleep(1)

driver.find_element(By.ID, "username").send_keys(EMAIL)
driver.find_element(By.CSS_SELECTOR, "input[placeholder*='senha']").send_keys(SENHA)

btn = driver.find_element(By.XPATH, "//button[contains(.,'Entrar') or contains(.,'Acessar')]")
btn.click()

print("⏳ Aguardando login...")
safe_sleep(4)


# =========================================================
# ABRE DIRETAMENTE A TELA FINAL (SEM REDIRECIONAR)
# =========================================================

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
"""
DESTINO = (
    "https://www.bling.com.br/exportacao.produtos.php?"
    "criterio=opc-ultimos&pesquisa=&idTag=0&dataValidadeIni=&dataValidadeFim=&"
    "fornecedor=0&marca=&ncm=&dataAlteracaoIni=&dataAlteracaoFim=&"
    "codigoProdutoNoFabricante=&tipoProduto=T&idComponente=0&"
    "filtroTipoItem=&situacaoVinculoLoja=&idLojaVinculo=&"
    "filtroEstoque=&categoria=T"
)

print("📦 Abrindo página de exportação...")
driver.get(DESTINO)

safe_sleep(2)


# =========================================================
# ENCONTRA O PRIMEIRO DOWNLOAD
# =========================================================

print("📄 Buscando links 'Download'...")

downloads = driver.find_elements(By.XPATH, "//a[contains(text(),'Download')]")

if len(downloads) < TOTAL_DOWNLOADS:
    print("⚠️ Não encontrou os 7 downloads — encontrado:", len(downloads))

primeiro = downloads[0]


# =========================================================
# PRIMEIRO DOWNLOAD → CLICK NORMAL
# =========================================================


print("➡️ Download 1/7 ...")

driver.execute_script("arguments[0].scrollIntoView(true);", primeiro)
driver.execute_script("arguments[0].click();", primeiro)

safe_sleep(INTERVALO_DOWNLOAD)


# =========================================================
# DEMAIS DOWNLOADS → TAB + ENTER
# =========================================================

print("⬇️ Iniciando sequência TAB → ENTER...")

for i in range(2, TOTAL_DOWNLOADS + 1):

    # impede que o Bling saia da página
    if not pagina_estavel("exportacao.produtos.php"):
        print("⚠️ Página desviou! Redirecionando de volta...")
        driver.get(DESTINO)
        safe_sleep(2)

    print(f"➡️ Download {i}/7 ... (TAB → ENTER)")

    body = driver.find_element(By.TAG_NAME, "body")
    body.send_keys(Keys.TAB)
    safe_sleep(0.3)
    body.send_keys(Keys.ENTER)

    safe_sleep(INTERVALO_DOWNLOAD)


print("✔️ Downloads finalizados.")
safe_sleep(2)
"""

# =========================================================
# PROCESSAMENTO DOS CSVs
# =========================================================

print("🛠️ Iniciando tratamento...")

csv_files = sorted(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
unificados = []

def converter_preco_BR(valor_float):
    """1234.5 → '1.234,50'"""
    return f"{valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def to_number(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s == "" or s == "0":
        return 0
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None


for csv_path in csv_files:

    nome = os.path.basename(csv_path)
    print("🔎 Processando:", nome)

    df = pd.read_csv(csv_path, sep=';', quotechar='"', dtype=str, encoding='latin-1')
    df.columns = [c.strip() for c in df.columns]

    if not {"NCM", "Observações", "Preço"}.issubset(df.columns):
        print("⚠️ Colunas esperadas não encontradas — salvando sem alterações.")
        df.to_csv(os.path.join(PROCESSED_DIR, "processed_" + nome), sep=';', encoding='utf-8-sig', index=False)
        continue

    atualizados = 0

    for idx in df.index:

        if df.at[idx, "NCM"].strip() != NCM_ALVO:
            continue

        obs = to_number(df.at[idx, "Observações"])

        if obs is None or obs == 0:
            continue

        novo_preco = obs * MULTIPLICADOR
        df.at[idx, "Preço"] = converter_preco_BR(novo_preco)
        atualizados += 1

    out = os.path.join(PROCESSED_DIR, "processed_" + nome)
    df.to_csv(out, sep=';', index=False, encoding='utf-8-sig')

    print(f"   ✔️ Atualizados: {atualizados}")
    unificados.append(df)


# =========================================================
# UNIFICAÇÃO FINAL
# =========================================================

if unificados:
    df_all = pd.concat(unificados, ignore_index=True)
    df_all.to_csv(os.path.join(PROCESSED_DIR, "processed_all.csv"), sep=';', index=False, encoding='utf-8-sig')
    print("📦 Arquivo unificado gerado: processed_all.csv")
else:
    print("⚠️ Nenhum arquivo atualizado para unificação.")


print("\n🏁 CONCLUÍDO COM SUCESSO! 🎉")
driver.quit()
