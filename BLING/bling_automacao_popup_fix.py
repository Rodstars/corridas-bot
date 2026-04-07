# bling_automacao_popup_fix.py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time, os

# === CONFIGURAÇÕES DO CHROME ===
chrome_options = Options()
chrome_options.add_argument("--start-maximized")

download_dir = r"C:\Users\rodri\Downloads\bling_export"

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)

# === LOGIN ===
driver.get("https://www.bling.com.br/login")

wait = WebDriverWait(driver, 20)

campo_email = wait.until(EC.presence_of_element_located((By.ID, "username")))
campo_email.send_keys("09839356100")

campo_senha = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Insira sua senha']")))
campo_senha.send_keys("@%Gabriel99")

driver.find_element(By.TAG_NAME, "button").click()

# === FECHAR POP-UP SE EXISTIR ===
time.sleep(5)
try:
    driver.execute_script("""
        let pop = document.querySelector('.modal, .popup, .intercom-lightweight-app-launcher');
        if(pop) pop.remove();
    """)
    print("Popup removido.")
except:
    print("Nenhum popup para remover.")

# === IR PARA A TELA DE EXPORTAÇÃO ===
driver.get("https://www.bling.com.br/exportacao.produtos.php?criterio=opc-ultimos&pesquisa=&idTag=0&dataValidadeIni=&dataValidadeFim=&fornecedor=0&marca=&ncm=&dataAlteracaoIni=&dataAlteracaoFim=&codigoProdutoNoFabricante=&tipoProduto=T&idComponente=0&filtroTipoItem=&situacaoVinculoLoja=&idLojaVinculo=&filtroEstoque=&categoria=T")
time.sleep(4)

# === BAIXAR TODAS AS 7 PLANILHAS ===
links = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.action-link")))

print(f"Encontrados {len(links)} botões de download.")

for i, link in enumerate(links):
    try:
        print(f"⏬ Baixando arquivo {i+1} de {len(links)}...")
        driver.execute_script("arguments[0].click();", link)
        time.sleep(2)  # pequena espera entre os downloads
    except Exception as e:
        print(f"Erro ao baixar {i+1}: {e}")

print("🏁 Todos downloads disparados.")

time.sleep(10)
driver.quit()
