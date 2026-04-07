from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time

# ----------------------------------------------------
# CONFIGURAÇÕES
# ----------------------------------------------------
EMAIL = "seu_email_aqui"
SENHA = "sua_senha_aqui"
URL_LOGIN = "https://www.bling.com.br/login"


# ----------------------------------------------------
# INICIAR NAVEGADOR
# ----------------------------------------------------
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()
wait = WebDriverWait(driver, 12)

print("🔐 Acessando Bling...")
driver.get(URL_LOGIN)

try:
    # ----------------------------------------------------
    # LOGIN
    # ----------------------------------------------------
    campo_email = wait.until(EC.presence_of_element_located((By.ID, "username")))
    campo_email.send_keys(EMAIL)

    campo_senha = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//input[@placeholder='Insira sua senha']")
    ))
    campo_senha.send_keys(SENHA)

    campo_senha.send_keys(Keys.ENTER)
    print("✅ Login realizado, aguardando carregamento...")

    time.sleep(5)

    # ----------------------------------------------------
    # FECHAR POP-UP DE PROPAGANDA DO BLING
    # ----------------------------------------------------
    print("🟡 Verificando se há pop-up...")

    popup_fechado = False

    # Tentativa 1: Ícone X padrão
    try:
        fechar1 = driver.find_element(By.XPATH, "//button[contains(., 'Fechar')]")
        fechar1.click()
        popup_fechado = True
        print("🟢 Pop-up fechado (botão 'Fechar').")
    except:
        pass

    # Tentativa 2: botão X com aria-label
    if not popup_fechado:
        try:
            fechar2 = driver.find_element(By.XPATH, "//button[@aria-label='Fechar']")
            fechar2.click()
            popup_fechado = True
            print("🟢 Pop-up fechado (aria-label='Fechar').")
        except:
            pass

    # Tentativa 3: qualquer ícone X dentro de pop-up
    if not popup_fechado:
        try:
            fechar3 = driver.find_element(By.XPATH, "//span[contains(text(), '×') or contains(text(),'x')]")
            fechar3.click()
            popup_fechado = True
            print("🟢 Pop-up fechado (ícone X).")
        except:
            pass

    if not popup_fechado:
        print("⚠️ Nenhum pop-up encontrado ou impossível de fechar.")
    else:
        time.sleep(2)

    print("🚀 Prosseguindo automação...")

    # ---------------------------------------------
    # EXEMPLO: abrir menu Produtos
    # ---------------------------------------------
    produtos_btn = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//a[contains(@href, 'produtos')]")
    ))
    produtos_btn.click()

    print("🎉 Automação funcionando sem bloqueios!")

except Exception as e:
    print("❌ Erro durante o processo:")
    print(e)

finally:
    print("⚙️ Processo finalizado.")
