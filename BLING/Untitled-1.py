print("🔐 Acessando Bling...")

# Campo de email/usuário
campo_usuario = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "username"))
)
campo_usuario.send_keys("SEU_EMAIL_AQUI")

# Botão "Continuar"
botao_continuar = driver.find_element(By.XPATH, "//button[contains(., 'Continuar')]")
botao_continuar.click()

# Agora esperar o campo de senha aparecer
campo_senha = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "password"))
)
campo_senha.send_keys("SUA_SENHA_AQUI")

# Botão entrar
botao_entrar = driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]")
botao_entrar.click()
