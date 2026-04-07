"""
bling_automacao_otimizado_v4.py
Automação Selenium (versão otimizada) para:
 - login no Bling
 - abrir a página de exportação de produtos
 - clicar no primeiro "Download" e então usar TAB+ENTER para acionar os demais
 - esperar confiavelmente os downloads (12s entre acionamentos)
 - processar os CSVs: para NCM == "7113.19.00" e Observações > 0 -> Preço = Observações * 1400
 - não multiplicar quando Observações == 0 ou não-numérico
 - formatar Preço com ponto de milhar e vírgula decimal (ex: 7.700,00)
 - salvar processed_*.csv em subpasta "processed" e gerar um CSV unificado "processed_all.csv"
"""

import os
import time
import glob
from pathlib import Path
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
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

# =========================
# CONFIGURAÇÃO (AJUSTE AQUI)
# =========================
EMAIL = "09839356100"
SENHA = "@%Gabriel99"

# Pasta de downloads e processed
DOWNLOAD_DIR = str(Path.home() / "Downloads" / "bling_exports")
PROCESSED_DIR = os.path.join(DOWNLOAD_DIR, "processed")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(PROCESSED_DIR, exist_ok=True)

# URL direta da tela de exportação (você informou que funciona)
EXPORT_PAGE_URL = ("https://www.bling.com.br/exportacao.produtos.php?criterio=opc-ultimos"
                   "&pesquisa=&idTag=0&dataValidadeIni=&dataValidadeFim=&fornecedor=0&marca="
                   "&ncm=&dataAlteracaoIni=&dataAlteracaoFim=&codigoProdutoNoFabricante="
                   "&tipoProduto=T&idComponente=0&filtroTipoItem=&situacaoVinculoLoja="
                   "&idLojaVinculo=&filtroEstoque=&categoria=T")

# Multiplicador e NCM alvo
MULTIPLICADOR = 1400.0
NCM_ALVO = "7113.19.00"

# Tempo entre acionamentos (10-12s recomendado)
CLICK_DELAY_SECONDS = 12

# Timeout para downloads (segundos)
DOWNLOAD_WAIT_TIMEOUT = 240

# =========================
# OPÇÕES DO CHROME
# =========================
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "profile.default_content_setting_values.notifications": 2,
}
chrome_options.add_experimental_option("prefs", prefs)
# chrome_options.add_argument("--start-maximized")  # útil para debug
# chrome_options.add_argument("--headless=new")    # habilite se quiser rodar sem janela (poderá mudar comportamentos)

# =========================
# INICIALIZAÇÃO DO DRIVER
# =========================
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 20)
actions = ActionChains(driver)

def safe_sleep(t):
    time.sleep(t)

def remove_overlays_js():
    """Remove overlays e elementos que possam bloquear clique."""
    js = """
    try {
      document.querySelectorAll('.modal, .modal-backdrop, .overlay, .ui-widget-overlay, .fancybox-overlay, .lightbox-overlay').forEach(el => el.remove());
      document.querySelectorAll('[style*=\"z-index\"]').forEach(el => {
        try { el.style.zIndex = '0'; } catch(e) {}
      });
    } catch(e){}
    """
    try:
        driver.execute_script(js)
    except Exception:
        pass

def try_close_popups():
    """Tenta fechar botões 'X' e similares."""
    xpaths = [
        "//button[contains(@class,'close')]",
        "//button[contains(text(),'×') or contains(text(),'X') or contains(text(),'Fechar') or contains(text(),'fechar')]",
        "//a[contains(@class,'close') or contains(text(),'×') or contains(text(),'Fechar')]",
        "//div[contains(@class,'modal')]//button[contains(.,'Fechar')]",
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
                        safe_sleep(0.8)
                except Exception:
                    pass
        except Exception:
            pass

def wait_for_no_crdownloads(timeout=DOWNLOAD_WAIT_TIMEOUT):
    """Aguarda até que não haja .crdownload na pasta de download (ou timeout)."""
    start = time.time()
    while True:
        cr = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
        if len(cr) == 0:
            return True
        if time.time() - start > timeout:
            return False
        time.sleep(1)

def wait_for_new_csvs(before_set, timeout=DOWNLOAD_WAIT_TIMEOUT):
    """Aguarda até que novos CSVs apareçam ou timeout. Retorna lista de novos caminhos."""
    start = time.time()
    while True:
        current = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
        new = current - before_set
        if new:
            # esperar evitar .crdownload finais
            time.sleep(1)
            if wait_for_no_crdownloads(timeout=10):
                return sorted(new)
        if time.time() - start > timeout:
            return sorted(new)

def format_brazilian_number(value_float):
    """Formata float para string: 1.234,56 (ponto de milhar e vírgula decimal)."""
    s = f"{value_float:,.2f}"            # ex: "7,700.00"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s

def parse_number_from_str(s):
    """Converte strings tipo '5,5' ou '5.5' ou '1.234,56' para float; retorna None se inválido."""
    if s is None:
        return None
    st = str(s).strip()
    if st == "":
        return None
    # remover espaços
    st = st.replace(" ", "")
    # se contem letras, considerar inválido (mas permitir vírgula e ponto)
    allowed = set("0123456789.,-")
    if any(ch not in allowed for ch in st):
        return None
    # heurística: se houver mais de one comma and one dot, remove thousands
    # tratar milhar: se '.' aparece e ',' aparece -> assume '.' milhar, ',' decimal
    try:
        if st.count(",") > 0 and st.count(".") > 0:
            st = st.replace(".", "")       # remover ponto de milhar
            st = st.replace(",", ".")      # converter decimal para ponto
        else:
            # só vírgula => trocar por ponto
            if st.count(",") > 0 and st.count(".") == 0:
                st = st.replace(",", ".")
            # só ponto => assume ponto decimal (ok)
        return float(st)
    except Exception:
        return None

# =========================
# FLUXO PRINCIPAL
# =========================
try:
    print("🔐 Abrindo Bling (login)...")
    driver.get("https://www.bling.com.br/login")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    safe_sleep(1)

    # localizar campo username
    try:
        campo_username = wait.until(EC.presence_of_element_located((By.ID, "username")))
    except TimeoutException:
        try:
            campo_username = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='usuário' i], input[placeholder*='e-mail' i]")
        except Exception:
            campo_username = None

    if not campo_username:
        raise Exception("Campo username não encontrado automaticamente. Ajuste o seletor.")

    campo_username.clear()
    campo_username.send_keys(EMAIL)
    safe_sleep(0.3)

    # localizar campo senha (pode ser input text com placeholder)
    try:
        campo_senha = driver.find_element(By.CSS_SELECTOR, "input[placeholder*='senha' i], input[type='password']")
    except Exception:
        campo_senha = None

    if not campo_senha:
        raise Exception("Campo senha não encontrado automaticamente. Ajuste o seletor.")

    campo_senha.clear()
    campo_senha.send_keys(SENHA)
    safe_sleep(0.4)

    # localizar botão de login: tentativas por id, button text, input[type=submit]
    botao_entrar = None
    try:
        botao_entrar = driver.find_element(By.ID, "login-submit")
    except Exception:
        pass
    if not botao_entrar:
        try:
            botao_entrar = driver.find_element(By.XPATH, "//button[contains(., 'Entrar') or contains(., 'Login') or contains(., 'Entrar no Bling')]")
        except Exception:
            pass
    if not botao_entrar:
        try:
            botao_entrar = driver.find_element(By.CSS_SELECTOR, "input[type=submit], button[type=submit]")
        except Exception:
            pass

    if not botao_entrar:
        raise Exception("Botão de login não encontrado automaticamente. Ajuste o seletor.")

    # clicar no botão de login (via JS para robustez)
    try:
        driver.execute_script("arguments[0].click();", botao_entrar)
    except Exception:
        botao_entrar.click()

    print("⏳ Fazendo login, aguardando dashboard...")
    # aguardar URL mudar e dashboard carregar
    time_limit = 20
    start = time.time()
    while time.time() - start < time_limit:
        if "bling.com.br" in driver.current_url and driver.current_url != "https://www.bling.com.br/login":
            break
        time.sleep(0.5)
    safe_sleep(3)
    try_close_popups()
    remove_overlays_js()
    safe_sleep(1)

    # ir para página de exportação direta
    print("📦 Acessando página de exportação de produtos...")
    driver.get(EXPORT_PAGE_URL)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    safe_sleep(2)
    try_close_popups()
    remove_overlays_js()
    safe_sleep(1)

    # localizar os links "Download"
    print("🔎 Procurando 'Download' links na página...")
    # esperamos aparecer pelo menos um link
    # (usa 'a.action-link' e texto 'Download')
    time.sleep(1)
    anchors = driver.find_elements(By.CSS_SELECTOR, "a.action-link")
    download_anchors = [a for a in anchors if "download" in a.text.strip().lower()]
    if not download_anchors:
        # fallback por xpath
        download_anchors = driver.find_elements(By.XPATH, "//a[contains(translate(., 'DOWNLOAD', 'download'), 'download')]")
    if not download_anchors:
        raise Exception("Nenhum link 'Download' encontrado. Certifique-se de estar na página correta e logado.")

    total = len(download_anchors)
    print(f"🔘 {total} links 'Download' encontrados.")

    # limpar CSVs antigos
    old_csvs = glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv"))
    for f in old_csvs:
        try:
            os.remove(f)
        except Exception:
            pass

    # --- Estratégia de clique:
    # 1) clicar no primeiro link via JS (garante iniciar o 1º download)
    # 2) em seguida, enviar TAB+ENTER repetidamente para acionar os demais (cada TAB foca o próximo botão)
    #    (VOCÊ confirmou que TAB pula exclusivamente entre links Download)
    before_csvs = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
    first = download_anchors[0]
    # garantir visibilidade e remover overlays
    try:
        driver.execute_script("arguments[0].scrollIntoView(true);", first)
        remove_overlays_js()
        safe_sleep(0.4)
        driver.execute_script("arguments[0].click();", first)
    except Exception:
        try:
            first.click()
        except Exception as ex:
            raise Exception("Falha ao clicar primeiro download: " + str(ex))

    print("➡️ Iniciando download 1/{} ...".format(total))
    # aguardar detecção do 1º arquivo (com timeout curto)
    new1 = wait_for_new_csvs(before_csvs, timeout=30)
    if new1:
        print("    ✅ Novo arquivo detectado:", os.path.basename(new1[0]))
    else:
        print("    ⚠️ Não detectei novo CSV imediatamente após 1º clique. Continuarei mesmo assim.")

    # agora usar TAB + ENTER para os remanescentes
    body = driver.find_element(By.TAG_NAME, "body")
    for idx in range(1, total):
        print(f"➡️ Iniciando download {idx+1}/{total} ... (TAB -> ENTER)")
        # enviar TAB para focar próximo botão (assumindo fluxo correto)
        actions = ActionChains(driver)
        actions.send_keys(Keys.TAB)
        actions.perform()
        safe_sleep(0.2)
        # dar ENTER
        actions = ActionChains(driver)
        actions.send_keys(Keys.ENTER)
        actions.perform()

        # esperar o intervalo e tentativa de detectar novo arquivo
        # coletar CSVs antes e aguardar
        before = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
        # aguardar o delay entre clicks para o Bling processar e salvar
        safe_sleep(CLICK_DELAY_SECONDS)
        new = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv"))) - before
        if new:
            fname = sorted(new)[0]
            print("    ✅ Novo arquivo detectado:", os.path.basename(fname))
        else:
            # tentar esperar um pouco mais
            more = wait_for_new_csvs(before, timeout=30)
            if more:
                print("    ✅ Novo arquivo detectado (após espera extra):", os.path.basename(more[0]))
            else:
                print("    ⚠️ Novo arquivo não detectado para o download", idx+1)

    # depois de todos os cliques, aguardar finalização de possíveis .crdownload
    print("🔁 Cliques finalizados. Aguardando término de quaisquer .crdownload remanescentes...")
    ok_no_cr = wait_for_no_crdownloads(timeout=DOWNLOAD_WAIT_TIMEOUT)
    if not ok_no_cr:
        print("⚠️ Timeout esperando arquivos terminarem de baixar. Verifique manualmente na pasta:", DOWNLOAD_DIR)
    else:
        print("✅ Downloads finalizados (sem .crdownload).")

    # listar CSVs encontrados
    csvs = sorted(glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv")))
    new_csvs = [c for c in csvs if os.path.basename(c).startswith("produtos_")]
    if not new_csvs:
        print("⚠️ Nenhum CSV novo encontrado para processar.")
    else:
        print(f"🔘 Encontrados {len(new_csvs)} novos arquivos CSV: ")
        for p in new_csvs:
            print("   ", os.path.basename(p))

    # =========================
    # TRATAMENTO AUTOMÁTICO DOS CSVs
    # =========================
    print("🛠️ Iniciando tratamento automático das planilhas baixadas...")
    processed_paths = []
    for csv_path in new_csvs:
        try:
            print("🔎 Processando:", os.path.basename(csv_path))
            # tentativas de leitura robusta
            df = None
            read_success = False
            for enc in ("utf-8-sig", "utf-8", "latin-1"):
                try:
                    # geralmente estes CSVs usam ';' como separador. tentamos com sep=';'
                    df = pd.read_csv(csv_path, sep=';', dtype=str, encoding=enc, engine='python', on_bad_lines='warn')
                    read_success = True
                    break
                except Exception:
                    pass
            if not read_success:
                # tentativa sem forçar separador
                df = pd.read_csv(csv_path, dtype=str, encoding='latin-1', engine='python', on_bad_lines='warn')

            # limpar BOM no nome da primeira coluna se houver
            df.columns = [c.strip().lstrip('\ufeff').lstrip('\ufeff').strip() for c in df.columns]

            # verificar colunas essenciais
            if "NCM" not in df.columns or "Observações" not in df.columns or "Preço" not in df.columns:
                print("    ⚠️ Colunas esperadas ausentes. Salvando cópia sem alteração.")
                out_path = os.path.join(PROCESSED_DIR, "copied_" + os.path.basename(csv_path))
                df.to_csv(out_path, index=False, sep=';', encoding='utf-8-sig')
                continue

            # iterar e aplicar regra
            updated = 0
            for idx in df.index:
                try:
                    ncm_val = str(df.at[idx, "NCM"]).strip()
                except Exception:
                    ncm_val = ""
                if ncm_val != NCM_ALVO:
                    continue

                obs_raw = df.at[idx, "Observações"]
                num = parse_number_from_str(obs_raw)
                # se Observações == 0 ou não numérico -> não multiplicar
                if num is None:
                    continue
                if abs(num) < 1e-9:  # igual a zero
                    continue

                # multiplicar
                new_price = num * MULTIPLICADOR
                # formatar com ponto de milhar e vírgula decimal
                formatted = format_brazilian_number(new_price)
                # salvar string formatada no campo 'Preço'
                df.at[idx, "Preço"] = formatted
                updated += 1

            out_name = "processed_" + os.path.basename(csv_path)
            out_path = os.path.join(PROCESSED_DIR, out_name)
            # salvar com ';' separador e encoding compatível
            df.to_csv(out_path, index=False, sep=';', encoding='utf-8-sig')
            processed_paths.append(out_path)
            print(f"    ✅ Salvo {os.path.basename(out_path)} — registros atualizados: {updated}")

        except Exception as e:
            print("    ❌ Erro ao processar", os.path.basename(csv_path), ":", e)

    # unir todos processed_*.csv em um único processed_all.csv
    if processed_paths:
        print("🔗 Unindo processed CSVs em único arquivo...")
        dfs = []
        for p in processed_paths:
            try:
                d = pd.read_csv(p, sep=';', dtype=str, encoding='utf-8-sig', engine='python', on_bad_lines='warn')
                dfs.append(d)
            except Exception:
                try:
                    d = pd.read_csv(p, sep=';', dtype=str, encoding='latin-1', engine='python', on_bad_lines='warn')
                    dfs.append(d)
                except Exception:
                    pass
        if dfs:
            combined = pd.concat(dfs, ignore_index=True, sort=False)
            combined_path = os.path.join(PROCESSED_DIR, "processed_all.csv")
            combined.to_csv(combined_path, index=False, sep=';', encoding='utf-8-sig')
            print("✅ Arquivo unificado salvo em:", combined_path)
        else:
            print("⚠️ Nenhuma tabela válida para concatenar.")
    else:
        print("⚠️ Nenhum arquivo processado para unir.")

    print("🏁 Todos os processos finalizados. Verifique as pastas:")
    print(" - downloads:", DOWNLOAD_DIR)
    print(" - processed:", PROCESSED_DIR)

except Exception as e:
    print("❌ Erro durante o processo:\n", e)

finally:
    print("🔚 Finalizando driver em 6s...")
    time.sleep(6)
    try:
        driver.quit()
    except Exception:
        pass
