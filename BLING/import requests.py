import requests
import pandas as pd
import time

# ===============================
# CONFIGURAÇÕES
# ===============================
API_TOKEN = "3c9e38a51dff87cba4951cb83953ce164d7b5644f9d8a01238d1d3f9c0e5023b0b61e689"
BASE_URL = "https://www.bling.com.br/Api/v3/produtos"

# Cabeçalho obrigatório da API v3
HEADERS = {
    "Authorization": f"Bearer {API_TOKEN}",
    "Accept": "application/json"
}

# ===============================
# FUNÇÃO PARA BUSCAR UMA PÁGINA
# ===============================
def buscar_pagina(pagina):
    params = {
        "pagina": pagina,        # define o número da página
        "limite": 100            # máximo permitido (100 itens por página)
    }
    
    response = requests.get(BASE_URL, headers=HEADERS, params=params)
    
    if response.status_code == 200:
        return response.json()
    else:
        print(f"❌ Erro ao buscar página {pagina}: {response.status_code}")
        print(response.text)
        return None

# ===============================
# LOOP PARA EXPORTAR 14 PÁGINAS
# ===============================
for pagina in range(1, 15):  # páginas 1 a 14
    print(f"🔄 Baixando página {pagina}...")
    
    dados = buscar_pagina(pagina)
    if not dados:
        continue

    produtos = dados.get("data", [])

    # Se não tiver produtos, parar o processo
    if len(produtos) == 0:
        print(f"⚠️ Página {pagina} sem produtos. Encerrando extração.")
        break

    # Converte para DataFrame
    df = pd.json_normalize(produtos)

    # Salvar CSV
    nome_arquivo = f"produtos_pagina_{pagina}.csv"
    df.to_csv(nome_arquivo, index=False, encoding="utf-8-sig")

    print(f"✅ Página {pagina} salva como: {nome_arquivo}")
    
    # Evita bloqueio por excesso de requisições
    time.sleep(1)

print("🏁 Processo concluído com sucesso! Todas as páginas foram exportadas.")
