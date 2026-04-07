from openpyxl import Workbook, load_workbook
import requests
from bs4 import BeautifulSoup
import sqlite3
import time
import hashlib
import os

# =========================
# CONFIG
# =========================
#TOKEN = "8643526470:AAEqdTHnTz1YNRP7aSntyS93ULbTnChkMjA"
#CHAT_ID = "1169857949"

TOKEN = os.getenv("8643526470:AAEqdTHnTz1YNRP7aSntyS93ULbTnChkMjA")
CHAT_ID = os.getenv("1169857949")

# =========================
# BANCO
# =========================
conn = sqlite3.connect("corridas.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS eventos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    hash TEXT UNIQUE,
    titulo TEXT,
    link TEXT,
    score INTEGER
)
""")

# =========================
# REGRAS
# =========================
PALAVRAS_DF = ["brasília", "df", "taguatinga", "ceilândia"]
PALAVRAS_GRATUITO = ["gratuito", "gratuita", "free", "sem custo"]
PALAVRAS_ORGAOS = ["gdf", "sesc", "sesi"]

# =========================
# FUNÇÕES
# =========================

def calcular_score(texto):
    t = texto.lower()
    score = 0

    for p in PALAVRAS_GRATUITO:
        if p in t:
            score += 100

    for p in PALAVRAS_ORGAOS:
        if p in t:
            score += 60

    if "r$" not in t:
        score += 20
    else:
        score -= 100

    return score


def eh_df(texto):
    t = texto.lower()
    return any(p in t for p in PALAVRAS_DF)


def gerar_hash(titulo, link):
    return hashlib.md5((titulo + link).encode()).hexdigest()


def enviar_telegram(msg):
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    requests.post(url, data={
        "chat_id": CHAT_ID,
        "text": msg,
        "parse_mode": "Markdown"
    })


# =========================
# SCRAPER
# =========================

def buscar_eventos():
    url = "https://www.corridasbr.com.br/df/calendario.asp"
    eventos = []

    try:
        response = requests.get(url, timeout=10)
        soup = BeautifulSoup(response.text, "html.parser")

        for item in soup.find_all("a"):
            titulo = item.text.strip()
            link = item.get("href")

            if titulo and link:
                eventos.append((titulo, link))

    except Exception as e:
        print("Erro:", e)

    return eventos


# =========================
# EXCEL
# =========================

def salvar_excel(titulo, link, score):
    arquivo = "corridas.xlsx"

    if not os.path.exists(arquivo):
        wb = Workbook()
        ws = wb.active
        ws.append(["Título", "Link", "Score"])
        wb.save(arquivo)

    wb = load_workbook(arquivo)
    ws = wb.active

    ws.append([titulo, link, score])
    wb.save(arquivo)


# =========================
# PROCESSAMENTO (TESTE)
# =========================

def processar():
    print("🔥 TESTE TELEGRAM NUVEM")

    enviar_telegram("🚀 FUNCIONANDO NA NUVEM")

# =========================
# LOOP
# =========================

while True:
    print("🔍 Rodando...")
    processar()
    time.sleep(600)