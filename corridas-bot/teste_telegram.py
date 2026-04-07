import requests

TOKEN = "8643526470:AAEqdTHnTz1YNRP7aSntyS93ULbTnChkMjA"
CHAT_ID = "1169857949"

url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"

response = requests.post(url, data={
    "chat_id": CHAT_ID,
    "text": "🚀 Agora vai!"
})

print(response.text)