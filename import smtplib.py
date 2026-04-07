import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from getpass import getpass

# Configurações do servidor SMTP
SMTP_SERVER = 'smtp.gmail.com'  # Altere para seu servidor SMTP
SMTP_PORT = 587
EMAIL_FROM = 'seu_email@gmail.com'  # Seu e-mail remetente
EMAIL_PASSWORD = getpass("Digite a senha do e-mail remetente: ")  # Seguro - não armazena em texto plano

# Carrega os dados da planilha
planilha = pd.read_excel('dados_usuarios.xlsx')  # Supondo colunas: 'email', 'usuario', 'senha'

# Template do e-mail (use f-strings para as variáveis)
TEMPLATE_EMAIL = """
Olá,

Segue suas credenciais de acesso ao sistema:

Usuário: {usuario}
Senha: {senha}

Por segurança, recomendamos alterar esta senha no primeiro acesso.

Atenciosamente,
Equipe de Suporte
"""

def enviar_email(destinatario, usuario, senha):
    """Função para enviar e-mail individual"""
    try:
        # Prepara a mensagem
        msg = MIMEMultipart()
        msg['From'] = EMAIL_FROM
        msg['To'] = destinatario
        msg['Subject'] = "Suas credenciais de acesso"
        
        # Personaliza o corpo
        corpo_email = TEMPLATE_EMAIL.format(usuario=usuario, senha=senha)
        msg.attach(MIMEText(corpo_email, 'plain'))
        
        # Envia o e-mail
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.send_message(msg)
        
        print(f"E-mail enviado para {destinatario} com sucesso!")
        return True
    except Exception as e:
        print(f"Erro ao enviar para {destinatario}: {str(e)}")
        return False

# Processa cada linha da planilha
for index, row in planilha.iterrows():
    enviar_email(
        destinatario=row['email'],
        usuario=row['usuario'],
        senha=row['senha']
    )

print("Processo de envio concluído!")