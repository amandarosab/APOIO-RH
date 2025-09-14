import os.path
import base64
from email.mime.text import MIMEText
import tkinter as tk
from tkinter import messagebox

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Se modificar esses escopos, delete o arquivo token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send',
          'https://www.googleapis.com/auth/gmail.compose',
          'https://www.googleapis.com/auth/gmail.modify']

# --- Cores da Interface ---
COR_FUNDO = '#2E1A47'  # Um tom de roxo escuro
COR_TEXTO = '#FFFFFF'  # Branco
COR_BOTAO = '#4A2B7A'

def autenticar_gmail():
    """
    Realiza a autenticação com a API do Gmail usando OAuth 2.0.
    Persiste as credenciais em 'token.json'.
    """
    creds = None
    # O arquivo token.json armazena os tokens de acesso e de atualização do usuário.
    # Ele é criado automaticamente na primeira vez que o fluxo de autorização é concluído.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # Se não houver credenciais válidas disponíveis, permite que o usuário faça login.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Altere 'credentials.json' para o nome do seu arquivo de segredos do cliente OAuth 2.0.
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Salva as credenciais para a próxima execução
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def enviar_email_teste(service):
    """
    Cria e envia um e-mail de teste simples.
    Retorna True se enviado com sucesso, False caso contrário.
    """
    try:
        message = MIMEText('Este é um e-mail de teste automático do sistema de RH Quintessa.')
        message['to'] = 'teste@exemplo.com' # Mude para um e-mail de destino válido
        message['from'] = 'me@example.com' # Será substituído pelo seu e-mail autenticado ('me')
        message['subject'] = 'Teste de Conexão - Sistema RH Python'
        
        # Codifica a mensagem em base64url
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {'raw': encoded_message}

        # Chama a API do Gmail
        send_message = (service.users().messages().send(userId="me", body=create_message).execute())
        print(f'E-mail enviado com sucesso. Message Id: {send_message["id"]}')
        return True
    except HttpError as error:
        print(f'Ocorreu um erro ao enviar o e-mail: {error}')
        return False
    except Exception as e:
        print(f'Ocorreu um erro inesperado: {e}')
        return False

def main_action():
    """Função principal acionada pelo botão da GUI."""
    try:
        # 1. Autenticar
        creds = autenticar_gmail()
        
        # 2. Construir o serviço da API
        service = build('gmail', 'v1', credentials=creds)
        
        # 3. Enviar o e-mail
        if enviar_email_teste(service):
            messagebox.showinfo("Sucesso", "E-mail de teste enviado com sucesso!")
        else:
            messagebox.showerror("Erro", "Falha ao enviar o e-mail. Verifique o console para mais detalhes.")

    except Exception as e:
        messagebox.showerror("Erro de Autenticação", f"Não foi possível autenticar ou conectar à API do Gmail.\nErro: {e}")

# --- Configuração da Interface Gráfica (GUI) com Tkinter ---
def setup_gui():
    window = tk.Tk()
    window.title("Sistema de RH Quintessa - Autenticação")
    window.geometry("400x200")
    window.configure(bg=COR_FUNDO)

    label = tk.Label(
        window,
        text="Teste de Conexão com Gmail API",
        fg=COR_TEXTO,
        bg=COR_FUNDO,
        font=("Helvetica", 16)
    )
    label.pack(pady=20)

    send_button = tk.Button(
        window,
        text="Autenticar e Enviar E-mail de Teste",
        command=main_action,
        bg=COR_BOTAO,
        fg=COR_TEXTO,
        font=("Helvetica", 12),
        padx=10,
        pady=10
    )
    send_button.pack(pady=20)

    window.mainloop()

if __name__ == '__main__':
    setup_gui()