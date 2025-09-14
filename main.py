import os
import json
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import messagebox, font as tkfont

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- Configurações e Constantes ---
SCOPES = ['https://www.googleapis.com/auth/gmail.send',
          'https://www.googleapis.com/auth/gmail.compose',
          'https://www.googleapis.com/auth/gmail.modify']
TEMPLATES_FILE = 'templates.json'

# --- Cores e Fontes da Interface ---
COR_FUNDO = '#2E1A47'
COR_TEXTO = '#FFFFFF'
COR_BOTAO = '#4A2B7A'
COR_INPUT_FUNDO = '#3B2F5B'
FONTE_TITULO = ("Helvetica", 16, "bold")
FONTE_PADRAO = ("Helvetica", 12)
FONTE_BOTAO = ("Helvetica", 12, "bold")

# --- Funções de Autenticação e API (semelhante ao anterior) ---

def autenticar_gmail():
    """Realiza a autenticação com a API do Gmail e retorna as credenciais."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                messagebox.showerror("Erro de Autenticação", f"Não foi possível atualizar o token de acesso. Por favor, autorize novamente.\nErro: {e}")
                os.remove('token.json')
                return None
        else:
            try:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            except FileNotFoundError:
                messagebox.showerror("Erro Crítico", "Arquivo 'credentials.json' não encontrado. Verifique a pasta do projeto.")
                return None
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def enviar_email(service, para, assunto, corpo_html, cc=None, cco=None):
    """Cria e envia um e-mail formatado em HTML."""
    try:
        message = MIMEMultipart('alternative')
        message['to'] = para
        message['subject'] = assunto
        if cc:
            message['cc'] = cc
        if cco:
            message['bcc'] = cco
        
        message.attach(MIMEText(corpo_html, 'html'))
        
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {'raw': raw_message}
        
        send_message = service.users().messages().send(userId="me", body=create_message).execute()
        print(f'E-mail enviado. Message Id: {send_message["id"]}')
        return True
    except HttpError as error:
        messagebox.showerror("Erro na API", f'Ocorreu um erro ao enviar o e-mail: {error}')
        return False
    except Exception as e:
        messagebox.showerror("Erro Inesperado", f'Ocorreu um erro inesperado: {e}')
        return False

# --- Funções de Gerenciamento de Templates ---

def carregar_templates():
    """Carrega os templates do arquivo JSON ou cria um arquivo padrão."""
    if not os.path.exists(TEMPLATES_FILE):
        templates_padrao = {
            "convite_processo_seletivo": "Olá [NOME_CANDIDATO],\n\nObrigado pelo seu interesse no Quintessa...",
            "envio_case": "Olá [NOME_CANDIDATO],\n\nComo próxima etapa do nosso processo seletivo, segue o case para desenvolvimento...",
            "agendamento_entrevista_gg": "Olá [NOME_CANDIDATO],\n\nParabéns pela resolução do case! Gostaríamos de agendar uma entrevista com a equipe de Gente & Gestão...",
            "agendamento_entrevista_gg_lider": "Olá [NOME_CANDIDATO],\n\nParabéns! Gostaríamos de agendar a próxima entrevista com GG e o Líder Técnico...",
            "agendamento_entrevista_lider": "Olá [NOME_CANDIDATO],\n\nParabéns! Gostaríamos de agendar a próxima entrevista com o Líder Técnico...",
            "feedback_negativo_case": "Olá [NOME_CANDIDATO],\n\nAgradecemos sua participação. Neste momento, não seguiremos com sua candidatura...",
            "feedback_negativo_entrevista": "Olá [NOME_CANDIDATO],\n\nAgradecemos sua participação na entrevista. No momento, optamos por seguir com outros candidatos...",
            "proposta_trabalho": "Olá [NOME_CANDIDATO],\n\nTemos ótimas notícias! Estamos muito felizes em te oferecer a posição de... Bem-vindo(a) ao Quintessa!"
        }
        with open(TEMPLATES_FILE, 'w', encoding='utf-8') as f:
            json.dump(templates_padrao, f, ensure_ascii=False, indent=4)
        return templates_padrao
    else:
        try:
            with open(TEMPLATES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo '{TEMPLATES_FILE}'. Verifique se ele não está corrompido.")
            return {}

def salvar_template(chave, conteudo):
    """Salva um template específico no arquivo JSON."""
    templates = carregar_templates()
    templates[chave] = conteudo
    try:
        with open(TEMPLATES_FILE, 'w', encoding='utf-8') as f:
            json.dump(templates, f, ensure_ascii=False, indent=4)
        messagebox.showinfo("Sucesso", "Template salvo com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o template.\nErro: {e}")


# --- Estrutura Principal da Aplicação com Frames ---

class App(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        self.title("Sistema de RH Quintessa")
        self.geometry("800x700")
        self.configure(bg=COR_FUNDO)
        
        self.title_font = tkfont.Font(family='Helvetica', size=18, weight="bold")

        container = tk.Frame(self, bg=COR_FUNDO)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (MainMenuFrame, RecruitmentFrame, EmailTypeSelectionFrame, EmailEditorFrame):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("MainMenuFrame")

    def show_frame(self, page_name, **kwargs):
        frame = self.frames[page_name]
        if kwargs:
            frame.on_show(**kwargs) # Passa argumentos para o frame que será exibido
        frame.tkraise()

class MainMenuFrame(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=COR_FUNDO)
        self.controller = controller
        
        label = tk.Label(self, text="Menu Principal", font=FONTE_TITULO, bg=COR_FUNDO, fg=COR_TEXTO)
        label.pack(side="top", fill="x", pady=20)

        button_recrutamento = tk.Button(self, text="Recrutamento e Seleção", font=FONTE_BOTAO, bg=COR_BOTAO, fg=COR_TEXTO, command=lambda: controller.show_frame("RecruitmentFrame"))
        button_embarque = tk.Button(self, text="Embarque", font=FONTE_BOTAO, bg=COR_BOTAO, fg=COR_TEXTO, state="disabled")
        button_gestao = tk.Button(self, text="Gestão Operacional", font=FONTE_BOTAO, bg=COR_BOTAO, fg=COR_TEXTO, state="disabled")
        button_endomarketing = tk.Button(self, text="Endomarketing", font=FONTE_BOTAO, bg=COR_BOTAO, fg=COR_TEXTO, state="disabled")
        button_desembarque = tk.Button(self, text="Desembarque", font=FONTE_BOTAO, bg=COR_BOTAO, fg=COR_TEXTO, state="disabled")

        for btn in [button_recrutamento, button_embarque, button_gestao, button_endomarketing, button_desembarque]:
            btn.pack(pady=10, padx=50, fill="x")

class RecruitmentFrame(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=COR_FUNDO)
        self.controller = controller
        label = tk.Label(self, text="Recrutamento e Seleção", font=FONTE_TITULO, bg=COR_FUNDO, fg=COR_TEXTO)
        label.pack(side="top", fill="x", pady=20)

        button_enviar_email = tk.Button(self, text="ENVIAR E-MAIL", font=FONTE_BOTAO, bg=COR_BOTAO, fg=COR_TEXTO, command=lambda: controller.show_frame("EmailTypeSelectionFrame"))
        button_enviar_email.pack(pady=15)

        button_voltar = tk.Button(self, text="Voltar ao Menu", font=FONTE_BOTAO, bg=COR_BOTAO, fg=COR_TEXTO, command=lambda: controller.show_frame("MainMenuFrame"))
        button_voltar.pack(pady=15)

class EmailTypeSelectionFrame(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=COR_FUNDO)
        self.controller = controller
        label = tk.Label(self, text="Selecione o Tipo de E-mail", font=FONTE_TITULO, bg=COR_FUNDO, fg=COR_TEXTO)
        label.pack(side="top", fill="x", pady=20)

        email_types = {
            "E-mail de convite para processo seletivo": "convite_processo_seletivo",
            "E-mail de confecção de case": "envio_case",
            "Retorno de case e agendamento (GG)": "agendamento_entrevista_gg",
            "Retorno de case e agendamento (GG e Líder)": "agendamento_entrevista_gg_lider",
            "Retorno de case e agendamento (Líder)": "agendamento_entrevista_lider",
            "Feedback negativo pós case": "feedback_negativo_case",
            "Feedback negativo pós entrevista": "feedback_negativo_entrevista",
            "Feedback positivo: Proposta de trabalho": "proposta_trabalho"
        }

        for display_text, key in email_types.items():
            button = tk.Button(self, text=display_text, font=FONTE_PADRAO, bg=COR_BOTAO, fg=COR_TEXTO,
                               command=lambda k=key, t=display_text: controller.show_frame("EmailEditorFrame", template_key=k, template_title=t))
            button.pack(pady=5, padx=50, fill="x")

        button_voltar = tk.Button(self, text="Voltar", font=FONTE_BOTAO, bg=COR_BOTAO, fg=COR_TEXTO, command=lambda: controller.show_frame("RecruitmentFrame"))
        button_voltar.pack(pady=20)

class EmailEditorFrame(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=COR_FUNDO)
        self.controller = controller
        self.template_key = None
        self.template_title = None
        self.templates = carregar_templates()

        # --- Widgets da tela ---
        self.title_label = tk.Label(self, text="", font=FONTE_TITULO, bg=COR_FUNDO, fg=COR_TEXTO)
        self.title_label.pack(pady=10)

        # Template Text Box
        self.text_editor = tk.Text(self, height=15, width=80, wrap=tk.WORD, bg=COR_INPUT_FUNDO, fg=COR_TEXTO, insertbackground=COR_TEXTO, font=("Consolas", 11))
        self.text_editor.pack(pady=5, padx=20, fill="both", expand=True)

        # Botão Salvar Template
        self.save_button = tk.Button(self, text="Salvar Template", font=FONTE_PADRAO, bg=COR_BOTAO, fg=COR_TEXTO, command=self.save_current_template)
        self.save_button.pack(pady=5)

        # Frame para os inputs
        input_frame = tk.Frame(self, bg=COR_FUNDO)
        input_frame.pack(pady=10, padx=20, fill="x")

        # Labels e Entradas
        tk.Label(input_frame, text="Nome do destinatário*:", font=FONTE_PADRAO, bg=COR_FUNDO, fg=COR_TEXTO).grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.entry_nome = tk.Entry(input_frame, width=50, bg=COR_INPUT_FUNDO, fg=COR_TEXTO, insertbackground=COR_TEXTO)
        self.entry_nome.grid(row=0, column=1, sticky='ew')

        tk.Label(input_frame, text="E-mail do destinatário*:", font=FONTE_PADRAO, bg=COR_FUNDO, fg=COR_TEXTO).grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.entry_email = tk.Entry(input_frame, width=50, bg=COR_INPUT_FUNDO, fg=COR_TEXTO, insertbackground=COR_TEXTO)
        self.entry_email.grid(row=1, column=1, sticky='ew')

        tk.Label(input_frame, text="Pessoas em cópia (Cc):", font=FONTE_PADRAO, bg=COR_FUNDO, fg=COR_TEXTO).grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.entry_cc = tk.Entry(input_frame, width=50, bg=COR_INPUT_FUNDO, fg=COR_TEXTO, insertbackground=COR_TEXTO)
        self.entry_cc.grid(row=2, column=1, sticky='ew')

        tk.Label(input_frame, text="Pessoas em cópia oculta (Bcc):", font=FONTE_PADRAO, bg=COR_FUNDO, fg=COR_TEXTO).grid(row=3, column=0, sticky='w', padx=5, pady=2)
        self.entry_cco = tk.Entry(input_frame, width=50, bg=COR_INPUT_FUNDO, fg=COR_TEXTO, insertbackground=COR_TEXTO)
        self.entry_cco.grid(row=3, column=1, sticky='ew')
        
        input_frame.grid_columnconfigure(1, weight=1)

        # Botões de Ação
        action_frame = tk.Frame(self, bg=COR_FUNDO)
        action_frame.pack(pady=10)
        self.send_button = tk.Button(action_frame, text="ENVIAR", font=FONTE_BOTAO, bg='green', fg=COR_TEXTO, command=self.send_final_email)
        self.send_button.pack(side="left", padx=10)
        self.back_button = tk.Button(action_frame, text="Voltar", font=FONTE_BOTAO, bg=COR_BOTAO, fg=COR_TEXTO, command=lambda: controller.show_frame("EmailTypeSelectionFrame"))
        self.back_button.pack(side="left", padx=10)

    def on_show(self, template_key, template_title):
        """Chamado quando o frame é exibido."""
        self.template_key = template_key
        self.template_title = template_title
        self.title_label.config(text=self.template_title)
        
        # Limpa os campos
        self.text_editor.delete('1.0', tk.END)
        self.entry_nome.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_cc.delete(0, tk.END)
        self.entry_cco.delete(0, tk.END)
        
        # Carrega o template
        self.templates = carregar_templates()
        content = self.templates.get(self.template_key, "Template não encontrado.")
        self.text_editor.insert('1.0', content)

    def save_current_template(self):
        """Salva o conteúdo atual do editor de texto no arquivo JSON."""
        if not self.template_key:
            messagebox.showerror("Erro", "Nenhum template selecionado para salvar.")
            return
        content = self.text_editor.get("1.0", tk.END).strip()
        salvar_template(self.template_key, content)

    def send_final_email(self):
        """Valida os campos e envia o e-mail."""
        nome = self.entry_nome.get().strip()
        email_para = self.entry_email.get().strip()
        cc = self.entry_cc.get().strip()
        cco = self.entry_cco.get().strip()
        corpo = self.text_editor.get("1.0", tk.END)
        assunto = self.template_title

        # Validação
        if not nome or not email_para:
            messagebox.showwarning("Campos Obrigatórios", "Os campos 'Nome do destinatário' e 'E-mail do destinatário' são obrigatórios.")
            return

        # Substituição de tags
        corpo_final = corpo.replace('[NOME_CANDIDATO]', nome).replace('\n', '<br>')
        
        confirm = messagebox.askyesno("Confirmar Envio", f"Você confirma o envio do e-mail para {email_para}?")
        if not confirm:
            return

        try:
            creds = autenticar_gmail()
            if not creds:
                # A função de autenticação já deve ter mostrado um erro
                return

            service = build('gmail', 'v1', credentials=creds)
            if enviar_email(service, email_para, assunto, corpo_final, cc, cco):
                messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")
                self.controller.show_frame("EmailTypeSelectionFrame") # Volta para a tela anterior
            # A função enviar_email já mostra a mensagem de erro
        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado no processo de envio: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()