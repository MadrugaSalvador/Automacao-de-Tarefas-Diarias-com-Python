# --- Bloco de Verificação e Instalação de Requisitos ---
try:
    import customtkinter as ctk
except ImportError:
    print("CustomTkinter não encontrado. Tentando instalar...")
    try:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "customtkinter"])
        import customtkinter as ctk
        print("CustomTkinter instalado com sucesso!")
    except Exception as e:
        print(f"Erro ao instalar CustomTkinter: {e}")
        print("Por favor, instale CustomTkinter manualmente: pip install customtkinter")
        sys.exit(1) # Sai do programa se a instalação falhar

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import customtkinter as ctk 
from tkinter import filedialog, messagebox
import json
import datetime

# --- CUSTOMTKINTER THEME SETTINGS ---
ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light")
ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "green", "dark-blue")
# --- END CUSTOMTKINTER THEME SETTINGS ---

# --- Caminhos dos arquivos para salvar configurações ---
DEFAULT_MESSAGE_FILE = "default_email_message.txt"
SENDER_EMAIL_FILE = "sender_email.txt"
RECURRING_CONFIG_FILE = "recurring_email_config.json"
APP_PASSWORD_FILE = "app_password.txt" # Arquivo para a senha de app
CC_EMAIL_FILE = "cc_email.txt" # NOVO: Arquivo para o email de cópia

# --- CONFIGURAÇÕES PADRÃO PARA RELATÓRIO DIÁRIO ---
# Estes serão usados APENAS se não houver um arquivo manual selecionado
# e nenhuma configuração de recorrência salva anteriormente.
DEFAULT_RECURRING_FOLDER = os.path.join(os.path.expanduser("~"), "RelatoriosDiarios") 
DEFAULT_RECURRING_FILE_NAME = "relatorio_gerencial_diario.xlsx"
# --- FIM DAS CONFIGURAÇÕES PADRÃO ---

# --- Funções para Carregar/Salvar Configurações ---
def load_default_message():
    """Carrega a mensagem padrão do arquivo ou retorna uma mensagem padrão."""
    if os.path.exists(DEFAULT_MESSAGE_FILE):
        with open(DEFAULT_MESSAGE_FILE, "r", encoding="utf-8") as f:
            return f.read()
    return """Prezado(a),

Segue em anexo o documento/planilha solicitado(a).

Atenciosamente,
Seu Nome
"""

def save_default_message(message):
    """Salva a mensagem atual como padrão no arquivo."""
    with open(DEFAULT_MESSAGE_FILE, "w", encoding="utf-8") as f:
        f.write(message)

def load_sender_email():
    """Carrega o e-mail do remetente do arquivo ou retorna uma string vazia."""
    if os.path.exists(SENDER_EMAIL_FILE):
        with open(SENDER_EMAIL_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "seu_email@gmail.com" # Valor padrão se não houver arquivo salvo

def save_sender_email(email):
    """Salva o e-mail do remetente no arquivo."""
    with open(SENDER_EMAIL_FILE, "w", encoding="utf-8") as f:
        f.write(email.strip())

def load_app_password():
    """Carrega a senha de app do arquivo."""
    if os.path.exists(APP_PASSWORD_FILE):
        with open(APP_PASSWORD_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return ""

def save_app_password(password):
    """Salva a senha de app no arquivo."""
    with open(APP_PASSWORD_FILE, "w", encoding="utf-8") as f:
        f.write(password.strip())

def load_cc_email():
    """Carrega o email de cópia (CC) do arquivo."""
    if os.path.exists(CC_EMAIL_FILE):
        with open(CC_EMAIL_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "" # Valor padrão vazio

def save_cc_email(email):
    """Salva o email de cópia (CC) no arquivo."""
    with open(CC_EMAIL_FILE, "w", encoding="utf-8") as f:
        f.write(email.strip())

def load_recurring_config():
    """Carrega as configurações de e-mail recorrente do arquivo."""
    if os.path.exists(RECURRING_CONFIG_FILE):
        with open(RECURRING_CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {} # Retorna um dicionário vazio se o arquivo não existir

def save_recurring_config(config):
    """Salva as configurações de e-mail recorrente no arquivo."""
    with open(RECURRING_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)

# --- Função de Envio de E-mail ---
def send_email_with_attachment(sender_email, app_password, receiver_email, subject, body, file_path, cc_email=None):
    """
    Envia um e-mail com um anexo usando uma Senha de App do Google.

    Args:
        sender_email (str): O endereço de e-mail do remetente (sua conta Gmail).
        app_password (str): A Senha de App de 16 caracteres gerada pelo Google.
        receiver_email (str): O endereço de e-mail do destinatário.
        subject (str): O assunto do e-mail.
        body (str): O corpo do texto do e-mail.
        file_path (str): O caminho completo para o arquivo a ser anexado.
        cc_email (str, optional): Endereços de e-mail em cópia (separados por vírgula). Defaults to None.
    """
    try:
        smtp_server = "smtp.gmail.com"
        smtp_port = 587 

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = subject
        if cc_email: # Adiciona CC se houver
            msg['Cc'] = cc_email

        msg.attach(MIMEText(body, 'plain'))

        if file_path and os.path.exists(file_path):
            with open(file_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {os.path.basename(file_path)}',
                )
                
                msg.attach(part)
        else:
            # Em caso de envio imediato via GUI, avisa. Para agendado, a mensagem é no console.
            if __name__ == "__main__": # Apenas se executando a GUI
                messagebox.showwarning("Aviso", f"O arquivo '{file_path}' não foi encontrado ou não foi selecionado. O e-mail será enviado sem anexo.")
            else: # Se executando no modo agendado
                print(f"Aviso: Arquivo de anexo '{file_path}' não encontrado. O e-mail será enviado sem anexo.")
            
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, app_password)
        # Combina TO e CC para o envio do e-mail
        recipients = [receiver_email]
        if cc_email:
            recipients.extend([email.strip() for email in cc_email.split(',') if email.strip()])
        
        server.sendmail(sender_email, recipients, msg.as_string())

        # No modo agendado, não usa messagebox
        if __name__ == "__main__":
            messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")
        else:
            print("E-mail agendado enviado com sucesso!")

    except Exception as e:
        if __name__ == "__main__":
            messagebox.showerror("Erro", f"Ocorreu um erro ao enviar o e-mail: {e}")
        else:
            print(f"Erro ao enviar e-mail agendado: {e}")
    finally:
        if 'server' in locals() and server:
            server.quit()

# --- Variáveis globais para os widgets CustomTkinter ---
# Devem ser definidas como globais para serem acessíveis pelas funções antes de serem totalmente criadas
send_button = None
stop_recurring_button = None
file_label = None
select_file_button = None 
recurring_time_hour_combobox = None
recurring_time_minute_combobox = None
# As entrys de pasta e nome do arquivo recorrente foram removidas da GUI,
# mas as variáveis ainda podem ser usadas internamente se necessário
# recurring_folder_entry = None
# recurring_file_name_entry = None
# select_recurring_folder_button = None
is_recurring_var = None # Variável para o Checkbox


# --- Funções da Interface Gráfica (CustomTkinter) ---
selected_file_path = "" # Para o último arquivo selecionado manualmente
recurring_config = load_recurring_config() # Carrega as configurações recorrentes ao iniciar

def select_file():
    """Abre uma janela para o usuário selecionar um arquivo para anexo."""
    global selected_file_path
    file_types = [
        ('Arquivos de Planilha', '*.xlsx *.xls *.csv'),
        ('Documentos PDF', '*.pdf'),
        ('Documentos de Texto', '*.txt *.doc *.docx'),
        ('Todos os Arquivos', '*.*')
    ]
    selected_file_path = filedialog.askopenfilename(
        title="Selecione o arquivo para anexar",
        filetypes=file_types
    )
    if selected_file_path:
        file_label.configure(text=f"Arquivo selecionado: {os.path.basename(selected_file_path)}")
        # A lógica de preencher os campos de recorrência foi simplificada,
        # pois agora o anexo recorrente é sempre o último selecionado.
        # Os campos explícitos de pasta/nome do arquivo foram removidos da GUI.
    else:
        file_label.configure(text="Nenhum arquivo selecionado.")

def toggle_recurring_fields():
    """Ativa/desativa os campos de recorrência com base no checkbox e ajusta labels/botões."""
    state = "normal" if is_recurring_var.get() else "disabled"
    
    recurring_time_hour_combobox.configure(state=state)
    recurring_time_minute_combobox.configure(state=state)
    
    if is_recurring_var.get():
        send_button.configure(text="Agendar Envio Recorrente")
        file_label.configure(text="O anexo para recorrência será o último selecionado no campo 'Anexo'.") 
        select_file_button.configure(state="normal") # Botão de selecionar arquivo sempre ativo
        # Habilita/desabilita o botão Parar Recorrência com base no estado salvo
        stop_recurring_button.configure(state="normal" if recurring_config.get("is_active", False) else "disabled")
        
        # Se está ativando a recorrência e não há configuração de hora/minuto salva, preenche com a hora/minuto atual
        if not recurring_config.get("recurring_time_hour"): # Se não há hora/minuto recorrente salvo, usa o atual
             recurring_time_hour_combobox.set(datetime.datetime.now().strftime("%H"))
             recurring_time_minute_combobox.set(datetime.datetime.now().strftime("%M"))

    else:
        send_button.configure(text="Enviar E-mail Agora")
        file_label.configure(text="Nenhum arquivo selecionado.") # Este é atualizado pela select_file()
        select_file_button.configure(state="normal")
        stop_recurring_button.configure(state="disabled") # Desativa o botão de parar se não for recorrente


def send_or_schedule_email():
    """Coleta os dados da GUI e decide se envia imediatamente ou agenda."""
    sender_email = sender_entry.get().strip()
    app_password = password_entry.get().strip()
    receiver_email = receiver_entry.get().strip()
    cc_email = cc_entry.get().strip() 
    subject = subject_entry.get().strip()
    email_body = body_text.get("1.0", ctk.END).strip() 

    if not sender_email or not app_password or not receiver_email or not subject:
        messagebox.showwarning("Campos Faltando", "Por favor, preencha todos os campos obrigatórios (Remetente, Senha de App, Destinatário, Assunto).")
        return

    # Salva as configurações básicas (e-mail remetente, senha de app e mensagem padrão, e CC)
    save_default_message(email_body)
    save_sender_email(sender_email)
    save_app_password(app_password) 
    save_cc_email(cc_email) 

    if is_recurring_var.get(): # Modo recorrente
        recurring_hour = recurring_time_hour_combobox.get().strip()
        recurring_minute = recurring_time_minute_combobox.get().strip()
        
        # O anexo recorrente será o último selecionado manualmente
        if not selected_file_path or not os.path.exists(selected_file_path):
             messagebox.showwarning("Anexo Recorrente Faltando", "Por favor, selecione um arquivo no campo 'Anexo' para que ele seja enviado recorrentemente.")
             return

        try: # Validação numérica da hora e minuto
            hour_int = int(recurring_hour)
            minute_int = int(recurring_minute)
            if not (0 <= hour_int <= 23 and 0 <= minute_int <= 59):
                raise ValueError("Hora ou minuto inválido.")
        except ValueError:
            messagebox.showwarning("Formato de Hora Inválido", "Por favor, selecione uma hora e minuto válidos.")
            return

        # Salva a configuração recorrente com a flag 'is_active'
        # Agora salva o caminho completo do arquivo selecionado para recorrência
        config_to_save = {
            "sender_email": sender_email,
            "receiver_email": receiver_email,
            "cc_email": cc_email, 
            "subject": subject,
            "email_body": email_body,
            "recurring_file_path": selected_file_path, # USA O CAMINHO COMPLETO DO ÚLTIMO ARQUIVO SELECIONADO
            "recurring_time_hour": recurring_hour,
            "recurring_time_minute": recurring_minute,
            "is_active": True # Ativa a recorrência ao agendar
        }
        save_recurring_config(config_to_save)
        
        # Atualiza o estado do botão "Parar Recorrência"
        stop_recurring_button.configure(state="normal")

        messagebox.showinfo(
            "Agendamento Salvo",
            "As configurações de e-mail recorrente foram salvas!\n\n"
            f"**O anexo para envio recorrente será: {os.path.basename(selected_file_path)}**\n\n"
            "**Para que o e-mail seja enviado automaticamente, você precisa configurar uma tarefa agendada no seu sistema operacional (Agendador de Tarefas do Windows ou Cron no Linux/macOS) para executar este script diariamente no horário configurado.**\n\n"
            "Quando o script for executado pelo agendador, ele lerá essas configurações e enviará o e-mail se a recorrência estiver ativa."
        )

    else: # Modo de envio imediato
        send_email_with_attachment(sender_email, app_password, receiver_email, subject, email_body, selected_file_path, cc_email)

def stop_recurring_email():
    """Define a flag 'is_active' como False no arquivo de configuração."""
    current_config = load_recurring_config()
    if not current_config:
        messagebox.showinfo("Informação", "Nenhuma recorrência agendada para parar.")
        return
    
    if not current_config.get("is_active", False):
        messagebox.showinfo("Informação", "A recorrência já está inativa.")
        return

    current_config["is_active"] = False
    save_recurring_config(current_config)
    stop_recurring_button.configure(state="disabled")
    messagebox.showinfo("Sucesso", "Recorrência desativada! O e-mail não será mais enviado automaticamente.")


# --- Função para Execução Agendada (Chamada quando o script é executado pelo agendador) ---
def run_scheduled_email():
    """
    Função para ser chamada quando o script é executado por um agendador externo.
    Lê a configuração e tenta enviar o e-mail.
    """
    config = load_recurring_config()
    
    # Verifica se há configuração e se a recorrência está ativa
    if not config or not config.get("is_active", False):
        print("Recorrência inativa ou nenhuma configuração encontrada. Saindo.")
        return

    sender_email = config.get("sender_email")
    receiver_email = config.get("receiver_email")
    cc_email = config.get("cc_email") 
    subject = config.get("subject")
    email_body = config.get("email_body")
    recurring_file_path = config.get("recurring_file_path") # Carrega o caminho do arquivo salvo para recorrência

    # Carrega a senha de app do arquivo para o envio agendado
    app_password_for_scheduled_run = load_app_password() 

    if not all([sender_email, receiver_email, subject, email_body, recurring_file_path, app_password_for_scheduled_run]):
        print("Configurações incompletas para envio agendado (ou Senha de App/Caminho do arquivo não encontrada). Saindo.")
        return
    
    print(f"Tentando enviar e-mail agendado para {receiver_email} (CC: {cc_email if cc_email else 'Nenhum'}) com anexo: {recurring_file_path}")
    
    # Chama a função de envio, sem messagebox para execução agendada
    send_email_with_attachment(sender_email, app_password_for_scheduled_run, receiver_email, subject, email_body, recurring_file_path, cc_email)
    
    # A função send_email_with_attachment já imprime sucesso/erro no console se não for via GUI
    # e lida com o anexo não encontrado.


# --- Lógica de Execução Principal ---
if __name__ == "__main__":
    import sys
    # Se o script está sendo executado no modo agendado, não inicialize a GUI
    if "--scheduled" in sys.argv:
        run_scheduled_email()
    else:
        # --- Configuração da Janela CustomTkinter ---
        root = ctk.CTk()
        root.title("Automatizador de E-mail com Anexo")
        root.geometry("600x750") # Altura ajustada para mais espaço

        # Criar o CTkScrollableFrame principal
        # Este frame irá conter todo o conteúdo da sua GUI, permitindo rolagem
        main_scroll_frame = ctk.CTkScrollableFrame(root, corner_radius=10)
        main_scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        # Configurar as colunas do main_scroll_frame para que o conteúdo se ajuste
        main_scroll_frame.grid_columnconfigure(0, weight=1)

        # --- Frames para organizar os widgets (AGORA DENTRO do main_scroll_frame) ---
        # Removido `grid_columnconfigure(1, weight=1)` pois não é mais necessário para o frame raiz.
        # Os widgets internos usarão suas próprias configurações de grid.

        input_frame = ctk.CTkFrame(main_scroll_frame, corner_radius=10)
        input_frame.pack(padx=10, pady=10, fill="x", expand=False)
        input_frame.grid_columnconfigure(1, weight=1) # Permite que a coluna de entrada expanda
        ctk.CTkLabel(input_frame, text="Dados do E-mail", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, columnspan=2, pady=(10,5)) 

        file_frame = ctk.CTkFrame(main_scroll_frame, corner_radius=10)
        file_frame.pack(padx=10, pady=5, fill="x", expand=False)
        ctk.CTkLabel(file_frame, text="Anexo", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(10,5)) 

        recurring_config_frame = ctk.CTkFrame(main_scroll_frame, corner_radius=10) # Frame para o checkbox e hora
        recurring_config_frame.pack(padx=10, pady=5, fill="x", expand=False)
        recurring_config_frame.grid_columnconfigure(1, weight=1) 
        ctk.CTkLabel(recurring_config_frame, text="Disparo Recorrente", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, columnspan=4, pady=(10,5))


        message_frame = ctk.CTkFrame(main_scroll_frame, corner_radius=10)
        message_frame.pack(padx=10, pady=5, fill="both", expand=True)
        ctk.CTkLabel(message_frame, text="Mensagem do E-mail", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(10,5)) 

        button_frame = ctk.CTkFrame(main_scroll_frame, fg_color="transparent") # Frame transparente para botões
        button_frame.pack(padx=10, pady=10, fill="x", expand=False)
        button_frame.grid_columnconfigure(0, weight=1) 
        button_frame.grid_columnconfigure(1, weight=1) 


        # --- Widgets no input_frame ---
        ctk.CTkLabel(input_frame, text="Seu E-mail (Gmail):").grid(row=1, column=0, sticky="w", padx=15, pady=5)
        sender_entry = ctk.CTkEntry(input_frame, width=300)
        sender_entry.grid(row=1, column=1, padx=15, pady=5, sticky="ew")
        sender_entry.insert(0, load_sender_email())

        ctk.CTkLabel(input_frame, text="Senha de App:").grid(row=2, column=0, sticky="w", padx=15, pady=5)
        password_entry = ctk.CTkEntry(input_frame, width=300, show="*")
        password_entry.grid(row=2, column=1, padx=15, pady=5, sticky="ew")
        password_entry.insert(0, load_app_password()) 

        ctk.CTkLabel(input_frame, text="E-mail Destinatário:").grid(row=3, column=0, sticky="w", padx=15, pady=5)
        receiver_entry = ctk.CTkEntry(input_frame, width=300)
        receiver_entry.grid(row=3, column=1, padx=15, pady=5, sticky="ew")
        receiver_entry.insert(0, "destinatario@example.com") 

        # Campo de CC
        ctk.CTkLabel(input_frame, text="Pessoas em Cópia (CC):").grid(row=4, column=0, sticky="w", padx=15, pady=5)
        cc_entry = ctk.CTkEntry(input_frame, width=300)
        cc_entry.grid(row=4, column=1, padx=15, pady=5, sticky="ew")
        cc_entry.insert(0, load_cc_email()) 

        ctk.CTkLabel(input_frame, text="Assunto:").grid(row=5, column=0, sticky="w", padx=15, pady=5)
        subject_entry = ctk.CTkEntry(input_frame, width=300)
        subject_entry.grid(row=5, column=1, padx=15, pady=5, sticky="ew")
        subject_entry.insert(0, "Relatório Diário/Semanal")

        # --- Widgets no file_frame (Anexo) ---
        select_file_button = ctk.CTkButton(file_frame, text="Selecionar Arquivo", command=select_file)
        select_file_button.pack(pady=10)

        file_label = ctk.CTkLabel(file_frame, text="Nenhum arquivo selecionado.")
        file_label.pack(pady=5)
        # Se houver uma recorrência ativa e um arquivo salvo, exibe o nome do arquivo recorrente
        # e preenche selected_file_path para garantir que o envio agendado o use.
        if recurring_config.get("is_active", False) and recurring_config.get("recurring_file_path"):
            selected_file_path = recurring_config.get("recurring_file_path")
            file_label.configure(text=f"Anexo para recorrência: {os.path.basename(selected_file_path)}")


        # --- Widgets no recurring_config_frame (Disparo Recorrente) ---
        is_recurring_var = ctk.BooleanVar(value=False)
        is_recurring_checkbox = ctk.CTkCheckBox(recurring_config_frame, text="Tornar Recorrente", variable=is_recurring_var, command=toggle_recurring_fields)
        is_recurring_checkbox.grid(row=1, column=0, columnspan=4, sticky="w", padx=15, pady=10) 

        # Hora e Minuto (usando CTkComboBox)
        hours = [f"{i:02d}" for i in range(24)]
        minutes = [f"{i:02d}" for i in range(60)]

        ctk.CTkLabel(recurring_config_frame, text="Hora (HH):").grid(row=2, column=0, sticky="w", padx=15, pady=5)
        recurring_time_hour_combobox = ctk.CTkComboBox(recurring_config_frame, values=hours, width=70, state="disabled")
        recurring_time_hour_combobox.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        recurring_time_hour_combobox.set("08") # Hora padrão
        
        ctk.CTkLabel(recurring_config_frame, text="Minuto (MM):").grid(row=2, column=2, sticky="w", padx=15, pady=5)
        recurring_time_minute_combobox = ctk.CTkComboBox(recurring_config_frame, values=minutes, width=70, state="disabled")
        recurring_time_minute_combobox.grid(row=2, column=3, sticky="w", padx=5, pady=5)
        recurring_time_minute_combobox.set("00") # Minuto padrão
        
        # Carrega e preenche outras configurações recorrentes se existirem
        if recurring_config:
            is_recurring_var.set(True)
            recurring_time_hour_combobox.set(recurring_config.get("recurring_time_hour", "08"))
            recurring_time_minute_combobox.set(recurring_config.get("recurring_time_minute", "00"))
            # `toggle_recurring_fields()` será chamado após a criação dos botões, no final do setup.

        # --- Widgets no message_frame ---
        body_text = ctk.CTkTextbox(message_frame, wrap="word", height=150, width=450)
        body_text.insert("0.0", load_default_message())
        body_text.pack(padx=10, pady=10, fill="both", expand=True)

        # --- Widgets no button_frame ---
        send_button = ctk.CTkButton(button_frame, text="Enviar E-mail Agora", command=send_or_schedule_email, font=ctk.CTkFont(size=14, weight="bold"))
        send_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        stop_recurring_button = ctk.CTkButton(button_frame, text="Parar Recorrência", command=stop_recurring_email, font=ctk.CTkFont(size=14, weight="bold"), fg_color="red", hover_color="darkred")
        stop_recurring_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        # FINALMENTE: Chama toggle_recurring_fields() para configurar o estado inicial dos widgets de recorrência
        # e dos botões "Enviar/Agendar" e "Parar", após todos terem sido criados.
        toggle_recurring_fields()

        root.mainloop()