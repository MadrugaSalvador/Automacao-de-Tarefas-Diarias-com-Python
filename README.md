🚀 Automação de E-mails com Python: Solução Completa para Envio Diário de Relatórios

Recentemente, desenvolvi uma aplicação em Python para automatizar o envio diário de e-mails com anexos, eliminando um processo manual e repetitivo. O sistema foi projetado para ser robusto e flexível, atendendo tanto a envios imediatos quanto a agendamentos recorrentes.

🔨 Tecnologias e Funcionalidades Implementadas
GUI Moderna: Interface intuitiva usando customtkinter (com temas personalizáveis).

Gerenciamento de Configurações: Persistência de dados (e-mail remetente, mensagens padrão, anexos) via arquivos JSON e txt.

Segurança: Integração com Gmail via OAuth2 (ou App Password para envio SMTP).

Recorrência: Sistema de agendamento interno + integração com Task Scheduler (Windows) e Cron (Linux).

Tratamento de Erros: Validações de arquivos, credenciais e formato de horários.

Features Extras:

Suporte a CC (cópia) em e-mails.

Pré-visualização do anexo selecionado.

Controle de recorrência (iniciar/parar via interface).
