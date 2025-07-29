🚀 Automação de E-mails com Python: Solução Completa para Envio Diário de Relatórios
Desenvolvi uma aplicação em Python para automatizar o envio diário de e-mails com anexos, eliminando processos manuais e repetitivos. Este sistema é robusto e flexível, suportando tanto envios imediatos quanto agendamentos recorrentes.

🔨 Tecnologias e Funcionalidades Implementadas
GUI Moderna: Interface intuitiva e personalizável utilizando customtkinter.

Gerenciamento de Configurações: Persistência de dados (e-mail remetente, mensagens padrão, anexos) via arquivos JSON e TXT.

Segurança: Integração segura com Gmail via OAuth2 (ou App Password para envio SMTP).

Recorrência: Sistema de agendamento interno, com integração para o Agendador de Tarefas (Windows) e Cron (Linux).

Tratamento de Erros: Validações robustas para arquivos, credenciais e formato de horários.

✨ Features Extras:
Suporte para CC (cópia) em e-mails.

Pré-visualização do anexo selecionado.

Controle de recorrência (iniciar/parar via interface).

⚠️ Importante: Configuração do Gmail (Senha de App)
Para que a aplicação funcione corretamente com o Gmail, é obrigatório o uso de uma Senha de App se você não estiver utilizando OAuth2. Isso exige que a verificação em duas etapas esteja ativada na sua conta Google.

📌 Passos para Gerar a Senha de App:
Acesse: https://myaccount.google.com/apppasswords (é necessário estar logado na sua conta Google).

Em "Selecionar aplicativo", escolha "Outro (Nome personalizado)".

Digite um nome para identificar a senha (ex.: Python_Email_Sender).

Clique em "Gerar".

Copie a senha de 16 caracteres exibida (ela não será mostrada novamente!).

Cole essa senha no campo apropriado dentro da interface da aplicação, logo após preencher suas outras informações.

⚙️ Configurando o Agendador de Tarefas do Windows
Após parametrizar as informações no aplicativo e gerar a Senha de App, o próximo passo é configurar o Agendador de Tarefas do Windows para executar seu script automaticamente no horário que você definiu na interface. Isso garante que seus e-mails recorrentes sejam enviados sem a necessidade de abrir o programa manualmente.

Passos Detalhados:
Abra o Agendador de Tarefas:

Clique no menu Iniciar do Windows.

Digite "Agendador de Tarefas" (ou "Task Scheduler") na barra de pesquisa e abra o aplicativo.

Crie uma Tarefa Básica:

No painel "Ações" (localizado no lado direito), clique em "Criar Tarefa Básica...". Isso iniciará um assistente.

Defina o Nome e a Descrição:

Nome: Dê um nome claro, por exemplo, "Enviar Relatório Diário Automatizado".

Descrição (Opcional): Adicione uma breve descrição, como "Envia o relatório diário automaticamente via script Python".

Clique em "Avançar".

Defina o Disparador (Quando a tarefa deve iniciar):

Selecione "Diariamente".

Clique em "Avançar".

Configure a Frequência Diária:

Data de início: Defina a data atual.

Hora: Insira a hora exata que você configurou na interface CustomTkinter do seu script (por exemplo, se você selecionou 08:00, coloque 08:00:00).

Clique em "Avançar".

Defina a Ação (O que a tarefa vai fazer):

Selecione "Iniciar um programa".

Clique em "Avançar".

Configure o Programa/Script a Ser Iniciado:

Programa/script: Este é o caminho completo para o executável do Python. Exemplo: C:\Users\pedro\AppData\Local\Programs\Python\Python312\python.exe

Adicionar argumentos (opcional): Este é o caminho completo para o seu script Python, seguido pelo argumento --scheduled. Coloque o caminho do script entre aspas duplas, especialmente se houver espaços no caminho. Exemplo: "c:\Users\pedro\Downloads\automação\e-mail.py" --scheduled

Iniciar em (opcional): Digite o caminho completo da pasta onde seu script e-mail.py está localizado. Isso é crucial para que o script encontre seus arquivos de configuração (.json, .txt). Exemplo: c:\Users\pedro\Downloads\automação\

Clique em "Avançar".

Conclua:

Por fim, clique em "Concluir".

Após a primeira parametrização e a configuração no Agendador de Tarefas, os próximos envios ocorrerão de maneira totalmente automática.
