# Relatório Automático de Consultas Médicas com Python

Este projeto é um script Python que permite:

- Selecionar um arquivo Excel com dados de consultas médicas;
- Processar o arquivo, desmesclando células e filtrando colunas importantes;
- Gerar gráficos (barras e pizza) baseados nos dados;
- Criar uma apresentação PowerPoint com os gráficos;
- Enviar a apresentação por e-mail automaticamente.

---

## Funcionalidades

- Interface simples para escolher o arquivo de entrada e onde salvar os arquivos gerados.
- Processamento automático dos dados para manter apenas colunas relevantes.
- Gráficos bonitos com visualização de quantidade de consultas por médico, status e motivo.
- Apresentação automática em PowerPoint com slides contendo os gráficos.
- Envio por e-mail usando Gmail com autenticação segura via senha de app.
- Mensagens de status para o usuário acompanhar o andamento.

---

## Pré-requisitos

- Python 3.6 ou superior
- Conta Gmail com senha de app gerada para envio via SMTP

---

## Instalação

1. Clone este repositório ou baixe o arquivo `excel_apresentacao.py`

2. Instale as bibliotecas necessárias executando no terminal:

   ```bash
   pip install pandas openpyxl matplotlib python-pptx


3. Configuração
No arquivo excel_apresentacao.py, configure suas credenciais de e-mail na seção:

python
   ```bash
  # Configurações de email Gmail
  EMAIL_REMETENTE = 'seu_email@gmail.com'
  SENHA_APP = 'sua_senha_app'  # senha gerada em myaccount.google.com
  EMAIL_DESTINATARIO = 'destino_email@gmail.com'  # email que receberá o relatório
 ```

Uso
Execute o script com:

 ```bash
python excel_apresentacao.py
 ```
---

## Siga as instruções na tela:

1. Escolha o arquivo Excel com os dados das consultas.
   <img width="598" height="288" alt="image" src="https://github.com/user-attachments/assets/03061728-2316-4937-be55-bcd3f1c56873" />

    Modelo de relátorio:
  
    <img width="882" height="338" alt="image" src="https://github.com/user-attachments/assets/4fb68790-0b7b-4653-aacd-5b1655536fd5" />


3. Escolha onde salvar o arquivo Excel processado.
   <img width="681" height="341" alt="image" src="https://github.com/user-attachments/assets/a04ec802-2103-4d58-bdb7-2b495b87583d" />

   Mensagem de aguarde:
   
   <img width="300" height="200" alt="image" src="https://github.com/user-attachments/assets/fb8f41a3-fd6d-4654-86f2-19bccc4f52fe" />

   Mensagem de salvo com sucesso:
   
   <img width="404" height="152" alt="image" src="https://github.com/user-attachments/assets/8dc04727-ae5d-4e90-991d-ec77401ab1e6" />




    Estrutura das colunas do Excel esperado
   
    O arquivo Excel deve conter (pelo menos) as colunas: Data | Médico | Motivo | Status

    <img width="466" height="280" alt="image" src="https://github.com/user-attachments/assets/b2bafdf0-d365-40b7-8fa6-f08a00343351" />


5. Escolha onde salvar a apresentação PowerPoint gerada.
   <img width="598" height="326" alt="image" src="https://github.com/user-attachments/assets/c233a7a6-f1df-45a4-90d6-75cd872f468f" />

    Mensagem de aguarde:
   
   <img width="300" height="200" alt="image" src="https://github.com/user-attachments/assets/6a5402bd-b843-45b8-85ec-b6f0ad539250" />

   Mensagem de salvo com sucesso:
   
   <img width="400" height="163" alt="image" src="https://github.com/user-attachments/assets/867be313-1dff-4eec-ac56-f44d650cfa4f" />



7. Quando solicitado, confirme se deseja enviar o relatório por e-mail para o destinatário configurado.

   <img width="300" height="200" alt="image" src="https://github.com/user-attachments/assets/f71d043d-9fbf-4592-aaab-c02710f47bb0" />

   Depois aparece mensagem de envio
   
   <img width="300" height="200" alt="image" src="https://github.com/user-attachments/assets/fe349791-3028-4e17-ace4-1ba8944caf68" />

   Mensagem de sucesso
   
   <img width="300" height="200" alt="image" src="https://github.com/user-attachments/assets/81a61c59-7d61-4b28-a758-906c1357e3f6" />





Relatorio final:

<img width="1527" height="866" alt="image" src="https://github.com/user-attachments/assets/ef22679d-533f-4a7d-a7dd-468dd43c3d68" />
<img width="915" height="688" alt="image" src="https://github.com/user-attachments/assets/ae2b9b24-a413-44e5-b39e-15cf225117bd" />
<img width="906" height="686" alt="image" src="https://github.com/user-attachments/assets/52b4d283-7507-4707-a526-bfc187916958" />
<img width="909" height="671" alt="image" src="https://github.com/user-attachments/assets/8213fe0f-4bc6-4f59-86e2-980cbcfa5273" />



