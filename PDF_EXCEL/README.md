# 🐍 PDF para Excel com Limpeza Automática

Este script em Python converte **tabelas de arquivos PDF** para **planilhas Excel** (`.xlsx`) de forma prática, permitindo que o usuário selecione o arquivo e o local de salvamento.  
Após a conversão, ele aplica uma **limpeza automática** para remover linhas indesejadas (como cabeçalhos repetidos com "Data" ou "Nome").

---

## ✨ Funcionalidades

- Interface gráfica simples com **Tkinter**.
- Seleção do arquivo PDF e escolha do local/nome do arquivo Excel final.
- Extração de **todas as tabelas** do PDF usando **Tabula**.
- Conversão automática para `.xlsx`.
- Limpeza do arquivo Excel:
  - Remove linhas repetidas de cabeçalho contendo "Data" ou "Nome".
- Mensagens de status e erros exibidas para o usuário.

---

## 📦 Dependências

Para rodar o projeto, você precisa instalar:

```bash
pip install pandas tabula-py openpyxl
````

Além disso, o **tabula-py** requer **Java** instalado no sistema.

---

## 🚀 Como Usar

1. Copie ou baixe o script:

```bash
Baixe https://github.com/SuellenCFerreira/AutomacaoPython/blob/main/PDF_EXCEL/pdf_excel_tabula.py
```

2. Instale as dependências:

```bash
pip install pandas tabula-py openpyxl
```

3. Execute o script:

```bash
python seu_script.py
```

---

### Passos no programa:

1. Escolha o arquivo PDF.
   
<img width="664" height="307" alt="image" src="https://github.com/user-attachments/assets/11616280-a162-4565-930d-2f95c88056ab" />

3. Aguarde a conversão e extração das tabelas.
<img width="298" height="226" alt="image" src="https://github.com/user-attachments/assets/d0c92c7e-5225-4181-93d0-c2fca0fb92e2" />

5. Escolha onde salvar o arquivo Excel final.

<img width="722" height="320" alt="image" src="https://github.com/user-attachments/assets/68fe09d0-ee7c-4a12-a18d-4c24e7fd19a0" />

5. Receba a mensagem de sucesso.
<img width="418" height="161" alt="image" src="https://github.com/user-attachments/assets/c5c76a6b-ca2f-4624-b04f-642c29bdc228" />

6.Arquivo convertido.

<img width="688" height="259" alt="image" src="https://github.com/user-attachments/assets/5c20e3f2-6690-47b4-8bbb-aa9eb8a557b8" />





