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
2. Aguarde a conversão e extração das tabelas.
3. Escolha onde salvar o arquivo Excel final.
4. Receba a mensagem de sucesso.


