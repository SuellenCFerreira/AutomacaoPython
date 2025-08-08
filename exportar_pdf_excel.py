import tabula  # Biblioteca para extrair tabelas de arquivos PDF
import pandas as pd  # Biblioteca para manipular dados em tabelas (DataFrames)
import tkinter as tk  # Biblioteca para criar janelas e diálogos gráficos
from tkinter import filedialog, messagebox  # Para abrir janelas de seleção de arquivos e mensagens
import os  # Biblioteca para manipular arquivos e pastas do sistema operacional

# Função para abrir uma janela para o usuário escolher um arquivo PDF
def selecionar_pdf():
    root = tk.Tk()  # Cria uma janela principal (mas não mostra)
    root.withdraw()  # Esconde a janela principal, porque só queremos o diálogo
    arquivo_pdf = filedialog.askopenfilename(  # Abre a janela para escolher um arquivo
        title="Selecione o arquivo PDF",  # Título da janela
        filetypes=[("PDF files", "*.pdf")]  # Mostrar apenas arquivos PDF
    )
    root.destroy()  # Fecha a janela invisível criada antes
    return arquivo_pdf  # Retorna o caminho do arquivo escolhido

# Função para abrir uma janela para o usuário escolher onde salvar o arquivo Excel e com qual nome
def selecionar_pasta_e_nome():
    root = tk.Tk()  # Cria janela invisível
    root.withdraw()
    caminho_completo = filedialog.asksaveasfilename(  # Abre janela para salvar arquivo
        defaultextension=".xlsx",  # Extensão padrão do arquivo
        filetypes=[("Arquivo Excel", "*.xlsx")],  # Mostrar só arquivos Excel
        title="Salvar arquivo Excel como"  # Título da janela
    )
    root.destroy()  # Fecha janela invisível
    return caminho_completo  # Retorna o caminho completo onde salvar o arquivo

# Função para mostrar uma caixa de diálogo perguntando se o usuário quer tentar novamente após erro
def confirmar_tentar_novamente():
    root = tk.Tk()  # Cria janela invisível
    root.withdraw()
    resposta = messagebox.askretrycancel(  # Caixa com botão "Tentar novamente" e "Cancelar"
        "Erro", "Nome ou pasta inválidos. Deseja tentar novamente?"
    )
    root.destroy()
    return resposta  # Retorna True se escolher tentar, False se cancelar

# Função para mostrar uma janela simples que diz "Aguarde, processando"
def mostrar_aguarde():
    aguarde = tk.Tk()  # Cria janela visível
    aguarde.title("Aguarde")  # Define título da janela
    aguarde.geometry("500x400")  # Tamanho da janela
    aguarde.resizable(False, False)  # Não permite redimensionar
    label = tk.Label(aguarde, text="Processando arquivo, aguarde...", font=("Arial", 12))
    label.pack(expand=True, padx=20, pady=30)  # Coloca texto centralizado na janela
    aguarde.update()  # Atualiza a janela para mostrar imediatamente
    return aguarde  # Retorna a janela para poder fechá-la depois

# Função para "limpar" o arquivo Excel gerado, removendo linhas que começam com "Data" ou "Nome"
def limpar_excel(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo)  # Abre o arquivo Excel numa tabela (DataFrame)
    except Exception as e:
        raise RuntimeError(f"Erro ao ler o arquivo Excel para limpeza: {e}")

    # Procurar se existe coluna chamada "Data" e "Nome" (ignorando maiúsculas/minúsculas)
    col_data = None
    col_nome = None
    for c in df.columns:
        c_lower = c.lower()
        if 'data' == c_lower:
            col_data = c
        if 'nome' == c_lower:
            col_nome = c

    # Se encontrar essas colunas, filtra para remover linhas que parecem cabeçalho repetido (que começam com "data" ou "nome")
    if col_data or col_nome:
        condicoes = []
        if col_data:
            condicoes.append((df.index >= 1) & df[col_data].astype(str).str.lower().str.startswith('data'))
        if col_nome:
            condicoes.append((df.index >= 1) & df[col_nome].astype(str).str.lower().str.startswith('nome'))

        # Junta as condições e filtra para excluir essas linhas
        if condicoes:
            cond_final = ~(condicoes[0] if len(condicoes) == 1 else condicoes[0] | condicoes[1])
            df = df[cond_final].reset_index(drop=True)  # Reseta o índice da tabela

    try:
        df.to_excel(caminho_arquivo, index=False)  # Salva o arquivo Excel limpo, sem o índice extra
    except Exception as e:
        raise RuntimeError(f"Erro ao salvar o arquivo Excel limpo: {e}")

# Função principal que executa o processo todo
def main():
    while True:  # Loop para repetir caso o usuário cancele ou dê erro
        arquivo_pdf = selecionar_pdf()  # Chama função para escolher PDF
        if not arquivo_pdf:  # Se o usuário cancelar (não escolher arquivo)
            print("Operação cancelada pelo usuário.")
            messagebox.showinfo("Cancelado", "Operação cancelada pelo usuário.")
            return  # Sai do programa

        janela_aguarde = mostrar_aguarde()  # Mostra a janela "Aguarde" para avisar que vai processar
        print("Arquivo selecionado:", arquivo_pdf)

        try:
            # Tenta extrair todas as tabelas do PDF numa lista de DataFrames
            dfs = tabula.read_pdf(arquivo_pdf, pages='all', multiple_tables=True)
            if not dfs:  # Se não encontrar nenhuma tabela
                raise ValueError("Nenhuma tabela encontrada no PDF.")
            df = pd.concat(dfs, ignore_index=True)  # Junta todas as tabelas numa só tabela grande
        except Exception as e:
            janela_aguarde.destroy()  # Fecha janela "Aguarde"
            print(f"Erro ao processar o PDF: {e}")
            messagebox.showerror("Erro", f"Erro ao processar o PDF:\n{e}")
            return  # Sai do programa

        janela_aguarde.destroy()  # Fecha janela "Aguarde" após processar

        while True:  # Loop para tentar salvar o arquivo Excel
            caminho_salvar = selecionar_pasta_e_nome()  # Pergunta onde salvar e nome do Excel
            if not caminho_salvar:  # Se cancelar
                tentar_novamente = confirmar_tentar_novamente()  # Pergunta se quer tentar de novo
                if tentar_novamente:
                    continue  # Recomeça esse loop
                else:
                    print("Operação cancelada pelo usuário.")
                    messagebox.showinfo("Cancelado", "Operação cancelada pelo usuário.")
                    return  # Sai do programa

            # Se o nome não terminar com .xlsx, adiciona essa extensão
            if not caminho_salvar.lower().endswith(".xlsx"):
                caminho_salvar += ".xlsx"

            try:
                df.to_excel(caminho_salvar, index=False)  # Salva o Excel com os dados extraídos

                limpar_excel(caminho_salvar)  # Abre o Excel salvo, limpa linhas repetidas e salva de novo

                print(f"Arquivo gerado com sucesso: {caminho_salvar}")
                messagebox.showinfo("Sucesso", f"Arquivo gerado com sucesso:\n{caminho_salvar}")
                return  # Sai do programa depois de sucesso
            except Exception as e:
                print(f"Erro ao salvar ou limpar o arquivo Excel: {e}")
                tentar_novamente = confirmar_tentar_novamente()  # Pergunta se quer tentar novamente
                if not tentar_novamente:
                    print("Operação cancelada pelo usuário.")
                    messagebox.showinfo("Cancelado", "Operação cancelada pelo usuário.")
                    return  # Sai do programa

# Ponto de entrada do programa: chama a função principal
if __name__ == "__main__":
    main()
