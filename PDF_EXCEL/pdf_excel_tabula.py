import tabula
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def selecionar_pdf():
    root = tk.Tk()
    root.withdraw()
    arquivo_pdf = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=[("PDF files", "*.pdf")]
    )
    root.destroy()
    return arquivo_pdf

def selecionar_pasta_e_nome():
    root = tk.Tk()
    root.withdraw()
    caminho_completo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivo Excel", "*.xlsx")],
        title="Salvar arquivo Excel como"
    )
    root.destroy()
    return caminho_completo

def confirmar_tentar_novamente():
    root = tk.Tk()
    root.withdraw()
    resposta = messagebox.askretrycancel("Erro", "Nome ou pasta inválidos. Deseja tentar novamente?")
    root.destroy()
    return resposta

def mostrar_aguarde():
    aguarde = tk.Tk()
    aguarde.title("Aguarde")
    aguarde.geometry("500x400")
    aguarde.resizable(False, False)
    label = tk.Label(aguarde, text="Processando arquivo, aguarde...", font=("Arial", 12))
    label.pack(expand=True, padx=20, pady=30)
    aguarde.update()
    return aguarde

def limpar_excel(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo)
    except Exception as e:
        raise RuntimeError(f"Erro ao ler o arquivo Excel para limpeza: {e}")

    # Verificar se as colunas 'Data' e 'Nome' existem para aplicar filtro
    col_data = None
    col_nome = None
    for c in df.columns:
        c_lower = c.lower()
        if 'data' == c_lower:
            col_data = c
        if 'nome' == c_lower:
            col_nome = c

    if col_data or col_nome:
        condicoes = []
        if col_data:
            condicoes.append((df.index >= 1) & df[col_data].astype(str).str.lower().str.startswith('data'))
        if col_nome:
            condicoes.append((df.index >= 1) & df[col_nome].astype(str).str.lower().str.startswith('nome'))

        if condicoes:
            cond_final = ~(condicoes[0] if len(condicoes) == 1 else condicoes[0] | condicoes[1])
            df = df[cond_final].reset_index(drop=True)

    try:
        df.to_excel(caminho_arquivo, index=False)
    except Exception as e:
        raise RuntimeError(f"Erro ao salvar o arquivo Excel limpo: {e}")

def main():
    while True:
        arquivo_pdf = selecionar_pdf()
        if not arquivo_pdf:
            print("Operação cancelada pelo usuário.")
            messagebox.showinfo("Cancelado", "Operação cancelada pelo usuário.")
            return

        janela_aguarde = mostrar_aguarde()
        print("Arquivo selecionado:", arquivo_pdf)

        try:
            # Extrair todas as tabelas e juntar num único DataFrame
            dfs = tabula.read_pdf(arquivo_pdf, pages='all', multiple_tables=True)
            if not dfs:
                raise ValueError("Nenhuma tabela encontrada no PDF.")
            df = pd.concat(dfs, ignore_index=True)
        except Exception as e:
            janela_aguarde.destroy()
            print(f"Erro ao processar o PDF: {e}")
            messagebox.showerror("Erro", f"Erro ao processar o PDF:\n{e}")
            return

        janela_aguarde.destroy()

        while True:
            caminho_salvar = selecionar_pasta_e_nome()
            if not caminho_salvar:
                tentar_novamente = confirmar_tentar_novamente()
                if tentar_novamente:
                    continue
                else:
                    print("Operação cancelada pelo usuário.")
                    messagebox.showinfo("Cancelado", "Operação cancelada pelo usuário.")
                    return

            if not caminho_salvar.lower().endswith(".xlsx"):
                caminho_salvar += ".xlsx"

            try:
                # Salvar Excel inicialmente (sem limpeza)
                df.to_excel(caminho_salvar, index=False)

                # Agora abrir o Excel salvo, limpar e salvar novamente
                limpar_excel(caminho_salvar)

                print(f"Arquivo gerado com sucesso: {caminho_salvar}")
                messagebox.showinfo("Sucesso", f"Arquivo gerado com sucesso:\n{caminho_salvar}")
                return
            except Exception as e:
                print(f"Erro ao salvar ou limpar o arquivo Excel: {e}")
                tentar_novamente = confirmar_tentar_novamente()
                if not tentar_novamente:
                    print("Operação cancelada pelo usuário.")
                    messagebox.showinfo("Cancelado", "Operação cancelada pelo usuário.")
                    return

if __name__ == "__main__":
    main()