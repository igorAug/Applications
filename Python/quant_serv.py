import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

COLUNA_SERVICO = "TipoServico"
COLUNA_LOCALIDADE = "NomeCidade"
COLUNA_MATRICULA = "MatriculaImovel"

def processar_dados():
    caminho_arquivo = entry_caminho_arquivo.get()
    if not caminho_arquivo:
        messagebox.showerror("Erro", "Selecione um arquivo primeiro.")
        return

    tree.destroy()
    criar_treeview()
    label_status.config(text="Processando...", fg="blue")
    root.update_idletasks()

    try:
        if caminho_arquivo.endswith('.xlsx'):
            df = pd.read_excel(caminho_arquivo, engine='openpyxl')
        elif caminho_arquivo.endswith('.csv'):
            df = pd.read_csv(caminho_arquivo, sep=None, engine='python')
        else:
            messagebox.showerror("Erro", "Formato de arquivo não suportado.")
            return
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        return

    colunas = [COLUNA_LOCALIDADE, COLUNA_MATRICULA, COLUNA_SERVICO]
    faltando = [c for c in colunas if c not in df.columns]
    if faltando:
        messagebox.showerror("Erro", f"Colunas faltando: {', '.join(faltando)}")
        return

    try:
        tabela = pd.pivot_table(
            df,
            values=COLUNA_MATRICULA,
            index=COLUNA_SERVICO,
            columns=COLUNA_LOCALIDADE,
            aggfunc='count',
            fill_value=0
        )
        tabela['Total Geral'] = tabela.sum(axis=1).astype(int)
        tabela_final = tabela.reset_index().rename_axis(columns=None)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar: {e}")
        return

    try:
        tabela_final.to_excel("relatorio_servico_x_cidade.xlsx", index=False, engine='openpyxl')
    except Exception as e:
        messagebox.showwarning("Aviso", f"Erro ao salvar arquivo: {e}")

    colunas = list(tabela_final.columns)
    tree["columns"] = colunas
    tree.column("#0", width=0, stretch=tk.NO)

    for c in colunas:
        largura = 100
        ancora = tk.CENTER
        if c == COLUNA_SERVICO:
            largura = 180
            ancora = tk.W
        elif c == "Total Geral":
            largura = 120
        tree.column(c, anchor=ancora, width=largura)
        tree.heading(c, text=c, anchor=ancora)

    for _, row in tabela_final.iterrows():
        tree.insert("", tk.END, values=tuple(row))

    label_status.config(text="Processamento concluído!", fg="green")

def selecionar_arquivo():
    tipos = (("Excel", "*.xlsx"), ("CSV", "*.csv"), ("Todos", "*.*"))
    caminho = filedialog.askopenfilename(title="Selecione a planilha", filetypes=tipos)
    if caminho:
        entry_caminho_arquivo.delete(0, tk.END)
        entry_caminho_arquivo.insert(0, caminho)
        label_status.config(text="Arquivo selecionado.", fg="black")

def criar_treeview():
    global tree
    frame_tabela = ttk.Frame(frame_resultados)
    frame_tabela.pack(fill=tk.BOTH, expand=True)
    scroll_y = ttk.Scrollbar(frame_tabela, orient=tk.VERTICAL)
    scroll_x = ttk.Scrollbar(frame_tabela, orient=tk.HORIZONTAL)
    tree = ttk.Treeview(frame_tabela, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
    scroll_y.config(command=tree.yview)
    scroll_x.config(command=tree.xview)
    scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    tree.pack(fill=tk.BOTH, expand=True)

root = tk.Tk()
root.title("Relatório de Serviços por Cidade")
root.geometry("800x500")

frame_principal = ttk.Frame(root, padding="10")
frame_principal.pack(fill=tk.BOTH, expand=True)

frame_arquivo = ttk.LabelFrame(frame_principal, text="Arquivo", padding="10")
frame_arquivo.pack(fill=tk.X, expand=True, pady=5)

entry_caminho_arquivo = ttk.Entry(frame_arquivo, width=60)
entry_caminho_arquivo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
btn_selecionar = ttk.Button(frame_arquivo, text="Selecionar", command=selecionar_arquivo)
btn_selecionar.pack(side=tk.LEFT)

btn_processar = ttk.Button(frame_principal, text="Gerar Relatório", command=processar_dados)
btn_processar.pack(pady=10, fill=tk.X)

frame_resultados = ttk.LabelFrame(frame_principal, text="Resultados", padding="10")
frame_resultados.pack(fill=tk.BOTH, expand=True, pady=5)

criar_treeview()

label_status = tk.Label(frame_principal, text="Pronto. Selecione um arquivo.", relief=tk.SUNKEN, anchor=tk.W)
label_status.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))

root.mainloop()
