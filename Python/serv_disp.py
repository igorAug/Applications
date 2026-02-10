import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys

def processar_planilha(arquivo_entrada, arquivo_saida, servico_desejado, grouping_column, debit_min_str, debit_max_str):
    try:
        try:
            df = pd.read_excel(arquivo_entrada)
        except:
            df = pd.read_csv(arquivo_entrada, sep=None, engine='python')

        colunas_necessarias = ['DataGeracao', 'TipoServico', 'NomeBairro', 'NomeCidade',
                               'MatriculaImovel', 'TotalDebitos', 'TotalDebitosParcelamento']
        colunas_faltando = [c for c in colunas_necessarias if c not in df.columns]
        if colunas_faltando:
            if grouping_column in colunas_faltando:
                raise ValueError(f"Coluna '{grouping_column}' não encontrada.")
            colunas_faltando = [c for c in colunas_faltando if c not in ['NomeBairro', 'NomeCidade']]
            if colunas_faltando:
                raise ValueError(f"Faltam colunas: {', '.join(colunas_faltando)}")

        df['TotalDebitos'] = pd.to_numeric(df['TotalDebitos'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df['TotalDebitosParcelamento'] = pd.to_numeric(df['TotalDebitosParcelamento'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df['DebitoTotal'] = df['TotalDebitos'] + df['TotalDebitosParcelamento']

        try:
            debit_min = float(debit_min_str.replace(',', '.')) if debit_min_str else 0.0
        except:
            debit_min = 0.0
        try:
            debit_max = float(debit_max_str.replace(',', '.')) if debit_max_str else float('inf')
        except:
            debit_max = float('inf')

        df = df[(df['DebitoTotal'] >= debit_min) & (df['DebitoTotal'] <= debit_max)]
        df['DataGeracao'] = pd.to_datetime(df['DataGeracao'], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['DataGeracao'])
        hoje = pd.Timestamp.now().normalize()
        df['IdadeDias'] = (hoje - df['DataGeracao']).dt.days
        df['IdadeDias'] = df['IdadeDias'].apply(lambda x: max(0, x))
        df_filtrado = df[df['TipoServico'] == servico_desejado].copy()

        if df_filtrado.empty:
            messagebox.showwarning("Sem Dados", f"Nenhum registro encontrado para {servico_desejado}")
            return False

        bins = [0,5,10,20,30,40,50,60,70,80,90,100,110,120,float('inf')]
        labels = ['0-4','5-9','10-19','20-29','30-39','40-49','50-59','60-69','70-79','80-89','90-99','100-109','110-119','120+']
        df_filtrado['FaixaIdade'] = pd.cut(df_filtrado['IdadeDias'], bins=bins, labels=labels, right=False)

        agg_total = df_filtrado.groupby(grouping_column).agg(
            TOTAL_SERV_DISP=('MatriculaImovel', 'count'),
            VALOR_RS=('DebitoTotal', 'sum')
        )
        agg_idade = pd.pivot_table(df_filtrado, values='MatriculaImovel', index=grouping_column,
                                   columns='FaixaIdade', aggfunc='count', fill_value=0)

        resultado = pd.concat([agg_total, agg_idade], axis=1)
        totais = resultado.sum().to_frame().T
        totais.index = ['TOTAIS']
        resultado = pd.concat([totais, resultado])
        resultado.index.name = grouping_column.upper()

        writer = pd.ExcelWriter(arquivo_saida, engine='xlsxwriter')
        resultado.to_excel(writer, sheet_name='Relatorio')
        ws = writer.sheets['Relatorio']
        wb = writer.book
        moeda = wb.add_format({'num_format': 'R$ #,##0.00'})
        ws.set_column('C:C', 18, moeda)
        ws.set_column('A:A', 20)
        ws.set_column('B:B', 15)
        ws.set_column('D:Z', 12)
        writer.close()
        return True

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        return False


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Analisador de Serviços")
        self.root.geometry("500x420")

        frame = ttk.Frame(root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Planilha de Entrada:").grid(row=0, column=0, sticky="w")
        self.entry_entrada_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.entry_entrada_var, width=50, state="disabled").grid(row=1, column=0, padx=5, sticky="we")
        ttk.Button(frame, text="Procurar...", command=self.selecionar_arquivo_entrada).grid(row=1, column=1, padx=5)

        ttk.Label(frame, text="Agrupar por:").grid(row=2, column=0, sticky="w")
        self.combo_agrupar_var = tk.StringVar()
        self.combo_agrupar = ttk.Combobox(frame, textvariable=self.combo_agrupar_var, width=48, state="readonly")
        self.combo_agrupar.grid(row=3, column=0, padx=5, sticky="we")

        ttk.Label(frame, text="Tipo de Serviço:").grid(row=4, column=0, sticky="w")
        self.combo_servico_var = tk.StringVar()
        self.combo_servico = ttk.Combobox(frame, textvariable=self.combo_servico_var, width=48, state="readonly")
        self.combo_servico.grid(row=5, column=0, padx=5, sticky="we")

        ttk.Label(frame, text="Filtro de Débito (R$):").grid(row=6, column=0, sticky="w")
        debit_frame = ttk.Frame(frame)
        debit_frame.grid(row=7, column=0, columnspan=2, padx=5, sticky="we")
        ttk.Label(debit_frame, text="Mínimo:").pack(side=tk.LEFT)
        self.entry_debit_min_var = tk.StringVar(value="0")
        ttk.Entry(debit_frame, textvariable=self.entry_debit_min_var, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(debit_frame, text="Máximo:").pack(side=tk.LEFT)
        self.entry_debit_max_var = tk.StringVar(value="999999999")
        ttk.Entry(debit_frame, textvariable=self.entry_debit_max_var, width=15).pack(side=tk.LEFT, padx=5)

        ttk.Label(frame, text="Salvar Relatório Como:").grid(row=8, column=0, sticky="w")
        self.entry_saida_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.entry_saida_var, width=50, state="disabled").grid(row=9, column=0, padx=5, sticky="we")
        ttk.Button(frame, text="Salvar...", command=self.selecionar_arquivo_saida).grid(row=9, column=1, padx=5)

        self.btn_gerar = ttk.Button(frame, text="Gerar Relatório", command=self.executar_processamento, state="disabled")
        self.btn_gerar.grid(row=10, column=0, columnspan=2, pady=20)

    def selecionar_arquivo_entrada(self):
        f = filedialog.askopenfilename(title="Selecionar Planilha", filetypes=(("Excel", "*.xlsx"), ("CSV", "*.csv")))
        if f:
            self.entry_entrada_var.set(f)
            self.carregar_opcoes_planilha(f)

    def carregar_opcoes_planilha(self, f):
        try:
            try:
                df = pd.read_excel(f, nrows=200)
            except:
                df = pd.read_csv(f, sep=None, engine='python', nrows=200)
            tipos = df['TipoServico'].dropna().unique().tolist()
            tipos.sort()
            self.combo_servico['values'] = tipos
            if tipos: self.combo_servico.current(0)
            grupos = []
            if 'NomeBairro' in df.columns: grupos.append('NomeBairro')
            if 'NomeCidade' in df.columns: grupos.append('NomeCidade')
            self.combo_agrupar['values'] = grupos
            if grupos: self.combo_agrupar.current(0)
            self.btn_gerar.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler: {e}")
            self.btn_gerar.config(state=tk.DISABLED)

    def selecionar_arquivo_saida(self):
        f = filedialog.asksaveasfilename(title="Salvar Como", defaultextension=".xlsx", filetypes=(("Excel", "*.xlsx"),))
        if f:
            self.entry_saida_var.set(f)

    def executar_processamento(self):
        inp = self.entry_entrada_var.get()
        out = self.entry_saida_var.get()
        serv = self.combo_servico_var.get()
        grp = self.combo_agrupar_var.get()
        dmin = self.entry_debit_min_var.get()
        dmax = self.entry_debit_max_var.get()

        if not all([inp, out, serv, grp]):
            messagebox.showerror("Erro", "Preencha todos os campos.")
            return

        self.btn_gerar.config(state=tk.DISABLED, text="Processando...")
        self.root.update_idletasks()
        try:
            ok = processar_planilha(inp, out, serv, grp, dmin, dmax)
            if ok:
                messagebox.showinfo("Sucesso", f"Relatório salvo em:\n{out}")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
        self.btn_gerar.config(state=tk.NORMAL, text="Gerar Relatório")


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
    root = tk.Tk()
    App(root)
    root.mainloop()
