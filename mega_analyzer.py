import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from itertools import combinations
from collections import Counter
import random
import os
import re
from datetime import datetime
from tkcalendar import DateEntry

class MegaAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mega-Sena Analyzer Pro")
        self.root.geometry("800x900")
        self.root.resizable(False, False)

        self.df = None
        self.bolas_cols = []
        self.date_col = None

        self._setup_ui()

    def _setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', padding=6, font=('Segoe UI', 10))
        style.configure('TLabel', font=('Segoe UI', 10))
        self.root.configure(bg='#f0f0f0')

        main = ttk.Frame(self.root, padding=15)
        main.pack(fill='both', expand=True)

        # Upload
        ttk.Label(main, text="Planilha de Sorteios", font=('Segoe UI', 12, 'bold')).pack(anchor='w')
        up_frame = ttk.Frame(main)
        up_frame.pack(fill='x', pady=(5, 15))
        self.path_var = tk.StringVar(value="Nenhuma planilha selecionada")
        ttk.Label(up_frame, textvariable=self.path_var, wraplength=550).pack(side='left', fill='x', expand=True)
        ttk.Button(up_frame, text="Selecionar Arquivo", command=self.load_file).pack(side='right')

        # Filtros
        filt_frame = ttk.LabelFrame(main, text="Filtros de Pesquisa", padding=10)
        filt_frame.pack(fill='x', pady=10)

        # Calendários
        ttk.Label(filt_frame, text="Período:").grid(row=0, column=0, sticky='w')
        
        self.cal_start = DateEntry(filt_frame, width=12, background='darkblue',
                                   foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy',
                                   year=2023, month=1, day=1)
        self.cal_start.grid(row=0, column=1, padx=5)
        
        ttk.Label(filt_frame, text="até").grid(row=0, column=2)
        
        self.cal_end = DateEntry(filt_frame, width=12, background='darkblue',
                                 foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy',
                                 year=2024, month=12, day=31)
        self.cal_end.grid(row=0, column=3, padx=5)

        ttk.Label(filt_frame, text="Números juntos:").grid(row=1, column=0, sticky='w', pady=(10,0))
        self.combo_size = ttk.Combobox(filt_frame, values=[2, 3, 4, 5, 6], state='readonly', width=5)
        self.combo_size.set(2)
        self.combo_size.grid(row=1, column=1, padx=5, pady=(10,0))

        ttk.Label(filt_frame, text="Faixa %:").grid(row=1, column=2, sticky='w', pady=(10,0))
        self.range_min = ttk.Entry(filt_frame, width=4)
        self.range_min.insert(0, "1")
        self.range_min.grid(row=1, column=3, padx=5, pady=(10,0))
        ttk.Label(filt_frame, text="até").grid(row=1, column=4, pady=(10,0))
        self.range_max = ttk.Entry(filt_frame, width=4)
        self.range_max.insert(0, "60")
        self.range_max.grid(row=1, column=5, padx=5, pady=(10,0))

        self.show_mode = tk.StringVar(value="repetidos")
        ttk.Checkbutton(filt_frame, text="Mostrar TODAS", 
                       variable=self.show_mode, onvalue="todos", offvalue="repetidos").grid(row=2, column=0, columnspan=6, sticky='w', pady=(10,0))

        # Botões
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill='x', pady=15)
        ttk.Button(btn_frame, text="Analisar", command=self.run_combo_analysis).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Faixa %", command=self.run_range_analysis).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Gerar Jogos", command=self.generate_games).pack(side='left', padx=5)

        # Resultados
        res_frame = ttk.LabelFrame(main, text="Resultados", padding=10)
        res_frame.pack(fill='both', expand=True, pady=10)
        self.text_res = tk.Text(res_frame, height=22, font=('Consolas', 9), state='disabled')
        scrollbar = ttk.Scrollbar(res_frame, orient="vertical", command=self.text_res.yview)
        self.text_res.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        self.text_res.pack(side='left', fill='both', expand=True)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")])
        if not file_path:
            return

        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.csv':
                self.df = pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8-sig')
            else:
                self.df = pd.read_excel(file_path)

            bola_cols = [c for c in self.df.columns if re.search(r'bola\d?', str(c), re.IGNORECASE)]
            if len(bola_cols) < 6:
                bola_cols = self.df.select_dtypes(include='number').columns[:6]
            
            if len(bola_cols) < 6:
                raise ValueError("Não encontrei 6 colunas de bolas.")
            
            self.bolas_cols = bola_cols
            self.df[self.bolas_cols] = self.df[self.bolas_cols].apply(pd.to_numeric, errors='coerce')
            self.df.dropna(subset=self.bolas_cols, inplace=True)

            date_candidates = ['Data do Sorteio', 'Data', 'Sorteio', 'Concurso', 'date']
            self.date_col = None
            for col in date_candidates:
                if col in self.df.columns:
                    self.date_col = col
                    break
            if not self.date_col:
                for col in self.df.columns:
                    if 'data' in str(col).lower():
                        self.date_col = col
                        break

            if self.date_col:
                self.df[self.date_col] = pd.to_datetime(self.df[self.date_col], format='%d/%m/%Y', errors='coerce')
                self.df.dropna(subset=[self.date_col], inplace=True)
                
                # Define datas padrão baseadas nos dados
                min_date = self.df[self.date_col].min()
                max_date = self.df[self.date_col].max()
                self.cal_start.set_date(min_date)
                self.cal_end.set_date(max_date)
                
                self._print(f"✅ {len(self.df)} sorteios carregados")
                self._print(f"📅 Período: {min_date.strftime('%d/%m/%Y')} até {max_date.strftime('%d/%m/%Y')}\n")
            else:
                self._print("Coluna de data não encontrada\n")

            self.path_var.set(os.path.basename(file_path))

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar:\n{e}")

    def _get_filtered_data(self):
        if self.df is None:
            return None
        df_f = self.df.copy()

        start_dt = self.cal_start.get_date()
        end_dt = self.cal_end.get_date()

        self._print(f"🔍 Período selecionado: {start_dt.strftime('%d/%m/%Y')} até {end_dt.strftime('%d/%m/%Y')}")
        
        if self.date_col:
            min_data = self.df[self.date_col].min()
            max_data = self.df[self.date_col].max()
            self._print(f"📊 Dados disponíveis: {min_data.strftime('%d/%m/%Y')} até {max_data.strftime('%d/%m/%Y')}")
            
            df_f = df_f[df_f[self.date_col] >= start_dt]
            df_f = df_f[df_f[self.date_col] <= end_dt]

        count = len(df_f)
        self._print(f"✅ Sorteios encontrados: {count}\n")
        
        if count == 0:
            self._print("❌ NENHUM sorteio neste período!")
            self._print("💡 Ajuste as datas ou verifique se a planilha tem dados nesse intervalo\n")
            return None

        return [sorted(row.astype(int)) for _, row in df_f[self.bolas_cols].iterrows()], count

    def run_combo_analysis(self):
        result = self._get_filtered_data()
        if not result: return
        bolas, total_jogos = result
        
        n = int(self.combo_size.get())
        mode = self.show_mode.get()
        
        self._print(f"Analisando {total_jogos} jogos | Combinações de {n} números")
        self._print(f"Modo: {'TODAS' if mode == 'todos' else 'Só REPETIDAS'}\n")

        counter = Counter()
        for jogo in bolas:
            for combo in combinations(jogo, n):
                counter[tuple(sorted(combo))] += 1

        if mode == "repetidos":
            filtered = {k: v for k, v in counter.items() if v >= 2}
            if not filtered:
                self._print(f"Nenhuma combinação de {n} se repetiu em {total_jogos} jogos.")
                return
            display = sorted(filtered.items(), key=lambda x: x[1], reverse=True)[:30]
        else:
            display = sorted(counter.items(), key=lambda x: x[1], reverse=True)[:50]

        for combo, freq in display:
            marker = " [REPETE]" if freq >= 2 else " [1x]"
            self._print(f"  {list(combo)} -> {freq}x{marker}")
        self._print("")

    def run_range_analysis(self):
        result = self._get_filtered_data()
        if not result: return
        bolas, total_jogos = result

        try:
            r_min, r_max = int(self.range_min.get()), int(self.range_max.get())
        except:
            messagebox.showwarning("Erro", "Faixa inválida")
            return

        total = total_jogos * 6
        na_faixa = sum(1 for jogo in bolas for n in jogo if r_min <= n <= r_max)
        
        self._print(f"Faixa {r_min}-{r_max}:\n")
        self._print(f"  Total: {na_faixa} de {total}")
        self._print(f"  Porcentagem: {(na_faixa/total)*100:.2f}%\n")

    def generate_games(self):
        result = self._get_filtered_data()
        if not result: return
        bolas, total_jogos = result

        self._print(f"Gerando 5 jogos:\n")
        freq = Counter()
        for jogo in bolas: freq.update(jogo)
        
        top_nums = [n for n, _ in freq.most_common(25)]
        jogos = set()
        while len(jogos) < 5:
            jogo = tuple(sorted(random.sample(top_nums, 6)))
            jogos.add(jogo)
        
        for i, jogo in enumerate(sorted(jogos), 1):
            self._print(f"  Jogo {i}: {list(jogo)}")
        self._print("")

    def _print(self, text):
        self.text_res.config(state='normal')
        self.text_res.insert('end', text + '\n')
        self.text_res.see('end')
        self.text_res.config(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = MegaAnalyzerApp(root)
    root.mainloop()
