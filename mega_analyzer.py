import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from itertools import combinations
from collections import Counter
import random
import os
import re
from datetime import datetime

class MegaAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎲 Mega-Sena Analyzer Pro")
        self.root.geometry("750x850")
        self.root.resizable(False, False)

        self.df = None
        self.bolas_cols = []
        self.date_col = None
        self.filtered_bolas = []

        self._setup_ui()

    def _setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', padding=6, font=('Segoe UI', 10))
        style.configure('TLabel', font=('Segoe UI', 10))
        style.configure('TEntry', padding=5)
        style.configure('TFrame', background='#f0f0f0')
        self.root.configure(bg='#f0f0f0')

        main = ttk.Frame(self.root, padding=15)
        main.pack(fill='both', expand=True)

        # 1. Upload
        ttk.Label(main, text="📁 Planilha de Sorteios", font=('Segoe UI', 12, 'bold')).pack(anchor='w')
        up_frame = ttk.Frame(main)
        up_frame.pack(fill='x', pady=(5, 15))
        self.path_var = tk.StringVar(value="Nenhuma planilha selecionada")
        ttk.Label(up_frame, textvariable=self.path_var, wraplength=500).pack(side='left', fill='x', expand=True)
        ttk.Button(up_frame, text="Selecionar Arquivo", command=self.load_file).pack(side='right')

        # 2. Filtros
        filt_frame = ttk.LabelFrame(main, text="⚙️ Filtros de Pesquisa", padding=10)
        filt_frame.pack(fill='x', pady=10)

        ttk.Label(filt_frame, text="Período (opcional):").grid(row=0, column=0, sticky='w')
        self.date_start = ttk.Entry(filt_frame, width=12)
        self.date_start.grid(row=0, column=1, padx=5)
        ttk.Label(filt_frame, text="até").grid(row=0, column=2)
        self.date_end = ttk.Entry(filt_frame, width=12)
        self.date_end.grid(row=0, column=3, padx=5)
        ttk.Label(filt_frame, text="(DD/MM/AAAA)").grid(row=0, column=4, sticky='w')

        ttk.Label(filt_frame, text="Números juntos:").grid(row=1, column=0, sticky='w', pady=(10,0))
        self.combo_size = ttk.Combobox(filt_frame, values=[2, 3, 4, 5, 6], state='readonly', width=5)
        self.combo_size.set(3)
        self.combo_size.grid(row=1, column=1, padx=5, pady=(10,0))

        ttk.Label(filt_frame, text="Faixa % personalizada:").grid(row=1, column=2, sticky='w', pady=(10,0))
        self.range_min = ttk.Entry(filt_frame, width=4)
        self.range_min.insert(0, "1")
        self.range_min.grid(row=1, column=3, padx=5, pady=(10,0))
        ttk.Label(filt_frame, text="até").grid(row=1, column=4, pady=(10,0))
        self.range_max = ttk.Entry(filt_frame, width=4)
        self.range_max.insert(0, "60")
        self.range_max.grid(row=1, column=5, padx=5, pady=(10,0))

        # 3. Botões
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill='x', pady=15)
        ttk.Button(btn_frame, text="🔍 Analisar Combinações", command=self.run_combo_analysis).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="📊 Analisar Faixa %", command=self.run_range_analysis).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="🎲 Gerar Jogos Otimizados", command=self.generate_games).pack(side='left', padx=5)

        # 4. Resultados
        res_frame = ttk.LabelFrame(main, text="📋 Resultados", padding=10)
        res_frame.pack(fill='both', expand=True, pady=10)
        self.text_res = tk.Text(res_frame, height=20, font=('Consolas', 10), state='disabled')
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
            self.df = pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8-sig') if ext == '.csv' else pd.read_excel(file_path)

            # Detectar colunas de bolas
            bola_cols = [c for c in self.df.columns if re.search(r'bola|n[º°]?|dezena', str(c), re.IGNORECASE)]
            if len(bola_cols) < 6:
                bola_cols = self.df.select_dtypes(include='number').columns[:6]
            if len(bola_cols) < 6:
                raise ValueError("Planilha não contém 6 colunas numéricas identificáveis.")
            
            self.bolas_cols = bola_cols
            self.df[self.bolas_cols] = self.df[self.bolas_cols].apply(pd.to_numeric, errors='coerce')
            self.df.dropna(subset=self.bolas_cols, inplace=True)

            # 🔧 CORREÇÃO: Detecção e conversão de data robusta
            date_keywords = ['data', 'sorteio', 'concurso', 'date', 'dia']
            possible_date_cols = [c for c in self.df.columns if any(k in str(c).lower() for k in date_keywords)]
            self.date_col = possible_date_cols[0] if possible_date_cols else None

            if self.date_col:
                # Força formato BR e ignora linhas com data inválida
                self.df[self.date_col] = pd.to_datetime(self.df[self.date_col], dayfirst=True, errors='coerce')
                self.df.dropna(subset=[self.date_col], inplace=True)
                self._print(f"📅 Coluna de data detectada: `{self.date_col}`")
            else:
                self._print("️ Nenhuma coluna de data encontrada. Filtros de período serão ignorados.")

            self.path_var.set(os.path.basename(file_path))
            self._print(f"✅ Planilha carregada: {len(self.df)} sorteios válidos.\n")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar planilha:\n{e}")

    def _get_filtered_data(self):
        if self.df is None:
            return None
        df_f = self.df.copy()

        start = self.date_start.get().strip()
        end = self.date_end.get().strip()

        if self.date_col and (start or end):
            try:
                mask = pd.Series([True] * len(df_f))
                if start:
                    start_dt = pd.to_datetime(start, dayfirst=True)
                    mask &= df_f[self.date_col] >= start_dt
                if end:
                    end_dt = pd.to_datetime(end, dayfirst=True)
                    mask &= df_f[self.date_col] <= end_dt
                df_f = df_f[mask]
            except Exception as e:
                messagebox.showwarning("Erro de Data", f"Formato de data inválido. Use DD/MM/AAAA.\n{e}")
                return None

        if df_f.empty:
            self._print("️ Nenhum sorteio encontrado no período selecionado.")
            return None

        return [sorted(row.astype(int)) for _, row in df_f[self.bolas_cols].iterrows()]

    def run_combo_analysis(self):
        bolas = self._get_filtered_data()
        if not bolas: return
        n = int(self.combo_size.get())
        self._print(f" Analisando combinações de {n} números...\n")

        counter = Counter()
        for jogo in bolas:
            for combo in combinations(jogo, n):
                counter[tuple(sorted(combo))] += 1

        top = counter.most_common(15)
        for combo, freq in top:
            self._print(f"  {list(combo)} → {freq} vezes")
        self._print("")

    def run_range_analysis(self):
        bolas = self._get_filtered_data()
        if not bolas: return

        try:
            r_min, r_max = int(self.range_min.get()), int(self.range_max.get())
            if not (1 <= r_min < r_max <= 60):
                raise ValueError
        except:
            messagebox.showwarning("Intervalo Inválido", "Informe números entre 1 e 60. Ex: 11 até 30")
            return

        self._print(f" Análise da faixa {r_min}-{r_max}:\n")
        total_jogos = len(bolas)
        total_nums = total_jogos * 6
        contagem = 0

        for jogo in bolas:
            for n in jogo:
                if r_min <= n <= r_max:
                    contagem += 1

        pct = (contagem / total_nums) * 100
        self._print(f"  • Total de números na faixa: {contagem}")
        self._print(f"  • Porcentagem sobre todos os números sorteados: {pct:.2f}%")
        self._print(f"  • Média por jogo: {contagem/total_jogos:.2f} números")
        self._print("")

    def generate_games(self):
        bolas = self._get_filtered_data()
        if not bolas: return

        n = int(self.combo_size.get())
        self._print(f"🎲 Gerando jogos baseados nos top números (combinações de {n})...\n")

        freq_all = Counter()
        for jogo in bolas:
            freq_all.update(jogo)
        
        top_numbers = [num for num, _ in freq_all.most_common(25)]

        jogos_gerados = set()
        attempts = 0
        while len(jogos_gerados) < 5 and attempts < 1000:
            jogo = sorted(random.sample(top_numbers, 6))
            if tuple(jogo) not in jogos_gerados:
                jogos_gerados.add(tuple(jogo))
            attempts += 1

        for i, jogo in enumerate(sorted(jogos_gerados), 1):
            self._print(f"  Jogo {i}: {list(jogo)}")
        self._print("\n💡 Dica: Use esses números como base e personalize conforme sua estratégia.")

    def _print(self, text):
        self.text_res.config(state='normal')
        self.text_res.insert('end', text + '\n')
        self.text_res.see('end')
        self.text_res.config(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = MegaAnalyzerApp(root)
    root.mainloop()
