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
        self.root.title("🎲 Mega-Sena Analyzer Pro")
        self.root.geometry("850x950")
        self.root.resizable(False, False)

        self.df = None
        self.bolas_cols = []
        self.date_col = None
        self.range_stats = None  # Armazena análise de faixas
        self.top_combos = {'2': [], '3': [], '4': []}  # Armazena combos frequentes

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
        ttk.Label(main, text="📁 Planilha de Sorteios", font=('Segoe UI', 12, 'bold')).pack(anchor='w')
        up_frame = ttk.Frame(main)
        up_frame.pack(fill='x', pady=(5, 15))
        self.path_var = tk.StringVar(value="Nenhuma planilha selecionada")
        ttk.Label(up_frame, textvariable=self.path_var, wraplength=600).pack(side='left', fill='x', expand=True)
        ttk.Button(up_frame, text="Selecionar Arquivo", command=self.load_file).pack(side='right')

        # Filtros
        filt_frame = ttk.LabelFrame(main, text="⚙️ Filtros de Pesquisa", padding=10)
        filt_frame.pack(fill='x', pady=10)

        ttk.Label(filt_frame, text="Período:").grid(row=0, column=0, sticky='w')
        self.cal_start = DateEntry(filt_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.cal_start.grid(row=0, column=1, padx=5)
        ttk.Label(filt_frame, text="até").grid(row=0, column=2)
        self.cal_end = DateEntry(filt_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
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
        ttk.Checkbutton(filt_frame, text="Mostrar TODAS as combinações", variable=self.show_mode, onvalue="todos", offvalue="repetidos").grid(row=2, column=0, columnspan=6, sticky='w', pady=(10,0))

        # Botões de Análise
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill='x', pady=15)
        ttk.Button(btn_frame, text="🔍 Analisar Repetições", command=self.run_combo_analysis).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="📊 Analisar Faixa %", command=self.run_range_analysis).pack(side='left', padx=5)

        # 🎲 Frame do Gerador
        gen_frame = ttk.LabelFrame(main, text="🎲 Gerador Inteligente", padding=10)
        gen_frame.pack(fill='x', pady=10)
        ttk.Button(gen_frame, text="⚙️ Configurar e Gerar Jogos", command=self._open_generator_config, width=25).pack()

        # Resultados
        res_frame = ttk.LabelFrame(main, text="📋 Resultados", padding=10)
        res_frame.pack(fill='both', expand=True, pady=10)
        self.text_res = tk.Text(res_frame, height=18, font=('Consolas', 9), state='disabled')
        scrollbar = ttk.Scrollbar(res_frame, orient="vertical", command=self.text_res.yview)
        self.text_res.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        self.text_res.pack(side='left', fill='both', expand=True)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")])
        if not file_path: return

        try:
            ext = os.path.splitext(file_path)[1].lower()
            self.df = pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8-sig') if ext == '.csv' else pd.read_excel(file_path)

            bola_cols = [c for c in self.df.columns if re.search(r'bola\d?', str(c), re.IGNORECASE)]
            if len(bola_cols) < 6:
                bola_cols = self.df.select_dtypes(include='number').columns[:6]
            if len(bola_cols) < 6:
                raise ValueError("Não encontrei 6 colunas de bolas.")
            
            self.bolas_cols = bola_cols
            self.df[self.bolas_cols] = self.df[self.bolas_cols].apply(pd.to_numeric, errors='coerce')
            self.df.dropna(subset=self.bolas_cols, inplace=True)

            date_candidates = ['Data do Sorteio', 'Data', 'Sorteio', 'Concurso', 'date']
            self.date_col = next((c for c in date_candidates if c in self.df.columns), None)
            if not self.date_col:
                self.date_col = next((c for c in self.df.columns if 'data' in str(c).lower()), None)

            if self.date_col:
                self.df[self.date_col] = pd.to_datetime(self.df[self.date_col], format='%d/%m/%Y', errors='coerce')
                self.df.dropna(subset=[self.date_col], inplace=True)
                self.cal_start.set_date(self.df[self.date_col].min())
                self.cal_end.set_date(self.df[self.date_col].max())
                self._print(f"✅ {len(self.df)} sorteios carregados | Período: {self.df[self.date_col].min().strftime('%d/%m/%Y')} até {self.df[self.date_col].max().strftime('%d/%m/%Y')}\n")
            else:
                self._print("⚠️ Coluna de data não encontrada\n")

            self.path_var.set(os.path.basename(file_path))
            self.range_stats = None  # Resetar stats ao carregar novo arquivo
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar:\n{e}")

    def _get_filtered_data(self):
        if self.df is None: return None
        df_f = self.df.copy()
        start_dt, end_dt = pd.Timestamp(self.cal_start.get_date()), pd.Timestamp(self.cal_end.get_date())
        
        if self.date_col:
            df_f = df_f[(df_f[self.date_col] >= start_dt) & (df_f[self.date_col] <= end_dt)]
        
        count = len(df_f)
        if count == 0:
            self._print("❌ Nenhum sorteio encontrado no período selecionado.\n")
            return None
        return [sorted(row.astype(int)) for _, row in df_f[self.bolas_cols].iterrows()], count

    def run_combo_analysis(self):
        result = self._get_filtered_data()
        if not result: return
        bolas, total = result
        n = int(self.combo_size.get())
        mode = self.show_mode.get()
        
        self._print(f"🔍 Analisando {total} jogos | Combinações de {n} números | Modo: {'TODAS' if mode=='todos' else 'REPETIDAS'}\n")
        counter = Counter()
        for jogo in bolas:
            for combo in combinations(jogo, n):
                counter[tuple(sorted(combo))] += 1

        # Armazenar top combos para o gerador
        self.top_combos[str(n)] = [c for c, _ in counter.most_common(30)]

        filtered = counter if mode == 'todos' else {k:v for k,v in counter.items() if v>=2}
        if not filtered:
            self._print("Nenhuma combinação encontrada.\n")
            return

        display = sorted(filtered.items(), key=lambda x: x[1], reverse=True)[:30]
        for combo, freq in display:
            self._print(f"  {list(combo)} -> {freq}x{' [REPETE]' if freq>=2 else ''}")
        self._print("")

    def run_range_analysis(self):
        result = self._get_filtered_data()
        if not result: return
        bolas, total_jogos = result

        # Cálculo automático 10 em 10
        ranges_10 = [(1,10), (11,20), (21,30), (31,40), (41,50), (51,60)]
        stats_10 = []
        for r in ranges_10:
            cnt = sum(1 for j in bolas for n in j if r[0] <= n <= r[1])
            stats_10.append({'range': f"{r[0]}-{r[1]}", 'avg': round(cnt/total_jogos, 2), 'total': cnt})

        # Cálculo automático 20 em 20
        ranges_20 = [(1,20), (21,40), (41,60)]
        stats_20 = []
        for r in ranges_20:
            cnt = sum(1 for j in bolas for n in j if r[0] <= n <= r[1])
            stats_20.append({'range': f"{r[0]}-{r[1]}", 'avg': round(cnt/total_jogos, 2), 'total': cnt})

        self.range_stats = {'step10': stats_10, 'step20': stats_20}

        self._print("📊 ANÁLISE AUTOMÁTICA DE DISTRIBUIÇÃO POR FAIXAS:\n")
        self._print("🔹 Faixas de 10 em 10 (Média por jogo):")
        for s in stats_10:
            self._print(f"   {s['range']}: {s['avg']} números")
        self._print("\n Faixas de 20 em 20 (Média por jogo):")
        for s in stats_20:
            self._print(f"   {s['range']}: {s['avg']} números")
        self._print("\n Clique em '⚙️ Configurar e Gerar Jogos' para usar esses padrões.\n")

    def _open_generator_config(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Carregue uma planilha primeiro.")
            return
        if not self.range_stats:
            if not messagebox.askyesno("Aviso", "Você ainda não clicou em 'Analisar Faixa %'. Deseja continuar com padrões padrão?"):
                return

        win = tk.Toplevel(self.root)
        win.title("⚙️ Configuração do Gerador")
        win.geometry("420x520")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text="Quantidade de Jogos:", font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=15, pady=(10,0))
        qtd_spin = ttk.Spinbox(win, from_=1, to=100, width=10)
        qtd_spin.set(5)
        qtd_spin.pack(anchor='w', padx=15)

        ttk.Label(win, text="Combinações base (sementes):", font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=15, pady=(15,0))
        u_dupla = tk.BooleanVar(value=True)
        u_tripla = tk.BooleanVar(value=True)
        u_quadra = tk.BooleanVar(value=False)
        ttk.Checkbutton(win, text="✅ Duplas (2 números)", variable=u_dupla).pack(anchor='w', padx=30)
        ttk.Checkbutton(win, text="✅ Triplas (3 números)", variable=u_tripla).pack(anchor='w', padx=30)
        ttk.Checkbutton(win, text="Quadras (4 números)", variable=u_quadra).pack(anchor='w', padx=30)

        ttk.Label(win, text="Distribuição por Faixas:", font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=15, pady=(15,0))
        dist_var = tk.StringVar(value="historico")
        ttk.Radiobutton(win, text=" Padrão histórico (mais frequente)", variable=dist_var, value="historico").pack(anchor='w', padx=30)
        ttk.Radiobutton(win, text="📏 Faixas de 10 em 10 (1-10, 11-20...)", variable=dist_var, value="step10").pack(anchor='w', padx=30)
        ttk.Radiobutton(win, text="📐 Faixas de 20 em 20 (1-20, 21-40...)", variable=dist_var, value="step20").pack(anchor='w', padx=30)

        ttk.Button(win, text="🎲 GERAR JOGOS AGORA", command=lambda: self.generate_smart_games(
            int(qtd_spin.get()), u_dupla.get(), u_tripla.get(), u_quadra.get(), dist_var.get(), win
        )).pack(pady=25, fill='x', padx=20)

    def generate_smart_games(self, qtd, use_dupla, use_tripla, use_quadra, dist_mode, config_win):
        result = self._get_filtered_data()
        if not result: return
        bolas, total_jogos = result
        config_win.destroy()

        self._print(f"🎲 GERANDO {qtd} JOGOS INTELIGENTES...\n")
        self._print(f"Base: {total_jogos} sorteios | Combinações: Dupla={use_dupla}, Tripla={use_tripla}, Quadra={use_quadra}")
        self._print(f"Distribuição: {dist_mode}\n")

        # 1. Pool de combinações
        combo_pool = []
        if use_dupla: combo_pool.extend(self.top_combos.get('2', []))
        if use_tripla: combo_pool.extend(self.top_combos.get('3', []))
        if use_quadra: combo_pool.extend(self.top_combos.get('4', []))
        if not combo_pool:
            self._print("⚠️ Nenhuma combinação selecionada. Usando frequência simples.\n")

        # 2. Definir alvos de faixa
        targets = []
        if dist_mode == "historico" and self.range_stats:
            # Pega o step que teve maior aderência histórica (soma dos averages mais próxima de 6)
            s10_sum = sum(s['avg'] for s in self.range_stats['step10'])
            s20_sum = sum(s['avg'] for s in self.range_stats['step20'])
            chosen = 'step10' if abs(s10_sum - 6) < abs(s20_sum - 6) else 'step20'
            stats = self.range_stats[chosen]
            targets = [s['avg'] for s in stats]
            ranges = [(1,10), (11,20), (21,30), (31,40), (41,50), (51,60)] if chosen=='step10' else [(1,20), (21,40), (41,60)]
            self._print(f"📊 Usando padrão histórico ({chosen}): {targets}\n")
        elif dist_mode == "step10":
            targets = [1.0]*6
            ranges = [(1,10), (11,20), (21,30), (31,40), (41,50), (51,60)]
        elif dist_mode == "step20":
            targets = [2.0, 2.0, 2.0]
            ranges = [(1,20), (21,40), (41,60)]
        else:
            targets = [1.0]*6
            ranges = [(1,10), (11,20), (21,30), (31,40), (41,50), (51,60)]

        # Frequência individual para fallback
        freq_individual = Counter()
        for j in bolas: freq_individual.update(j)
        hot_numbers = [n for n, _ in freq_individual.most_common(30)]

        # 3. Geração
        jogos_gerados = set()
        tentativas = 0
        while len(jogos_gerados) < qtd and tentativas < qtd * 200:
            tentativas += 1
            jogo = []

            # Semente de combinação
            if combo_pool and random.random() < 0.75:
                seed = random.choice(combo_pool)
                jogo = list(seed)

            # Preencher até 6 números respeitando faixas
            while len(jogo) < 6:
                # Determinar qual faixa precisa de número
                current_dist = [0]*len(targets)
                for n in jogo:
                    for i, (min_v, max_v) in enumerate(ranges):
                        if min_v <= n <= max_v:
                            current_dist[i] += 1
                
                # Escolher faixa com maior déficit
                deficits = [targets[i] - current_dist[i] for i in range(len(targets))]
                best_idx = deficits.index(max(deficits))
                min_v, max_v = ranges[best_idx]

                # Buscar número quente nessa faixa
                candidatos = [n for n in hot_numbers if min_v <= n <= max_v and n not in jogo]
                if not candidatos:
                    candidatos = [n for n in range(1, 61) if min_v <= n <= max_v and n not in jogo]
                
                if candidatos:
                    jogo.append(random.choice(candidatos))
                else:
                    break # Fallback se não achar

            if len(jogo) == 6:
                t_jogo = tuple(sorted(jogo))
                if t_jogo not in jogos_gerados:
                    jogos_gerados.add(t_jogo)

        # 4. Saída
        if jogos_gerados:
            self._print(f"✅ {len(jogos_gerados)} JOGOS GERADOS:\n")
            for i, j in enumerate(sorted(jogos_gerados), 1):
                self._print(f"  Jogo {i:2d}: {list(j)}")
            self._print(f"\n Tentativas: {tentativas} | Jogos únicos: {len(jogos_gerados)}")
        else:
            self._print(" Não foi possível gerar jogos únicos com essas restrições.\n")

    def _print(self, text):
        self.text_res.config(state='normal')
        self.text_res.insert('end', text + '\n')
        self.text_res.see('end')
        self.text_res.config(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = MegaAnalyzerApp(root)
    root.mainloop()
