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
from openpyxl.utils import get_column_letter

class MegaAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎲 Mega-Sena Analyzer Pro")
        self.root.geometry("950x950")
        self.root.resizable(False, False)

        self.df = None
        self.bolas_cols = []
        self.date_col = None
        
        # Dados armazenados para exportação
        self.range_stats = None
        self.top_combos = {'2': [], '3': [], '4': []}
        self.freq_tiers = None  # {'high': [], 'medium': [], 'low': []}
        self.freq_pattern = None  # Padrão histórico (ex: [2,2,2])
        self.session_data = {
            'combos': None,
            'ranges': None,
            'games': [],
            'frequency': None,
            'freq_comparison': None,
            'meta': {}
        }

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
        ttk.Label(up_frame, textvariable=self.path_var, wraplength=680).pack(side='left', fill='x', expand=True)
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

        # Botões
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill='x', pady=15)
        ttk.Button(btn_frame, text="🔍 Repetições", command=self.run_combo_analysis).pack(side='left', padx=5)
        ttk.Button(btn_frame, text=" Faixas", command=self.run_range_analysis).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="🔢 Frequência", command=self.run_frequency_analysis).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="📊 Comparar Freq", command=self.run_frequency_comparison).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="🎲 Gerador", command=self._open_generator_config).pack(side='left', padx=5)
        ttk.Button(btn_frame, text=" Exportar Excel", command=self.export_to_excel).pack(side='right', padx=5)

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
            self.session_data = {'combos': None, 'ranges': None, 'games': [], 'frequency': None, 'freq_comparison': None, 'meta': {}}
            self.freq_tiers = None
            self.freq_pattern = None
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
        
        self.session_data['meta'].update({
            "Periodo_Inicio": start_dt.strftime("%d/%m/%Y"),
            "Periodo_Fim": end_dt.strftime("%d/%m/%Y"),
            "Total_Sorteios_Analisados": count,
            "Data_Analise": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        })
        
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

        self.top_combos[str(n)] = counter.most_common(30)
        self.session_data['combos'] = counter.most_common(50)

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

        ranges_10 = [(1,10), (11,20), (21,30), (31,40), (41,50), (51,60)]
        stats_10 = [{'Faixa': f"{r[0]}-{r[1]}", 'Média_por_Jogo': round(sum(1 for j in bolas for n in j if r[0] <= n <= r[1])/total_jogos, 2)} for r in ranges_10]

        ranges_20 = [(1,20), (21,40), (41,60)]
        stats_20 = [{'Faixa': f"{r[0]}-{r[1]}", 'Média_por_Jogo': round(sum(1 for j in bolas for n in j if r[0] <= n <= r[1])/total_jogos, 2)} for r in ranges_20]

        self.range_stats = {'step10': stats_10, 'step20': stats_20}
        self.session_data['ranges'] = {'step10': stats_10, 'step20': stats_20}

        self._print("📊 ANÁLISE AUTOMÁTICA DE DISTRIBUIÇÃO POR FAIXAS:\n")
        self._print("🔹 Faixas de 10 em 10 (Média por jogo):")
        for s in stats_10: self._print(f"   {s['Faixa']}: {s['Média_por_Jogo']} números")
        self._print("\n🔹 Faixas de 20 em 20 (Média por jogo):")
        for s in stats_20: self._print(f"   {s['Faixa']}: {s['Média_por_Jogo']} números")
        self._print("")

    def run_frequency_analysis(self):
        result = self._get_filtered_data()
        if not result: return
        bolas, total_jogos = result

        freq = Counter()
        for jogo in bolas:
            freq.update(jogo)

        full_freq = {n: freq.get(n, 0) for n in range(1, 61)}
        sorted_freq = sorted(full_freq.items(), key=lambda x: x[1], reverse=True)
        total_nums = total_jogos * 6

        self._print(f"🔢 FREQUÊNCIA POR NÚMERO ({total_jogos} jogos | {total_nums} números sorteados)\n")
        self._print("🔥 TOP 10 MAIS FREQUENTES:")
        for num, count in sorted_freq[:10]:
            pct = (count / total_nums) * 100
            self._print(f"   Nº {num:02d} → {count:3d} vezes ({pct:.2f}%)")

        self._print("\n❄️ TOP 10 MENOS FREQUENTES:")
        for num, count in sorted_freq[-10:]:
            pct = (count / total_nums) * 100
            self._print(f"   Nº {num:02d} → {count:3d} vezes ({pct:.2f}%)")

        self._print("\n📊 LISTA COMPLETA (Ordenada por frequência):")
        for num, count in sorted_freq:
            pct = (count / total_nums) * 100
            self._print(f"   Nº {num:02d} → {count:3d} vezes ({pct:.2f}%)")
        self._print("")

        self.session_data['frequency'] = sorted_freq
        self.session_data['meta']['Total_Numeros_Analisados'] = total_nums

    def run_frequency_comparison(self):
        """Análise comparativa: quantos números de cada tier (alta/média/baixa) aparecem nos sorteios"""
        result = self._get_filtered_data()
        if not result: return
        bolas, total_jogos = result

        # 1. Calcular frequência e criar tiers
        freq = Counter()
        for jogo in bolas:
            freq.update(jogo)
        
        sorted_nums = sorted(freq.items(), key=lambda x: x[1], reverse=True)
        
        # Dividir em 3 tiers: top 20, middle 20, bottom 20
        high_tier = [n for n, _ in sorted_nums[:20]]
        medium_tier = [n for n, _ in sorted_nums[20:40]]
        low_tier = [n for n, _ in sorted_nums[40:]]
        
        self.freq_tiers = {'high': set(high_tier), 'medium': set(medium_tier), 'low': set(low_tier)}

        # 2. Analisar cada sorteio histórico
        pattern_counter = Counter()
        for jogo in bolas:
            h = sum(1 for n in jogo if n in self.freq_tiers['high'])
            m = sum(1 for n in jogo if n in self.freq_tiers['medium'])
            l = sum(1 for n in jogo if n in self.freq_tiers['low'])
            pattern_counter[(h, m, l)] += 1

        # 3. Encontrar padrão mais comum
        most_common = pattern_counter.most_common(5)
        self.freq_pattern = most_common[0][0] if most_common else (2, 2, 2)

        # 4. Exibir resultados
        self._print(f"📊 COMPARATIVO DE FREQUÊNCIA POR TIER ({total_jogos} jogos)\n")
        self._print("🔹 Classificação dos números:")
        self._print(f"   🔥 Alta (Top 20): {sorted(high_tier)}")
        self._print(f"   ⚡ Média (20-40): {sorted(medium_tier)}")
        self._print(f"   ❄️ Baixa (40-60): {sorted(low_tier)}\n")

        self._print("🔹 Padrões encontrados nos sorteios (High, Medium, Low):")
        for pattern, count in most_common:
            pct = (count / total_jogos) * 100
            marker = " ✅ MAIS COMUM" if pattern == self.freq_pattern else ""
            self._print(f"   {pattern} → {count} jogos ({pct:.1f}%){marker}")

        self._print(f"\n💡 Padrão selecionado para o gerador: {self.freq_pattern}")
        self._print(f"   (Ex: {self.freq_pattern[0]} altos + {self.freq_pattern[1]} médios + {self.freq_pattern[2]} baixos)\n")

        # Salvar para exportação
        self.session_data['freq_comparison'] = {
            'tiers': {'high': high_tier, 'medium': medium_tier, 'low': low_tier},
            'patterns': most_common,
            'selected_pattern': self.freq_pattern
        }

    def _open_generator_config(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Carregue uma planilha primeiro.")
            return

        win = tk.Toplevel(self.root)
        win.title("⚙️ Configuração do Gerador")
        win.geometry("500x680")
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

        focus_top_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(win, text="🔒 Focar apenas nas TOP combinações", variable=focus_top_var).pack(anchor='w', padx=30, pady=(5,0))

        ttk.Label(win, text="Distribuição por Faixas:", font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=15, pady=(15,0))
        dist_var = tk.StringVar(value="historico")
        ttk.Radiobutton(win, text="🤖 Padrão histórico (mais frequente)", variable=dist_var, value="historico").pack(anchor='w', padx=30)
        ttk.Radiobutton(win, text="📏 Faixas de 10 em 10", variable=dist_var, value="step10").pack(anchor='w', padx=30)
        ttk.Radiobutton(win, text="📐 Faixas de 20 em 20", variable=dist_var, value="step20").pack(anchor='w', padx=30)

        # 🆕 Novo: Filtro por padrão de frequência
        use_freq_pattern_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(win, text="🎯 Usar padrão de frequência (Alta/Média/Baixa)", 
                       variable=use_freq_pattern_var).pack(anchor='w', padx=30, pady=(10,0))
        if self.freq_pattern:
            ttk.Label(win, text=f"   Padrão detectado: {self.freq_pattern[0]}🔥 + {self.freq_pattern[1]}⚡ + {self.freq_pattern[2]}❄️", 
                     font=('Segoe UI', 9)).pack(anchor='w', padx=45)

        ttk.Button(win, text="🎲 GERAR JOGOS AGORA", command=lambda: self.generate_smart_games(
            int(qtd_spin.get()), u_dupla.get(), u_tripla.get(), u_quadra.get(), 
            dist_var.get(), focus_top_var.get(), use_freq_pattern_var.get(), win
        )).pack(pady=25, fill='x', padx=20)

    def generate_smart_games(self, qtd, use_dupla, use_tripla, use_quadra, dist_mode, focus_top, use_freq_pattern, config_win):
        result = self._get_filtered_data()
        if not result: return
        bolas, total_jogos = result
        config_win.destroy()

        self._print(f"🎲 GERANDO {qtd} JOGOS INTELIGENTES...\n")
        
        # Pool ponderado de combinações
        combo_pool = []
        weights = []
        limit = 5 if focus_top else 30
        
        if use_dupla and self.top_combos.get('2'):
            for c, f in self.top_combos['2'][:limit]: combo_pool.append(c); weights.append(f)
        if use_tripla and self.top_combos.get('3'):
            for c, f in self.top_combos['3'][:limit]: combo_pool.append(c); weights.append(f)
        if use_quadra and self.top_combos.get('4'):
            for c, f in self.top_combos['4'][:limit]: combo_pool.append(c); weights.append(f)

        # Configurar alvos de faixa numérica
        targets, ranges = [], []
        if dist_mode == "historico" and self.range_stats:
            s10_sum = sum(s['Média_por_Jogo'] for s in self.range_stats['step10'])
            s20_sum = sum(s['Média_por_Jogo'] for s in self.range_stats['step20'])
            chosen = 'step10' if abs(s10_sum - 6) < abs(s20_sum - 6) else 'step20'
            stats = self.range_stats[chosen]
            targets = [s['Média_por_Jogo'] for s in stats]
            ranges = [(1,10), (11,20), (21,30), (31,40), (41,50), (51,60)] if chosen=='step10' else [(1,20), (21,40), (41,60)]
        elif dist_mode == "step10": targets, ranges = [1.0]*6, [(1,10), (11,20), (21,30), (31,40), (41,50), (51,60)]
        elif dist_mode == "step20": targets, ranges = [2.0, 2.0, 2.0], [(1,20), (21,40), (41,60)]
        else: targets, ranges = [1.0]*6, [(1,10), (11,20), (21,30), (31,40), (41,50), (51,60)]

        # Frequência individual para fallback
        freq_individual = Counter()
        for j in bolas: freq_individual.update(j)
        hot_numbers = [n for n, _ in freq_individual.most_common(30)]

        # Configurar tiers de frequência se ativado
        tier_targets = None
        if use_freq_pattern and self.freq_tiers and self.freq_pattern:
            tier_targets = list(self.freq_pattern)  # Ex: [2, 2, 2]
            self._print(f"🎯 Filtro de frequência ativado: {tier_targets[0]}🔥 + {tier_targets[1]}⚡ + {tier_targets[2]}❄️\n")

        # Geração dos jogos
        jogos_gerados = set()
        generated_list = []
        tentativas = 0

        while len(jogos_gerados) < qtd and tentativas < qtd * 400:
            tentativas += 1
            jogo = []
            seed_combo = None
            seed_freq = 0

            # Selecionar semente de combinação
            if combo_pool and random.random() < 0.80:
                chosen_idx = random.choices(range(len(combo_pool)), weights=weights, k=1)[0]
                seed_combo = combo_pool[chosen_idx]
                seed_freq = weights[chosen_idx]
                jogo = list(seed_combo)

            # Preencher jogo respeitando filtros
            while len(jogo) < 6:
                # 1. Verificar filtro de tiers de frequência (se ativado)
                if tier_targets and self.freq_tiers:
                    current_tiers = [
                        sum(1 for n in jogo if n in self.freq_tiers['high']),
                        sum(1 for n in jogo if n in self.freq_tiers['medium']),
                        sum(1 for n in jogo if n in self.freq_tiers['low'])
                    ]
                    # Escolher tier com maior déficit
                    deficits = [tier_targets[i] - current_tiers[i] for i in range(3)]
                    if max(deficits) > 0:
                        best_tier_idx = deficits.index(max(deficits))
                        if best_tier_idx == 0:
                            candidatos = [n for n in self.freq_tiers['high'] if n not in jogo]
                        elif best_tier_idx == 1:
                            candidatos = [n for n in self.freq_tiers['medium'] if n not in jogo]
                        else:
                            candidatos = [n for n in self.freq_tiers['low'] if n not in jogo]
                        
                        if candidatos:
                            # Priorizar números quentes dentro do tier
                            hot_in_tier = [n for n in candidatos if n in hot_numbers]
                            escolha = random.choice(hot_in_tier if hot_in_tier else candidatos)
                            jogo.append(escolha)
                            continue

                # 2. Verificar filtro de faixas numéricas
                if ranges:
                    current_dist = [0]*len(targets)
                    for n in jogo:
                        for i, (min_v, max_v) in enumerate(ranges):
                            if min_v <= n <= max_v: current_dist[i] += 1
                    
                    deficits = [targets[i] - current_dist[i] for i in range(len(targets))]
                    if max(deficits) > 0:
                        best_idx = deficits.index(max(deficits))
                        min_v, max_v = ranges[best_idx]
                        candidatos = [n for n in hot_numbers if min_v <= n <= max_v and n not in jogo]
                        if not candidatos:
                            candidatos = [n for n in range(1, 61) if min_v <= n <= max_v and n not in jogo]
                        if candidatos:
                            jogo.append(random.choice(candidatos))
                            continue

                # 3. Fallback: qualquer número quente disponível
                candidatos = [n for n in hot_numbers if n not in jogo]
                if not candidatos:
                    candidatos = [n for n in range(1, 61) if n not in jogo]
                if candidatos:
                    jogo.append(random.choice(candidatos))

            # Validar e salvar jogo
            if len(jogo) == 6:
                t_jogo = tuple(sorted(jogo))
                if t_jogo not in jogos_gerados:
                    jogos_gerados.add(t_jogo)
                    generated_list.append({
                        "Jogo": len(generated_list) + 1,
                        "Bola1": t_jogo[0], "Bola2": t_jogo[1], "Bola3": t_jogo[2],
                        "Bola4": t_jogo[3], "Bola5": t_jogo[4], "Bola6": t_jogo[5],
                        "Semente": str(list(seed_combo)) if seed_combo else "Aleatória",
                        "Freq_Semente": seed_freq if seed_combo else 0,
                        "Distrib_Freq": f"{sum(1 for n in t_jogo if n in (self.freq_tiers['high'] or set()))}🔥+{sum(1 for n in t_jogo if n in (self.freq_tiers['medium'] or set()))}⚡+{sum(1 for n in t_jogo if n in (self.freq_tiers['low'] or set()))}❄️" if use_freq_pattern else "-"
                    })

        # Salvar para exportação
        self.session_data['games'] = generated_list
        self.session_data['meta']['Jogos_Gerados'] = qtd
        self.session_data['meta']['Modo_Distribuicao'] = dist_mode
        self.session_data['meta']['Usar_Padrao_Frequencia'] = "Sim" if use_freq_pattern else "Não"

        # Exibir resultados
        if generated_list:
            self._print(f"✅ {len(generated_list)} JOGOS GERADOS:\n")
            for item in generated_list:
                freq_info = f" | {item['Distrib_Freq']}" if use_freq_pattern and item['Distrib_Freq'] != "-" else ""
                self._print(f"  Jogo {item['Jogo']:2d}: [{item['Bola1']}, {item['Bola2']}, {item['Bola3']}, {item['Bola4']}, {item['Bola5']}, {item['Bola6']}]{freq_info}")
            self._print(f"\n💾 Clique em 'Exportar Excel' para salvar TUDO.\n")
        else:
            self._print("❌ Não foi possível gerar jogos únicos com essas restrições.\n")

    def export_to_excel(self):
        if all(v is None or v == [] for v in [self.session_data['combos'], self.session_data['ranges'], self.session_data['games'], self.session_data['frequency'], self.session_data['freq_comparison']]):
            messagebox.showwarning("Exportar", "Nenhuma análise foi realizada nesta sessão.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile=f"Relatorio_MegaSena_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        if not file_path: return

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Aba Configuração
                if self.session_data['meta']:
                    pd.DataFrame(list(self.session_data['meta'].items()), columns=['Parâmetro', 'Valor']).to_excel(writer, sheet_name='Configuração', index=False)
                
                # Aba Jogos Gerados
                if self.session_data['games']:
                    pd.DataFrame(self.session_data['games']).to_excel(writer, sheet_name='Jogos Gerados', index=False)
                
                # Aba Combinações
                if self.session_data['combos']:
                    pd.DataFrame([(str(list(c)), f) for c, f in self.session_data['combos']], columns=['Combinação', 'Frequência']).to_excel(writer, sheet_name='Análise Combinações', index=False)
                
                # Aba Faixas
                if self.session_data['ranges']:
                    pd.DataFrame(self.session_data['ranges']['step10']).to_excel(writer, sheet_name='Análise Faixas', index=False, startrow=0)
                    pd.DataFrame(self.session_data['ranges']['step20']).to_excel(writer, sheet_name='Análise Faixas', index=False, startrow=8)
                
                # Aba Frequência Individual
                if self.session_data['frequency']:
                    total = self.session_data['meta'].get('Total_Numeros_Analisados', 1)
                    df_freq = pd.DataFrame(self.session_data['frequency'], columns=['Número', 'Frequência'])
                    df_freq['Porcentagem'] = (df_freq['Frequência'] / total * 100).round(2)
                    df_freq.to_excel(writer, sheet_name='Frequência_Números', index=False)
                
                # 🆕 Aba Comparativo de Frequência por Tier
                if self.session_data['freq_comparison']:
                    fc = self.session_data['freq_comparison']
                    # Tiers
                    df_tiers = pd.DataFrame([
                        {'Tier': '🔥 Alta', 'Números': ', '.join(map(str, fc['tiers']['high'])), 'Qtd': 20},
                        {'Tier': '⚡ Média', 'Números': ', '.join(map(str, fc['tiers']['medium'])), 'Qtd': 20},
                        {'Tier': '❄️ Baixa', 'Números': ', '.join(map(str, fc['tiers']['low'])), 'Qtd': 20}
                    ])
                    df_tiers.to_excel(writer, sheet_name='Comparativo_Freq', index=False, startrow=0)
                    # Padrões
                    df_patterns = pd.DataFrame([
                        {'Padrão (H,M,L)': str(p), 'Ocorrências': c, '%': round(c/self.session_data['meta'].get('Total_Sorteios_Analisados',1)*100, 1)}
                        for p, c in fc['patterns']
                    ])
                    df_patterns.to_excel(writer, sheet_name='Comparativo_Freq', index=False, startrow=5)
                    # Padrão selecionado
                    pd.DataFrame([{'✅ Padrão Selecionado': str(fc['selected_pattern']), 'Descrição': f"{fc['selected_pattern'][0]} altos + {fc['selected_pattern'][1]} médios + {fc['selected_pattern'][2]} baixos"}]).to_excel(writer, sheet_name='Comparativo_Freq', index=False, startrow=12)

                # Ajustar largura de colunas
                for sheet in writer.sheets.values():
                    for idx, col in enumerate(sheet.columns, 1):
                        max_length = max(len(str(cell.value)) for cell in col if cell.value)
                        sheet.column_dimensions[get_column_letter(idx)].width = min(max_length + 4, 50)

            messagebox.showinfo("Sucesso", f"Relatório completo salvo!\n📂 {file_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao exportar:\n{e}")

    def _print(self, text):
        self.text_res.config(state='normal')
        self.text_res.insert('end', text + '\n')
        self.text_res.see('end')
        self.text_res.config(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = MegaAnalyzerApp(root)
    root.mainloop()
