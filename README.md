# 🎲 Mega-Sena Analyzer

Ferramenta desktop para análise estatística dos sorteios da Mega-Sena. Interface gráfica, portátil e sem instalação.

## 📦 Download
Baixe a versão mais recente em **[Releases](../../releases)**.

## 🖥️ Como usar
1. Extraia ou rode `MegaSenaAnalyzer.exe`
2. Clique em `Selecionar Arquivo` e carregue sua planilha (`.xlsx`, `.xls` ou `.csv`)
3. Filtre por período, escolha quantos números analisar e clique nos botões de ação
4. Gere jogos otimizados ou exporte os resultados

## 🛠️ Gerar o EXE localmente
```bash
git clone https://github.com/SEU-USUARIO/megasena-analyzer.git
cd megasena-analyzer
pip install -r requirements-build.txt
pyinstaller --onefile --noconsole --name "MegaSenaAnalyzer" mega_analyzer.py
