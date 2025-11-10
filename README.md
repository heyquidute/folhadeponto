# 🧾 Folha de Ponto – Processador

Um aplicativo desktop em **Python** que automatiza a extração e análise de **folhas de ponto em PDF**, gerando planilhas Excel organizadas e verificadas automaticamente.

Desenvolvido com **Tkinter**, o projeto possui uma interface simples e moderna que permite selecionar o arquivo PDF, processá-lo e gerar um relatório detalhado com possíveis erros ou inconsistências de jornada.

---

## 🚀 Funcionalidades

✅ Conversão automática de folhas de ponto em **Excel (.xlsx)**  
✅ Análise de jornada de trabalho com detecção de erros:
   - Jornadas superiores a 10 horas  
   - Falta de marcação de entrada/saída  
   - Ocorrências irregulares  
✅ Geração de uma aba “RESUMO” com os resultados da análise  
✅ Interface gráfica moderna e intuitiva (Tkinter + ttk)  
✅ Barra de progresso e botão de cancelamento  
✅ Suporte a múltiplas páginas (um funcionário por aba)

---

## 🧰 Tecnologias utilizadas

- **Python 3.10+**
- **Tkinter** (interface gráfica)
- **pdfplumber** (extração de tabelas do PDF)
- **openpyxl** (manipulação de planilhas Excel)
- **pandas** (tratamento de dados)
- **threading** (processamento assíncrono)
- **os / re / datetime** (operações utilitárias)

---

## 💻 Estrutura do projeto

Folha_de_Ponto/
│
├── main.py # Interface principal (Tkinter)
├── analisar_folha.py # Lógica de análise e verificação das jornadas
├── extrair_tabela_pdfplumber.py # Extração dos dados do PDF e geração do Excel
├── icone.ico # (Opcional) Ícone do aplicativo
└── README.md # Documentação do projeto
