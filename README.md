# 🧾 Analisador da Folha de Ponto

![Screenshot da Interface](/assets/imagem.png)

Interface gráfica em **Python (Tkinter)** para automatizar a extração e análise de folhas de ponto em PDF, permitindo gerar relatórios detalhados de **atestados**, **horários** e **batidas** de funcionários.

---

## 📌 Visão Geral

O **Analisador da Folha de Ponto** é uma ferramenta desenvolvida para facilitar a leitura e interpretação de folhas de ponto emitidas em formato PDF.

O sistema extrai automaticamente as informações, gera um arquivo Excel organizado e realiza análises específicas conforme o tipo de relatório escolhido.

A interface é simples, moderna e intuitiva — basta selecionar o PDF e escolher o tipo de relatório desejado.

---

## ⚙️ Funcionalidades Principais

* 📂 **Leitura automática de PDFs** de folha de ponto.

* 📊 **Geração de relatórios personalizados** em Excel:

    - **Relatório de horários:** analisa entradas, saídas e períodos de trabalho.

    - **Relatório de atestados:** identifica e registra atestados médicos.

    - **Relatório de divergência:** detecta e lista dias com marcações incorretas.

* 🎨 **Interface gráfica amigável (Tkinter)**, sem necessidade de comandos no terminal.

* 📁 **Arquivos de saída organizados** com sufixos descritivos:

    - `_horarios.xlsx` → Relatório de horários

    - `_atestados.xlsx` → Relatório de atestados

    - `_batidasdeponto.xlsx` → Relatório de inconsistências

    - `_processado.xlsx` → Caso o tipo de relatório não seja reconhecido

* 🧠 **Processamento seguro e não bloqueante**, com barra de progresso e opção de cancelamento.

* 🖋️ **Formatação automática no Excel** (cabeçalhos, cores, alinhamento e colunas ajustadas).

---

## 🖥️ Como Usar

1. Execute o aplicativo:

```bash
python main.py
```

2. Na janela que abrir:

   * Clique em **“Selecionar PDF”** e escolha o arquivo da folha de ponto.

   * Escolha o tipo de relatório desejado:

       - **Relatório de horários**
       - **Relatório de atestados**
       - **Relatório de inconsistências**

   * Clique em **“Processar”**.

3. Aguarde o processamento (a barra de progresso mostrará o andamento).

4. O arquivo Excel será salvo automaticamente na mesma pasta do PDF selecionado.

---

## 📂 Estrutura do Projeto

```
📁 AnalisadorFolhaPonto/
├── main.py                 # Interface principal (Tkinter)
├── extrair_tabela.py       # Responsável por extrair dados do PDF e gerar Excel
├── analisar_folha.py       # Gera o relatório de horários
├── analisar_atestados.py   # Gera o relatório de atestados
├── analisar_batidas.py # Gera o relatório de inconsistências
├── icone.ico               # Ícone da aplicação
└── README.md               # Documentação do projeto
```

---

## 🧠 Lógica de Funcionamento

1. O **usuário seleciona o PDF** e o **tipo de relatório**.
2. O aplicativo chama a função `gerar_excel()` (em `extrair_tabela.py`) para extrair e converter o conteúdo do PDF em Excel.
3. Dependendo do tipo de relatório selecionado:

   * Chama `analisar_folha()` → gera arquivo `_horarios.xlsx`
   * Chama `analisar_atestados()` → gera arquivo `_atestados.xlsx`
   * Chama `analisar_batidas()` → gera arquivo `_batidasdeponto.xlsx`
4. Caso o tipo de relatório não seja reconhecido, o programa gera um arquivo `_processado.xlsx`.

---

## 🧾 Relatórios Gerados

### 🕐 Relatório de Horários

Analisa os dados de ponto (entrada, almoço, saída) e calcula totais e diferenças de horários por funcionário.

### 🩺 Relatório de Atestados

* Cria uma aba chamada **ATESTADOS** no início da planilha.
* Lista o **nome do funcionário**, **data** e **detalhe** (texto completo da ocorrência).
* As linhas correspondentes a atestados são **pintadas de verde** nas abas individuais dos funcionários.
* O cabeçalho é formatado com **negrito, centralização e borda inferior**.
* As colunas têm **largura ajustada automaticamente**, e **as linhas de grade são ocultadas**.

### ⚠️ Relatório de Divergências

* Gera uma aba chamada **DIVERGENCIA** no início da planilha.
* Identifica e lista dias com quantidade de **batidas incompletas**.
* Cada linha apresenta o **nome do funcionário** e **data**.
* As células com erro são destacadas com **fundo azul claro** para fácil visualização.
* O layout segue o mesmo padrão visual dos outros relatórios (formatação automática, cabeçalhos e colunas ajustadas).

---

## 🧩 Tecnologias Utilizadas

* **Python 3**
* **Tkinter** → Interface gráfica
* **openpyxl** → Manipulação e formatação de planilhas Excel
* **pdfplumber** → Leitura e extração de dados de PDFs
* **threading** → Processamento paralelo (mantém a interface fluida)

---

## 🧠 Boas Práticas

* Use **arquivos PDF legíveis (não escaneados)** para garantir extração correta.
* Mantenha o nome das colunas originais no Excel extraído.
* Não modifique manualmente o arquivo Excel gerado antes de finalizar a análise.
* Sempre revise o relatório para identificar e corrigir possíveis falhas.

---

## 👩‍💻 Créditos

**Desenvolvido por Ana Clara Quidute**

Projeto: **“Analisador da Folha de Ponto”**
