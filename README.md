# **Processador de Folha de Ponto**

Sistema interno da **Comando Auto Peças** para leitura, análise e geração automática de relatórios de ponto.

---

## 📌 **Descrição**

Sistema que lê arquivos **PDF de folha de ponto** e gera automaticamente:

* Arquivo Excel consolidado
* Relatório de Verificação
* Relatório de Não Conformidade

Desenvolvido para facilitar a conferência de ponto dos colaboradores e agilizar o processo interno.

---

## 🖥️ **Como usar o programa (.EXE)**

1. Execute o arquivo **"Processador Folha de Ponto.exe"** e selecione o **Tipo de Relatório** que você quer:

  ![Screenshot do passo 1](/assets/passo1.png)

2. Na janela, clique em:

   * **“Selecionar PDF”** → escolha o arquivo de ponto que deseja analisar

   ![Screenshot do passo 2](/assets/passo2.png)

   * O nome do arquivo selecionado aparecerá na tela

   ![Screenshot do passo 3](/assets/passo3.png)

3. Em seguida clique em **“Processar”** e aguarde. Você consegue acompanhar o processamento pela barra de progresso.

   ![Screenshot do passo 4](/assets/passo4.png)

4. Quando terminar, aparecerá uma mensagem informando que o arquivo foi salvo.

    ![Screenshot do passo 5](/assets/passo5.png)

5. O excel será salvo **na mesma pasta onde está o pdf selecionado**.

    ![Screenshot do passo 6](/assets/passo6.png)

## 📂 **Estrutura do Projeto**

```
📦 Processador de Folha de Ponto
│
├── main.py              → Interface gráfica (Tkinter)
├── extrair_tabela.py    → Lê os PDFs e gera tabelas em Excel
├── verificacao.py       → Gera o relatório de verificação
├── naoconformidade.py   → Gera o relatório de não conformidade
├── cria_link.py         → Cria links e navegação entre abas no Excel
├── convert.py           → Funções de conversão de horário e números
├── dist       → Funções de conversão de horário e números
└── outros arquivos de suporte
```
> Essa é a organização interna dos arquivos do sistema, caso seja necessário manutenção ou consulta técnica.
---

## 📘 **Detalhamento das Regras de Análise**

A seguir estão as regras **detalhadas** utilizadas nos dois principais relatórios:

---

# 📝 **Relatório de Verificação (`verificacao.py`)**

Este relatório verifica **ocorrências específicas**, gerando uma aba Resumo para ser analisada.

### ✔️ Lógica da condição:

### **1. Atestados médicos**

Indica quantidade de atestados médicos no período. 
>O programa verifica se o texto da coluna de Ocorrências começa com **"007"** ou **"ATESTADO"**. Se sim, cria uma linha na aba Resumo com os dados.

### **2. Banco de horas**

Indica saídas antecipadas onde as horas vão como saldo negativo para o banco de horas.
>O programa verifica se o texto da coluna de Ocorrências começa com **"008"** ou com **"BANCO DE HORAS"**. Se sim, cria uma linha na aba Resumo com os dados.

### **3. Abono**

Indica saídas antecipadas onde as horas NÃO vão como saldo negativo para o banco de horas.
> Verifica se o texto da coluna de Ocorrências começa com **"004"** ou com **"ABONO"**. Se sim, cria uma linha na aba Resumo com os dados.

### **4. Saída antecipada**

Indica saídas antecipadas usando as horas que tem na casa.
> Verifica se o texto da coluna de Ocorrências começa com **"014"**. Se sim, cria uma linha na aba Resumo com os dados.

### **5. Compensação de horas**

> Verifica se o texto da coluna de Ocorrências começa com **"434"**. Se sim, cria uma linha na aba Resumo com os dados.

### **6. Suspensão**

Verifica se há alguma suspensão identificada na folha de ponto
> Verifica se o texto da coluna de Ocorrências começa com **"010"** ou com **"SUSPENS"**. Se sim, cria uma linha na aba Resumo com os dados.


---

# 🛑 **Relatório de Não Conformidade (`naoconformidade.py`)**

Este relatório aponta **inconsistências nos horários de batidas** que precisam de uma atenção maior.
### ✔️ Lógica da condição:

### **1. Almoço < 1h**

Verifica se o tempo de almoço foi menor que 1 hora. 
>O programa verifica se o valor na **coluna "T Almoço" é menor que 1:00**. Se sim, cria uma linha na aba Resumo com os dados.

### **2. Almoço > 1h20min**

Verifica se o tempo de almoço foi maior que 1 hora e 20 minutos. 
>O programa verifica se o valor na **coluna "T Almoço" é maior que 1:20**. Se sim, cria uma linha na aba Resumo com os dados.

### **3. Período da Manhã/Tarde > 6h**

Verifica se o tempo de um dos períodos foi maior que 6 horas. 
>O programa verifica se o valor na **coluna "Turno Manhã" ou na coluna "Turno Tarde" é maior que 6:00**. Se sim, cria uma linha na aba Resumo com os dados.

### **4. Jornada > 10h**

Verifica se o tempo da jornada diário foi maior que 10 horas. 
>O programa verifica se o valor na **coluna "Total" é maior que 10:00**. Se sim, cria uma linha na aba Resumo com os dados.

### **5. Saiu depois de 22h**

Verifica se o funcionário saiu após 22h. 
>O programa verifica se o valor na **coluna "Hr Sai T" é maior que 22:00**. Se sim, cria uma linha na aba Resumo com os dados.

### **6. Saldo de hora negativo**

Verifica se o saldo atual de horas está negativo. 
>O programa verifica se é negativo o valor da **última célula da coluna I** *(é onde está a informação do saldo atual, seguindo a formatação da folha de ponto)*. Se sim, cria uma linha na aba Resumo com os dados.

---

## 🔗 **Criação de Links**
#### Ao clicar nos nomes na coluna "Funcionário" da aba Resumo, você será redirecionado à aba do funcionário. E na célula A1 de cada aba de funcionário tem o link que retorna para a aba Resumo.

O módulo `cria_link.py` cria automaticamente:

* Link de cada colaborador → aba individual
* Link de retorno → aba RESUMO
* Navegação organizada entre relatórios

---

## ⏱️ **Conversões Internas**

O módulo `convert.py` trata:

* Conversão de texto para horário
* Conversão de horas para número decimal
* Ajustes de formatação

---

## 🏷️ **Licença**

Este projeto **não possui licença aberta**.

✔️ Uso interno exclusivo da **Comando Auto Peças**.

---

## 👩‍💻 **Créditos**

**Desenvolvido por Ana Clara Quidute**
