# 🤖 BOT de Monitoramento: Dados de Crédito INVESTSUS

Este projeto é um *script* Python (implementado em um Jupyter Notebook) focado na **automação** e **consolidação** de dados do painel INVESTSUS (Ministério da Saúde do Brasil). O bot automatiza o *download*, o tratamento e a estruturação de dados de **Crédito Financeiro** e **Modalidade 1** para gerar um arquivo Excel (`.xlsx`) pronto para análise.

-----

## 🎯 Funcionalidade Principal

O objetivo é eliminar o trabalho manual de coleta, limpeza e estruturação, garantindo que o monitoramento seja feito sobre uma base de dados unificada e padronizada.

O script executa as seguintes tarefas:

  * **Coleta Automática:** Acessa o painel INVESTSUS via automação web (Selenium).
  * **Tratamento de Dados:** Aplica rotinas de limpeza, cálculos e correções manuais de status de propostas (Aprovado/Cancelado).
  * **Consolidação:** Gera um arquivo Excel final, seguindo o *template* `MODELO`, pronto para monitoramento.

-----

## ⚙️ Pré-requisitos Técnicos

Para rodar o projeto, certifique-se de ter instalado o seguinte:

### 1\. Ambiente

  * **Python 3.x**
  * **Microsoft Edge:** O script utiliza o Selenium para automação.
  * **`msedgedriver.exe`:** O **driver do Edge** (compatível com sua versão do navegador) deve ser colocado na pasta **`web/`** (ex: `web/msedgedriver.exe`).

### 2\. Dependências Python

Instale as bibliotecas necessárias usando o `pip`. Se houver um arquivo `requirements.txt` no repositório, use o comando:

```bash
pip install -r requirements.txt
# Ou, instale manualmente:
# pip install pandas openpyxl selenium
```

### 3\. Estrutura de Diretórios

As seguintes pastas são essenciais e devem existir no diretório raiz do projeto:

| Pasta | Conteúdo Esperado |
| :--- | :--- |
| **`downloads/`** | Onde os arquivos `.xlsx` baixados serão temporariamente salvos. |
| **`model/`** | Onde deve estar o arquivo *template* (modelo) de saída: **`MONITORAMENTO DE COMPONENTE.xlsx`**. |
| **`saida/`** | Onde o arquivo Excel final e datado será gerado. |
| **`web/`** | Onde o executável **`msedgedriver.exe`** deve ser colocado. |

-----

## 🚀 Guia de Execução Rápida

Siga estes passos para rodar o bot de monitoramento:

### 1\. Clonagem do Repositório

Use o Git para baixar o código e entrar no diretório do projeto:

```bash
git clone https://github.com/otavioaugust1/BOT_MONITORAMENTO_COMPONETE_GERAL
cd BOT_MONITORAMENTO_COMPONETE_GERAL
```

### 2\. Configuração e Dependências

Certifique-se de que a estrutura de pastas e o `msedgedriver.exe` (na pasta `web/`) estão configurados e que as dependências Python foram instaladas (conforme a seção Pré-requisitos).

### 3\. Execução

O processo é totalmente gerenciado pelo Jupyter Notebook:

1.  Abra o arquivo **`credito_modalidade.ipynb`**.
2.  **Execute todas as células em sequência.**

O script fará o restante, desde o acesso ao INVESTSUS até o salvamento do arquivo final datado na pasta **`saida/`**.

-----

## 💡 Detalhamento do Script (`credito_modalidade.ipynb`)

O Notebook está logicamente dividido em três seções principais, que refletem o fluxo de trabalho do bot:

### 1\. PARTE 1: Automação e Downloads (Selenium)

Esta seção configura e utiliza a biblioteca **Selenium** para a interação web.

  * **Coleta de Dados:** Acessa o painel INVESTSUS e executa a rotina de *download* para um total de **6 arquivos** (3 de "Crédito Financeiro" e 3 de "Modalidade 1"), renomeando-os e salvando-os na pasta `downloads/`.

### 2\. PARTE 2: Tratamento e Estruturação (Pandas)

O foco aqui é carregar e manipular os dados usando o **Pandas**.

  * **Correção de Status:** Aplica regras de negócio para forçar o status de propostas específicas para **'Aprovado'** ou **'Cancelado'**.
  * **Criação de Matrizes:** Calcula e estrutura as matrizes de Oferta, calculando o `VALOR_TOTAL_MES`.
  * **Consolidação Simplificada:** Cria as abas de resumo simplificado (`SIMP`), consolidando os valores calculados de Cirurgia e OCI para criar as colunas de valor total.

### 3\. PARTE 3: Geração da Saída (OpenPyXL)

Utiliza a biblioteca **OpenPyXL** para inserir os dados tratados no *template* Excel.

  * **Sobrescrita:** Carrega o modelo (`MONITORAMENTO DE COMPONENTE.xlsx`) e insere os dados de cada DataFrame **a partir da linha 3** de suas abas correspondentes (preservando o cabeçalho original).
  * **Saída Final:** Salva o arquivo final com um nome datado (ex: `saida/YYYYMMDD_MONITORAMENTO DE COMPONENTE.xlsx`), garantindo que o modelo original nunca seja sobrescrito.