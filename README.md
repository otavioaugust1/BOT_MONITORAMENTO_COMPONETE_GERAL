# 🤖 BOT de Monitoramento de Componente Geral - INVESTSUS

Este projeto é um *script* Python, implementado em um Jupyter Notebook, que automatiza o *download*, tratamento e consolidação de dados de **Crédito Financeiro** e **Modalidade 1** do painel INVESTSUS (do Ministério da Saúde do Brasil). O objetivo é gerar um arquivo Excel final (`.xlsx`) atualizado para monitoramento e análise.

## 🎯 Objetivo

Automatizar o processo de coleta, limpeza e estruturação dos dados de propostas de Crédito Financeiro e Modalidade 1 do Painel INVESTSUS, aplicando correções manuais de status de propostas (Aprovado/Cancelado) e gerando um arquivo Excel consolidado em um formato padronizado (`MODELO`).

## ⚙️ Pré-requisitos

Para executar este *script* com sucesso, você precisará dos seguintes softwares e bibliotecas:

1.  **Python 3.x**
2.  **Microsoft Edge:** O *script* utiliza o Selenium para automação do navegador Edge.
3.  **`msedgedriver.exe`:** O executável do *driver* do Edge, compatível com a sua versão do navegador, deve estar localizado na pasta `web` (i.e., `web/msedgedriver.exe`).
4.  **Bibliotecas Python:**
    ```bash
    pip install pandas openpyxl selenium
    ```
5.  **Estrutura de Pastas:** O projeto deve estar organizado com as seguintes pastas:
      * `downloads/`: Onde os arquivos `.xlsx` baixados serão salvos.
      * `model/`: Onde o arquivo **`MONITORAMENTO DE COMPONENTE.xlsx`** (o *template* de saída) deve estar localizado.
      * `saida/`: Onde o arquivo Excel final e datado será gerado.
      * `web/`: Onde o `msedgedriver.exe` deve estar.

## 🚀 Como Executar o Script

1.  **Clone o Repositório:**
    ```bash
    git clone https://www.dio.me/articles/enviando-seu-projeto-para-o-github
    cd [pasta do repositório]
    ```
2.  **Verifique os Pré-requisitos:** Certifique-se de que todas as bibliotecas estão instaladas e as pastas `web/` e `model/` (com o arquivo modelo) existem.
3.  **Execute o Notebook:**
    Abra o arquivo `credito_modalidade.ipynb` e execute todas as células em sequência. O *script* irá:
      * Iniciar o navegador Edge.
      * Acessar o painel INVESTSUS.
      * Baixar os 6 arquivos (`aba1`, `aba2`, `aba3`) das abas **"Crédito Financeiro"** e **"Modalidade 1"**.
      * Realizar o tratamento dos dados (limpeza, cálculo, alteração de status).
      * Consolidar os DataFrames no arquivo Excel modelo e salvá-lo na pasta `saida/`.

## 📦 Estrutura do Script (Notebook)

O *script* está dividido em três partes principais:

### 1\. PRIMEIRA PARTE - Downloads Manual dos Arquivos do INVESTSUS

Esta seção utiliza a biblioteca **Selenium** para automação do navegador Edge.

  * **Configuração:** Define o diretório de *downloads* (`downloads/`) e inicializa o `webdriver`.
  * **Acesso e Navegação:** Acessa a URL do painel INVESTSUS.
  * **Rotina de Download:** Define a função `baixar_e_renomear` que clica nos botões de download, lida com alertas e renomeia o arquivo baixado para um nome padronizado (ex: `credito_financeiro_aba1.xlsx`).
  * **Execução:** Realiza o download de 3 arquivos para a aba **"Crédito Financeiro"** e 3 para a aba **"Modalidade 1"**.

### 2\. SEGUNDA PARTE, Realizar Tratamento dos Dados

Esta seção carrega e manipula os dados usando a biblioteca **Pandas**.

  * **Carregamento:** Carrega os 6 arquivos Excel baixados em DataFrames separados (`df_cf_aba1`, `df_m1_aba3`, etc.).
  * **Tratamento de Status:** Aplica correções manuais de status em propostas específicas:
      * Marca propostas em `propostas_aprovada` como **'Aprovado'**.
      * Marca propostas em `propostas_canceladas` como **'Cancelado'**.
  * **Criação da Matriz de Oferta (M\_OFERTA):**
      * Cria e estrutura os DataFrames de Matriz de Oferta (Cirurgias - `_CC` e OCI - `_OCI`) para ambas as modalidades (`df_cf_aba3`, `df_m1_aba3`, `df_cf_aba2`, `df_m1_aba2`).
      * Calcula a coluna `VALOR_TOTAL_MES` (quantidade mensal \* valor máximo/procedimento).
      * Reorganiza a ordem e renomeia colunas.
  * **Criação da Aba Simplificada (SIMP):**
      * Cria os DataFrames simplificados (`df_simp_cc`, `df_simp_m1`) com base nas abas `aba1`.
      * Faz o *merge* dos valores totais calculados (soma de `VALOR_TOTAL_MES` para Cirurgia e OCI) para cada Proposta de Referência.
      * Calcula as colunas `VALOR_TOTAL_MES_COMP+OCI` e `VALOR_TOTAL_ANO_COMP+OCI`.

### 3\. TERCEIRA PARTE, Carregar os DataFrames para a Tabela MODELO

Esta seção usa a biblioteca **OpenPyXL** para inserir os dados tratados no arquivo modelo Excel.

  * **Mapeamento:** Define o mapeamento entre os DataFrames tratados e as abas específicas do arquivo modelo (`MAPPING`).
  * **Sobrescrita de Abas:** A função `sobrescrever_aba` carrega o arquivo `MONITORAMENTO DE COMPONENTE.xlsx` e insere os dados de cada DataFrame a partir da **linha 3** de suas abas correspondentes (mantendo o cabeçalho original do modelo).
  * **Atualização de Metadados:** Atualiza a aba **`INFO`** com a data/hora da execução e a sigla "BR".
  * **Saída:** Salva o resultado final com um nome datado (ex: `saida/20251009_MONITORAMENTO DE COMPONENTE.xlsx`), garantindo que o modelo original permaneça intacto.

