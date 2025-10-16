# ==============================================================================
# PRIMEIRA PARTE - downloads manual dos arquivos do INVESTSUS
# ==============================================================================

# 📚 BIBLIOTECAS
import os
import time
import warnings
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

# 🔕 Oculta alertas de warnings
warnings.filterwarnings('ignore')

# Contagem de TEMPO de processamento
inicio = datetime.now()
print(f"🔵 Início da execução: {inicio.strftime('%H:%M:%S')}")

# 📁 Diretório de downloads
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
print(f"📁 Diretório de downloads configurado: {DOWNLOAD_DIR}")

# ⚙️ Configuração do Edge
# ATENÇÃO: Verifique se o caminho para o msedgedriver.exe está correto
driver_path = os.path.join(os.getcwd(), "web", "msedgedriver.exe")
service = Service(executable_path=driver_path)
edge_options = Options()
edge_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
})

# 🚀 Inicializa o navegador
try:
    driver = webdriver.Edge(service=service, options=edge_options)
    wait = WebDriverWait(driver, 20)
    print("🚀 Navegador Edge iniciado com sucesso.")
except Exception as e:
    print(f"❌ ERRO ao iniciar o Edge: {e}")
    print("Certifique-se de que o msedgedriver.exe está no caminho correto e compatível com sua versão do Edge.")
    # Define driver como None para pular o resto da execução se o driver falhar
    driver = None 

if driver:
    # 🌐 Acessa a página alvo
    url = "https://investsuspaineis.saude.gov.br/extensions/CGIN_PMAE/CGIN_PMAE.html#"
    driver.get(url)
    print(f"🌐 Página acessada: {url}")
    time.sleep(5)

    # 🔁 Função para baixar e renomear arquivos
    def baixar_e_renomear(xpath_botao, nome_destino):
        print(f"📥 Iniciando download para: {nome_destino}")
        
        # Clica no botão de download
        try:
            botao = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_botao)))
            botao.click()
            time.sleep(2)

            # Aceita o alerta
            WebDriverWait(driver, 10).until(EC.alert_is_present())
            alerta = driver.switch_to.alert
            alerta.accept()

            # Aguarda o download finalizar
            # ATENÇÃO: Este tempo pode precisar de ajuste dependendo do tamanho do arquivo
            time.sleep(5)

            # Renomeia o arquivo mais recente .xlsx
            arquivos_xlsx = [os.path.join(DOWNLOAD_DIR, f) for f in os.listdir(DOWNLOAD_DIR) if f.endswith(".xlsx")]
            if not arquivos_xlsx:
                print(f"⚠️ Aviso: Nenhum arquivo .xlsx encontrado em {DOWNLOAD_DIR} para renomear.")
                return

            arquivo_mais_recente = max(arquivos_xlsx, key=os.path.getctime)
            caminho_novo = os.path.join(DOWNLOAD_DIR, nome_destino)

            # Se já existir um arquivo com o nome de destino, exclui
            if os.path.exists(caminho_novo):
                os.remove(caminho_novo)

            # Renomeia o novo arquivo
            os.rename(arquivo_mais_recente, caminho_novo)
            print(f"📦 Arquivo renomeado para: {nome_destino}\n")
        except Exception as e:
            print(f"❌ Erro ao baixar ou renomear {nome_destino}: {e}\n")


    # 📊 Aba Crédito
    print("📂 Acessando aba: Crédito Financeiro")
    try:
        aba_credito = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="menu_abas"]/a[3]')))
        aba_credito.click()
        time.sleep(5)
    
        baixar_e_renomear('//*[@id="QV3-02574e8688-0e17-49d8-8ca9-c037abb4a5f7"]', "credito_financeiro_aba1.xlsx")
        baixar_e_renomear('//*[@id="QV3-03a27c0e0b-cac7-45c5-92c1-ec1cb34f3828"]', "credito_financeiro_aba2.xlsx")
        baixar_e_renomear('//*[@id="QV3-04541ab7f8-9dda-4cbd-82cf-72840cb4ac2d"]', "credito_financeiro_aba3.xlsx")
    except Exception as e:
        print(f"❌ Erro na aba Crédito Financeiro: {e}")

    # 📊 Aba Modalidade 1
    print("📂 Acessando aba: Modalidade 1")
    try:
        aba_modalidade = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="menu_abas"]/a[4]')))
        aba_modalidade.click()
        time.sleep(5)
    
        baixar_e_renomear('//*[@id="QV4-02yfhvCp"]', "modalidade_1_aba1.xlsx")
        baixar_e_renomear('//*[@id="QV4-03JQRjW"]', "modalidade_1_aba2.xlsx")
        baixar_e_renomear('//*[@id="QV4-04gvqxmPC"]', "modalidade_1_aba3.xlsx")
    except Exception as e:
        print(f"❌ Erro na aba Modalidade 1: {e}")

    # 🛑 Fecha o navegador
    driver.quit()
    print("-----------------------------------------------------------------")
    print("Navegador fechado. Iniciando processamento de dados.")
    print("-----------------------------------------------------------------")


# ==============================================================================
# SEGUNDA PARTE - REALIZAR TRATAMENTO DOS DADOS.
# ==============================================================================

if not driver:
    print("\n⚠️ A automação de download foi pulada devido a um erro. Tentando carregar arquivos existentes.")

try:
    # Ajusta a opção de exibição para mostrar todas as colunas
    pd.set_option('display.max_columns', None)

    # Função auxiliar para ler Excel com caminho de download
    def ler_excel_download(nome_arquivo):
        caminho = os.path.join(DOWNLOAD_DIR, nome_arquivo)
        return pd.read_excel(caminho)

    # Créditos financeiros
    df_cf_aba1 = ler_excel_download("credito_financeiro_aba1.xlsx")
    df_cf_aba2 = ler_excel_download("credito_financeiro_aba2.xlsx")
    df_cf_aba3 = ler_excel_download("credito_financeiro_aba3.xlsx")

    # Modalidade 1
    df_m1_aba1 = ler_excel_download("modalidade_1_aba1.xlsx")
    df_m1_aba2 = ler_excel_download("modalidade_1_aba2.xlsx")
    df_m1_aba3 = ler_excel_download("modalidade_1_aba3.xlsx")
    
    print("\n✅ Arquivos Excel carregados com sucesso.")

    # --------------------------------------------------------------------------
    # Tratamento Comum (Aprovação/Cancelamento Manual)
    # --------------------------------------------------------------------------

    # Lista de propostas a aprovar (Status Manual)
    propostas_aprovada = [
        781831300012025501   # APROVAÇÃO DE SOBRAL (MANUALMENTE)
    ]
    propostas_aprovada = [str(p) for p in propostas_aprovada]

    # Atualiza o status nos DataFrames de Crédito Financeiro
    print("\n--- Atualizando Status (Aprovado) ---")
    for i, df in enumerate([df_cf_aba1, df_cf_aba2, df_cf_aba3], start=1):
        df['Proposta de Referência'] = df['Proposta de Referência'].astype(str)
        df.loc[df['Proposta de Referência'].isin(propostas_aprovada), 'Status da Proposta'] = 'Aprovado'
        aprovadas = df[df['Proposta de Referência'].isin(propostas_aprovada)]
        print(f"✅ df_cf_aba{i}: {len(aprovadas)} propostas marcadas como 'Aprovado'.")


    # Lista de propostas a cancelar (Status Manual)
    propostas_canceladas = [
        1086978200012025503,    # PE-RECIFE 
        8862568600242025502,    # RS-PORTO ALEGRE
        2870053000032025501,    # SC-TIMBE DO SUL
        2870053000022025502,    # SC-SOMBRIO
        2870053000022025503,    # SC-SOMBRIO
        4605648700012025501,    # SP-VALINHOS
        24712500022025504,      # RJ-NITEROI
        353343200012025502,     # CAMPO GRANDE
    ]
    propostas_canceladas = [str(p) for p in propostas_canceladas]

    # Atualiza o status nos DataFrames de Crédito Financeiro
    print("\n--- Atualizando Status (Cancelado) ---")
    for i, df in enumerate([df_cf_aba1, df_cf_aba2, df_cf_aba3], start=1):
        df['Proposta de Referência'] = df['Proposta de Referência'].astype(str)
        df.loc[df['Proposta de Referência'].isin(propostas_canceladas), 'Status da Proposta'] = 'Cancelado'
        canceladas = df[df['Proposta de Referência'].isin(propostas_canceladas)]
        print(f"✅ df_cf_aba{i}: {len(canceladas)} propostas marcadas como 'Cancelado'.")
    print("-----------------------------------------------------------------")

    # --------------------------------------------------------------------------
    # Tratamento para string em numeros das proposta da modalidade 1 (abas 1,2,3) 
    # --------------------------------------------------------------------------
    print("🛠 Tratamento para string os numeros das proposta da modalidade 1 (abas 1,2,3)")
    for i, df in enumerate([df_m1_aba1, df_m1_aba2, df_m1_aba3], start=1):
    # Garante que a coluna esteja no tipo string
        df['Proposta de Referência'] = df['Proposta de Referência'].astype(str)

    # --------------------------------------------------------------------------
    # Tratamento MATRIZ DE OFERTA - CRÉDITO FINANCEIRO - CIRURGIAS (df_cf_aba3)
    # --------------------------------------------------------------------------
    print("🛠 Tratando Matriz de Oferta: Crédito Financeiro - Cirurgias (df_cf_aba3)")
    df_cf_aba3.drop(columns='TP_COMPLEXIDADE', inplace=True)
    df_cf_aba3['VALOR_TOTAL_MES'] = df_cf_aba3['QT_ATENDIMENTO_MES'] * df_cf_aba3['VL_TOTAL_COMPLEMENTACAO_MAXIMA']
    
    # Adicionar coluna 'Entidade' do df_cf_aba1
    df_cf_aba3 = df_cf_aba3.merge(
        df_cf_aba1[['Proposta de Referência', 'Entidade']],                                   
        on='Proposta de Referência',
        how='left')
    
    df_cf_aba3.rename(columns={
        'Entidade': 'ENTIDADE',
        'TX_COMPLEMENTACAO_MAXIMA': '% COMPLEMENTACAO_MAXIMA'                                        
    }, inplace=True)
    
    # Reorganizar as colunas
    df_cf_aba3 = df_cf_aba3[[                                                                                               
        'Proposta de Referência', 'Status da Proposta', 'UF', 'Município', 'CNES', 'ENTIDADE',
        'CO_PROCEDIMENTO_SIGTAP', 'NO_GRUPO', 'NO_PROCEDIMENTO', '% COMPLEMENTACAO_MAXIMA',
        'VL_TABELA_SUS', 'VL_TOTAL_COMPLEMENTACAO_MAXIMA', 'QT_ATENDIMENTO_MES', 'VALOR_TOTAL_MES'
    ]]


    # --------------------------------------------------------------------------
    # Tratamento MATRIZ DE OFERTA - MODALIDADE 1 - CIRURGIAS (df_m1_aba3)
    # --------------------------------------------------------------------------
    print("🛠 Tratando Matriz de Oferta: Modalidade 1 - Cirurgias (df_m1_aba3)")
    df_m1_aba3.drop(columns='TP_COMPLEXIDADE', inplace=True)                                                                
    df_m1_aba3['VALOR_TOTAL_MES'] = df_m1_aba3['QT_ATENDIMENTO_MES'] * df_m1_aba3['VL_TOTAL_COMPLEMENTACAO_MAXIMA']         
    
    # Adicionar coluna 'Entidade' do df_m1_aba1
    df_m1_aba3 = df_m1_aba3.merge(
        df_m1_aba1[['Proposta de Referência', 'Entidade']],                                   
        on='Proposta de Referência',
        how='left')
    
    df_m1_aba3.rename(columns={
        'Entidade': 'ENTIDADE',
        'TX_COMPLEMENTACAO_MAXIMA': '% COMPLEMENTACAO_MAXIMA'                                        
    }, inplace=True)
    
    # Reorganizar as colunas
    df_m1_aba3 = df_m1_aba3[[                                                                                               
        'Proposta de Referência', 'Status da Proposta', 'UF', 'Município', 'CNES', 'ENTIDADE',
        'CO_PROCEDIMENTO_SIGTAP', 'NO_GRUPO', 'NO_PROCEDIMENTO', '% COMPLEMENTACAO_MAXIMA',
        'VL_TABELA_SUS', 'VL_TOTAL_COMPLEMENTACAO_MAXIMA', 'QT_ATENDIMENTO_MES', 'VALOR_TOTAL_MES'
    ]]


    # --------------------------------------------------------------------------
    # Tratamento MATRIZ DE OFERTA - CRÉDITO FINANCEIRO - OCI (df_cf_aba2)
    # --------------------------------------------------------------------------
    print("🛠 Tratando Matriz de Oferta: Crédito Financeiro - OCI (df_cf_aba2)")
    df_cf_aba2.drop(columns=['TP_SEXO','NU_IDADE_MINIMA','NU_IDADE_MAXIMA'], inplace=True)                   
    df_cf_aba2['VALOR_TOTAL_MES'] = df_cf_aba2['QT_ATENDIMENTO_MES'] * df_cf_aba2['VL_PROCEDIMENTO']         
    
    # Adicionar coluna 'Entidade' do df_cf_aba1
    df_cf_aba2 = df_cf_aba2.merge(df_cf_aba1[['Proposta de Referência', 'Entidade']],                        
        on='Proposta de Referência',
        how='left')
    
    # Reorganizar as colunas
    df_cf_aba2= df_cf_aba2[[                                                                                 
        'Proposta de Referência', 'Status da Proposta', 'UF', 'Município', 'CNES', 'Entidade',
        'NU_PROCEDIMENTO', 'NO_GRUPO', 'DS_PROCEDIMENTO', 'QT_ATENDIMENTO_MES',
        'VL_PROCEDIMENTO', 'VALOR_TOTAL_MES'
    ]]


    # --------------------------------------------------------------------------
    # Tratamento MATRIZ DE OFERTA - MODALIDADE 1 - OCI (df_m1_aba2)
    # --------------------------------------------------------------------------
    print("🛠 Tratando Matriz de Oferta: Modalidade 1 - OCI (df_m1_aba2)")
    df_m1_aba2.drop(columns=['TP_SEXO', 'NU_IDADE_MINIMA', 'NU_IDADE_MAXIMA'], inplace=True)
    df_m1_aba2['VALOR_TOTAL_MES'] = df_m1_aba2['QT_ATENDIMENTO_MES'] * df_m1_aba2['VL_PROCEDIMENTO']
    
    # Adicionar coluna 'Entidade' do df_m1_aba1
    df_m1_aba2 = df_m1_aba2.merge(
        df_m1_aba1[['Proposta de Referência', 'Entidade']],
        on='Proposta de Referência',
        how='left'
    )
    
    # Reorganizar as colunas
    df_m1_aba2 = df_m1_aba2[[
        'Proposta de Referência', 'Status da Proposta', 'UF', 'Município', 'CNES', 'Entidade',
        'NU_PROCEDIMENTO', 'NO_GRUPO', 'DS_PROCEDIMENTO', 'QT_ATENDIMENTO_MES',
        'VL_PROCEDIMENTO', 'VALOR_TOTAL_MES'
    ]]


    # --------------------------------------------------------------------------
    # Tratamento SIMPLIFICADA - CRÉDITO FINANCEIRO (df_simp_cc)
    # --------------------------------------------------------------------------
    print("🛠 Tratando Matriz Simplificada: Crédito Financeiro (df_simp_cc)")
    df_simp_cc = df_cf_aba1.copy()
    df_simp_cc.drop(columns=[
        'Dt. Cadastro', 'Dt. Atualização', 'Dívida Aprox.', 'VL_SALDO_DEVEDOR', 
        'VL_TRIBUTO_FEDERAL_ESTIMADO'], inplace=True)

    # CIRURGIAS: Soma 'VALOR_TOTAL_MES' do df_cf_aba3
    soma_por_proposta_cc = df_cf_aba3.groupby('Proposta de Referência')['VALOR_TOTAL_MES'].sum().reset_index()
    soma_por_proposta_cc.rename(columns={'VALOR_TOTAL_MES': 'VL_TOTAL_COMP_CIRUGICO'}, inplace=True)
    df_simp_cc = df_simp_cc.merge(soma_por_proposta_cc, on='Proposta de Referência', how='left')

    # OCI: Soma 'VALOR_TOTAL_MES' do df_cf_aba2
    soma_por_proposta_co = df_cf_aba2.groupby('Proposta de Referência')['VALOR_TOTAL_MES'].sum().reset_index()
    soma_por_proposta_co.rename(columns={'VALOR_TOTAL_MES': 'VL_TOTAL_OCI'}, inplace=True)
    df_simp_cc = df_simp_cc.merge(soma_por_proposta_co, on='Proposta de Referência', how='left')

    # Cálculos Finais
    df_simp_cc['VL_TOTAL_COMP_CIRUGICO'].fillna(0, inplace=True)
    df_simp_cc['VL_TOTAL_OCI'].fillna(0, inplace=True)
    df_simp_cc['VALOR_TOTAL_MES_COMP+OCI'] = df_simp_cc['VL_TOTAL_COMP_CIRUGICO'] + df_simp_cc['VL_TOTAL_OCI']
    df_simp_cc['VALOR_TOTAL_ANO_COMP+OCI']= df_simp_cc['VALOR_TOTAL_MES_COMP+OCI']*12


    # --------------------------------------------------------------------------
    # Tratamento SIMPLIFICADA - MODALIDADE 1 (df_simp_m1)
    # --------------------------------------------------------------------------
    print("🛠 Tratando Matriz Simplificada: Modalidade 1 (df_simp_m1)")
    df_simp_m1 = df_m1_aba1.copy()
    df_simp_m1.drop(columns=['Dt. Cadastro', 'Dt. Atualização'], inplace=True)

    # CIRURGIAS: Soma 'VALOR_TOTAL_MES' do df_m1_aba3
    soma_por_proposta_mc = df_m1_aba3.groupby('Proposta de Referência')['VALOR_TOTAL_MES'].sum().reset_index()
    soma_por_proposta_mc.rename(columns={'VALOR_TOTAL_MES': 'VL_TOTAL_COMP_CIRUGICO'}, inplace=True)
    df_simp_m1 = df_simp_m1.merge(soma_por_proposta_mc, on='Proposta de Referência', how='left')

    # OCI: Soma 'VALOR_TOTAL_MES' do df_m1_aba2
    soma_por_proposta_mo = df_m1_aba2.groupby('Proposta de Referência')['VALOR_TOTAL_MES'].sum().reset_index()
    soma_por_proposta_mo.rename(columns={'VALOR_TOTAL_MES': 'VL_TOTAL_OCI'}, inplace=True)
    df_simp_m1 = df_simp_m1.merge(soma_por_proposta_mo, on='Proposta de Referência', how='left')

    # Cálculos Finais
    df_simp_m1['VL_TOTAL_COMP_CIRUGICO'].fillna(0, inplace=True)
    df_simp_m1['VL_TOTAL_OCI'].fillna(0, inplace=True)
    df_simp_m1['VALOR_TOTAL_MES_COMP+OCI'] = df_simp_m1['VL_TOTAL_COMP_CIRUGICO'] + df_simp_m1['VL_TOTAL_OCI']
    df_simp_m1['VALOR_TOTAL_ANO_COMP+OCI']= df_simp_m1['VALOR_TOTAL_MES_COMP+OCI']*12


    # ==============================================================================
    # TERCEIRA PARTE - Carregar os dataFrame para a tabela MODELO
    # ==============================================================================
    print("\n-----------------------------------------------------------------")
    print("Iniciando a atualização do arquivo modelo Excel.")
    print("-----------------------------------------------------------------")


    # Diretórios e arquivos
    MODELO_FILENAME = "MONITORAMENTO DE COMPONENTE.xlsx"
    MODELO_DIR = os.path.join(os.getcwd(), "model")
    MODELO_PATH = os.path.join(MODELO_DIR, MODELO_FILENAME)

    # Mapeamento dos DataFrames para as abas
    MAPPING = {
        "df_cf_aba1": "CREDITO_FINANCEIRO",
        "df_cf_aba2": "M_OFERTA_CF_OCI",
        "df_cf_aba3": "M_OFERTA_CF_CC",
        "df_simp_cc": "SIMP_CF",
        "df_m1_aba1": "MODALIDADE_1",
        "df_m1_aba2": "M_OFERTA_M1_OCI",
        "df_m1_aba3": "M_OFERTA_M1_CC",
        "df_simp_m1": "SIMP_M1",
    }

    # Função para sobrescrever a aba a partir da linha 3
    def sobrescrever_aba(workbook, aba_nome, df):
        if aba_nome in workbook.sheetnames:
            ws = workbook[aba_nome]
            
            # Limpar o conteúdo anterior (a partir da linha 3)
            # Determina a última linha antes de adicionar novos dados
            max_row = ws.max_row
            for row_index in range(3, max_row + 1):
                for col_index in range(1, ws.max_column + 1):
                    ws.cell(row=row_index, column=col_index, value=None)
            
            # Escrever novos dados a partir da linha 3
            for i, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=3):
                for j, value in enumerate(row, start=1):
                    ws.cell(row=i, column=j, value=value)
            print(f"✅ Aba '{aba_nome}' atualizada com {len(df)} linhas.")
        else:
            print(f"⚠️ Aba '{aba_nome}' não encontrada no modelo.")

    # Função para carregar os DataFrames (acessa variáveis globais)
    def carregar_dados_do_excel(nome_df):
        try:
            return globals()[nome_df]
        except KeyError:
            print(f"⚠️ DataFrame '{nome_df}' não está definido.")
            return None

    # Execução principal de carregamento no Excel
    if not os.path.exists(MODELO_PATH):
        print(f"❌ Erro: O arquivo modelo esperado '{MODELO_FILENAME}' não foi encontrado em '{MODELO_DIR}'.")
        print("Verifique se o caminho do arquivo modelo está correto.")
    else:
        try:
            # Carrega o workbook e mantém fórmulas
            modelo_wb = openpyxl.load_workbook(MODELO_PATH)
            print("📂 Arquivo modelo carregado.")

            # Atualizar abas conforme mapeamento
            for df_nome, aba_nome in MAPPING.items():
                df = carregar_dados_do_excel(df_nome)
                if df is not None:
                    sobrescrever_aba(modelo_wb, aba_nome, df)

            # Atualizar aba INFO
            print("\n🛠 Atualizando aba 'INFO'...")
            sigla_uf = "BR"

            if 'INFO' in modelo_wb.sheetnames:
                aba_info = modelo_wb['INFO']
                aba_info['H2'] = datetime.now().strftime('%d/%m/%Y %H:%M')
                aba_info['G1'] = sigla_uf
                print(f"   -> Aba 'INFO' atualizada (H2: Data/Hora, G1: {sigla_uf}).")
            else:
                print("   ⚠️ Aba 'INFO' não encontrada no arquivo modelo.")

            # Salvar novo arquivo
            os.makedirs("saida", exist_ok=True)
            novo_nome = os.path.join("saida", f"{datetime.today().strftime('%Y%m%d')}_{MODELO_FILENAME}")
            modelo_wb.save(novo_nome)

            print("\n-----------------------------------------------------------------")
            print(f"🎉 Sucesso! O novo arquivo '{novo_nome}' foi criado.")
            print("   As abas 'VISÃO_GERAL' e 'VISÃO_GERAL_M1' devem ter sido recalculadas pelo Excel.")
            print("-----------------------------------------------------------------")

        except Exception as e:
            print(f"❌ Erro fatal ao manipular o Excel: {e}")

except FileNotFoundError as fnfe:
    print(f"\n❌ ERRO: Um ou mais arquivos Excel não foram encontrados no diretório 'downloads'.")
    print(f"Detalhe: {fnfe}")
    print("Certifique-se de que a automação do download funcionou ou que os arquivos existem localmente.")

except Exception as e:
    print(f"\n❌ Erro geral durante o processamento dos dados: {e}")

fim = datetime.now()
tempo_total = fim - inicio

horas, resto = divmod(tempo_total.total_seconds(), 3600)
minutos, segundos = divmod(resto, 60)

print(f"✅ Tempo total de execução: {int(horas)}h {int(minutos)}min {int(segundos)}s")