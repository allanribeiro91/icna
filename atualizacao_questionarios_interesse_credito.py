from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import shutil
import pandas as pd
import pyautogui
import time
from datetime import datetime
import subprocess

def countdown_timer(seconds):
    for i in range(seconds, 0, -1):
        print(f"Tempo restante: {i} segundos", end='\r')
        time.sleep(1)

def baixar_dados_questionarios_ateg():
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)
    url_intranet = "https://www.cna.org.br/intranet/#"
    navegador.get(url_intranet)

    #Entrar na INTRANET
    navegador.find_element(By.ID, 'TXT_LOGIN').send_keys(usuario)
    navegador.find_element(By.ID, 'TXT_SENHA').send_keys(senha)
    navegador.find_element(By.CSS_SELECTOR, '.btn.btn-success.pull-right').click()

    #Entrar na ATEG
    time.sleep(3)
    url_ateg = 'https://ateg.cna.org.br/ateg/public/painelpbi'
    navegador.get(url_ateg)

    #Ir até a página de questionários
    url_questionarios = 'https://ateg.cna.org.br/ateg/public/ateg/questionario-vinculo/relatorioQuestionarioExpandido'
    navegador.get(url_questionarios)
    time.sleep(1)

    #Selecionar o questionário
    dropdown_field = navegador.find_element(By.XPATH, '//*[@id="codQuestionario_chzn"]/a/span')
    dropdown_field.click()
    time.sleep(1)
    option = navegador.find_element(By.XPATH, '//*[@id="codQuestionario_chzn"]//li[text()="Interesse do produtor por Crédito - Custeio"]')
    option.click()
    navegador.find_element(By.ID, 'btnFiltrarQuestionario').click()
    countdown_timer(80)

    #Baixar os dados
    navegador.find_element(By.XPATH, '//*[@id="gridContainer"]/div/div[4]/div/div/div[3]/div[1]/div/div/div').click()
    time.sleep(5)

def get_latest_file(path):
    # Lista todos os arquivos no diretório
    files = [os.path.join(path, f) for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
    
    # Retorna o arquivo com a data de criação mais recente
    return max(files, key=os.path.getctime)

def mover_arquivo():
    #parâmetros
    download_folder = 'C:\\Users\\allan.ribeiro\\Downloads\\'
    destination_folder = 'G:\\Meu Drive\\Instituto CNA\\FIAGRO\\Questionário Crédito'
    new_filename = 'ateg_questionario_fonte_dados'

    #1. Identificar o arquivo mais recente
    latest_file = get_latest_file(download_folder)

    #2. Renomear o arquivo
    base, ext = os.path.splitext(latest_file)  # Separa o nome do arquivo e sua extensão
    renamed_file = os.path.join(download_folder, new_filename + ext)
    os.rename(latest_file, renamed_file)

    #3. Mover (e substituir) o arquivo renomeado para a pasta de destino
    destination_file = os.path.join(destination_folder, new_filename + ext)
    if os.path.exists(destination_file):  # Se o arquivo já existe, remova-o
        os.remove(destination_file)
    shutil.move(renamed_file, destination_folder)

def formatacao_dados():
    
    lista_campos = [
        "prop_id", "prop_municipio", "prop_nome", "prop_endereco", "produtor_cpf", 
        "produtor_nome", "produtor_sexo", "produtor_telefone", "produtor_nascimento_data", 
        "produtor_nascimento_local", "produtor_nacionalidade", "produtor_estado_civil", 
        "produtor_identidade_orgao", "produtor_identidade_numero", "produtor_identidade_data", 
        "produtor_email", "produtor_endereco_corresp", "produtor_endereco_corresp_cep", 
        "produtor_conjuge", "prop_area_produtiva_irrigada", "prop_area_produtiva_propria", 
        "prop_area_produtiva_total", "atvp_atividade_principal", "atvp_produto", 
        "atvp_receita_produtividade_ult_periodo", "atvp_receita_preco_ult_periodo", 
        "atvp_receita_bruta_anual", "atvp_custo_coe", "atvp_custo_cot", 
        "atvp_receita_preco_prox_periodo", "leite_cbt", "leite_ccs", "credito_interesse", 
        "credito_custeio_valor_interesse", "credito_cna", "tecnico_nome", "ciencia", 
        "autorizacao"
    ]

    ######PARTE 1 - MESCLAR OS DADOS
    # Carregar os dados do arquivo "ateg_questionario_fonte_dados.xlsx"
    df1 = pd.read_excel("ateg_questionario_fonte_dados.xlsx", sheet_name="Sheet")

    # Ler o arquivo "estrutura_dados.xlsx" e a primeira planilha (por padrão)
    df2 = pd.read_excel("estrutura_dados.xlsx")

    # Realizar a junção dos DataFrames usando a coluna "Pergunta"
    merged_df = pd.merge(df1, df2, how="left", left_on="Pergunta", right_on="Pergunta")

    # Selecionar e reordenar as colunas desejadas
    final_df = merged_df[["Id. Propriedade", "ID Pergunta", "Resposta"]]

    # Remover linhas duplicadas considerando "Id. Propriedade" e "ID Pergunta"
    final_df.drop_duplicates(subset=["Id. Propriedade", "ID Pergunta"], inplace=True)

    # Salvar o DataFrame final como um novo arquivo Excel
    final_df.to_excel("dados_mesclados.xlsx", index=False)

    ######PARTE 2 - TRANSPOR OS DADOS
    # Carregar o arquivo Excel
    df = pd.read_excel("dados_mesclados.xlsx")

    # Transpor os dados
    df_pivot = df.pivot(index="Id. Propriedade", columns="ID Pergunta", values="Resposta").reset_index()

    # Renomear a coluna "Id. Propriedade" para "prop_id"
    df_pivot.rename(columns={"Id. Propriedade": "prop_id"}, inplace=True)

    # Reordenar as colunas com base na lista de campos
    # Isso também remove quaisquer colunas que não estejam na lista
    df_pivot = df_pivot[lista_campos]

    # Remover linhas duplicadas baseadas na coluna "prop_id"
    df_pivot.drop_duplicates(subset=["prop_id"], keep='first', inplace=True)

    # Salvar o DataFrame em um novo arquivo Excel
    df_pivot.to_excel("dados_formatados.xlsx", index=False)

def atualizar_dados_pbi():
    #Iniciando
    data_hora = datetime.now()

    # Executar o Power BI Desktop com o arquivo como argumento
    pbidesktop_path = "C:\\Program Files\\Microsoft Power BI Desktop\\bin\\PBIDesktop.exe"
    pbix_file_path = "G:\\Meu Drive\\Instituto CNA\\FIAGRO\\Questionário Crédito\\FIAGRO_Questionário Crédito_Interesse.pbix"
    subprocess.Popen([pbidesktop_path, pbix_file_path])
    countdown_timer(45)

    #Clicar em atualizar dados
    pyautogui.click(x=716, y=95)
    countdown_timer(120)

    #Clicar em salvar
    pyautogui.click(x=22, y=15)
    countdown_timer(20)

    #Clicar em publicar
    pyautogui.click(x=1169, y=98)
    countdown_timer(6)

    #Selecionar DATEG
    pyautogui.click(x=758, y=509)
    countdown_timer(3)

    #Clicar em Selecionar
    pyautogui.click(x=1116, y=662)
    countdown_timer(6)

    #Clicar em Substituir
    pyautogui.click(x=1050, y=630)
    countdown_timer(60)

    #Clicar em Substituir
    pyautogui.click(x=1149, y=652)
    countdown_timer(10)


#Executar processo de atualização
print("INÍCIO")
data_hora = datetime.now()
print("Data/hora: " + data_hora.strftime('%Y-%m-%d %H:%M:%S'))
baixar_dados_questionarios_ateg()
mover_arquivo()
formatacao_dados()
atualizar_dados_pbi()
data_hora = datetime.now()
print("Data/hora: " + data_hora.strftime('%Y-%m-%d %H:%M:%S'))
print("FIM")
