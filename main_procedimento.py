from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import os
from dotenv import load_dotenv
from utils import remover_acentos
import re

load_dotenv()

url = os.getenv("URL")
login_sistema = os.getenv("LOGIN")
senha_sistema = os.getenv("SENHA")
local_planilha = os.getenv("PLANILHA")

navegador = webdriver.Chrome()
acoes = ActionChains(navegador)

dados_filtrados = pd.DataFrame()

def fazer_login_sistema(login_sistema, senha_sistema, url):
    navegador.get(url)
    time.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="conta"]').send_keys(login_sistema)
    navegador.find_element(By.XPATH, '//*[@id="password"]').send_keys(senha_sistema)
    navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/form/div[2]').click()

    while len(navegador.find_elements(By.XPATH,'//*[@id="lookup_key_ppi_definicao_pactuacao_ram_lista_municipio_sede_id"]')) < 1:
        time.sleep(2)

def inserir_dados_padroes():
    input_ano = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="ppi_definicao_pactuacao_ano"]'))
    )

    municipio_sede = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="lookup_key_ppi_definicao_pactuacao_ram_lista_municipio_sede_id"]'))
    )

    input_ano.click()
    time.sleep(0.1)
    acoes.send_keys(Keys.BACKSPACE * 4).perform()
    time.sleep(0.1)
    input_ano.send_keys('2025')
    time.sleep(0.3)

    acoes.double_click(municipio_sede).perform()
    time.sleep(0.3)
    municipio_sede.send_keys('2')
    time.sleep(0.3)

def inserir_dados_variaveis(municipio, procedimento, valor_total, dados_planilha):
    input_valor = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="ppi_definicao_pactuacao_valor_pactuado"]'))
    )

    input_procedimento = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="lookup_key_ppi_definicao_pactuacao_adm_procedimento_id"]'))
    )

    selc_municipio = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="select2-lookup_name_ppi_definicao_pactuacao_ram_lista_municipio_referencia_id-container"]'))
    )  

    #inserir municipios
    selc_municipio.click()
    input_municipio = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/span/span/span[1]/input'))
    )
    input_municipio.send_keys(municipio)
    time.sleep(1)
    acoes.send_keys(Keys.ENTER).perform()
    time.sleep(1)

    campo_municipio = navegador.find_element(By.XPATH, '//*[@id="select2-lookup_name_ppi_definicao_pactuacao_ram_lista_municipio_referencia_id-container"]')
    texto_campo_municipio = campo_municipio.text[1:].strip()

    texto_campo_municipio_formatado = remover_acentos(texto_campo_municipio).upper()
    texto_municipio = remover_acentos(municipio).upper()

    if texto_municipio != texto_campo_municipio_formatado:
        mensagem_erro = f'Municipio de {municipio} está diferente do que o sistema selecionou: {texto_campo_municipio}'
        print(mensagem_erro)

        mensagem_sanitizada = re.sub(r'[\\/*?:"<>|]', '-', mensagem_erro)
        nome_arquivo = f'{municipio}_{procedimento}_{mensagem_sanitizada}'
        nome_arquivo = re.sub(r'\s+', '_', nome_arquivo)  

        os.makedirs('./err_procedimento', exist_ok=True)
        primeira_linha_df = pd.DataFrame([dados_planilha.iloc[0]])
        primeira_linha_df.to_excel(f'./err_procedimento/{nome_arquivo}.xlsx', index=False)
        dados_planilha.drop(index=dados_planilha.index[0], inplace=True)

        return True 
    #inserir valor
    acoes.double_click(input_valor).perform()
    input_valor.send_keys(valor_total)

    #inserir procedimento
    acoes.double_click(input_procedimento).perform()
    input_procedimento.send_keys(procedimento)
    acoes.send_keys(Keys.ENTER).perform()
    time.sleep(2)
    procedimento_selecionado = navegador.find_element(By.XPATH, '//*[@id="select2-lookup_name_ppi_definicao_pactuacao_adm_procedimento_id-container"]/span').text

    if 'Selecione...' in procedimento_selecionado:
        mensagem_erro = f'{municipio}-Procedimento {procedimento} não encontrado'
        #criar planilha com erro procedimento
        mensagem_sanitizada = re.sub(r'[\\/*?:"<>|]', '-', mensagem_erro)
        nome_arquivo = f'{municipio}_{procedimento}_{mensagem_sanitizada}'
        nome_arquivo = re.sub(r'\s+', '_', nome_arquivo)  

        os.makedirs('./err_procedimento', exist_ok=True)
        primeira_linha_df = pd.DataFrame([dados_planilha.iloc[0]])
        primeira_linha_df.to_excel(f'./err_procedimento/{nome_arquivo}.xlsx', index=False)
        dados_planilha.drop(index=dados_planilha.index[0], inplace=True)

        return True
    
    btn_incluir = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_adicionar"]'))
    )

    btn_incluir.click()

    time.sleep(2)

    btn_ok = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="fwk_show_dialog_modal"]/div/div/div[3]/div/button[1]'))
        )
    
    btn_ok.click()

def main_procedimento():
    colunas_texto = ['GRUPO', 'SUBGRUPO', 'FORMA DE ORGANIZACAO', 'PROCEDIMENTO']
    dados_planilha = pd.read_excel(local_planilha, dtype={coluna: str for coluna in colunas_texto})
    dados_planilha['COTAS'] = dados_planilha['COTAS'].fillna(0)
    dados_planilha = dados_planilha.loc[dados_planilha['COTAS'] != 0]
    dados_planilha['TOTAL'] = ((dados_planilha['VALOR'] * 100) * dados_planilha['COTAS']) / 100
    dados_planilha['TOTAL'] = dados_planilha['TOTAL'].round(2)
    condicao_para_excluir = (dados_planilha['GRUPO'] == '02') & (dados_planilha['SUBGRUPO'].isin(['02', '03']))
    filtro = ~condicao_para_excluir
    dados_planilha = dados_planilha[filtro]
    dados_planilha['PROCEDIMENTO'] = dados_planilha['PROCEDIMENTO'].astype(str).str.strip()
    dados_planilha[['CODIGO', 'DESCRICAO']] = dados_planilha['PROCEDIMENTO'].str.strip().str.split(' ', n=1, expand=True)
    print(dados_planilha)
    dados_planilha.to_excel('sem_zero_procedimento.xlsx', index=False)

    fazer_login_sistema(login_sistema,senha_sistema,url)
    inserir_dados_padroes()

    while not dados_planilha.empty:
        dados_priemeira_linha = dados_planilha.iloc[0].values
        procedimento = dados_priemeira_linha[8]
        valor_total=dados_priemeira_linha[7]
        municipio = dados_priemeira_linha[6]

        erro_ao_inserir = inserir_dados_variaveis(municipio, procedimento, valor_total, dados_planilha)

        if erro_ao_inserir:
            continue
        
        print(f'{municipio} - {procedimento} - {valor_total} Inserido com SUCESSO!')
        dados_planilha.drop(index=dados_planilha.index[0], inplace=True)
        dados_planilha.to_excel('sem_zero_procedimento.xlsx', index=False)

main_procedimento()
