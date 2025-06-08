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

ERROS = []
SUCESSO = []
JA_INSERIDOS = []

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


def inserir_dados_variaveis(valor_total_grupo, grupo, subgrupo, forma_organizacao, municipio, filtro, dados_planilha):
    global dados_filtrados

    input_grupo = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="lookup_key_ppi_definicao_pactuacao_adm_grupo_proced_id"]'))
    )
    input_subgrupo = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="lookup_key_ppi_definicao_pactuacao_adm_subgrupo_proced_id"]'))
    )
    input_forma_organizacao = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="lookup_key_ppi_definicao_pactuacao_adm_forma_organizacao_proced_id"]'))
    )
    input_valor = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="ppi_definicao_pactuacao_valor_pactuado"]'))
    )
    selc_municipio = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="select2-lookup_name_ppi_definicao_pactuacao_ram_lista_municipio_referencia_id-container"]'))
    )

    input_grupo.send_keys(grupo)
    acoes.send_keys(Keys.ENTER).perform()

    input_subgrupo.click()
    time.sleep(0.3)
    input_subgrupo.send_keys(subgrupo)
    acoes.send_keys(Keys.ENTER).perform()
    time.sleep(1)

    valor_subgrupo = navegador.find_element(By.XPATH, '//*[@id="select2-lookup_name_ppi_definicao_pactuacao_adm_subgrupo_proced_id-container"]/span').text

    while 'Selecione...' in valor_subgrupo:
        input_subgrupo.click()
        time.sleep(0.1)
        input_subgrupo.send_keys(subgrupo)
        acoes.send_keys(Keys.ENTER).perform()
        time.sleep(3)
        valor_subgrupo = navegador.find_element(By.XPATH, '//*[@id="select2-lookup_name_ppi_definicao_pactuacao_adm_subgrupo_proced_id-container"]/span').text
        print(valor_subgrupo)

    input_forma_organizacao.click()
    time.sleep(0.1)
    input_forma_organizacao.send_keys(str(forma_organizacao))
    acoes.send_keys(Keys.ENTER).perform()
    time.sleep(0.7)

    valor_forma_organizacao = navegador.find_element(By.XPATH, '//*[@id="select2-lookup_name_ppi_definicao_pactuacao_adm_forma_organizacao_proced_id-container"]/span').text

    while 'Selecione...' in valor_forma_organizacao:
        input_forma_organizacao.click()
        time.sleep(0.1)
        input_forma_organizacao.send_keys(str(forma_organizacao))
        acoes.send_keys(Keys.ENTER).perform()
        time.sleep(3)
        valor_forma_organizacao = navegador.find_element(By.XPATH, '//*[@id="select2-lookup_name_ppi_definicao_pactuacao_adm_forma_organizacao_proced_id-container"]/span').text

    input_valor.send_keys(valor_total_grupo)
    time.sleep(0.3)

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
        nome_arquivo = f'{municipio}-{mensagem_sanitizada}-{grupo}_{subgrupo}_{forma_organizacao}'
        nome_arquivo = re.sub(r'\s+', '_', nome_arquivo)  

        os.makedirs('./err', exist_ok=True)
        dados_filtrados = dados_planilha[filtro]
        dados_filtrados.to_excel(f'./err/{nome_arquivo}.xlsx', index=False)

        dados_planilha.drop(dados_planilha[filtro].index, inplace=True)

        return True

    btn_incluir = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_adicionar"]'))
    )

    btn_incluir.click()

    time.sleep(1)

def limpar_campos_grupo():
    input_valor = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="ppi_definicao_pactuacao_valor_pactuado"]'))
    )
    input_grupo = WebDriverWait(navegador, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="lookup_key_ppi_definicao_pactuacao_adm_grupo_proced_id"]'))
    )

    acoes.double_click(input_grupo).perform()
    acoes.send_keys(Keys.BACKSPACE).perform()
    acoes.double_click(input_valor).perform()
    acoes.send_keys(Keys.BACKSPACE).perform()
    time.sleep(1)

def main(login_sistema, senha_sistema, url):
    global dados_filtrados
    global dados_planilha
    fazer_login_sistema(login_sistema, senha_sistema, url)

    colunas_texto = ['GRUPO', 'SUBGRUPO', 'FORMA DE ORGANIZACAO']
    dados_planilha = pd.read_excel(local_planilha, dtype={coluna: str for coluna in colunas_texto})
    dados_planilha['COTAS'] = dados_planilha['COTAS'].fillna(0)
    dados_planilha = dados_planilha.loc[dados_planilha['COTAS'] != 0]
    dados_planilha['TOTAL'] = dados_planilha['VALOR'] * dados_planilha['COTAS']
    filtro_grupo = (dados_planilha['GRUPO'] == '02') & ((dados_planilha['SUBGRUPO'] == '03') | (dados_planilha['SUBGRUPO'] == '02'))

    dados_planilha = dados_planilha[filtro_grupo]

    print(dados_planilha)
    dados_planilha.to_excel('sem_zero.xlsx', index=False)

    while not dados_planilha.empty:
        dados_priemeira_linha = dados_planilha.iloc[0].values
        grupo = dados_priemeira_linha[0]
        subgrupo = dados_priemeira_linha[1]
        forma_organizacao = dados_priemeira_linha[2]
        municipio = dados_priemeira_linha[6]
        valor_total_grupo = 0

        filtro = (dados_planilha['GRUPO'] == grupo) & (dados_planilha['SUBGRUPO'] == subgrupo) & (dados_planilha['FORMA DE ORGANIZACAO'] == forma_organizacao) & (dados_planilha['MUNICIPIO'] == municipio)
        dados_filtrados = dados_planilha[filtro]

        if dados_filtrados.empty:
            print(f"Nenhum dado encontrado para {grupo}, {subgrupo}, {forma_organizacao}, {municipio}. Pulando.")
            dados_planilha = dados_planilha.drop(dados_planilha[filtro].index)
            continue

        inserir_dados_padroes()

        for elemento in dados_filtrados.values:
            valor_exame = float(elemento[4]) * 100
            qtd_exames = elemento[5]
            valor_total_exame = valor_exame * qtd_exames
            valor_total_grupo += valor_total_exame

        valor_total_grupo_formatado = valor_total_grupo / 100
        limpar_campos_grupo()

        erro_ao_inserir = inserir_dados_variaveis(valor_total_grupo_formatado, grupo, subgrupo, forma_organizacao, municipio, filtro, dados_planilha)

        if erro_ao_inserir:
            limpar_campos_grupo()
            print("teve erro e pulei o loop")
            continue

        mensagem_aviso = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="fwk_show_dialog_modal"]/div/div/div[2]/div'))
        )
        texto = mensagem_aviso.text
        time.sleep(0.5)
        btn_ok = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="fwk_show_dialog_modal"]/div/div/div[3]/div/button[1]'))
        )

        if "Registro inserido com sucesso!" not in texto:
            mensagem_erro = f'{texto} Item: {grupo, subgrupo, forma_organizacao} - {valor_total_grupo_formatado} - {municipio}'
            btn_ok.click()

            if "Já existe uma regra cadastrada para esses paramêtros." not in texto:
                ERROS.append(mensagem_erro)
                dados_planilha = dados_planilha.drop(dados_planilha[filtro].index)
                dados_filtrados = dados_planilha[filtro]

                nome_arquivo = f'{municipio}-{mensagem_erro}-{grupo, subgrupo, forma_organizacao}'
                nome_arquivo = re.sub(r'\s+', '_', nome_arquivo)  

                dados_filtrados.to_excel(f'./err/{nome_arquivo}.xlsx', index=True)
                print(mensagem_erro)
                continue

            JA_INSERIDOS.append(mensagem_erro)
            print(mensagem_erro)
            limpar_campos_grupo()
            dados_planilha = dados_planilha.drop(dados_planilha[filtro].index)
            continue

        btn_ok.click()
        SUCESSO.append(f'{grupo, subgrupo, forma_organizacao} - {municipio}')
        print(f'{grupo, subgrupo, forma_organizacao} - {municipio}')
        dados_planilha = dados_planilha.drop(dados_planilha[filtro].index)

main(login_sistema, senha_sistema, url)
