from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import unicodedata

navegador = webdriver.Chrome()
acoes = ActionChains(navegador)

def apagar_campo():
    acoes.send_keys(Keys.BACKSPACE).perform()
    acoes.send_keys(Keys.BACKSPACE).perform()
    acoes.send_keys(Keys.BACKSPACE).perform()
    acoes.send_keys(Keys.BACKSPACE).perform()
    print('Campo apagado com sucesso!')

def deletar_linha_cota_zerada(cotas, dados_planilha, i):
    if cotas == 0:
            dados_planilha.drop(i, inplace=True)
            dados_planilha.to_excel('exames_sem_cota_zero.xlsx', index=False)

def apagar_primeira_linha(dados_planilha):
    dados_planilha = dados_planilha.drop(0).reset_index(drop=True)
    dados_planilha.to_excel('exames_sem_cota_zero.xlsx', index=False)

def remover_acentos(texto):
    texto_normalizado = unicodedata.normalize('NFD', texto)
    texto_sem_acentos = ''.join(c for c in texto_normalizado if unicodedata.category(c) != 'Mn')
    return texto_sem_acentos