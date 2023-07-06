from selenium import webdriver
from selenium.webdriver.common.by import By
#navegador = webdriver.Chrome()
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from Password import login, senha

##comando para tirar o autorizar download

options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\natha\Downloads",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

navegador = webdriver.Chrome(options= options)

#Entrar na página de login (no nosso caso é login.html)
import os

caminho = os.getcwd()

arquivo = caminho + r'\login.html'
navegador.get(arquivo)

# login e senha
login = login
senha = senha

input_login = navegador.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys(login)
input_senha = navegador.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys(senha)

##clicar no botão de fazer login

botao_login = navegador.find_element(By.XPATH, '/html/body/div/form/button').click()

##importar a base de clientes

import pandas as pd
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

tabela = pd.read_excel('NotasEmitir.xlsx')

##para cada cliente - rodar o processo de emissão de nota fiscal

for linha in tabela.index:

  input_razao = navegador.find_element(By.NAME, 'nome').send_keys(tabela.loc[(linha, 'Cliente')])

  input_end = navegador.find_element(By.NAME, 'endereco').send_keys(tabela.loc[linha, 'Endereço'])

  input_bairro = navegador.find_element(By.NAME, 'bairro').send_keys(tabela.loc[linha, 'Bairro'])

  input_municipio = navegador.find_element(By.NAME, 'municipio').send_keys(tabela.loc[linha, 'Municipio'])

  input_cep = navegador.find_element(By.NAME, 'cep').send_keys(str(tabela.loc[linha, 'CEP']))

  estados = navegador.find_element(By.TAG_NAME, 'select')
  estados_select = Select(estados)
  estados_select.select_by_visible_text(tabela.loc[linha, 'UF'])

  input_cnpj = navegador.find_element(By.NAME, 'cnpj').send_keys(str(tabela.loc[linha, 'CPF/CNPJ']))

  input_inscricao = navegador.find_element(By.NAME, 'inscricao').send_keys(str(tabela.loc[linha, 'Inscricao Estadual']))

  input_desc = navegador.find_element(By.NAME, 'descricao').send_keys(tabela.loc[linha, 'Descrição'])

  input_quant = navegador.find_element(By.NAME, 'quantidade').send_keys(str(tabela.loc[linha, 'Quantidade']))

  input_vu = navegador.find_element(By.NAME, 'valor_unitario').send_keys(str(tabela.loc[linha, 'Valor Unitario']))

  input_valortotal = navegador.find_element(By.NAME, 'total').send_keys(str(tabela.loc[linha, 'Valor Total']))

# clicar em emitir nota fiscal

  click_emitir = navegador.find_element(By.CLASS_NAME, 'registerbtn').click()

  ##Recarregar a página para limpar o formulário

  navegador.refresh()

##para fechar o navegador utilizamos  o navegador.quit()
navegador.quit()

