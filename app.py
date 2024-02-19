from chave import chave_api
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from anticaptchaofficial.recaptchav2proxyless import *
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.print_page_options import PrintOptions
from base64 import b64decode
from openpyxl import Workbook

import pandas as pd
import time

# ----------------------------------------------------------------------------------------------------------
#     Automação para gerar boletos de 2023 no site da Funesbom. Realizado em Python usando selenium, base64
# openpyxl, webdriver Chrome e anticaptcha do https://anti-captcha.com/
#     Consistem em abrir o site da Funesbom, preencher os campos: número do CBMERJ e digito CBMERJ, com o
# que consta numa planilha de excel. Após isso, envia as informações para a resolução do captcha. Com a tela
# dos boletos acessada, seleciona o botão de 2023 para gerar o boleto. Com o boleto em tela é dado o comando
# para transforma-lo em base64 e posterior é realizado a decodificação e escrito (salvo) na pasta onde o
# executavel do programa está armazenado.
#
# Feito por Melchior Passos Araújo em fevereiro de 2024.
#
# ----------------------------------------------------------------------------------------------------------

print_options = PrintOptions()
print_options.page_ranges = ['1']


def find_window(url: str):
    for window in wids:
        navegador.switch_to.window(window)
        if url in navegador.current_url:
            break

link = 'https://www.funesbom.rj.gov.br/sistema/imovel'
captcha_id = "6LdZ1EQkAAAAAAKRMr1Dhld-8WiyW5Qt0HgCXyfa"
navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
tabela = pd.read_excel("numeros_cbmerj.xlsx")

for i, numero in enumerate(tabela["CBMERJ"]):
    numero_dv = tabela.loc[i, "DV"]

    print("Contador: " + str(i))
    nome_arquivo = str(numero) + "-" + str(numero_dv) + ".pdf"
    print(nome_arquivo)

    navegador.get(link)

    form_numero = navegador.find_element(
        By.ID, "cbmerj").find_element(By.ID, "cbmerj")
    form_dv = navegador.find_element(
        By.ID, "cbmerj").find_element(By.ID, "cbmerj_dv")

    form_numero.send_keys(numero)
    form_dv.send_keys(str(numero_dv))

    n_cbmerj = numero
    dv_cbmerj = numero_dv

    solver = recaptchaV2Proxyless()
    # verbose(0) não imprime nada, verbose(1) imprime a cada 3 segundos o status da requisicao
    solver.set_verbose(1)
    solver.set_key(chave_api)
    solver.set_website_url(link)
    solver.set_website_key(captcha_id)

    resposta = solver.solve_and_return_solution()

    if resposta != 0:

        # preencher o text que está escondido g-recaptcha-response
        navegador.execute_script(
            f"document.getElementById('g-recaptcha-response').innerHTML = '{resposta}'")
        navegador.find_element(By.ID, "btnEnviar").click()
        time.sleep(2)
        try:
            navegador.find_element(By.CLASS_NAME, "avisosTP").find_element(
                By.NAME, "botao-2023").click()
            time.sleep(2)
            # Trocando para a aba do Boleto
            navegador.current_window_handle
            wids = navegador.window_handles
            find_window('boleto')

            # imprimindo um boleto em pdf
            base64code = navegador.print_page(print_options)

            b64 = base64code
            bytes = b64decode(b64, validate=True)
            if bytes[0:4] != b'%PDF':
                raise ValueError('Missing the PDF file signature')

            f = open(nome_arquivo, 'wb')
            f.write(bytes)
            f.close()

            if (i % 10 == 0):
                navegador.quit()
                navegador = webdriver.Chrome(
                    service=Service(ChromeDriverManager().install()))
                link = 'https://www.funesbom.rj.gov.br/sistema/imovel'
                captcha_id = "6LdZ1EQkAAAAAAKRMr1Dhld-8WiyW5Qt0HgCXyfa"

        except:
            print("O número " + str(numero) + "-" +
                  str(numero_dv) + " não possui botão de 2023.")

    else:
        print(solver.err_string)

navegador.quit()
