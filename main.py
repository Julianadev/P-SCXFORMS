from selenium import webdriver as controle
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
import traceback
import logging
import os

name_nome = '168680934'
name_email = '168680946_1229936298'
name_telefone = '168681004_1229936585Country'
xpath_sexo_feminino = '//*[@id="168680984_1229936473_label"]/span[1]'
xpath_sexo_masculino = '//*[@id="168680984_1229936474_label"]/span[1]'
name_sobre = '168681061'
xpath_botao_enviar = '//*[@id="patas"]/main/article/section/form/div[2]/button'

def preenche_planilha(navegador, nome, email, telefone, sexo, sobre):
    navegador.get('https://pt.surveymonkey.com/r/YP6FH5Z')

    try:
        WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.NAME, name_nome)))
        nome = navegador.find_element(By.NAME, name_nome).send_keys(nome)
        email = navegador.find_element(By.NAME, name_email).send_keys(email)
        telefone = navegador.find_element(By.NAME, name_telefone).send_keys(telefone)
        if sexo == 'masculino':
            navegador.find_element(By.XPATH, xpath_sexo_masculino).click()
        else:
            navegador.find_element(By.XPATH, xpath_sexo_feminino).click()
        sobre = navegador.find_element(By.NAME, name_sobre).send_keys(sobre)

        navegador.find_element(By.XPATH, xpath_botao_enviar).click()

        return nome, email, telefone, sexo, sobre
    except Exception as e:
        logging.error('Ocorreu um erro: %s', e)
        traceback.print_exc()
        return None, None, None, None, None

def main():

    try:
        caminho_arquivo = input('Digite o caminho do arquivo: ').strip('"')

        if caminho_arquivo.lower().endswith(('.xls', '.xlsx', '.csv')):
            caminho_arquivo = os.path.abspath(caminho_arquivo)
            df_leitura = pd.read_excel(caminho_arquivo)

            options = controle.EdgeOptions()
            options.add_argument('--headless')
            navegador = controle.Edge(options=options)

            if set(['NOME', 'EMAIL', 'TELEFONE', 'SEXO', 'SOBRE']).issubset(df_leitura.columns):
                df = pd.DataFrame(columns=['NOME', 'EMAIL', 'TELEFONE', 'SEXO', 'SOBRE'])
            else:
                print('Colunas necessárias não encontradas no arquivo Excel.')
                return

            for _, row in df_leitura.iterrows():
                nome_atual = row['NOME']
                email_atual = row['EMAIL']
                telefone_atual = row['TELEFONE']
                sexo_atual = row['SEXO']
                sobre_atual = row['SOBRE']

                nome, email, telefone, sexo, sobre = preenche_planilha(
                    navegador,
                    nome_atual,
                    email_atual,
                    telefone_atual,
                    sexo_atual,
                    sobre_atual
                )

                df.loc[len(df.index)] = [nome_atual, email_atual, telefone_atual, sexo_atual, sobre_atual]

            navegador.close()

            df.to_excel(caminho_arquivo, index=False)
            print('Arquivo enviado com sucesso para o formulário')
            os.startfile(caminho_arquivo)
        else:
            print('Formato de arquivo inválido. Suportado: .xls, .xlsx, .csv')
    except FileNotFoundError as e:
        print('Caminho não encontrado: ', e)

if __name__ == "__main__":
    main()

