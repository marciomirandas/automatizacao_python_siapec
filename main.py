# Importa as bibliotecas utilizadas
import time
from datetime import datetime, timedelta
from glob import glob

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


corretos = 0
nao_processados = 0
erro = 0

# Concatena e abre todos os arquivos excel na pasta base
arquivos = sorted(glob('arquivos/base/*.xlsx'))
todos_arquivos = pd.concat((pd.read_excel(cont)
                           for cont in arquivos), ignore_index=True)


# Instala a extensão para o chrome
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)


# Cria um dataframe com os dados certos
dados_corretos = pd.DataFrame(
    columns=["COD", "TIME", "NOTES", "LATITUDE", "LONGITUDE"])

# Cria um dataframe com os dados errados
dados_errados = pd.DataFrame(
    columns=["COD", "TIME", "NOTES", "LATITUDE", "LONGITUDE", "ERRO"])

# Cria um dataframe com os dados não processados
dados_nao_processados = pd.DataFrame(
    columns=["COD", "TIME", "NOTES", "LATITUDE", "LONGITUDE"])


# Entra no site
navegador.get("https://siapec3.emdagro.se.gov.br/siapec3/login.wsp")

time.sleep(5)


# Faz login
navegador.find_element(
    By.XPATH, '//*[@id="formLogin:user:input"]').send_keys("**********")
time.sleep(1)

navegador.find_element(
    By.XPATH, '//*[@id="formLogin:pass:input"]').send_keys("********")
time.sleep(1)

navegador.find_element(
    By.XPATH, '//*[@id="formLogin:captchaDigitado"]').click()
time.sleep(10)

navegador.find_element(
    By.XPATH, '//*[@id="formLogin:btnFazerLogin:button"]').click()
time.sleep(5)


# Define o horário de encerramento do programa
minuto239 = datetime.now() + timedelta(minutes=239)


# Navega no site
navegador.find_element(
    By.XPATH, '//*[@id="accordionOne:group10"]/div/div').click()
time.sleep(1)

navegador.find_element(
    By.XPATH, '//*[@id="accordionOne:group10:itens:it_040:card"]').click()
time.sleep(5)


# Pega o XPath do iframe e atribui a uma variável
iframe = navegador.find_element(
    By.XPATH, '//*[@id="gerenciadorSessoes"]/div[2]/iframe')


# Muda o foco para o iframe
navegador.switch_to.frame(iframe)


# Navega no menu pop-up
navegador.find_element(
    By.XPATH, '//*[@id="id_int_85_PROPRIEDADES"]').click()
time.sleep(1)

navegador.find_element(
    By.XPATH, '//*[@id="id_int_299_Propriedades"]').click()
time.sleep(1)

# Percorre os registros do dataframe
for i in range(0, len(todos_arquivos.index)):

    # Limita o tempo de execução do código
    if (datetime.now() < minuto239):

        # Caso encontre algum erro que precise reiniciar a página
        if erro == 1:
            navegador.refresh()

            try:
                # Navega no site
                navegador.find_element(
                    By.XPATH, '//*[@id="accordionOne:group10"]/div/div').click()
                time.sleep(1)

                navegador.find_element(
                    By.XPATH, '//*[@id="accordionOne:group10:itens:it_040:card"]').click()
                time.sleep(5)

                # Pega o XPath do iframe e atribui a uma variável
                iframe = navegador.find_element(
                    By.XPATH, '//*[@id="gerenciadorSessoes"]/div[2]/iframe')

                # Muda o foco para o iframe
                navegador.switch_to.frame(iframe)
            except:
                dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                         todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                         todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], "Etapa Erro"]
                continue

            erro = 0

        # Clica em propriedade rural
        while True:
            tentativa = 0
            try:
                navegador.find_element(
                    By.XPATH, '//*[@id="id_310_PropriedadeRural"]').click()
                break
            except:
                time.sleep(3)
                if (tentativa == 3):
                    erro = 1
                    break
                tentativa += 1

        if (tentativa == 3):
            erro = 1
            dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                     todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                     todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], "Etapa Propriedade Rural"]
            continue

        # Caso encontre algum erro que precise reiniciar a página
        if erro == 1:
            navegador.refresh()

            # Navega no site
            navegador.find_element(
                By.XPATH, '//*[@id="accordionOne:group10"]/div/div').click()
            time.sleep(1)

            navegador.find_element(
                By.XPATH, '//*[@id="accordionOne:group10:itens:it_040:card"]').click()
            time.sleep(5)

            # Pega o XPath do iframe e atribui a uma variável
            iframe = navegador.find_element(
                By.XPATH, '//*[@id="gerenciadorSessoes"]/div[2]/iframe')

            # Muda o foco para o iframe
            navegador.switch_to.frame(iframe)

            # Navega no menu pop-up
            navegador.find_element(
                By.XPATH, '//*[@id="id_int_85_PROPRIEDADES"]').click()
            time.sleep(1)

            navegador.find_element(
                By.XPATH, '//*[@id="id_int_299_Propriedades"]').click()
            time.sleep(1)

            erro = 0

            # Clica em propriedade rural
            while True:
                tentativa = 0
                try:
                    navegador.find_element(
                        By.XPATH, '//*[@id="id_310_PropriedadeRural"]').click()
                    break
                except:
                    time.sleep(3)
                    if (tentativa == 3):
                        erro = 1
                        break
                    tentativa += 1

        time.sleep(5)

        # Escreve o código da propriedade
        try:
            navegador.find_element(
                By.XPATH, '//*[@id="tmp.propriedade.codigo"]').send_keys(str(todos_arquivos["COD"][i]))
            navegador.find_element(By.XPATH, '//*[@id="btnPesq"]').click()
        except:
            dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                     todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                     todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], "Etapa Propriedade"]
            continue
        time.sleep(3)

        j = 2
        # Clica no código da propriedade
        while True:
            try:
                proprietario = navegador.find_element(
                    By.XPATH, f'//*[@id="smlista"]/div[14]/table/tbody/tr[{j}]/td[4]/font').text

                if 'Proprietário' in proprietario:
                    navegador.find_element(
                        By.XPATH, f'//*[@id="smlista"]/div[14]/table/tbody/tr[{j}]/td[1]/a').click()
                    break

                j += 1

            except:
                dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                         todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                         todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], "Etapa Clique Código"]
                erro = 1
                break

        if erro == 1:
            erro = 0
            continue

        time.sleep(2)

        # Verifica se a propriedade tem área válida
        try:
            area = navegador.find_element(
                By.XPATH, '//*[@id="tmp.area"]').get_attribute('value')

            if (area == '0,0000' or area == ''):
                navegador.find_element(
                    By.XPATH, '//*[@id="tmp.area"]').clear()

                navegador.find_element(
                    By.XPATH, '//*[@id="tmp.area"]').send_keys('0,0001')

                time.sleep(1)

        except:
            dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                     todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                     todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], "Etapa de verificação da área"]
            continue

        # Dados da latitude
        try:
            navegador.find_element(
                By.XPATH, '//*[@id="latitude_span"]/div/div[1]/input').clear()

            navegador.find_element(
                By.XPATH, '//*[@id="latitude_span"]/div/div[1]/input').send_keys(todos_arquivos["LATITUDE"][i].split("°")[0])

            navegador.find_element(
                By.XPATH, '//*[@id="latitude_span"]/div/div[2]/input').clear()

            navegador.find_element(
                By.XPATH, '//*[@id="latitude_span"]/div/div[2]/input').send_keys(todos_arquivos["LATITUDE"][i].split("°")[1].split("'")[0])

            navegador.find_element(
                By.XPATH, '//*[@id="latitude_span"]/div/div[3]/input').clear()

            navegador.find_element(
                By.XPATH, '//*[@id="latitude_span"]/div/div[3]/input').send_keys(round(float(todos_arquivos["LATITUDE"][i].split("°")[1].split("'")[1].split('"')[0]), 2))
            time.sleep(1)

        except:
            dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                     todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                     todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], "Etapa Latitude"]
            continue

        # Dados da longitude
        try:
            navegador.find_element(
                By.XPATH, '//*[@id="longitude_span"]/div/div[1]/input').clear()

            navegador.find_element(
                By.XPATH, '//*[@id="longitude_span"]/div/div[1]/input').send_keys(todos_arquivos["LONGITUDE"][i].split("°")[0])

            navegador.find_element(
                By.XPATH, '//*[@id="longitude_span"]/div/div[2]/input').clear()

            navegador.find_element(
                By.XPATH, '//*[@id="longitude_span"]/div/div[2]/input').send_keys(todos_arquivos["LONGITUDE"][i].split("°")[1].split("'")[0])

            navegador.find_element(
                By.XPATH, '//*[@id="longitude_span"]/div/div[3]/input').clear()

            navegador.find_element(
                By.XPATH, '//*[@id="longitude_span"]/div/div[3]/input').send_keys(round(float(todos_arquivos["LONGITUDE"][i].split("°")[1].split("'")[1].split('"')[0]), 2))
            time.sleep(1)

        except:
            dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                     todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                     todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], "Etapa Longitude"]
            continue

        # Envia o formulário
        try:
            navegador.find_element(
                By.XPATH, '//*[@id="btnGravar"]').click()
        except:
            dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                     todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                     todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], "Etapa Salvar Dados"]
            continue
        time.sleep(5)

        # Tenta localizar o pop-up
        tentativa = 0
        while True:
            try:
                navegador.find_element(By.XPATH, '//*[@id="bodyPage"]/div[7]')

                texto = navegador.find_element(
                    By.XPATH, '//*[@id="bodyPage"]/div[7]/p').get_attribute("textContent")

                break
            except:
                time.sleep(3)
                if (tentativa == 3):
                    break
                tentativa += 1

        if (tentativa == 3):
            erro = 1
            dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                     todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                     todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], "Etapa Pop-up"]
            continue

        # Verifica se os dados foram salvos
        if (texto == "Propriedade atualizada com sucesso."):
            dados_corretos.loc[len(dados_corretos)] = [todos_arquivos["COD"][i],
                                                       todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                       todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i]]
            corretos += 1
        else:
            erro = 1
            dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                     todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                     todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], texto]
            continue

        tentativa = 0
        while True:
            try:
                navegador.find_element(
                    By.XPATH, '//*[@id="bodyPage"]/div[7]/div[7]/div/button').click()
                break
            except:
                time.sleep(3)
                if (tentativa == 3):
                    break
                tentativa += 1

        if (tentativa == 3):
            erro = 1
            dados_errados.loc[len(dados_errados)] = [todos_arquivos["COD"][i],
                                                     todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                     todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i], "Etapa Confirmação"]
            continue

        time.sleep(2)

    # Else do horário
    else:
        dados_nao_processados.loc[len(dados_nao_processados)] = [todos_arquivos["COD"][i],
                                                                 todos_arquivos["TIME"][i], todos_arquivos["NOTES"][i],
                                                                 todos_arquivos["LATITUDE"][i], todos_arquivos["LONGITUDE"][i]]
        nao_processados += 1
# Fim do for principal


# Salva os dados em excel
dados_corretos.to_excel(
    f'arquivos/corretos/corretos_{datetime.now().strftime("%d_%m_%Y_%H_%M_%S")}.xlsx', index=False)

dados_errados.to_excel(
    f'arquivos/errados/errados_{datetime.now().strftime("%d_%m_%Y_%H_%M_%S")}.xlsx', index=False)

dados_nao_processados.to_excel(
    f'arquivos/nao_processados/nao_processados_{datetime.now().strftime("%d_%m_%Y_%H_%M_%S")}.xlsx', index=False)

errados = len(todos_arquivos.index) - corretos - nao_processados
# Apresenta no terminal a quantidade de registros atualizados
print()
print()
print('-----------------------')
print(f'Dados Corretos: {corretos} / {len(todos_arquivos.index)}')
print('-----------------------')
print(f'Dados Errados: {errados} / {len(todos_arquivos.index)}')
print('-----------------------')
print(
    f'Dados Não Processados: {nao_processados} / {len(todos_arquivos.index)}')
print('-----------------------')
print()
