from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from agenda import Agenda
from agendaN import a
from datetime import datetime



def run_call(email, password):

    try:

        veiculo = a.extract_data()[1][0][4]
        date = a.extract_data()[1][0][2]
        date_obj = datetime.strptime(date, "%Y-%m-%d")
        format_date = date_obj.strftime("%d/%m/%Y")

        driver = webdriver.Chrome()

        driver.get('https://sgcplus.fieb.org.br/') 

        sleep(2)

        barra_de_usuario = driver.find_element(By.XPATH, '//*[@id="usuario"]')
        barra_de_usuario.send_keys(email)

        barra_senha = driver.find_element(By.XPATH, '//*[@id="senha"]')
        barra_senha.send_keys(password)

        botao = driver.find_element(By.XPATH, '//*[@id="main"]/div/div/form/div[3]/button')
        botao.click()

        sleep(3.5)

        elements = driver.find_elements(By.XPATH, "//*")
        for el in elements:
            try:
                if el.is_displayed() and el.tag_name == 'button':
                    print(f"Tag: {el.tag_name}, ID: {el.get_attribute('id')}, Class: {el.get_attribute('class')}, Text: {el.text[:30]}")
                    el.click()
            except:
                continue

        sleep(1)

        novo_chamado = driver.find_element(By.XPATH, '//*[@id="content"]/nav/ol/li[1]')
        novo_chamado.click()

        sleep(2.5)

        aba_cimatec = driver.find_element(By.XPATH, '//*[@id="senaicimatec-tab"]')
        aba_cimatec.click()
        sleep(3)


        aba_nsi = driver.find_element(By.XPATH, '//*[@id="senaicimatec"]/div/div/p[7]/a')
        aba_nsi.click()
        sleep(4)

        area_texto = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[3]/form/div[4]/textarea')
        area_texto.send_keys(f"Olá, gostaria de realizar uma reserva do seguinte tipo de veículo: {veiculo} para o dia {format_date}")

        sleep(6)

        driver.quit()
        return 1
    except:
        driver.quit()
        return 0



# run_call('mateus.freire@fbest.org.br' , "Meufelho12*")