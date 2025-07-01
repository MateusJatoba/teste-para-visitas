from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from agendaN import a
from datetime import datetime



def run_call(email, password):
    try:

        visitors = a.extract_data()[2] # tupla de visitantes
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
        sleep(1)

        input_nsp = driver.find_element(By.XPATH, '//*[@id="senaicimatec"]/div/div/p[13]/a')
        input_nsp.click()
        sleep(4)

        print('passei daqui')


        # -----------------------------------------------
        # seleção área
        # -------------------------------------------------

        # select_element = WebDriverWait(driver, 10).until(
        #     EC.presence_of_element_located((By.ID, 'AREA_ID_SELECT'))
        # )
        # driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_element)
        print("morri aqui")


        driver.execute_script("""
            const select = document.getElementById('AREA_ID_SELECT');
            select.value = '10092';
            select.dispatchEvent(new Event('change'));
        """)
        # -----------------------------------------------
        # seleção área
        # -------------------------------------------------
        sleep(4)
      #---------------------------------------------------
        #seleção classificação
        #---------------------------------------------------
        script = """
        var select = document.getElementById('CLASS_ID');
        select.selectedIndex = 5;
        select.dispatchEvent(new Event('change'));
        """
        driver.execute_script(script)
        #---------------------------------------------------
        #seleção classificação
        #--------------------------------------------------
        sleep(4)
        #----------------------------------------------------
        #seleção produto
        #----------------------------------------------------

        script2 = """
        var select = document.getElementById('PROD_ID');
        select.selectedIndex = 13;
        select.dispatchEvent(new Event('change'));
        """
        driver.execute_script(script2)

        sleep(3)


        descricao = driver.find_element(By.XPATH , '//*[@id="OCOR_DESCRICAO"]')
        descricao.send_keys(f'Solicito acesso dos seguintes visitantes no dia {format_date}: ')
        for item in visitors:

            descricao.send_keys(f'\nNome: {item[2]} - Cargo: {item[3]}')

        sleep(10)
        driver.quit()

        return 1

    except:
        driver.quit()
        return 0
    

run_call('mateus.freire@fbest.org.br' , 'Meufelho12*')