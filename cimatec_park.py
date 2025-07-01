from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

driver = webdriver.Chrome()

driver.get('https://sgcplus.fieb.org.br/') 

sleep(2)

barra_de_usuario = driver.find_element(By.XPATH, '//*[@id="usuario"]')
barra_de_usuario.send_keys('mateus.freire@fbest.org.br')

barra_senha = driver.find_element(By.XPATH, '//*[@id="senha"]')
barra_senha.send_keys('Minhasenha12*')

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

sleep(2)


novo_chamado = driver.find_element(By.XPATH, '//*[@id="content"]/nav/ol/li[1]')
novo_chamado.click()

sleep(2.5)

aba_cimatecPark = driver.find_element(By.XPATH, '//*[@id="senaicimatecpark-tab"]')
aba_cimatecPark.click()
sleep(1)

aba_eventoCerimonial = driver.find_element(By.XPATH, '//*[@id="senaicimatecpark"]/div/div/p[1]/a')
aba_eventoCerimonial.click()

sleep(2)
#----------------------------------------------------------------------


select_element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, 'AREA_ID_SELECT'))
)
driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_element)


driver.execute_script("""
    const select = document.getElementById('AREA_ID_SELECT');
    select.value = '10162';
    select.dispatchEvent(new Event('change'));
""")

sleep(2)




script = """
var select = document.getElementById('CLASS_ID');
select.value = 11017;
select.dispatchEvent(new Event('change'));
"""
driver.execute_script(script)

sleep(4)



script2 = """
var select = document.getElementById('PROD_ID');
select.value = 12459;
select.dispatchEvent(new Event('change'));
"""
driver.execute_script(script2)

sleep(3)


descricao = driver.find_element(By.XPATH , '//*[@id="OCOR_DESCRICAO"]')
descricao.send_keys('Será conduzida uma visita')

sleep(5)