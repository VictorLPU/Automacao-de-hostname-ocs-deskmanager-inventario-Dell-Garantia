from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from selenium.webdriver.chrome.options import Options


def main_desk(hostname, data_inicio):

    servico = Service(ChromeDriverManager().install())

    navegador = webdriver.Chrome(service=servico)
    navegador.maximize_window()
    navegador.get('Link do Deskmanager') #Inserir o Link do Deskmanager "Nosso inventario"
    sleep(7)
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ModalDirecionaPrefixo"]/div/div/div[2]/form/div[3]/button')))
    navegador.find_element(By.XPATH, '//*[@id="ModalDirecionaPrefixo"]/div/div/div[2]/form/div[3]/button').click()
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ModalLoginTicket"]/div/div/div[2]/form/div[1]/input')))
    navegador.find_element(By.XPATH, '//*[@id="ModalLoginTicket"]/div/div/div[2]/form/div[1]/input').send_keys('USUARIO') #Colocar o Usuario que acessa o Deskmanager
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ModalLoginTicket"]/div/div/div[2]/form/div[2]/input')))
    navegador.find_element(By.XPATH, '//*[@id="ModalLoginTicket"]/div/div/div[2]/form/div[2]/input').send_keys('SENHA') #Colocar a senha que acessa o DeskManager
    sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="ModalLoginTicket"]/div/div/div[2]/form/div[3]/button').click()
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="side-menu"]/li[2]/a')))
    navegador.get('Link do Deskmanager') #Inserir o Link do Deskmanager "Nosso inventario"
    sleep(1)
    navegador.get('Link do Deskmanager') #Inserir o Link do Deskmanager "Nosso inventario"
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="PanelICs"]/div[2]/div[2]/div/div[1]/div/div/div/input')))
    navegador.find_element(By.XPATH, '//*[@id="PanelICs"]/div[2]/div[2]/div/div[1]/div/div/div/input').send_keys(hostname)
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="PanelICs"]/div[2]/div[2]/div/div[1]/div/div/div/div/button[1]')))
    navegador.find_element(By.XPATH, '//*[@id="PanelICs"]/div[2]/div[2]/div/div[1]/div/div/div/div/button[1]').click()
    sleep(3)
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="PanelICs"]/div[2]/div[2]/div/div[2]')))
    tabela = navegador.find_elements(By.XPATH, '//*[@id="PanelICs"]/div[2]/div[2]/div/div[2]')
    for elemento in tabela:
        texto = elemento.text

    texto = str(texto).split('\n')
    texto = ' '.join(texto)
    texto.split(' ')

    sleep(3)

    if not 'DESKTOP' in texto:
        return 'Ooooops... Nenhum Registro encontrado...'
    for elemento in tabela:
        elemento.click()


    sleep(4)
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="heading1ICs"]/h4/a')))
    navegador.find_element(By.XPATH, '//*[@id="heading1ICs"]/h4/a').click()
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="headingDatas"]/h4/a')))
    navegador.find_element(By.XPATH, '//*[@id="headingDatas"]/h4/a').click()

    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="collapseDatas"]/div/div[1]/div[2]/div/input')))
    navegador.find_element(By.XPATH, '//*[@id="collapseDatas"]/div/div[1]/div[2]/div/input').send_keys(data_inicio)
    WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ModalICs"]/div/div/div[3]/div/div/div/button[2]')))
    navegador.find_element(By.XPATH, '//*[@id="ModalICs"]/div/div/div[3]/div/div/div/button[2]').click()
    sleep(5)
    return 'Ok'
    # Campo Data aquisição - //*[@id="collapseDatas"]/div/div[1]/div[2]/div/input


if __name__ == '__main__':
    print(main_desk('RJ-020-NTK-LAT', '01-01-2020'))
