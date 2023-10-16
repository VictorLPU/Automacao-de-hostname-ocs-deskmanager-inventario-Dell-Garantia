from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from selenium.webdriver.chrome.options import Options

def main_ocs(elemento_escolhido):
  
    servico = Service(ChromeDriverManager().install())

    navegador = webdriver.Chrome(service=servico)
    navegador.maximize_window()
    navegador.get('Link do OCS') #Colocar o link que leva direto ao OCS
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="LOGIN"]')))
    navegador.find_element(By.XPATH, '//*[@id="LOGIN"]').send_keys('LOGIN') #"LOGIN" onde deve ir o user de Login do OCS

    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="PASSWD"]')))
    navegador.find_element(By.XPATH, '//*[@id="PASSWD"]').send_keys('SENHA') #"SENHA" onde deve ir a senha de Login do OCS

    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="btn-logon"]')))
    navegador.find_element(By.XPATH, '//*[@id="btn-logon"]').click()

    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ocs-navbar"]/ul[1]/li[1]/a')))
    navegador.find_element(By.XPATH, '//*[@id="ocs-navbar"]/ul[1]/li[1]/a').click()

    try:
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="list_show_all_filter"]/label/input')))
        navegador.find_element(By.XPATH, '//*[@id="list_show_all_filter"]/label/input').send_keys(elemento_escolhido)
    except:
        sleep(3)
        WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="list_show_all_filter"]/label/input')))
        navegador.find_element(By.XPATH, '//*[@id="list_show_all_filter"]/label/input').send_keys(elemento_escolhido)
    sleep(3)
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="list_show_all_wrapper"]/div/div[3]/div[2]')))
    elementos = navegador.find_element(By.XPATH, f'//*[@id="list_show_all_wrapper"]/div/div[3]/div[2]')
    try:
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/form[1]/div[2]/div/div/div[3]/div[2]/table/tbody/tr/td[4]/a')))
        navegador.find_element(By.XPATH, '/html/body/div/form[1]/div[2]/div/div/div[3]/div[2]/table/tbody/tr/td[4]/a').click()
    except:
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="list_show_all"]/tbody/tr/td')))
        serial_number = navegador.find_element(By.XPATH, '//*[@id="list_show_all"]/tbody/tr/td').text
        return serial_number
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[1]/ul/li[2]/a')))
    navegador.find_element(By.XPATH, '/html/body/div/div[1]/ul/li[2]/a').click()

    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="affich_bios"]/tbody/tr/td[1]')))
    serial_number = navegador.find_element(By.XPATH, '//*[@id="affich_bios"]/tbody/tr/td[1]').text
    return serial_number


if __name__ == '__main__':
    print(main_ocs('RJ-118-NTK-LAT'))
