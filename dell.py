from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep


def main_dell(serial_number):
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)
    navegador.maximize_window()

    # DELL
    navegador.get('https://www.dell.com/support/home/pt-br')
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="mh-search-input"]')))
    navegador.find_element(By.XPATH, '//*[@id="mh-search-input"]').send_keys(serial_number)

    try:
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="unified-masthead"]/div[1]/div[1]/div[2]/div/button[2]')))
        sleep(1)
        navegador.find_element(By.XPATH, '//*[@id="unified-masthead"]/div[1]/div[1]/div[2]/div/button[2]').click()
    except:
        sleep(3)
        navegador.find_element(By.XPATH, '//*[@id="unified-masthead"]/div[1]/div[1]/div[2]/div/button[2]').click()
    try:
        WebDriverWait(navegador, 5).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="null-result-text"]')))
        return navegador.find_element(By.XPATH, '//*[@id="null-result-text"]').text
    except:
        pass

    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="viewDetailsWarranty"]')))
    navegador.find_element(By.XPATH, '//*[@id="viewDetailsWarranty"]').click()
    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="dsk-purchaseDt"]')))
    data_inicio = navegador.find_element(By.XPATH, '//*[@id="dsk-purchaseDt"]').text

    WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="dsk-expirationDt"]')))
    data_vencimento = navegador.find_element(By.XPATH, '//*[@id="dsk-expirationDt"]').text
    return data_inicio, data_vencimento


if __name__ == '__main__':
    print(main_dell('RJ-048-NTK-LAT'))
