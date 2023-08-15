from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import openpyxl

planilha = openpyxl.load_workbook("C:\\Users\\rafae\\OneDrive\\Área de Trabalho\\PRG\\python\\challenge.xlsx")
tabela = planilha["Sheet1"]

#Abrindo Edge e navegando até a pagina
driver = webdriver.Chrome()
url = "https://rpachallenge.com/"
driver.get(url)
time.sleep(5)

buttonSubmit = driver.find_element(By.XPATH, '//input[@type="submit"]')
buttonStart = driver.find_element(By.TAG_NAME, 'button')
buttonStart.click()

for row in tabela.iter_rows(min_row=2, max_row=11, values_only=True):
    firstName = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelFirstName"]')
    firstName.send_keys(row[0])
    
    lastName = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelLastName"]')
    lastName.send_keys(row[1])
    
    email = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelEmail"]')
    email.send_keys(row[5])
    
    phone = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelPhone"]')
    phone.send_keys(row[6])

    adress = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelAddress"]')
    adress.send_keys(row[4])
   
    role = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelRole"]')
    role.send_keys(row[3])

    company = driver.find_element(By.XPATH, '//*[@ng-reflect-name="labelCompanyName"]')
    company.send_keys(row[2])

    buttonSubmit.click()

time.sleep(60)