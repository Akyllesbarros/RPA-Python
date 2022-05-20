
from time import sleep
from openpyxl import load_workbook
import os, csv, sys
from requests import request
from selenium import webdriver      
from selenium.webdriver.common.by import By
import pyautogui
from autoglass import verificar_elemento, gerar_log
from seleniumrequests import Chrome


# Faz o login no site
try:
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")
    global driver 
    driver = webdriver.Chrome(chrome_options=chrome_options)

    driver.get('http://www.eletropar.net/WebSite/index.php')

    login = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/ul/li[4]/a[1]").click()

    username = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/form/input[1]").send_keys("07571746006992")
    pyautogui.press("tab")
    pyautogui.write("MGVIDR")

    submit = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/form/input[4]").click()
    print("Realizado Login, Site aberto!")
   
    FOLDER = "M:/RPA/PYTHON"

    for nomeArquivo in os.listdir("M:/RPA/COMERCIO/CECOM/PRINCING/ELETROPAR/01.NOVOS"):

        if nomeArquivo.endswith(".xlsx"):
            
            print(f"Abrindo o arquivo {nomeArquivo}")

            wb = load_workbook(filename = "M:/RPA/COMERCIO/CECOM/PRINCING/ELETROPAR/01.NOVOS/" + nomeArquivo)
            print(wb.worksheets)
            
            for a in wb.worksheets:
                sheet_ranges = wb[a.title]
                for linha in sheet_ranges:
                    if linha[0].value != "":
                        codigoProduto = linha[0].value 

                        if codigoProduto != "" and codigoProduto != "." :

                            url = "http://www.eletropar.net/csp/coronel/Site/catalogopesqava.csp?cod=" + str(codigoProduto) + "&ini=&fim=&des=&Token=%A9%A2%8B%AF%A8%9B%96%9C%97%99%A6%89%A6%98%B4%87%BB%B6%C3q%92%B2%B2%C3%97%B0%94"
                            
                            Resultado = driver.get(url)
                            print(" Vamos pesquisar nessa URL: " + url)
                            verificarElemento = verificar_elemento(driver, "xpath", "/html/body/div/div/table/tbody[1]/tr/td[3]")
                            print(verificarElemento)

                            if verificarElemento == True: 
                                codigoFornecedor = driver.find_element(By.XPATH, "/html/body/div/div/table/tbody[1]/tr/td[2]").text
                                descricaoProduto = driver.find_element(By.XPATH, "/html/body/div/div/table/tbody[1]/tr/td[3]").text
                                marcaProduto = driver.find_element(By.XPATH, "/html/body/div/div/table/tbody[1]/tr/td[4]").text
                                precoProduto = driver.find_element(By.XPATH, "/html/body/div/div/table/tbody[1]/tr/td[5]").text

                                print(f"\nEscreve: {codigoProduto} 👉 {codigoFornecedor}👉 {descricaoProduto}👉 {marcaProduto}👉 {precoProduto}\n")

                                vCsv = open('M:/RPA/COMERCIO/CECOM/PRINCING/ELETROPAR/02.RESULTADO/resultado_'+ nomeArquivo.replace(".xlsx", "") + '.csv', 'a+', newline='', encoding='utf-8')
                                csv.writer(vCsv, delimiter=';').writerow([codigoProduto, codigoFornecedor, descricaoProduto, marcaProduto, precoProduto])
                                vCsv.close()

                            else:
                                print(f"\nEscreve: {codigoProduto} 👉 NULL 👉 NULL 👉 NULL 👉 NULL\n")
                                vCsv = open('M:/RPA/COMERCIO/CECOM/PRINCING/ELETROPAR/02.RESULTADO/resultado_'+ nomeArquivo.replace(".xlsx", "") + '.csv', 'a+', newline='', encoding='utf-8')
                                csv.writer(vCsv, delimiter=';').writerow([codigoProduto,"NULL","NULL","NULL","NULL"])
                                vCsv.close()                        
except Exception as e:

#COLETA DA LINHA DE ERRO E MENSAGEM DE ERRO
    exc_type, exc_value, exc_traceback = sys.exc_info()
    vLinhaErro = str(exc_traceback.tb_lineno)
    vErro = str(e)
    
    gerar_log("Eletropar",'Iniciar',False,vErro,vLinhaErro)

