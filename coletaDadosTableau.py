from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui, csv
from os.path import exists


# Chamando e abrindo driver
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")
chrome_options.add_experimental_option("prefs", {
  "download.default_directory": r"\\autoglass.com.br\Compartilhado$\RPA\PYTHON\COMERCIO\CECOM\ColetaDadosTableau\assets"
  
})
global driver 
caminhoCsv = "M:/RPA/PYTHON/COMERCIO/CECOM/ColetaDadosTableau/assets/Painelcoletadepreos.csv"
driver = webdriver.Chrome(chrome_options=chrome_options)
driver.get('https://tableau.autoglass.com.br/#/views/ProdutoFornecedorNF/Painelcoletadepreos.csv')
sleep(2)

# Realiza Login

username = driver.find_element(By.XPATH, '//*[@id="ng-app"]/div/div[1]/span/div[2]/span/div/div[2]/span[1]/form/div[1]/div[1]/div/div/input').send_keys("sa.32uyfeyy")
pyautogui.press("tab")
senha = driver.find_element(By.XPATH, '//*[@id="ng-app"]/div/div[1]/span/div[2]/span/div/div[2]/span[1]/form/div[1]/div[2]/div/div/input').send_keys("t4(HCOneXtcON")
submit = driver.find_element(By.XPATH, "//*[@id='ng-app']/div/div[1]/span/div[2]/span/div/div[2]/span[1]/form/button").click()

# Verifica se o arquivo existe na pasta
bVerificacaoArquivo = exists(caminhoCsv)
while bVerificacaoArquivo == False:
    bVerificacaoArquivo = exists(caminhoCsv)
    sleep(1)

if bVerificacaoArquivo == True :
    with open(caminhoCsv, newline='', encoding="utf-8") as csvfile:
        vCsv = csv.reader(csvfile, delimiter=';')
        for linha in vCsv:
            tipoProduto = linha[1]
            # codigoCHG = linha[2].replace('"', '').replace(" ", "")
            # codigoDPK = linha[3]
            # codigoEletropar = linha[4]
            # codigoJocar = linha[5]
            # codigoPassarella = linha[6]
            # codigoPellegrino = linha[7]
            # codigoProduto = linha[8]
            # codigoPython = linha[9]
            # codigoReiDoFFarol = linha[10]
            # descricaoProduto = linha[11]
            # codigoCilia = linha[14]

            
            # Separa dados por tipo de produto

            if "RETROVISOR" in tipoProduto:
                tipoProduto = "retrovisor"
            if "FAROL" in tipoProduto:
                tipoProduto = "iluminacao"
            if "PARABRISA" in tipoProduto:
                tipoProduto = "parabrisa"
            if "VIGIA" in tipoProduto:
                tipoProduto = "vigia"
            if "ACESSÓRIOS" in tipoProduto:
                tipoProduto = "acessorios"
            if "ARREFECIMENTO" in tipoProduto: 
                tipoProduto = "arrefecimento"
            if "LATARIA" in tipoProduto:
                tipoProduto = "lataria"
            if "DIREÇÃO" in tipoProduto:
                tipoProduto = "direcao"
            if "LATERAIS" in tipoProduto:
                tipoProduto = "laterais"
            if "PARACHOQUE" in tipoProduto:
                tipoProduto = "parachoque"
            if "SUSPENSÃO" in tipoProduto:
                tipoProduto = "suspensao"
            
            # Lista com concorrente, caminho e valor dos itens
            lista = [
                [
                    "CHG",
                    "M:/RPA/COMERCIO/CECOM/PRINCING/CHG/01.NOVOS/",
                    linha[2],
                    ".xlsx"
                ],
                [
                    "CILIA",
                    "M:/RPA/COMERCIO/CECOM/PRINCING/INFO PRODUTOS/01.NOVOS/",
                    linha[14],
                    ".csv"
                ],
                [
                    "DPK",
                    "M:/RPA/COMERCIO/CECOM/PRINCING/DPK(KDpeca)/01.NOVOS/",
                    linha[3],
                    ".xlsx"
                ],
                [
                    "ELETROPAR",
                    "M:/RPA/COMERCIO/CECOM/PRINCING/ELETROPAR/01.NOVOS/",
                    linha[4],
                    ".xlsx"
                ],
                [
                    "PELLEGRINO",
                    "M:/RPA/COMERCIO/CECOM/PRINCING/PELLEGRINO/01.NOVOS/",
                    linha[7],
                    ".xlsx"
                ],
                [
                    "PASSARELLA",
                    "M:/RPA/COMERCIO/CECOM/PRINCING/PASSARELLA/01.NOVOS/",
                    linha[6],
                    ".xlsx"
                ],
                [
                    "JOCAR",
                    "M:/RPA/COMERCIO/CECOM/PRINCING/INFO PRODUTOS/01.NOVOS/",
                    linha[5],
                    ".csv"
                ],
                [
                    "PYTHON",
                    "M:/RPA/COMERCIO/CECOM/PRINCING/PYTHON/01.NOVOS/",
                    linha[9],
                    ".xlsx"
                ],
                [
                    "REIDOFAROL",
                    "M:/RPA/COMERCIO/CECOM/PRINCING/INFO PRODUTOS/01.NOVOS/",
                    linha[10],
                    ".csv"
                ]
            ]
            
            # Escreve no CSV em cada pasta
            for item in lista:
                if ("NÃO" not in item[2]) and ( "Cód" not in item[2]):
                    codigoProd = item[2].replace('"', '').replace(" ", "")
                    nTamanho = len(codigoProd)
                    if nTamanho > 4:
                        print('ESCREVENDO: ' + codigoProd +' no arquivo' + item[1])
                        vCsv = open(item[1] + tipoProduto + '_'+ item[0] + item[3], 'a+', newline='', encoding='utf-8')
                        csv.writer(vCsv, delimiter=';').writerow([codigoProd])
                        vCsv.close()

            






                


                




