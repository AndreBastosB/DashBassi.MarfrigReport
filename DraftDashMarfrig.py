from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from datetime import date
import clipboard
import gspread
import sys
from oauth2client.service_account import ServiceAccountCredentials
from SenhasDraftDashMarfrig import senhas
import time

#CRIADO POR DECOGOL

planilha = "Vendas Bassi.Marfrig"
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

credenciais = (r"C:\Users\Take4\Desktop\creds.json")

creds = ServiceAccountCredentials.from_json_keyfile_name(credenciais, scope)

client = gspread.authorize(creds)

sheet = client.open(planilha).worksheet("Transações")

data = (date.today())

data2 = str(data)
ano = data2.split('-', 1)[0]

mes1 = data2.split('-', 1)[1]
mes = mes1.split('-', 1)[0]

dia1 = data2.split('-', 1)[1]
dia = dia1.split('-', 1)[1]

dataFinal = dia + '/' + mes + '/' + ano

print (dataFinal)

opt = webdriver.ChromeOptions() 
# opt.add_argument("--auto-open-devtools-for-tabs")
opt.add_experimental_option("excludeSwitches", ["enable-automation"])
opt.add_experimental_option('useAutomationExtension', False)
opt.add_argument("--start-maximized")

login = senhas[0]
senha = senhas[1]

loginZP = senhas[2]
senhaZP = senhas[3]

ser = Service(r"C:\Users\Take4\Documents\Saves\chromedriver_win32\chromedriver.exe")

driver = webdriver.Chrome(service=ser, options=opt)
driver2 = webdriver.Chrome(service=ser, options=opt)

driver2.get('https://dashboard.zoop.com.br/login')

try:
    WebDriverWait(driver2, 160).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/login-screen/div/section/div/form/div/div[2]/div/div/fieldset/div[1]/input"))).send_keys(loginZP)
    WebDriverWait(driver2, 160).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/login-screen/div/section/div/form/div/div[2]/div/div/fieldset/div[2]/input"))).send_keys(senhaZP)
    WebDriverWait(driver2, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/section/ui-view/login-screen/div/section/div/form/div/div[2]/div/div/button'))).click()
    WebDriverWait(driver2, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div/div[3]/button'))).click()

except:
    sys.exit("Erro na autenticação da Zoop")

driver.get ("https://dashboard.marfrig.app/")

def transformToFloatFromZoop(X):
    L11 = X.split(' ', 1)[1]
    L11 = L11.replace(',','.')
    L11 = float(L11)
    return (L11)

def readingAttributes(xpath):
    try:
        item = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, xpath))).text
    except:
        item = "Erro"
    return (item)

def readingAttributes2(xpath):
    try:
        item = WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, xpath))).text
    except:
        item = "Erro"
    return (item)

def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list)+1)

linha = next_available_row(sheet)

#ERRO AO AUTENTICAR NO DASH - ALTERAÇÃO NO XPATH DOS INPUTS E DO BOTAO DE LOGAR.

# checkSair = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/ul/li[7]/a/div/span'))).text
# if checkSair == 'Sair':
#     WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/ul/li[7]/a/div/span'))).click()
#     time.sleep(5)
#     WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/div/main/div/div/div/div/div[2]/div/div[2]/div/form/div[1]/input"))).send_keys(login)
#     WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/div/main/div/div/div/div/div[2]/div/div[2]/div/form/div[2]/input"))).send_keys(senha)
#     WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div/div/main/div/div/div/div/div[2]/div/div[2]/div/form/div[3]/button'))).click()
# else:
#     pass

try: 
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/div/main/div/div/div/div/div[2]/div/div/form/div[1]/input"))).send_keys(login)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div/div/div/main/div/div/div/div/div[2]/div/div/form/div[2]/input"))).send_keys(senha)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div/div/main/div/div/div/div/div[2]/div/div/form/div[3]/button'))).click()    
except:
    sys.exit("Erro na autenticação do dashboard.")

time.sleep(10)
driver.get ("https://dashboard.marfrig.app/orders?iop=1&aop=1&limit=100")

index = 1

checkDate2 = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[' + str(index) + ']/td[4]/time'))).text    
checkDate = checkDate2.split (' ', 1)[0]

while checkDate != dataFinal:

    if index == 85:
        driver.close()
        driver2.close()
        sys.exit("Sem pedidos no dia de hoje.")

    checkDate2 = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[' + str(index) + ']/td[4]/time'))).text    
    checkDate = checkDate2.split (' ', 1)[0]
    index += 1

if index == 1:
    pass
else:
    index -= 1
    

checkDate2 = WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[' + str(index) + ']/td[4]/time'))).text    
checkDate = checkDate2.split (' ', 1)[0]

while dataFinal in checkDate2:
    
    checkDate2 = WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr['+ str(index) +']/td[4]/time'))).text    
    checkDate = checkDate2.split (' ', 1)[0]
    checkStatus = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr['+ str(index) +']/td[2]'))).text
    checkNumber = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr['+ str(index) +']/td[1]'))).text    
    quantidadeProdutos = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr['+str(index)+']/td[3]')
    
    if checkStatus == 'Produto(s) entregue' and checkDate == dataFinal:
        
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr['+str(index)+']/td[8]/div/button'))).click()
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[' + str(index) + ']/td[8]/div/div/button/span'))).click()
        
        dataEntrega2 = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[9]/div/div[2]/ul/li[1]/div[2]/span[2]/time') 
        dataEntrega3 = str(dataEntrega2)
        dataEntrega = dataEntrega3.split(' ', 1)[0]
        print ('Data da Entrega: ' + str(dataEntrega))
        
        IdPedido2 = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[1]/div[1]/div[1]/div')
        IdPedido3 = str(IdPedido2)
        IdPedido = IdPedido3.split(' ', 1)[1]
        print ('ID do Pedido: ' + str(IdPedido))
        print ('Quantidade de Produtos: ' + str(quantidadeProdutos))
        
        try:
            IdTransacao2 = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td/div/div/button'))).click()       
            IdTransacao = clipboard.paste()
        except: 
            IdTransacao = "Erro"
        
        nomeCliente = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[3]/div/div[2]/table/tbody/tr[3]/td')
        print ('Nome do Cliente: ' + str(nomeCliente))
        
        index2 = 1
        verificaçãoLoja2 = WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[5]/div/div[2]/table/tbody/tr['+str(index2)+']/td[1]'))).text
        
        while ('confirm_checkout') not in verificaçãoLoja2:
            verificaçãoLoja2 = WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[5]/div/div[2]/table/tbody/tr['+str(index2)+']/td[1]'))).text    
            index2 += 1
            
        verificaçãoLoja =  readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[5]/div/div[2]/table/tbody/tr['+str(index2)+']/th')
        print ("Loja: " + str(verificaçãoLoja))
        
        verificaçãoDistancia = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[5]/div/div[2]/table/tbody/tr['+str(index2)+']/td[2]')
        print ('KM Percorrido: ' + str(verificaçãoDistancia))
        
        produtosPrePesagem = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[2]/td')
        print ('Pre Pesagem: ' + str(produtosPrePesagem))
        
        produtosPosPesagem = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[3]/td')
        print ('Pos Pesagem: ' + str(produtosPosPesagem))
        
        checkLinha5 = WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[5]/th'))).text
        
        if 'Desconto' in checkLinha5:
            
            valorDesconto = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[5]/td')
            print ('Desconto: ' + str(valorDesconto))
            
            valorFrete = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[7]/td')
            print ('Frete: ' + str(valorFrete))
            
            valorTaxas = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[8]/td')
            print ('Taxas: ' + str(valorTaxas))
            
            valorAntiFraude = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[9]/td')
            print ('Anti-Fraude: ' + str(valorAntiFraude))
            
            valorTotalDash = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[10]/td')
            print ('Valor total Dash: ' + str(valorTotalDash))
            
        elif 'Frete' in checkLinha5:
            
            valorDesconto = 'R$ 0,00'
            print ('Desconto: ' + str(valorDesconto))
            
            valorFrete = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[5]/td')
            print ('Frete: ' + str(valorFrete))
            
            valorTaxas = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[6]/td')
            print ('Taxas: ' + str(valorTaxas))
            
            valorAntiFraude = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[7]/td')
            print ('Anti-Fraude: ' + str(valorAntiFraude))
            
            valorTotalDash = readingAttributes('/html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/table/tbody/tr[8]/td')
            print ('Valor total Dash: ' + str(valorTotalDash))
        
        else: 
            print ("Este pedido teve algum problema:" + (str(checkNumber)))
                   
        #----------------------------- Puxando dados da Zoop --------------------------
        
        
        index3 = 1
        index4 = 1

        driver2.get('https://dashboard.zoop.com.br/')
        
        WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[1]/div[3]/div[1]/ul/li[2]/div[1]/button'))).click()
        WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[1]/div[3]/div[1]/ul/li[2]/div[1]/ul/li[2]/div/button[5]'))).click()
        
        if IdTransacao == "Erro":
            
            dataTransacao = readingAttributes2('/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div['+str(index4)+']/div/div[1]/div/p[1]')
            valorTransacao = readingAttributes2('/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div['+str(index4)+']/div/div[6]')
            statusTransacao = readingAttributes2('/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div['+str(index4)+']/div/div[2]/div')
                        
            while (valorTotalDash != valorTransacao and statusTransacao != 'Aprovado') and (index4 != 40):
                
                dataTransacao = readingAttributes2('/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div['+str(index4)+']/div/div[1]/div/p[1]')
                valorTransacao = readingAttributes2('/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div['+str(index4)+']/div/div[6]')
                statusTransacao = readingAttributes2('/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div['+str(index4)+']/div/div[2]/div')
                index4 += 1

            if index4 == 1:
                WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div['+str(index4)+']/div/div[4]/div/span'))).click()
                
                valorVenda = readingAttributes2("/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[2]/div/div/div/div[2]/div[1]/div/div[3]/strong")
                print ("Valor da Venda Zoop: " + str(valorVenda))
                
                valorCapturado = readingAttributes2("/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[2]/div/div/div/div[2]/div[3]/div/div[3]/strong")
                print ("Valor da Captura Zoop: " + str(valorCapturado))
                
                IdTransacao2 = readingAttributes2('/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[2]/div/div[3]/div[1]/div[2]/div[4]/div/strong')
                
                try:
                    valorL1 = WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div[1]/div/div[7]/span"))).text
                    valorL2 = WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div[2]/div/div[7]/span"))).text
                    liquidoFinal = transformToFloatFromZoop(valorL1) + transformToFloatFromZoop(valorL2)
                except: 
                    valorL3 = WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div/div/div[7]/span"))).text
                    liquidoFinal = transformToFloatFromZoop(valorL3)
                
                print ('Valor Liquido Zoop: ' + str(liquidoFinal))
                print ('ID Transacao Zoop: ' + str(IdTransacao2))
                
                print ('---------------------')
            
            elif index4 != 40:
                
                index4 -= 1
                WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div['+str(index4)+']/div/div[4]/div/span'))).click()
                
                valorVenda = readingAttributes2("/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[2]/div/div/div/div[2]/div[1]/div/div[3]/strong")
                print ("Valor da Venda Zoop: " + str(valorVenda))
                
                valorCapturado = readingAttributes2("/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[2]/div/div/div/div[2]/div[3]/div/div[3]/strong")
                print ("Valor da Captura Zoop: " + str(valorCapturado))
                
                IdTransacao2 = readingAttributes2('/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[2]/div/div[3]/div[1]/div[2]/div[4]/div/strong')
                
                try:
                    valorL1 = WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div[1]/div/div[7]/span"))).text
                    valorL2 = WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div[2]/div/div[7]/span"))).text
                    liquidoFinal = transformToFloatFromZoop(valorL1) + transformToFloatFromZoop(valorL2)
                except: 
                    valorL3 = WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div/div/div[7]/span"))).text
                    liquidoFinal = transformToFloatFromZoop(valorL3)
                
                print ('Valor Liquido Zoop: ' + str(liquidoFinal))
                print ('ID Transacao Zoop: ' + str(IdTransacao2))
                
                print ('---------------------')
            
            else:
                valorVenda = "Erro"
                print ("Valor da Venda Zoop: " + str(valorVenda))
                
                valorCapturado = "Erro"
                print ("Valor da Captura Zoop: " + str(valorCapturado))
                liquidoFinal = 'Erro'
                IdTransacao2 = 'Erro'
                print ('Valor Liquido Zoop: ' + str(liquidoFinal))
                print ('ID Transacao Zoop: ' + str(IdTransacao2))
                print ('---------------------')
                
                                                                                              
        else:
            checkTranscation = WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div["+str(index3)+"]/div/div[4]/div/span"))).text
            
            while checkTranscation != IdTransacao:
                checkTranscation = WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div["+str(index3)+"]/div/div[4]/div/span"))).text
                index3 += 1
            
            if index3 == 1:
                pass
            else:
                index3 -= 1
            
            WebDriverWait(driver2, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/section/ui-view/ui-view/div/transactions-container/div/div/div[3]/div/div[2]/div/div[1]/div[2]/div['+str(index3)+']/div/div[4]/div/span'))).click()
            
            valorVenda = WebDriverWait(driver2, 30).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[2]/div/div/div/div[2]/div[1]/div/div[3]/strong"))).text
            print ("Valor da Venda Zoop: " + str(valorVenda))
            
            valorCapturado = WebDriverWait(driver2, 30).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[2]/div/div/div/div[2]/div[3]/div/div[3]/strong"))).text
            print ("Valor da Captura Zoop: " + str(valorCapturado))
            
            
            try:
                valorL1 = readingAttributes2("/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div[1]/div/div[7]/span")
                valorL2 = readingAttributes2("/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div[2]/div/div[7]/span")
                liquidoFinal = transformToFloatFromZoop(valorL1) + transformToFloatFromZoop(valorL2)
            except: 
                valorL3 = readingAttributes("/html/body/div[1]/section/ui-view/ui-view/section/details-transaction-container/div/div[2]/section/div/div[3]/div[3]/div[2]/div/div/div/div[2]/div/div/div[7]/span")
                liquidoFinal = transformToFloatFromZoop(valorL3)
            
            print ('Valor Liquido Zoop: ' + str(liquidoFinal))
            print ('ID Transacao Zoop: ' + str(IdTransacao))
            
            print ('---------------------')
        
        data2 = dataFinal[0] + dataFinal[1]
        data3 = dataFinal[3] + dataFinal[4]
        data5 = dataFinal[6]+dataFinal[7]+dataFinal[8]+dataFinal[9] 
        
        dataMesDia = data3 + '/' + data2 + '/' + data5
        
        if IdTransacao == "Erro":
            IdTransacao = IdTransacao2

        sheet.update_cell(linha, 1, dataMesDia)
        sheet.update_cell(linha, 3, IdTransacao)
        sheet.update_cell(linha, 4, IdPedido)
        sheet.update_cell(linha, 5, nomeCliente)
        sheet.update_cell(linha, 6, verificaçãoLoja)
        sheet.update_cell(linha, 7, valorVenda)
        sheet.update_cell(linha, 8, valorCapturado)
        sheet.update_cell(linha, 9, valorCapturado)
        sheet.update_cell(linha, 10, liquidoFinal)
        sheet.update_cell(linha, 11, 'R$ 0,00')
        sheet.update_cell(linha, 12, verificaçãoDistancia)
        sheet.update_cell(linha, 13, quantidadeProdutos)
        sheet.update_cell(linha, 14, produtosPrePesagem)
        sheet.update_cell(linha, 15, produtosPosPesagem)
        sheet.update_cell(linha, 16, valorFrete)
        sheet.update_cell(linha, 17, valorDesconto)
        sheet.update_cell(linha, 18, valorAntiFraude)
        sheet.update_cell(linha, 19, valorTaxas)
        sheet.update_cell(linha, 20, valorTotalDash)
        
        linha = int(linha)
        linha += 1

        driver.get ("https://dashboard.marfrig.app/orders?iop=1&aop=1&limit=100")
    
    else:
        pass
    
    index += 1

driver.close()
driver2.close()

