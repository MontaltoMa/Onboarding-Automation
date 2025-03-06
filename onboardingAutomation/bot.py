# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

#Import for Pandas
import pandas as pd

#Import for Clicknium
from clicknium import clicknium as cc, locator, ui

#Import for Os
import os

#Import Sleep
from time import sleep 

#Import for Select
from selenium.webdriver.support.ui import Select

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False


def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    # bot.browser = Browser.FIREFOX

    var_strEmail = 'marcos.martins@t2cgroup.com.br'
    var_strSenha = '@12041999Ma'

    # Uncomment to set the WebDriver path
    bot.driver_path = r"C:\BotCity\chromedriver.exe"
    # bot.driver_path = "<path to your WebDriver binary>"

    # Opens the BotCity website.
    tab =cc.chrome.open("https://pathfinder.automationanywhere.com/challenges/automationanywherelabs-customeronboarding.html")

    # Implement here your logic...

    #Verificando se já estou logado
    if tab.wait_appear_by_xpath(xpath="//button[text()='Community login']" ,wait_timeout=10):
        print("Botão de Community login apareceu. Realizando login...")

        #Localiza o botão de Community login e clica
        print("Community login")
        tab.find_element_by_xpath("//button[text()='Community login']").click()

        #Inserindo E-mail
        print("Email")
        tab.find_element_by_xpath("//input[@type='text']").set_text(var_strEmail)

        #Clicando no botão next
        print("Next")
        tab.find_element_by_xpath("//button[text()='Next']").click()

        #Inserindo a senha
        print("Senha")
        tab.find_element_by_xpath("//input[@type='password']").set_text(var_strSenha)

        #Clicando no botão de Log In
        print("Log In")
        tab.find_element_by_xpath("//button[text()='Log in']").click()
        
    else:
        print("Botão de Community login não apareceu. Pulando login...")

    #Caminho final do Arquivo     
    var_strCaminhoArquivo = r"C:\Users\marcos.martins\Downloads\customer-onboarding-challenge.csv"

    #Verificando se já existe o arquivo baixado, se sim excluindo o mesmo
    if os.path.exists(var_strCaminhoArquivo):
        os.remove(var_strCaminhoArquivo)
        print(f"Arquivo existente excluído: {var_strCaminhoArquivo}")
    else:
        print("Nenhum arquivo anterior encontrado!")

    #Localiza o botão de Download
    tab.find_element_by_xpath("//a[text()='Download CSV']").click()
    print("Fazendo o download do Excel")

    sleep(5)

    #Lê o arquivo excel
    var_dfClientes = pd.read_csv(var_strCaminhoArquivo)

    """
    For para percorrer todas as linhas preenchidas do excel!
    Localiza o XPath de cada elemento a ser preenchido
    E escreve ele em cada espaço especifico
    """
    print("Iniciado Loop de preenchimento")
    for index, linha in var_dfClientes.iterrows():
        var_strCompanyName = linha['Company Name']
        var_strCustomerId = str(linha['Customer ID'])
        var_strPrimaryContact = linha['Primary Contact']
        var_strStreetAddress = linha['Street Address']
        var_strCity = linha['City']
        var_strState = linha['State']
        var_strZip = str(linha['Zip'])
        var_strEmailAddress = linha['Email Address']
        var_strOffersDiscounts = linha['Offers Discounts']
        var_strNonDisclosure = linha['Non-Disclosure On File']

        print(f"Preenchendo Customer Name: {var_strCompanyName}")
        tab.find_element_by_xpath("//input[@name='customerName']").set_text(var_strCompanyName)

        print(f"Preenchendo Customer ID: {var_strCustomerId}")
        tab.find_element_by_xpath("//input[@name='customerID ']").set_text(var_strCustomerId)

        print(f"Preenchendo Primary Contact {var_strPrimaryContact}")
        tab.find_element_by_xpath("//input[@name='contact ']").set_text(var_strPrimaryContact)

        print(f"Preenchendo Street Address: {var_strStreetAddress}")
        tab.find_element_by_xpath("//input[@name='street ']").set_text(var_strStreetAddress)

        print("Preenchendo City")
        tab.find_element_by_xpath("//input[@name='city ']").set_text(var_strCity)

        print("Selecionando o State")
        tab.find_element_by_xpath("//select[@name='state']").select_item(var_strState)

        print(f"Preenchendo Zip: {var_strZip}")
        tab.find_element_by_xpath("//input[@name='zip ']").set_text(var_strZip)

        print(f"Preenchendo Email Address {var_strEmailAddress}")
        tab.find_element_by_xpath("//input[@name='email']").set_text(var_strEmailAddress)

        #Verifica o valor da variável var_strOffersDiscounts e confirma se vai ter desconto ou não
        if var_strOffersDiscounts == 'YES':
            print("Selecionado Yes em Desconto")
            tab.find_element_by_xpath("//input[@value='option1']").click()
        else:
            tab.find_element_by_xpath("//input[@value='option2']").click()
            print("Selecionado No em Desconto")

        #Verifica o valor da variável var_strNonDisclosure e confirma se concorda com o termo de confidencialidade
        if var_strNonDisclosure == 'YES':
            print("Selecionado Acordo de não divulgação")
            tab.find_element_by_xpath("//input[@type='checkbox']").click()

        print(f"Enviando {index+1} linha(s) formulario preenchido")
        tab.find_element_by_xpath("//button[text()='Register']").click()

    
    sleep(5)

    tab.close()

    # Uncomment to mark this task as finished on BotMaestro
    maestro.finish_task(
        task_id=execution.task_id,
        status=AutomationTaskFinishStatus.SUCCESS,
        message="Task Finished OK.",
        total_items=0,
        processed_items=0,
        failed_items=0
    )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
