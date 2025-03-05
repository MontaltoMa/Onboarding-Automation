# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

#Import for Pandas
import pandas as pd

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

    bot.download_folder_path = r"C:\Users\marcos.martins\Downloads"

    # Opens the BotCity website.
    bot.browse("https://pathfinder.automationanywhere.com/challenges/automationanywherelabs-customeronboarding.html")

    bot.maximize_window()

    # Implement here your logic...
    ...

    #Localiza o botão de Aceitando os cookies
    bot.find_element("//button[text()='Aceitar cookies']", By.XPATH).click()
    print("Aceitando os cookies")

    #Localiza o botão de Community login
    bot.find_element("//button[text()='Community login']", By.XPATH).click()
    print("Community login")

    #Localiza o botão de Community login
    bot.find_element("//input[@type='text']", By.XPATH).send_keys(var_strEmail)
    print("Email")

    bot.enter()

    #Inserindo a senha
    bot.find_element("//input[@type='password']", By.XPATH).send_keys(var_strSenha)
    print("Senha")

    bot.enter()

    #Localiza o botão de Download
    bot.find_element("//a[text()='Download CSV']", By.XPATH).click()
    print("Feito o download do Excel")

    #Caminho final do Arquivo     
    var_strCaminhoArquivo = r"C:\Users\marcos.martins\Downloads\customer-onboarding-challenge.csv"
    print("Transferido para a pasta de downloads")

    var_intTempoEspera = 60

    #Espera o donwload ser finalizado
    bot.wait_for_downloads(timeout= var_intTempoEspera)

    #Volta para a página do RPAChallenge
    bot.back()

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

        bot.find_element("//input[@name='customerName']", By.XPATH).send_keys(var_strCompanyName)
        print("Preenchendo Customer Name")

        bot.find_element("//input[@name='customerID ']", By.XPATH).send_keys(var_strCustomerId)
        print("Preenchendo Customer ID")

        bot.find_element("//input[@name='contact ']", By.XPATH).send_keys(var_strPrimaryContact)
        print("Preenchendo Primary Contact")

        bot.find_element("//input[@name='street ']", By.XPATH).send_keys(var_strStreetAddress)
        print("Preenchendo Street Address")

        bot.find_element("//input[@name='city ']", By.XPATH).send_keys(var_strCity)
        print("Preenchendo City")

        bot.find_element("//select[@name='state']", By.XPATH).send_keys(var_strState)
        print("Selecionando o State")

        bot.find_element("//input[@name='zip ']", By.XPATH).send_keys(var_strZip)
        print("Preenchendo Zip")

        bot.find_element("//input[@name='email']", By.XPATH).send_keys(var_strEmailAddress)
        print("Preenchendo Email Address")

        #Verifica o valor da variável var_strOffersDiscounts e confirma se vai ter desconto ou não
        if var_strOffersDiscounts == 'YES':
            bot.find_element("//input[@value='option1']", By.XPATH).click()
            print("Selecionado Yes em Desconto")
        else:
            bot.find_element("//input[@value='option2']", By.XPATH).click()
            print("Selecionado No em Desconto")

        #Verifica o valor da variável var_strNonDisclosure e confirma se concorda com o termo de confidencialidade
        if var_strNonDisclosure == 'YES':
            bot.find_element("//input[@type='checkbox']", By.XPATH).click()
            print("Selecionado Acordo de não divulgação")

        bot.find_element("//button[text()='Register']", By.XPATH).click()
        print(f"Enviando {index} linha formulario preenchido")

    
    # Wait 5 seconds before closing
    bot.wait(5000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

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
