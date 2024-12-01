# WEB-SCRIPT
from selenium.webdriver.common.by import By
import pandas as pd
from selenium import webdriver
#from selenium.webdriver.common.keys import Keys
from time import sleep
from datetime import date
from openpyxl import load_workbook
from openpyxl import Workbook
import schedule
def encontrar_proxima_linha_vazia_na_coluna(planilha, coluna):# criei uma função para rodar automação 
    max_row = planilha.max_row
    for row in range(1, max_row + 1):
        cell_value = planilha[f'{coluna}{row}'].value
        if cell_value is None or cell_value == '':
            return row
    return max_row + 1
def formatar_valor_com_virgula(valor):
    # Converte o valor para string e substitui o ponto decimal por vírgula
    valor_str = f'{valor:.2f}'  # Formata o número para duas casas decimais
    return valor_str.replace('.', ',')

data_atual = date.today()
dia_atual = date.today().day
data_tabatinga = date.today().day
data_atual1 = date.today()

def automatizacao():
    url = 'https://proamanaus.com.br/servicos-proa/nivel-dos-rios'
    sleep(5)
    # Configuração do navegador
    #options = webdriver.ChromeOptions()
    #options.add_argument('--incognito')
    #options.add_argument('--headless')  # Para executar o navegador em segundo plano (opcional)
    driver = webdriver.Chrome()#driver_path, options=options)
    driver.get(url)
    tag = driver.find_element(By.XPATH,f'/html/body/main/div/article/div/table/tbody/tr[{dia_atual+1}]/td[2]')
    tag2 = driver.find_element(By.XPATH,f'/html/body/main/div/article/div/table/tbody/tr[{dia_atual+1}]/td[3]')
    tag3 = driver.find_element(By.XPATH,f'/html/body/main/div/article/div/table/tbody/tr[{dia_atual+1}]/td[5]')
    tag4 = driver.find_element(By.XPATH,f'/html/body/main/div/article/div/table/tbody/tr[{dia_atual+1}]/td[7]')
    texto = (tag.text)
    texto2 = (tag2.text)
    texto3 = (tag3.text)
    texto4 = (tag4.text)
    print(texto)
    print(texto2)
    print(texto3)
    print(texto4)
   # Fechar o navegado  
    #https://bemol.sharepoint.com/sites/TimedeRiscosSA/Documentos%20Compartilhados/Riscos%20e%20Controles%20Internos/12.%20Comit%C3%AA%20Seca/Dados%20-%20Dashboard%20dos%20Rios/Base_de_dados_N%C3%ADveis_do_rio.xlsx?web=1
    arquivo_excel = r'C:\Users\16205\OneDrive - BEMOL S A\Dados - Dashboard dos Rios\Base_de_dados_Níveis_do_rio.xlsx'
    data_formatada = data_atual.strftime('%d/%m/%Y')
    wb = load_workbook(arquivo_excel)
    planilha = wb.active
    
    valor_formatado = formatar_valor_com_virgula(float(texto))
    proxima_linha = planilha.max_row + 1 
    planilha[f'A{proxima_linha}'] = data_formatada
    planilha[f'B{proxima_linha}'] = valor_formatado
    planilha[f'C{proxima_linha}'] = 'manaus'   

    valor_formatado2 = formatar_valor_com_virgula(float(texto2))
    proxima_linha += 1
    planilha[f'A{proxima_linha}'] = data_formatada
    planilha[f'B{proxima_linha}'] = valor_formatado2
    planilha[f'C{proxima_linha}'] = 'Itacoatiara'

    valor_formatado3 = formatar_valor_com_virgula(float(texto3))
    proxima_linha += 1
    planilha[f'A{proxima_linha}'] = data_formatada
    planilha[f'B{proxima_linha}'] = valor_formatado3
    planilha[f'C{proxima_linha}'] = 'Tabatinga'

    valor_formatado4 = formatar_valor_com_virgula(float(texto4))
    proxima_linha += 1
    planilha[f'A{proxima_linha}'] = data_formatada
    planilha[f'B{proxima_linha}'] = valor_formatado4
    planilha[f'C{proxima_linha}'] = 'Coari'
    wb.save(arquivo_excel)
    wb.close()
    driver.quit()
automatizacao()
def agendar_automacao():
    # Horário desejado para executar a automação (por exemplo, às 8:00 todos os dias)
    horario_execucao = "08:00"

    # Agendar a execução da automação diariamente no horário especificado
    schedule.every().day.at(horario_execucao).do(automatizacao)

    # Loop para manter o programa em execução para que o agendamento funcione
    while True:
        schedule.run_pending()
        sleep(60)  # Verifica a cada minuto se há tarefas agendadas para serem executadas

# Iniciar o agendamento da automação
agendar_automacao()

