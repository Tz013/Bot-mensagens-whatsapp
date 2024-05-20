# Biblotecas
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 
#import sys

# Abrindo navegador
webbrowser.open('https://web.whatsapp.com/')
sleep(10)

# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('chamados.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    chamado = linha[2].value
    dt_abertura = linha[3].value
    status = linha[4].value
    
    mensagem = f'Olá {nome}, tudo bem? Sou Matheus da TI da Captamed Santos, estou com o seu chamado N°{chamado} em aberto, em que posso lhe ajudar? '
    if nome is None or nome == "":
        print('Celula nome vazia, programa finalizado.')
        exit()
    if status.lower() == "não iniciado":
        # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
        # com base nos dados da planilha
        try:
            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_mensagem_whatsapp)
            sleep(30)
            seta = pyautogui.locateCenterOnScreen('seta.png')
            sleep(5)
            pyautogui.click(seta[0],seta[1])
            sleep(5)
            pyautogui.hotkey('ctrl','w')
            sleep(5)
        except:
            print(f'Não foi possível enviar mensagem para {nome}')
            with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone}{os.linesep}')
                pyautogui.hotkey('ctrl','w')
    else:
        print(f'O atendimento do(a) {nome} já está {status}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
            pyautogui.hotkey('ctrl','w')