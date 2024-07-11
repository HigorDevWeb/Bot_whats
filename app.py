import openpyxl
from urllib.parse import quote
import webbrowser
import time
import pyautogui

# Open WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
time.sleep(30)  # wait for 30 seconds to load WhatsApp Web

# Load Excel file and worksheet
workbook = openpyxl.load_workbook('bot_whats/clientes.xlsx')
pagina_clientes = workbook['Sheet1']

# Iterate over rows in the worksheet
for linha in pagina_clientes.iter_rows(min_row=2):
    # Extract values from each row
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    # Check if vencimento is not empty
    if vencimento:
        mensagem = f"Olá {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar no link https://www.link_do_pagamento.com"
        link_personalizado = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
        print(f"Sending message to {nome} with link {link_personalizado}")
        
        # Open WhatsApp Web with the personalized link
        webbrowser.open(link_personalizado)
        time.sleep(10)  # wait for 10 seconds to load the WhatsApp Web page
        
        try:
            # Locate the send button on the screen
            seta = pyautogui.locateCenterOnScreen('bot_whats/image.png')
            time.sleep(5)  # wait for 5 seconds to ensure the button is loaded
            pyautogui.click(seta[0], seta[1])  # click the send button
            time.sleep(5)  # wait for 5 seconds to ensure the message is sent
            pyautogui.hotkey("ctrl", "w")  # close the WhatsApp Web tab
            time.sleep(5)  # wait for 5 seconds to ensure the tab is closed
        except Exception as e:
            print(f"Error sending message to {nome}: {e}")
            with open('bot_whats/erros.csv','a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f"{nome},{telefone},{vencimento}\n")
    else:
        print(f"Skipping row for {nome} because vencimento is empty")