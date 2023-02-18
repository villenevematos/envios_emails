#!/usr/bin/env python
# coding: utf-8
# importar bibliotecas
import pandas as pd
import pyautogui
import pyperclip
import time
import smtplib
import email.message
import mimetypes
from email.message import EmailMessage

# abrir uma nova janela
# colar o link na barra de pesquisa
pyautogui.PAUSE = 1
pyautogui.hotkey('ctrl', 'n')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(16)

# clicar na pasta "Exportar"
pyautogui.doubleClick(442, 280)

# fazer o download da planilha de vendas
for item in [(459, 418), (1203, 157), (943, 636)]:
    pyautogui.click(item)
    time.sleep(0.5)
time.sleep(5)
pyautogui.hotkey('ctrl', 'w')

# importar a base de dados
tabela = pd.read_excel('Vendas - Dez.xlsx')

# faturamento
faturamento = tabela['Valor Final'].sum()

# quantidade de produtos
qtde_produtos = tabela['Quantidade'].sum()

# enviar e-mail
import smtplib
import email.message
import mimetypes
from email.message import EmailMessage

def enviar_email():  
    # lista destinatarios
    destinarios = [('Renan', 'cecilia.chaves@benites.com.br'), ('Juliana', 'ferreira.rodrigo@yahoo.com'), ('Cristiano', 'janaina.soto@pereira.org')]
    
    for nome, email in destinarios:
        corpo_email = f"""
        <p>Prezado {nome}, bom dia</p>
        <br>O faturamento de ontem foi de R${faturamento:,.2f}</br>
        <br>A quantidade de produtos foi de {qtde_produtos:,}</br>
        <br></br>
        <br>Att,</br>
        <br>VILLENEVE VERARDO DE MATOS</br>
        """

        # Destinatario
        msg = EmailMessage()
        msg['Subject'] = "Relatório de Vendas de Ontem"
        msg['From'] = 'ville@gmail.com'
        msg['To'] = email
        
        password = 'qabtmturmviaxpxc' 
        msg.add_header('Content-Type', 'text/html')
        msg.set_content(corpo_email, subtype='html')

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Credenciais de login para enviar e-mail
        s.login(msg['From'], password)
        s.send_message(msg)
        pyautogui.alert(f'Email enviado para {nome}', timeout=2_000)

enviar_email()