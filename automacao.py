#!/usr/bin/env python
# coding: utf-8

# instalando a biblioteca pyautogui
#!pip install pyautogui

import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

#Entrar no sistema da empresa (no caso, o google drive)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# Pagina carregando
time.sleep(5)

# Navegar no sistema e encontrar a base de dados (entrar na pasta "Exportar")
pyautogui.click(x=412, y=274, clicks=2)
time.sleep(2)

# Download da base de dados
pyautogui.click(x=390, y=328)
pyautogui.click(x=1159, y=154)
pyautogui.click(x=1013, y=567)
time.sleep(7)

# Calcular os indicadores (faturamento, quantidade de produtos)
import pandas as pd

tabela = pd.read_excel(r"C:\Users\win10\Downloads\Vendas - Dez.xlsx")
# Mostrar tabela
display(tabela)

# Soma da coluna quantidade
quantidade = tabela["Quantidade"].sum()

# Soma da coluna faturamento
faturamento = tabela["Valor Final"].sum()

print(quantidade)
print(faturamento)

# Enviando pro email da diretoria
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(7)

# Mandar um email pra diretoria com os indicadores
# Clicar no botão "+ Escrever"
pyautogui.click(x=78, y=175)
time.sleep(2)

# Escrever o detinatário (para quem eu vou mandar)
pyautogui.write("email.exemplo@gmail.com")
# selecionando o email
pyautogui.press("tab")
#Passar pro campo de Assunto
pyautogui.press("tab")

# Escrever o assunto
pyperclip.copy("Relatório de vendas")
pyautogui.hotkey("ctrl", "v")
# Passar para o corpo do email
pyautogui.press("tab")

# Escrever o corpo do email
texto = f""" Prezados, bom dia! 

O faturamento da empresa foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Atenciosamente,

Roberto Amaral.
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# Enviar email
pyautogui.hotkey("ctrl", "enter")


# Use esse código para descobrir qual a posição de um item que queira clicar 
# A posição na sua tela é diferente da posição na minha tela
import time

#Tempo (5 segundos) para reconhecer a posição do mouse
time.sleep(5)
# Retorna a posição do mouse
pyautogui.position()