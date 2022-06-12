#pip install pyautogui
#pip install pyperclip
#pip install pandas
#pip install numpy
#pip install openpyxl

import time
# automatiza mouse, teclado e monitor do computador
import pyautogui
# Permite copiar e colar
import pyperclip
# pandas permite trabalhar com base de dados de forma eficiente
import pandas as pd

#pyautogui.hotkey -> Conjunto de teclas
#pyautogui.write -> Escrever um texto
#pyautogui.press -> Aprender 1 tecla
#pyautogui.click -> Clica com o mouse

pyautogui.PAUSE = 1

#Passo 1: Entrar no sistema da empresa(no nosso caso vai ser o link do drive)
pyautogui.hotkey("win")
pyautogui.write("chrome")
pyautogui.press("enter")

pyautogui.hotkey("ctrl", "t")
link = "https://drive.google.com/drive/folders/14oLE59U1RqyRqlBbKpsyymW-mitvbtoh"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

#demora alguns segundos
time.sleep(20)

#Passo 2: Navegar no sistemas e encontrar a base de dados (entrar na pasta exportar)
print(pyautogui.position())
#pyautogui.click(x= , y=)
pyautogui.click(x=392 , y=466, button='right')
time.sleep(10)


#Passo 3: Exportar/Fazer Download da Base de Dados
print(pyautogui.position())
pyautogui.click(x=534 , y=899)
time.sleep(6) # Esperar fazer download

#Passo 4: Importar a base de Dados para o Python
tabela = pd.read_excel(r"C:/Users/USER/Downloads/Vendas - Dez.xlsx")
print(tabela)

#Passo 5: Calcular os indicadores
#faturamento  - soma das colunas valor final
faturamento = tabela["Valor Final"].sum()

#quantidade de produtos  - soma da coluna Quantidade
quantidade = tabela["Quantidade"].sum()

#Passo 6: Enviar um e-mail para a diretoria com o relatório
#abrir o e-mail ("link")

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("mail.google.com")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(20)

#clicar no escrever
print(pyautogui.position())
pyautogui.click(x=110, y=249)
time.sleep(10)

#escreve o email do destinatario
pyautogui.write("gustavo.dev97@gmail.com")
pyautogui.press("enter")
time.sleep(4)

#pyautogui.press("tab") # -> passar para o campo de assunto
pyautogui.press("tab")
time.sleep(4)

#escreve o assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
time.sleep(4)

#escreve o corpo do email

#f = diz ao python que o texto será formatado permitindo valores dinâmicos (sempre que quiser inserir uma variável dentro de uma string)

texto = f"""
Prezados Bom dia!

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,} 

Abs!
Gustavo!"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

#envia o email
pyautogui.hotkey("ctrl", "enter")