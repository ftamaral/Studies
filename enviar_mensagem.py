# Está automção envia uma mensagem via Teams para usuários listado em uma planilha excel.

#--------- BIBLIOTECAS ----------------------

import pyautogui
from time import sleep
import pandas as pd
from pynput.keyboard import Controller


#--------- VARIÁVEIS -------------------------

keyboard = Controller()
df = pd.read_excel('C:\\planilha_excel.xlsx')
teste_auto = 'É uma mensagem de teste, não precisa ser respondida.\n'
hatual = datetime.now()
heditada = float(hatual.strftime("%H"))
meio_dia = float(12)

#--------- AUTOMAÇÃO -------------------------

for a in df["Usuario"]:                             # Loop
    
    sleep(1)
    pyautogui.moveTo(x=560, y=34)
    sleep(1)
    pyautogui.leftClick(x=560, y=34)
    sleep(1)
    keyboard.type(a + "@sicredi.com.br")
    sleep(3)
    pyautogui.leftClick(x=484, y=147)
    sleep(3)
    pyautogui.moveTo(x=663, y=835)
    sleep(3)
    pyautogui.leftClick(x=663, y=835)
    sleep(3)
    if heditada < meio_dia:                         # Condicional
        keyboard.type('Olá bom dia\n')
    else:
        keyboard.type('Olá boa tarde\n')
    sleep(3)
    keyboard.type(msn)
