#--------- BIBLIOTECAS ----------------------

import pyautogui
from time import sleep
import pandas as pd
from pynput.keyboard import Controller


#--------- VARIÁVEIS -------------------------

keyboard = Controller()
df = pd.read_excel('C:\\planilha_excel.xlsx')
bd = 'Olá, bom dia\n'
teste_auto = 'É uma mensagem de teste, não precisa ser respondida.\n'

#--------- AUTOMAÇÃO -------------------------

for u in df["Usuario"]:
    sleep(2)
    pyautogui.moveTo(x=560, y=24)
    sleep(2)
    pyautogui.leftClick(x=560, y=24)
    keyboard.type(u + "@dominio.com.br")
    sleep(2)
    pyautogui.leftClick(x=484, y=147)
    sleep(2)
    pyautogui.moveTo(x=663, y=835)
    sleep(2)
    pyautogui.leftClick(x=663, y=835)
    sleep(2)
    keyboard.type(bd)
