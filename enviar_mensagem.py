#--------- BIBLIOTECAS ----------------------

import pyautogui
from time import sleep
import pandas as pd
from pynput.keyboard import Controller


#--------- VARIÁVEIS -------------------------
keyboard = Controller()
df = pd.read_excel('C:\\Vscode\\bitlocker\\primeiros15\\primeiros15.xlsx')
msn = 'Em conjunto com o time de Segurança da Informação, estamos avaliando algumas estações que estão com a criptografia desativada, gerando risco de perda de dados corporativos. A sua estação foi reportada com esta falha e precisamos corrigir.\n'
bd = 'Olá, bom dia\n'


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
    sleep(1)
    keyboard.type(msn)
