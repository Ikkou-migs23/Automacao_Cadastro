import pandas as pd
import os
import pyautogui as pg
from time import sleep

# Bloco base, acessa o site e vai para a aba clientes!
"""def base():
    pg.press('win')
    sleep(5)
    pg.write('Brave')
    sleep(5)
    pg.press('enter')
    sleep(5)
    pg.write('http://127.0.0.1:5500/site-example/index.html')
    sleep(5)
    pg.press('enter')
    sleep(5)

def cad_clientes():
    pg.press('tab')
    pg.press('tab')
    sleep(5)
    pg.press('enter')


base()
cad_clientes()"""

caminho = r'/home/migs/Documentos/estudos/Automacao_Cadastro/data/base_erp_completa.xlsx'

leitor = pd.read_excel(caminho)

print(leitor)