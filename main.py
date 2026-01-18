from openpyxl import load_workbook
import os
import pyautogui as pg
from time import sleep
import pandas as pd

# Mover o mouse para o canto superior esquerdo para interromper
pg.FAILSAFE = True  

# 1. ABRIR NAVEGADOR E ACESSAR SITE
def abrir_sistema():
    pg.press('win')
    sleep(2)
    pg.write('Brave')
    sleep(2)
    pg.press('enter')
    sleep(3)
    pg.write('http://127.0.0.1:5500/site-example/index.html')
    sleep(2)
    pg.press('enter')
    sleep(5)
    
    # Navegar até a aba Clientes
    pg.press('tab')
    pg.press('tab')
    sleep(2)
    pg.press('enter')
    sleep(3)  # Aguarda carregar a tela de clientes

# 2. CLICAR NO BOTÃO "NOVO CLIENTE"
def abrir_formulario_cadastro():
    print("Agora, posicione o mouse no botão 'Novo Cliente'...")
    sleep(3)
    pg.click()  
    sleep(3)  

# 3. CADASTRAR CLIENTES
def cadastrar_clientes():

    # Ler dados do Excel
    caminho = r'/home/migs/Documentos/estudos/Automacao_Cadastro/data/base_erp_completa.xlsx'
    df = pd.read_excel(caminho, sheet_name="Clientes")
    
    print("Posicione o cursor no PRIMEIRO CAMPO (Nome) em 5 segundos...")
    sleep(5)
    
    for indice, linha in df.iterrows():
        # Tratar possíveis valores NaN
        nome = str(linha['Nome']) if pd.notna(linha['Nome']) else ''
        email = str(linha['Email']) if pd.notna(linha['Email']) else ''
        telefone = str(linha['Telefone']) if pd.notna(linha['Telefone']) else ''
        
        print(f"Cadastrando cliente {indice+1}: {nome}")
        
        # Preecnher formulario na ordem correta

        # nome
        sleep(0.5)
        pg.write(nome)      
        pg.press('tab')     
        sleep(0.5)
        
        # email
        pg.write(email)    
        pg.press('tab')    
        sleep(0.5)
        
        # telefone
        pg.write(telefone) 
        pg.press('tab')     
        sleep(0.5)
        
        # salvar
        pg.press('enter')
        
        print(f"   {nome} salvo! Aguardando 2 segundos...")
        sleep(2)  # Aguarda o salvamento

# 4. Ir para produtos
def abrir_formulario_produtos():
    pg.press('tab')
    sleep(3)

    pg.press('tab')
    sleep(3)

    pg.press('tab')
    sleep(3)

    pg.press('tab')
    sleep(3)

    pg.press('tab')
    sleep(3)

    pg.press('tab')
    sleep(3)

    pg.press('tab')
    sleep(3)

    pg.press('tab')
    sleep(3)

    pg.press('tab')
    sleep(3)

    pg.press('tab')
    sleep(3)

    pg.press('enter')
    sleep(3)


    pg.press('tab')
    sleep(3)
    print("Agora, posicione o mouse no botão 'Novo Cliente'...")
    sleep(3)
    pg.click()  
    sleep(3)  

# 5. CADASTRAR CLIENTES
def cadastrar_produtos():
    # Ler dados do Excel
    caminho = r'/home/migs/Documentos/estudos/Automacao_Cadastro/data/base_erp_completa.xlsx'
    df = pd.read_excel(caminho, sheet_name="Produtos")
    
    print("Posicione o cursor no PRIMEIRO CAMPO (Nome) em 5 segundos...")
    sleep(5)
    
    for indice, linha in df.iterrows():
        # Tratar possíveis valores NaN
        nome = str(linha['Nome']) if pd.notna(linha['Nome']) else ''
        preco = str(linha['Preco']) if pd.notna(linha['Preco']) else ''
        estoque = str(linha['Estoque']) if pd.notna(linha['Estoque']) else ''
        
        print(f"Cadastrando produtos {indice+1}: {nome}")
        
        # Preecnher formulario na ordem correta

        # nome
        sleep(0.5)
        pg.write(nome)      
        pg.press('tab')     
        sleep(0.5)
        
        # email
        pg.write(preco)    
        pg.press('tab')    
        sleep(0.5)
        
        # telefone
        pg.write(estoque) 
        pg.press('tab')     
        sleep(0.5)
        
        # salvar
        pg.press('enter')
        
        print(f"   {nome} salvo! Aguardando 2 segundos...")
        sleep(2)  # Aguarda o salvamento

# EXECUTAR TUDO NA ORDEM CORRETA
abrir_sistema()              
abrir_formulario_cadastro() 
abrir_formulario_produtos()  
cadastrar_clientes()  
cadastrar_produtos()       

print("Cadastro automatizado concluído!")

#11