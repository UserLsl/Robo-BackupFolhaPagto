# Detalhes -----------------------------------------------------------------------------------------------------------
#
# Autor:      Lucas S. Leinatti
# Data:       30/06/2024
# Objetivo:   Script para automatizar o processo de backup de um sistema de folha de pagamento especifico.
# Requisitos: Executar na máquina onde o sistema está instalado; Outlook instalado e configurado com uma conta.
#
# Fim Detalhes -------------------------------------------------------------------------------------------------------

# Bibliotecas --------------------------------------------------------------------------------------------------------
print("Carregando bibliotecas...")

import pyautogui                # Automação de tela
import time                     # Timers
import ctypes                   # Verificar NumLock
# import psutil                   # Verificar programas abertos
import socket                   # Verificando informações da máquina
import os                       # Acessando ficheiros e arquivos
from datetime import datetime
import win32com.client as win32 # Manipulando programas de terceiros
import pygetwindow as gw        # Manipulando janelas
# Fim Bibliotecas ----------------------------------------------------------------------------------------------------

# Parâmetros ---------------------------------------------------------------------------------------------------------
print("Carregando parâmetros...")

pyautogui.FAILSAFE = False # Desabilitar esse recurso de segurança que evita que o mouse interrompa a execução.
hostname = socket.gethostname()
log_file = r'C:\caminho\log.txt'
# Fim Parâmetros -----------------------------------------------------------------------------------------------------

# Funções ------------------------------------------------------------------------------------------------------------
print("Carregando funções...")

# Adicionando registros no arquivo de log
def add_log(msg):
    try:
        with open(log_file, 'a') as file:
            file.write(f'{datetime.now().strftime("%Y-%m-%d %H:%M:%S")} - {msg}\n')
    except Exception as e:
        error_message = str(e)
        print(f"Ocorreu o seguinte erro na edição do arquivo de log:\n{error_message}")
        input("Pressione Enter para encerrar...")

# Enviando e-mail
def sending_email(message):
    try:
        print("Enviando email...")
        
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.To = "destination@email.com"
        email.Subject = "Backup Automático Folha"
        email.HTMLBody = f"""
        <p>{message}</p>
        <p>Atenciosamente,</p>
        <p>X</p>
        """

        hoje = datetime.today()
        nome_arquivo = f"FP{hoje.day:02d}{hoje.month:02d}{hoje.year % 100:02d}.SCZ"
        
        anexo1 = r'C:\caminho\log.txt'
        email.Attachments.Add(anexo1)
        caminho_pasta = r"C:\caminho"
            
        caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)
        if os.path.isfile(caminho_arquivo):
            anexo2 = caminho_arquivo
            email.Attachments.Add(anexo2)
            
        email.Send()
    except Exception as e:
        error_message = str(e)
        add_log(f"Ocorreu o seguinte erro no envio do e-mail:\n{error_message}")
        print(f"Ocorreu o seguinte erro no envio do e-mail:\n{error_message}")
        input("Pressione Enter para encerrar...")  

# Verificando NumLock
def is_num_lock_on():
    try:
        return ctypes.windll.user32.GetKeyState(0x90) & 1 != 0
    except Exception as e:
        error_message = str(e)
        sending_email(error_message)
        add_log(f"Ocorreu o seguinte erro na declaração da função para desabilitar a NumLock:\n{error_message}")
        print(f"Ocorreu o seguinte erro na declaração da função para desabilitar a NumLock:\n{error_message}")
        input("Pressione Enter para encerrar...")

# Verificando processo
def verify_window(janela):
    try:
        window = gw.getWindowsWithTitle(janela)
        return len(window) > 0
    except Exception as e:
        error_message = str(e)
        sending_email(error_message)
        add_log(f"Ocorreu o seguinte erro na declaração da função para verificar janela:\n{error_message}")
        print(f"Ocorreu o seguinte erro na declaração da função para verificar janela:\n{error_message}")
        input("Pressione Enter para encerrar...") 
#Fim Funções --------------------------------------------------------------------------------------------------------

# Log ----------------------------------------------------------------------------------------------------------------
print("Carregando arquivo de log...")

try:
    if os.path.exists(log_file):
        with open(log_file, 'a') as file:
            file.write(f'{datetime.now().strftime("%Y-%m-%d %H:%M:%S")} - Backup iniciado\n')
    else:
        with open(log_file, 'w') as file:
            file.write(f'{datetime.now().strftime("%Y-%m-%d %H:%M:%S")} - Backup iniciado\n')
except Exception as e:
    error_message = str(e)
    sending_email(error_message)
    print(f"Ocorreu o seguinte erro na abertura do arquivo de log:\n{error_message}")
    input("Pressione Enter para encerrar...")
# Fim Log ------------------------------------------------------------------------------------------------------------

# Realizando backup -------------------------------------------------------------------------------------------------
add_log("Realizando backup...")
print("Realizando backup...")

try:
    time.sleep(3)
    pyautogui.hotkey('win')
    time.sleep(3)
    pyautogui.write("folha SCI Visual Practice", interval=0.1)
    time.sleep(2)
    pyautogui.press('enter')
    cont = 0
    while not gw.getWindowsWithTitle('folha SCI Visual Practice'):
        time.sleep(1)
        cont += 1
        print('Esperando abertura do programa...')
        if cont == 50:
            raise ValueError("Tempo limite excedido. Interrompendo execução...")
    add_log(f"[{cont}] Esperando abertura do programa...")
    add_log("Programa aberto.")
    print('Programa aberto.')
    time.sleep(20)
    # time.sleep(50)
    pyautogui.click(960, 580) # Fechando possível mensagem de aviso de atualização
    time.sleep(1)
    pyautogui.write("user", interval=0.1)
    time.sleep(2)
    pyautogui.press('tab')
    time.sleep(2)
    pyautogui.write("password", interval=0.1)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.click(1645, 20) # Fechando possível tela de aviso
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press('alt')
    time.sleep(2)
    pyautogui.press('u')
    if is_num_lock_on():
        pyautogui.press('numlock')
        time.sleep(1)
    for _ in range(11):
        pyautogui.press('down')
        time.sleep(0.1)
    for _ in range(3):
        pyautogui.press('enter')
        time.sleep(2)
    pyautogui.click(487, 279) # Clicando na opção total para avisar bug do suporte
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.click(58, 131) # Clicando em Iniciar Backup
    time.sleep(2)
    pyautogui.press('enter')
    add_log("Salvando arquivo de backup...")
    print("Salvando arquivo de backup...")
    time.sleep(300)
    for _ in range(3):
        pyautogui.press('enter')
        time.sleep(2)
    add_log("Encerrando backup...")
    print("Encerrando backup...")
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(5)
except Exception as e:
    error_message = str(e)
    sending_email(error_message)
    add_log(f"Ocorreu o seguinte erro no processo automatizado de backup:\n{error_message}")
    print(f"Ocorreu o seguinte erro no processo automatizado de backup:\n{error_message}")
    input("Pressione Enter para encerrar...")

add_log("Backup finalizado.")
print('Backup finalizado.')
# Fim Realizando backup ---------------------------------------------------------------------------------------------

# Verificando sucesso no backup -------------------------------------------------------------------------------------
add_log("Verificando backup...")
print("Verificando backup...")

try:
    if verify_window('folha SCI Visual Practice.exe'):
        raise ValueError("Programa SCI Visual Practice ainda aberto!")
    else:
        add_log("Programa SCI encerrado.")
        print("Programa SCI encerrado.")
except Exception as e:
    error_message = str(e)
    sending_email(error_message)
    add_log(f"Ocorreu o seguinte erro na verificação da finalização do programa SCI:\n{error_message}")
    print(f"Ocorreu o seguinte erro na verificação da finalização do programa SCI:\n{error_message}")
    input("Pressione Enter para encerrar...")

try:
    hoje = datetime.today()
    nome_arquivo = f"FP{hoje.day:02d}{hoje.month:02d}{hoje.year % 100:02d}.SCZ"
    caminho_pasta = r"C:\caminho"
    caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)
    if os.path.isfile(caminho_arquivo):
        add_log(f"O arquivo {nome_arquivo} existe na pasta.")
        add_log("Backup realizado com sucesso!")
        print(f"O arquivo {nome_arquivo} existe na pasta.")
        print("Backup realizado com sucesso!")
        sending_email("O backup do sistema SCI foi realizado com sucesso!")
    else:
        raise ValueError("Arquivo de backup não encontrado na pasta de destino!")
except Exception as e:
    error_message = str(e)
    sending_email(error_message)
    add_log(f"Ocorreu o seguinte erro na verificação do arquivo de backup:\n{error_message}")
    print(f"Ocorreu o seguinte erro na verificação do arquivo de backup:\n{error_message}")
    input("Pressione Enter para encerrar...")
# Fim Verificando sucesso no backup ---------------------------------------------------------------------------------




