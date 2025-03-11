# Projeto para automatizar a tarefa de trocar a senha dos cofres de todos os postos que estão no sistema do B2Click

import pyautogui as py
import time

time.sleep(5)

empresas = [150, 156, 159, 167, 440, 450, 482, 487, 488, 490,
            491, 492, 497, 498, 499, 505, 513, 514, 515]

# OBS: 161 não tem Cofre

def acessar_tela_opnf():
    py.hotkey("ctrl", "f")
    time.sleep(2)
    py.write("OPNF")
    time.sleep(1)
    py.press("enter")
    time.sleep(2)
    py.press("enter")
    time.sleep(4)
    py.hotkey("alt", "o")
    time.sleep(2)
    py.press("enter")
    time.sleep(2)
    py.press(["tab"] * 5)  # 5 tabs
    time.sleep(1)
    py.press("up")
    time.sleep(2)

acessar_tela_opnf()

for empresa in empresas:
    py.write(str(empresa))  # Digita o ID da Empresa
    time.sleep(1.5)
    py.press(["tab"] * 9)  # 9 tabs para chegar no usuario!
    time.sleep(1.5)
    py.write("Login")  # Altera o Usúario
    time.sleep(1.5)
    py.press("tab")
    time.sleep(1.5)
    py.write("Senha")  # Altera a Senha
    time.sleep(1.5)

    py.hotkey("alt", "v")  # alt + V (salva os dados)
    time.sleep(1.5)
    py.press("enter")
    time.sleep(1.5)

    # Volta para o campo de empresas
    py.hotkey(["shift", "tab"] * 9)  # 9 shift tabs para voltar para a empresa e um up
    time.sleep(1)
    py.press("up")
    time.sleep(1.5)

print("Processo Finalizado!")