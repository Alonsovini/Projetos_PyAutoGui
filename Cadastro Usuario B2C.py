""" Automação para cadastrar usuario no sistema da RedeSoft de forma automatica
Preciso ajustar no caso de ter colaborador como caixa e gerente para vincular código de usuario ao agente
- Adicionar a informação de IBGE para os colaboradores para o codigo funcionar
- Alterar o caminho da sua planilha.

"""

import time
import pandas as pd
import pyautogui as py

df = pd.read_excel(r"C:\Users\vinicius.alonso\Downloads\Bartolomeu.xlsx")

print(df.columns)

colunas = ['Nome', 'CPF', 'Tipo', 'IBGE', 'CEP', 'UF', 'Bairro', 'Rua', 'Numero']
df = df[colunas]

time.sleep(3)
py.press("a")  # Acessar aba de agente
py.press("enter")
time.sleep(1)
py.press(["tab"] * 11)  # Acessar box de é funcionario
time.sleep(0.5)
py.press("space")  # Vendedor
time.sleep(0.5)
py.press("tab")
time.sleep(0.5)
py.press("space")  # Funcionario

def cadastro(Nome,UF,CEP,IBGE,Bairro,Tipo,Rua,Numero,CPF):
    time.sleep(2)
    py.hotkey("alt", "n")  # Vai para o campo de nome
    time.sleep(1)
    py.write(Nome)
    time.sleep(0.5)
    py.press("tab")
    py.write("Empresa")  # Colocar o nome do Posto que está fazendo
    py.press(["tab"] * 2)
    py.write(UF)  # Vem no arquivo que o RH manda
    time.sleep(0.5)
    py.press("tab")
    py.write(str(CEP))  # Vem no arquivo que o Rh manda
    time.sleep(0.5)
    py.press("tab")
    time.sleep(0.5)
    py.write(str(IBGE))  # Precisa adicionar de acordo com a cidade
    time.sleep(0.5)
    # py.press("backspace")
    py.press(["tab"] * 2)
    py.write(Bairro)  # Preenche o bairro que vem por padrão
    time.sleep(0.5)
    py.press("tab")
    py.write(Tipo)  # Precisa editar na planilha (Rua,Avenida)
    time.sleep(0.5)
    py.press("tab")
    py.write(Rua)  # Ajustar na planilha para apenas o endereço sem o TIPO
    time.sleep(0.5)
    py.press("tab")
    py.write(str(Numero))
    time.sleep(0.5)
    py.hotkey("alt", "t")  # Trocar o Tipo de CNPJ para CPF
    time.sleep(0.5)
    py.press("tab")
    py.press("enter")
    time.sleep(0.5)
    py.press(["tab"] * 10)
    time.sleep(2)
    py.write(str(CPF))  # Tirar . e - do CPF
    time.sleep(5)
    py.hotkey("alt", "s")  # Salva o login
    time.sleep(3)
    py.press(["enter"] * 2)
    time.sleep(5)
    py.hotkey("alt", "l")  # Limpa os dados para o proximo preenchimento
    time.sleep(5)


# Iterar pelas linhas do DataFrame
for index, row in df.iterrows():
    Nome = row['Nome']
    CPF = row['CPF']
    Tipo = row['Tipo']
    CEP = row['CEP']
    IBGE = row["IBGE"]
    UF = row["UF"]
    Bairro = row["Bairro"]
    Rua = row["Rua"]
    Numero = row["Numero"]

    # Chamar a função de cadastro
    cadastro(Nome, UF, CEP, IBGE, Bairro, Tipo, Rua, Numero, CPF)

    py.hotkey(["shift", "tab"] * 5)  # Volta para salvar as opções de Funcionario e Vendedor
    py.press("space")
    py.press("tab")
    py.press("space")


print("Finalizado!")