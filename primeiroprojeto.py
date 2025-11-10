# Criando um projeto de automação

# Importando bibliotecas
import pyautogui
import time
import pandas
import pyperclip

# Entrar no sistema da empresa (Acessando o link)
pyautogui.PAUSE = 2

pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")


time.sleep(15)

link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyautogui.write(link)
pyautogui.press("enter")

# Navegar no sistema para encontrar o banco de dados
time.sleep(15)
pyautogui.click(x=610, y=553, clicks=2) # Acessando a pasta aonde está o arquivo

# Exportar a base de dados (baixar o arquivo)
pyautogui.click(x=610, y=553) # Dando um clique no arquivo
pyautogui.click(x=846, y=441) # Baixando o arquivo

# Calcular os indicadores (faturamento e a quantidade de produtos vendidos)
import pandas

caminho = r"C:\Users\Kelvin Rodrigues\Downloads\Vendas - Dez.xlsx" # procurando o arquivo baixado
tabela = pandas.read_excel(caminho) # lendo o arquivo
print(tabela) # imprimindo as informações

faturamento = tabela["Valor Final"].sum() # somando os valores na coluna selecionada
qtde_produtos = tabela["Quantidade"].sum() # somando os valores na coluna selecionada

# Enviar informações por e-mail
import pyperclip

pyautogui.hotkey("ctrl", "t") # abrindo uma nova aba
time.sleep(5)

pyautogui.write("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox") # acessando o Gmail
pyautogui.press("enter")
time.sleep(15)
pyautogui.click(x=127, y=237) # abrindo a caixa de mensagem
time.sleep(5)

pyautogui.write("olxconta963@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
pyperclip.copy("Relatório de vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

texto = f"""
Prezados,

Segue o relatório de vendas da data de hoje:

Faturamento: R${faturamento: ,.2f}
Quantidade de itens vendidos: {qtde_produtos:,}

Fico à disposição para qualquer esclarecimento.

atenciosamente,
Kelvin Rodrigues
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")
time.sleep(5)
pyautogui.click(x=414, y=964)