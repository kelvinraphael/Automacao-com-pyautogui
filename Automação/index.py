import pyautogui
import time
import pandas as pd
import openpyxl
import pyperclip

pyautogui.PAUSE = 0.5

# Abrir o navegador (Chrome)
pyautogui.press('win')
pyautogui.write('google chrome')
pyautogui.press('enter')

time.sleep(1)

# Entrar no link do Google Drive
linkdrive = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyautogui.write(linkdrive)
pyautogui.press('enter')

time.sleep(5)

# Realizar cliques para baixar o arquivo (posições baseadas na tela do autor)
pyautogui.click(x=357, y=333, clicks=2)
time.sleep(2)
pyautogui.click(x=495, y=334)
time.sleep(1)
pyautogui.click(x=562, y=409)

time.sleep(7)

# Ler a tabela baixada
caminho = r"C:\Users\kelvi\Downloads\Vendas - Dez.xlsx"
tabela = pd.read_excel("Vendas - Dez.xlsx")

print(tabela)

# Calcular indicadores
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

print(faturamento)
print(quantidade)

# Abrir nova aba
pyautogui.hotkey("ctrl", "t")
time.sleep(2)