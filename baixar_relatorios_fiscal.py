import pyautogui
import time
from tkinter import simpledialog
import time
import os
import tkinter as tk



# Função para obter o código da empresa, o mês e o ano do usuário
def get_user_input():
    root = tk.Tk()
    root.withdraw()  # Esconder a janela principal
    
    # Solicitar o nome da empresa
    company_name = simpledialog.askstring(title="Nome da Empresa", prompt="Digite o nome da empresa:")

    # Solicitar o código da empresa
    company_code = simpledialog.askstring(title="Código da Empresa", prompt="Digite o código da empresa:")

    # Solicitar o mês e o ano
    month_year = simpledialog.askstring(title="Mês e Ano", prompt="Digite o mês e o ano (MMYYYY):")

    return company_code, month_year, company_name

# Obter o código da empresa e o mês e o ano do usuário
company_code, month_year, company_name = get_user_input()

day_month_year = '01' + month_year

time.sleep(5)

pyautogui.press('alt')
time.sleep(0.5)
pyautogui.press('r')
time.sleep(0.5)
pyautogui.press('q')
time.sleep(2)
pyautogui.write('22')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.write(company_code)
pyautogui.press('enter')
time.sleep(1)
pyautogui.write(day_month_year)
time.sleep(1)
pyautogui.press('tab')
pyautogui.write(month_year)

time.sleep(2)

pyautogui.leftClick(96, 123)
time.sleep(1)
pyautogui.leftClick(99, 184)    
time.sleep(1)
pyautogui.leftClick(695, 399)
time.sleep(2)
pyautogui.doubleClick(628, 698)
time.sleep(1)
pyautogui.write(f"C:\\relatorios_fiscal")
time.sleep(1)
pyautogui.press('enter')