from ipi_a_recup_utils import perform_actions, extract_saldo_credor, compare_and_write_to_excel, perform_actions_cofins_a_recup
from bs4 import BeautifulSoup
import openpyxl
import pyautogui
import tkinter as tk
from tkinter import simpledialog
import time
import os
import pygetwindow as gw

pyautogui.FAILSAFE = False

url_pis_e_cofins = r"C:\fiscal\html\temp_pis_e_cofins.htm"

# Função para obter o código da empresa, o mês e o ano do usuário
def get_user_input():
    root = tk.Tk()
    root.withdraw()  # Esconder a janela principal

    # Solicitar o código da empresa
    company_code = simpledialog.askstring(title="Código da Empresa", prompt="Digite o código da empresa:")

    # Solicitar o mês e o ano
    month_year = simpledialog.askstring(title="Mês e Ano", prompt="Digite o mês e o ano (MMYYYY):")

    return company_code, month_year

# Obter o código da empresa e o mês e o ano do usuário
company_code, month_year = get_user_input()

day_month_year = '01' + month_year

network_path = r'\\ap05\modulos\UNICO.EXE'

# Pressionar Win + R
pyautogui.hotkey('win', 'r')
time.sleep(1)  # Esperar um momento para a janela de execução abrir

pyautogui.typewrite(network_path)
pyautogui.press('enter')

# Esperar 12 segundos para o executável carregar
time.sleep(12)

# Digitar "contabil"
pyautogui.typewrite('contabil')

# Pressionar "tab"
pyautogui.press('tab')

# Digitar "1234"
pyautogui.typewrite('1234')

# Pressionar "enter"
pyautogui.press('enter')

# Esperar um momento para a próxima ação
time.sleep(5)

# Pressionar "Ctrl + 0"
pyautogui.hotkey('ctrl', '1')


html_path_cofins_recup = "C:\\fiscal\\html\\temp_cofins_recup.htm"
perform_actions_cofins_a_recup(company_code, month_year, html_path_cofins_recup)
