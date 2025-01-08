import pyautogui
import time
import tkinter as tk
from tkinter import simpledialog
import pandas as pd
import os
import pygetwindow as gw
from bs4 import BeautifulSoup
import openpyxl

html_path = "C:\\fiscal\\html\\temp_ipi_recup.htm"

def perform_actions(empresa, mes, html_path):
    time.sleep(5)

    pyautogui.press('alt')
    time.sleep(1)
    pyautogui.press('o')
    time.sleep(1)
    pyautogui.press('p')
    time.sleep(2)
    pyautogui.write(mes)
    time.sleep(1)
    pyautogui.press('pgdn')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.leftClick(61, 121)
    time.sleep(3)
    pyautogui.write(mes)
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.leftClick(93, 164)
    time.sleep(1)
    pyautogui.leftClick(101, 226)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    pyautogui.write(html_path)
    time.sleep(1)
    pyautogui.press('enter')

    # Ensure the file is saved before attempting to open it
    time.sleep(5)


def extract_saldo_credor(html_path):
    try:
        with open(html_path, 'r', encoding='utf-8') as file:
            html_content = file.read()

        soup = BeautifulSoup(html_content, 'html.parser')
        saldo_credor_element = soup.find('td', string="SALDO CREDOR PERIODO SEGUINTE")

        if saldo_credor_element:
            ipi_a_recup = saldo_credor_element.find_next_sibling('td').get_text(strip=True)
            ipi_a_recup = float(ipi_a_recup.replace('.', '').replace(',', '.'))
            return ipi_a_recup
        else:
            print("Variável 'SALDO CREDOR PERIODO SEGUINTE' não encontrada.")
            return None
    except FileNotFoundError:
        print(f"Erro: O arquivo {html_path} não existe. Continuando a execução...")
        return None


def compare_and_write_to_excel(empresa, mes, ipi_a_recup):
    excel_path = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{empresa}_{mes}.xlsx'
    if not os.path.exists(excel_path):
        print(f"Erro: O arquivo {excel_path} não existe.")
        return

    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active

        for row in ws.iter_rows(min_row=2):
            cell_a = row[0].value  # Coluna A (índice 0)
            if cell_a == 42:
                valor_coluna_h = row[7].value  # Coluna H (índice 7)
                print(f"Número {cell_a} encontrado: Valor na coluna H = {valor_coluna_h}")

                if abs(valor_coluna_h - ipi_a_recup) <= 0.10:
                    row[8].value = "OK"
                else:
                    row[8].value = "Verificar"
                break

        wb.save(excel_path)
        print(f"Resultado da comparação escrito na planilha de conciliação.")
    except Exception as e:
        print(f"Erro ao abrir ou processar o arquivo de conciliação: {e}")


if __name__ == "__main__":
    html_path = "C:\\fiscal\\html\\temp_ipi_recup.htm"

    # Execute automation steps
    perform_actions(empresa, mes, html_path)

    # Extract and display result
    ipi_a_recup = extract_saldo_credor(html_path)
    if ipi_a_recup is not None:
        print(f"IPI a recuperar: {ipi_a_recup:.2f}")
        compare_and_write_to_excel(empresa, mes, ipi_a_recup)
        
os.remove("C:\\fiscal\\html\\temp_ipi_recup.htm")
