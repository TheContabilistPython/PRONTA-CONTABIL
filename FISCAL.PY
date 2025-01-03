from bs4 import BeautifulSoup
import openpyxl
import pyautogui
import tkinter as tk
from tkinter import simpledialog
import time

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

# Esperar 10 segundos para a próxima ação
time.sleep(5)

pyautogui.press('alt')
time.sleep(1)

pyautogui.press('E')
time.sleep(1)

pyautogui.leftClick(788, 506)
time.sleep(1)
pyautogui.leftClick(874, 117)

pyautogui.write(company_code)
time.sleep(2)
pyautogui.press('enter') 
time.sleep(1)

pyautogui.write(day_month_year)
time.sleep(1)

pyautogui.press('tab')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)

pyautogui.click(99, 165)
time.sleep(1)
pyautogui.click(251, 223)
time.sleep(3)
pyautogui.press('enter')
time.sleep(5)

arquivo_html_RET = r"C:\fiscal\html\temp.html"

pyautogui.write(arquivo_html_RET)
time.sleep(3)
pyautogui.press('enter')
time.sleep(5)

pyautogui.hotkey('alt', 'tab')

pyautogui.press('alt')
time.sleep(1)
pyautogui.press('o')
time.sleep(1)
pyautogui.press('l')
time.sleep(1)
pyautogui.write(month_year)
time.sleep(1)
pyautogui.press('enter')
pyautogui.press('pgdn')
time.sleep(1)
pyautogui.press('enter')
pyautogui.leftClick(63, 123)
time.sleep(3)
pyautogui.leftClick(95, 122)
time.sleep(2)
pyautogui.leftClick(95, 186)
time.sleep(2)
pyautogui.press('enter')
time.sleep(5)
pyautogui.write(url_pis_e_cofins)
time.sleep(2)
pyautogui.press('enter')

time.sleep(5)
"C:\fiscal\html\temp_pis_e_cofins.htm"

with open(r"C:\fiscal\html\temp.html.htm", 'r', encoding='utf-8') as file:
    html_content = file.read()

# Parsear o HTML
soup = BeautifulSoup(html_content, 'html.parser')

linha = soup.find_all('td', class_='s8')  # Busca todas as células com classe 's8'

# Filtrar com base no texto 'Total Geral'
for td in linha:
    if 'Total Geral' in td.get_text():
        linha_completa = td.find_parent('tr')  # Encontra a linha (<tr>) associada à célula

# Extrair valores dinâmicos, ignorando a primeira observação e formatando os valores
valores = []
for td in linha_completa.find_all('td')[1:]:
    try:
        valor = float(td.get_text(strip=True).replace('.', '').replace(',', '.'))
        valores.append(round(valor, 2))
    except ValueError:
        continue

# Atribuir valores às variáveis
Pis_retido = valores[1]
Cofins_retido = valores[2]
csll_retido = valores[3]
irrf_retido = valores[4]
iss_retido = valores[5]
inss_retido = valores[6]

# Calcular o total de CSRF retido
csrf_retido = Pis_retido + Cofins_retido + csll_retido

for td in linha_completa.find_all('td'):
    print(td.get_text(strip=True))  # Remove espaços extras

# Exibir valores retidos
print(f"PIS Retido: {Pis_retido:.2f}")
print(f"Cofins Retido: {Cofins_retido:.2f}")
print(f"CSLL Retido: {csll_retido:.2f}")
print(f"IRRF Retido: {irrf_retido:.2f}")
print(f"ISS Retido: {iss_retido:.2f}")
print(f"INSS Retido: {inss_retido:.2f}")
print(f"CSRF Retido: {csrf_retido:.2f}")

# Abrir a planilha e procurar pelos números na coluna A
excel_path = r"C:\projeto\planilhas\balancete\CONCILIACAO_8801_112024.xlsx"
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

# Números a serem procurados
numeros_procurados = [617, 2707, 185, 187, 186]

# Procurar os números na coluna A e mapear para a coluna H
for row in ws.iter_rows(min_row=2):
    cell_a = row[0].value  # Coluna A (índice 0)
    if cell_a in numeros_procurados:
        valor_coluna_h = row[7].value  # Coluna H (índice 7)
        print(f"Número {cell_a} encontrado: Valor na coluna H = {valor_coluna_h}")
        
        # Comparar os valores e escrever "OK" ou "Verificar" na coluna I
        if cell_a == 617:
            if valor_coluna_h == csrf_retido:
                row[8].value = "OK"
            else:
                row[8].value = "Verificar"
        elif cell_a == 2707:
            if valor_coluna_h == inss_retido:
                row[8].value = "OK"
            else:
                row[8].value = "Verificar"
        elif cell_a == 185:
            if valor_coluna_h == irrf_retido:
                row[8].value = "OK"
            else:
                row[8].value = "Verificar"
        elif cell_a == 186:
            if valor_coluna_h == iss_retido:
                row[8].value = "OK"
            else:
                row[8].value = "Verificar"

# Salvar as alterações de volta no arquivo Excel
wb.save(excel_path)
