from bs4 import BeautifulSoup
import openpyxl
import pyautogui
import tkinter as tk
from tkinter import simpledialog
import time
import os

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
pyautogui.hotkey('alt', 'tab')
time.sleep(1)
pyautogui.press('alt')
time.sleep(1)
pyautogui.press('o')
time.sleep(1)
pyautogui.press('i')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.write(month_year)
time.sleep(1)
pyautogui.press('enter')
pyautogui.press('enter')
pyautogui.press('enter')
time.sleep(3)
pyautogui.leftClick(63, 123)
time.sleep(3)
pyautogui.write(month_year)
time.sleep(2)
pyautogui.press('enter')
time.sleep(1)
pyautogui.leftClick(95, 165)
time.sleep(1)
pyautogui.leftClick(110, 225)
time.sleep(1)
pyautogui.press('enter')

dir_icms_recup = r"C:\fiscal\html\icms_a_recup.htm"
caminho_html_pis_recup = r"C:\fiscal\html\pis_a_recup.htm"

time.sleep(3)
pyautogui.write(dir_icms_recup)
time.sleep(1)
pyautogui.press('enter')

time.sleep(5)

# Caminho do arquivo HTML
file_path = dir_icms_recup

pyautogui.hotkey('alt', 'tab')
time.sleep(1)
pyautogui.press('alt')
time.sleep(1)
pyautogui.press('o')
time.sleep(1)
pyautogui.press('l')
time.sleep(1)
pyautogui.write(month_year)
time.sleep(1)
pyautogui.press('pgdn')
time.sleep(1)
pyautogui.press('pgdn')
time.sleep(1)
pyautogui.press('enter')
time.sleep(5)

pyautogui.leftClick(80, 187)
time.sleep(1)
pyautogui.leftClick(152, 522)
time.sleep(1)
pyautogui.leftClick(64, 131)
time.sleep(1)
pyautogui.leftClick(120, 150)
time.sleep(1)
pyautogui.leftClick(126, 214)
time.sleep(1)
pyautogui.press('enter')
time.sleep(2)
pyautogui.write(r"C:\fiscal\html\pis_a_recup.htm")
time.sleep(2)
pyautogui.press('enter')

# Ensure the file is saved before attempting to open it
time.sleep(5)

# Soma os valores e adiciona uma label de PIS_recup
total_pis_recup = 0
with open(caminho_html_pis_recup, 'r', encoding='utf-8') as file:
    content = file.read()

# Skip analysis if the file has less than 100 lines
if len(content.splitlines()) < 100:
    print("HTML file has less than 100 lines, skipping analysis.")
else:
    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "01 – Crédito decorrente de operações próprias" in text)

    for result in results:
        row = result.find_parent('tr')
        next_row = row.find_next_sibling('tr')
        if next_row:
            next_next_row = next_row.find_next_sibling('tr')
            if next_next_row:
                observations = [td.get_text(strip=True) for td in next_next_row.find_all('td')]
                valid_observations = []
                for obs in observations:
                    try:
                        valid_observations.append(float(obs.replace('.', '').replace(',', '.')))
                    except ValueError:
                        continue
                total_pis_recup += sum(valid_observations)
                print(valid_observations)

    total_pis_recup = round(float(total_pis_recup), 2)
    if total_pis_recup == 0:
        total_pis_recup = None
    else:
        try:
            if total_pis_recup != 0:
                print(f"PIS_recup: {total_pis_recup:.2f}".replace('.', ','))
        except TypeError:
            print(f"PIS_recup: {total_pis_recup}")

# Lê o conteúdo do arquivo HTML
with open(file_path, 'r', encoding='utf-8') as file:
    content = file.read()

# Skip analysis if the file has less than 100 lines
if len(content.splitlines()) < 100:
    print("HTML file has less than 100 lines, skipping analysis.")
else:
    # Cria um objeto BeautifulSoup
    soup = BeautifulSoup(content, 'html.parser')

    # Busca a linha que começa com <td colspan="2" class="s9">190</td>
    result = soup.find('td', {'colspan': '2', 'class': 's9'}, string='190')

    # Exibe a linha inteira
    if result:
        row = result.find_parent('tr')
        observations = row.find_all('td')
        if len(observations) > 1:
            ICMS_recup = observations[2].text
            ICMS_recup = ICMS_recup.replace('.', '').replace(',', '.')
            ICMS_recup = float(ICMS_recup)
            print(f"ICMS_recup: {ICMS_recup:.2f}".replace('.', ','))

with open(r"C:\fiscal\html\temp.html.htm", 'r', encoding='utf-8') as file:
    html_content = file.read()

# Parsear o HTML
soup = BeautifulSoup(html_content, 'html.parser')

linha = soup.find_all('td', class_='s8')  # Busca todas as células com classe 's8'

# Inicializa linha_completa como None
linha_completa = None

# Filtrar com base no texto 'Total Geral'
for td in linha:
    if 'Total Geral' in td.get_text():
        linha_completa = td.find_parent('tr')  # Encontra a linha (<tr>) associada à célula

# Verifica se linha_completa foi definida antes de usá-la
if linha_completa:
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

# Caminho do arquivo HTML
url_pis_e_cofins = r"C:\fiscal\html\temp_pis_e_cofins.htm"

# Função para extrair valores de PIS e COFINS do HTML
def extrair_valores_pis_cofins(caminho_html):
    with open(caminho_html, 'r', encoding='utf-8') as file:
        html_content = file.read()
    
    soup = BeautifulSoup(html_content, 'html.parser')
    valores = soup.find_all('td', string=lambda text: text and "Valor da contribuição não cumulativa a recolher/pagar" in text)
    
    if len(valores) >= 2:
        pis_valor = valores[0].find_next_sibling('td').get_text(strip=True).replace('.', '').replace(',', '.')
        cofins_valor = valores[1].find_next_sibling('td').get_text(strip=True).replace('.', '').replace(',', '.')
        return float(pis_valor), float(cofins_valor)
    else:
        return None, None

# Extrair valores de PIS e COFINS
pis_a_recolher, cofins_a_recolher = extrair_valores_pis_cofins(url_pis_e_cofins)

# Verificar se os valores foram extraídos corretamente
if pis_a_recolher is not None and cofins_a_recolher is not None:
    print(f"PIS a recolher: {pis_a_recolher}")
    print(f"COFINS a recolher: {cofins_a_recolher}")
else:
    print("Erro ao extrair valores de PIS e COFINS do HTML.")

# Abrir a planilha e procurar pelos números na coluna A
excel_path = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}.xlsx'
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

# Números a serem procurados
numeros_procurados = [617, 2707, 185, 187, 186, 41, 46, 197, 196]

# Procurar os números na coluna A e mapear para a coluna H
for row in ws.iter_rows(min_row=2):
    cell_a = row[0].value  # Coluna A (índice 0)
    if cell_a in numeros_procurados:
        valor_coluna_h = row[7].value  # Coluna H (índice 7)
        print(f"Número {cell_a} encontrado: Valor na coluna H = {valor_coluna_h}")
        
        # Comparar os valores e escrever "OK" ou "Verificar" na coluna I
        try:
            if cell_a == 617:
                if abs(valor_coluna_h - csrf_retido) <= 0.10:
                    row[8].value = "OK"
                else:
                    row[8].value = "Verificar"
            elif cell_a == 2707:
                if abs(valor_coluna_h - inss_retido) <= 0.10:
                    row[8].value = "OK"
                else:
                    row[8].value = "Verificar"
            elif cell_a == 185:
                if abs(valor_coluna_h - irrf_retido) <= 0.10:
                    row[8].value = "OK"
                else:
                    row[8].value = "Verificar"
            elif cell_a == 186:
                if abs(valor_coluna_h - iss_retido) <= 0.10:
                    row[8].value = "OK"
                else:
                    row[8].value = "Verificar"
            elif cell_a == 41:
                if abs(valor_coluna_h - ICMS_recup) <= 0.10:
                    row[8].value = "OK"
                else:
                    row[8].value = "Verificar"
            elif cell_a == 46 and total_pis_recup is not None:
                valor_coluna_h = float(str(valor_coluna_h))
                if abs(valor_coluna_h - total_pis_recup) <= 0.10:
                    row[8].value = "OK"
                else:
                    row[8].value = "Verificar"
            elif cell_a == 197:
                if abs(valor_coluna_h - pis_a_recolher) <= 0.10:
                    row[8].value = "OK"
                else:
                    row[8].value = "Verificar"
            elif cell_a == 196:
                if abs(valor_coluna_h - cofins_a_recolher) <= 0.10:
                    row[8].value = "OK"
                else:
                    row[8].value = "Verificar"
        except TypeError:
            continue

# Salvar as alterações de volta no arquivo Excel
wb.save(excel_path)

# Excluir os arquivos HTML

os.remove(url_pis_e_cofins)
os.remove(dir_icms_recup)
os.remove(caminho_html_pis_recup)
os.remove(r"C:\fiscal\html\temp.html.htm")

print('hello world!')