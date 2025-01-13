from bs4 import BeautifulSoup
import openpyxl
import pyautogui
import tkinter as tk
from tkinter import simpledialog
import time
import os
import pygetwindow as gw

pyautogui.FAILSAFE = False

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

#####################

caminho_html_icms_a_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração DimeSC Sequência 22 - Ordem 1.htm"
total_icms_a_recolher = 0
try:
    with open(caminho_html_icms_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string=lambda text: text and "(=)Imposto a recolher" in text)

        for result in results:
            row = result.find_parent('tr')
            last_value = row.find_all('td')[-1].get_text(strip=True)
            try:
                total_icms_a_recolher += float(last_value.replace('.', '').replace(',', '.'))
            except ValueError:
                continue

        total_icms_a_recolher = round(float(total_icms_a_recolher), 2)
        if total_icms_a_recolher == 0:
            total_icms_a_recolher = None
        else:
            try:
                if total_icms_a_recolher != 0:
                    print(f"icms_a_recolher: {total_icms_a_recolher}")
            except TypeError:
                print(f"erro ao formatar icms_a_recolher do html")
        if total_icms_a_recolher is None:
            total_icms_a_recolher = 0
        print(f"icms_a_recolher: {total_icms_a_recolher}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_icms_a_recolher} não existe. Continuando a execução...")

#####################

caminho_html_ipi_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração IPI Sequência 22 - Ordem 2.htm"
total_ipi_recolher = 0
try:
    with open(caminho_html_ipi_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string=lambda text: text and "Valor do saldo devedor do IPI a recolher" in text)

        for result in results:
            row = result.find_parent('tr')
            last_value = row.find_all('td')[-1].get_text(strip=True)
            try:
                total_ipi_recolher += float(last_value.replace('.', '').replace(',', '.'))
            except ValueError:
                continue

        total_ipi_recolher = round(float(total_ipi_recolher), 2)
        if total_ipi_recolher == 0:
            total_ipi_recolher = None
        else:
            try:
                if total_ipi_recolher != 0:
                    print(f"ipi_recolher: {total_ipi_recolher:.2f}".replace('.', ','))
            except TypeError:
                print(f"erro ao formatar ipi_recolher do html")
        if total_ipi_recolher is None:
            total_ipi_recolher = 0
        print(f"ipi_recolher: {total_ipi_recolher}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_ipi_recolher} não existe. Continuando a execução...")


#####################

caminho_html_cofins_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de PISpasep e Cofins Sequência 22 - Ordem 4.htm"
total_cofins_recolher = 0
try:
    with open(caminho_html_cofins_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string=lambda text: text and "Valor total da contribuição a recolher/pagar no período (08 + 12)" in text)

        for result in results:
            row = result.find_parent('tr')
            last_value = row.find_all('td')[-2].get_text(strip=True)
            try:
                total_cofins_recolher += float(last_value.replace('.', '').replace(',', '.'))
            except ValueError:
                continue

        total_cofins_recolher = round(float(total_cofins_recolher), 2)
        if total_cofins_recolher == 0:
            total_cofins_recolher = None
        else:
            try:
                if total_cofins_recolher != 0:
                    print(f"cofins_recolher: {total_cofins_recolher:.2f}".replace('.', ','))
            except TypeError:
                print(f"erro ao formatar cofins_recolher do html")
        if total_cofins_recolher is None:
            total_cofins_recolher = 0
        print(f"cofins_recolher: {total_cofins_recolher}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_cofins_recolher} não existe. Continuando a execução...")

#####################

caminho_html_pis_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de PISpasep e Cofins Sequência 22 - Ordem 4.htm"
total_pis_recolher = 0
try:
    with open(caminho_html_pis_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string=lambda text: text and "Valor total da contribuição a recolher/pagar no período (08 + 12)" in text)

        for result in results:
            row = result.find_parent('tr')
            last_value = row.find_all('td')[-2].get_text(strip=True)
            try:
                total_pis_recolher += float(last_value.replace('.', '').replace(',', '.'))
            except ValueError:
                continue

        total_pis_recolher = round(float(total_pis_recolher), 2)
        if total_pis_recolher == 0:
            total_pis_recolher = None
        else:
            try:
                if total_pis_recolher != 0:
                    print(f"PIS_recolher: {total_pis_recolher:.2f}".replace('.', ','))
            except TypeError:
                print(f"erro ao formatar PIS_recolher do html")
        if total_pis_recolher is None:
            total_pis_recolher = 0
        print(f"PIS_recolher: {total_pis_recolher}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_pis_recolher} não existe. Continuando a execução...")

#####################

with open(f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de retenções Sequência 22 - Ordem 5.htm", 'r', encoding='utf-8') as file:
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

    # Exibir valores retidos
    print(f"IRRF Retido: {irrf_retido:.2f}")
    print(f"ISS Retido: {iss_retido:.2f}")
    print(f"INSS Retido: {inss_retido:.2f}")
    print(f"CSRF Retido: {csrf_retido:.2f}")

#####################

caminho_html_icms_cp_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração DimeSC Sequência 22 - Ordem 1.htm"
total_icms_cp_recup = 0
try:
    with open(caminho_html_icms_cp_recup, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string=lambda text: text and "(=) Saldo credor das antecipações para o mês seguinte" in text)

        for result in results:
            row = result.find_parent('tr')
            last_value = row.find_all('td')[-2].get_text(strip=True)
            try:
                total_icms_cp_recup += float(last_value.replace('.', '').replace(',', '.'))
            except ValueError:
                continue

        total_icms_cp_recup = round(float(total_icms_cp_recup), 2)
        if total_icms_cp_recup == 0:
            total_icms_cp_recup = None
        else:
            try:
                if total_icms_cp_recup != 0:
                    print(f"ICMS_cp_recup: {total_icms_cp_recup:.2f}".replace('.', ','))
            except TypeError:
                print(f"erro ao formatar COFINS_recup do html")
        print(f"ICMS_cp_recup: {total_icms_cp_recup}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_icms_cp_recup} não existe. Continuando a execução...")

#####################

caminho_html_cofins_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de PISpasep e Cofins Sequência 22 - Ordem 4.htm"
total_cofins_recup = 0
try:
    with open(caminho_html_cofins_recup, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string=lambda text: text and "Totais Cofins" in text)

        for result in results:
            row = result.find_parent('tr')
            last_value = row.find_all('td')[-1].get_text(strip=True)
            try:
                total_cofins_recup += float(last_value.replace('.', '').replace(',', '.'))
            except ValueError:
                continue

        total_cofins_recup = round(float(total_cofins_recup), 2)
        if total_cofins_recup == 0:
            total_cofins_recup = None
        else:
            try:
                if total_cofins_recup != 0:
                    print(f"COFINS_recup: {total_cofins_recup:.2f}".replace('.', ','))
            except TypeError:
                print(f"erro ao formatar COFINS_recup do html")
        print(f"COFINS_recup: {total_cofins_recup}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_cofins_recup} não existe. Continuando a execução...")

#####################


caminho_html_pis_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de PISpasep e Cofins Sequência 22 - Ordem 4.htm"
total_pis_recup = 0
try:
    with open(caminho_html_pis_recup, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string=lambda text: text and "Totais Pis/Pasep" in text)

        for result in results:
            row = result.find_parent('tr')
            last_value = row.find_all('td')[-1].get_text(strip=True)
            try:
                total_pis_recup += float(last_value.replace('.', '').replace(',', '.'))
            except ValueError:
                continue

        total_pis_recup = round(float(total_pis_recup), 2)
        if total_pis_recup == 0:
            total_pis_recup = None
        else:
            try:
                if total_pis_recup != 0:
                    print(f"PIS_recup: {total_pis_recup:.2f}".replace('.', ','))
            except TypeError:
                print(f"erro ao formatar PIS_recup do html")
        print(f"PIS_recup: {total_pis_recup}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_pis_recup} não existe. Continuando a execução...")
    
    
########################


html_path_icms_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração DimeSC Sequência 22 - Ordem 1.htm"

try:
    with open(html_path_icms_recup, 'r', encoding='utf-8') as file:
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
except FileNotFoundError:
    print(f"Erro: O arquivo {html_path_icms_recup} não existe. Continuando a execução...")

# Abrir a planilha e procurar pelos números na coluna A
excel_path = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}.xlsx'
wb = openpyxl.load_workbook(excel_path)
ws = wb.active
    
numeros_procurados = [41, 42, 46, 47, 2654, 617, 185, 2707, 186, 197, 196, 198, 195]

def extract_saldo_credor(html_path_ipi_a_recup):
    try:
        with open(f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração IPI Sequência 22 - Ordem 2.htm", 'r', encoding='utf-8') as file:
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
        print(f"Erro: O arquivo {html_path_ipi_a_recup} não existe. Continuando a execução...")
        return None

html_path_ipi_a_recup = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}.xlsx'
ipi_a_recup = extract_saldo_credor(html_path_ipi_a_recup)
print(f"IPI a recup: {ipi_a_recup:.2f}".replace('.', ','))

for row in ws.iter_rows(min_row=2):
    cell_a = row[0].value  # Coluna A (índice 0)
    if cell_a in numeros_procurados:
        valor_coluna_h = row[7].value  # Coluna H (índice 7)
        print(f"Número {cell_a} encontrado: Valor na coluna H = {valor_coluna_h}")
        
        # Comparar os valores e escrever "OK" ou "Verificar" na coluna I
        try:
            if cell_a == 41:
                try:
                    if abs(valor_coluna_h - ICMS_recup) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: ICMS_recup não definido. Continuando a execução...")
                    row[8].value = "Verificar"
            elif cell_a == 42:
                try:
                    if abs(valor_coluna_h - ipi_a_recup) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: IPI_a_recup não definido. Continuando a execução...")
                    row[8].value = "Verificar"
            elif cell_a == 46:
                try:
                    if abs(valor_coluna_h - total_pis_recup) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: total_pis_recup não definido. Continuando a execução...")
                    row[8].value = "Verificar"
            elif cell_a == 47:
                try:
                    if abs(valor_coluna_h - total_cofins_recup) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: total_cofins_recup não definido. Continuando a execução...")
                    row[8].value = "Verificar"
            elif cell_a == 2654:
                try:
                    if abs(valor_coluna_h - total_icms_cp_recup) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: total_icms_cp_recup não definido. Continuando a execução...")
                    row[8].value = "Verificar"
            elif cell_a == 617:
                try:
                    if abs(valor_coluna_h - csrf_retido) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: csrf_retido não definido. Continuando a execução...")
            elif cell_a == 185:
                try:
                    if abs(valor_coluna_h - irrf_retido) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: irrf_retido não definido. Continuando a execução...") 
            elif cell_a == 2707:
                try:
                    if abs(valor_coluna_h - inss_retido) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: inss_retido não definido. Continuando a execução...")
            elif cell_a == 186:
                try:
                    if abs(valor_coluna_h - iss_retido) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: iss_retido não definido. Continuando a execução...")
            elif cell_a == 197:
                try:
                    if abs(valor_coluna_h - total_pis_recolher) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: pis_recolher não definido. Continuando a execução...")
            elif cell_a == 196:
                try:
                    if abs(valor_coluna_h - total_cofins_recolher) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: cofins_recolher não definido. Continuando a execução")
            elif cell_a == 198:
                try:
                    if abs(valor_coluna_h - total_ipi_recolher) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: ipi_recolher não definido. Continuando a execução...")
            elif cell_a == 195:
                try:
                    if abs(valor_coluna_h - total_icms_a_recolher) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: icms_a_recolher não definido. Continuando a execução...")
        except TypeError:
            continue

# Salvar as alterações de volta no arquivo Excel
wb.save(excel_path)
