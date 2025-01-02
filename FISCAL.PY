from bs4 import BeautifulSoup
import pandas as pd
import openpyxl

url = r"C:\Users\contabil18\Downloads\E008801_Relatório de retenções_01112024_30112024.htm"

with open(url, 'r', encoding='utf-8') as file:
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
