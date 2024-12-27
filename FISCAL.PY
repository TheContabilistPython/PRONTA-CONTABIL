from bs4 import BeautifulSoup
import pandas as pd

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

for td in linha_completa.find_all('td'):
    print(td.get_text(strip=True))  # Remove espaços extras
