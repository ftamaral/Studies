#-------IMPORTAÇÃO-----------
import pandas as pd


#-------VARIÁVEIS------------
datf1 = pd.read_excel('C:\\Vscode\\caminho\\sua_planilha.xlsx')
datf2 = pd.read_excel('C:\\Vscode\\caminho\\sua_planilha.xlsx')

# variável que mantem a mesclagem das planilhas
datf3 = pd.merge(datf1,datf2,how='outer').drop_duplicates(subset=['Usuario'],keep='first')
datf3.to_excel('C:\\Vscode\\bitlocker\\novo.xlsx', index=False)
