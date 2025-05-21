#LÓGICA DO PROJETO

#Importar a base de dado

#Visualizar a base de dados 

# Faturamento por loja

#Quantidade de produtos vendidos por loja

#Ticket médio por produto em cada loja

#Enviar um email com o relatório


#Importar a base de dado
import pandas as pd

tabela_vendas = pd.read_excel("Vendas.xlsx")

#print(tabela_vendas)

#Visualizar a base de dados 
pd.set_option("display.max_columns", None)
print(tabela_vendas)

# Faturamento por loja

faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
print(faturamento)
print("-" * 50)
#Quantidade de produtos vendidos por loja

quantidade = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()
print(quantidade)
print("-" * 50)
#Ticket médio por produto em cada loja

ticket_medio = (faturamento["Valor Final"] / quantidade["Quantidade"]).to_frame()
print(ticket_medio)
print("-" * 50)

#Enviar um email com o relatório
#instalar pywin32

import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'camilaaugustasnt@outlook.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = '''Prezados, 

Segue o relatório de vendas por cada loja.

Faturamento:
{}

Quantidade vendida:
{}

Ticket médio do produtos em cada loja:
{}

Qualquer dúvuda, estou a disposição.

Att.,

Camila.

'''
mail.Send()

print("Email enviado")
