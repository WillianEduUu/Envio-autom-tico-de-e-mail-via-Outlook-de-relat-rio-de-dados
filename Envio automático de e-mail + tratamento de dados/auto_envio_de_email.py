import pandas as pd
import win32com.client as win32

df = pd.read_excel("Vendas_novas_lojas.xlsx")  
print('-' * 50)

######-----------------MOSTRAR TODAS AS COLUNAS-----------------######

pd.set_option('display.max_columns', None)
print(df)

######----------------------------------------------------------######

faturamento = df[['ID Loja','Valor Final']].groupby('ID Loja').sum() # CÁLCULO DO FATURAMENTO
quantidade = df[['ID Loja', 'Quantidade']].groupby('ID Loja').sum() # QUANTIDADE DE PRODUTOS VENDIDOS POR LOJA
ticket_medio = (faturamento['Valor Final']/quantidade['Quantidade']).to_frame()# TICKET MÉDIO POR PRODUTO (faturamento / quantidade)
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'}) # NOMEANDO A COLUNA QUE TEM O VALOR DO TICKET MÉDIO

print(faturamento)
print('-' * 50)
print(quantidade)
print('-' * 50)
print(ticket_medio)
print('-' * 50)

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email.adress@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Lira</p>
'''

mail.Send()

print('Email Enviado')
