import pandas as pd
import win32com.client as win32

# Importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')


# Visualizar a base de dados e ver se precisa de tratamento

pd.set_option('display.max_columns', None )


# Faturamento por Loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

print('Faturamento: \n')
print(faturamento)
print('-' *50)
# Quantidade de produtos vendidos por Loja
print('\nQuantidade: \n')

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' *50)
# Ticket Medio por produto em cada Loja ( Faturamento / quantidade prod Vendidos (por loja) )
print('\nTicket Medio: \n')

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)

# Envio de relatorio por email

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'kimadzn@gmail.com'
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
<p>Tutu</p>
'''

mail.Send()

print('Email Enviado')