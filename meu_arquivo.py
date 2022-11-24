import pandas as pd
import win32com.client as win32

faturamento = 0
tabela_vendas = pd.read_excel('Vendas.xlsx')

pd.set_option('display.max_columns', None)

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

print(faturamento)


quantidade = 0

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()


print(quantidade)

print('-' * 100)
##Ticket médio porm produto em cada loja. Faturamento divididos pela quantidade de vendas

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'brunaravelli13@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p> 

<p>Segue o Relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p> 
{quantidade.to_html()}

<p>Ticket médio dos produtos em casa loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}


<p>Qualquer dúvida estou à disposição</p>

<p>Bruna Ravelli</p>
<p>Desenvolvedora de software Júnior</p>
'''

mail.Send()
print('E-mail enviado')
