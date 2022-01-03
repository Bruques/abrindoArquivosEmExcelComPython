import pandas as pd
import pywin32 as win32

'''Importando nossa base de dados (tabela no excel)'''

tabelaVendas = pd.read_excel('Vendas.xlsx')

'''Visualizando a base de dados inteira'''
'''Com essa linha de codigo abaixo falamos para o python que queremos que ele nos mostre todas as colunas
que estao presentes na tebela que pedimos que ele abrisse'''

pd.set_option('display.max_columns', None)

#print(tabelaVendas[['ID Loja', 'Valor Final']])

'''Faturamento por loja'''
'''[[Aqui dentro dos dois colchetes fica o que queremos filtrar na nossa variavel, que no caso e a tabela]]'''
'''A funcao groupby() serve para agruparmos uma certa coluna, que no caso e a "ID Loja"'''
'''A funcao .sum() serve para somarmos as lojas'''

faturamento = tabelaVendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
#print(faturamento)

'''Quantidade de produtos vendidos por loja'''
'''
Aqui a inves de selecionarmos o valor final apenas mudamos para filtrar pela quantidade, o resto
do codigo fica a mesma coisa'''

qtdProdutos = tabelaVendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
#print(qtdProdutos)

'''Ticket medio por produto em cada loja (faturamento / quantidade)'''
'''
Sempre que e feito contas entre colunas no final do codigo devemos colocar ".to_frame()"
'''
ticketMedio = (faturamento['Valor Final'] / qtdProdutos['Quantidade']).to_frame()
print(ticketMedio)

'''Enviar um email com o relatorio'''

'''E necessario instalar a biblioteca pywin32 '''

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'digiteoemailaqui@gmail.com'
mail.Subject = 'Messagem ao sujeito'
mail.HTMLBody = f'''
<p>Prezados</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html()} #temos que colocar .to_html para transformar a tabela em HTML

<p>Quantidade de produtos vendidos por loja:</p>
{qtdProdutos.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{ticketMedio.to_html()}

<p>Qualquer dúvida fico à disposição.</p>

'''
mail.Send()