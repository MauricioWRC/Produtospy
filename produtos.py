import pandas as pd
import time
from random import randint


rand = randint(1,1000)
print(rand)

Excel = str(input(r"""
    CRIE UMA PLAILHA EM EXCEL
    E COLE AQUI O SEU CAMINHO
    Ex:('C:\Users\joao\Desktop\Produtos')
    DEPOIS COLOQUE UM CONTRA BARRA '\'
    JUNTO COM O NOME DE SUA PLANILA;
    CASO A PLANILHA ESTIVER NA MESMA PASTA
    QUE ESTE PROGRAMA, APENAS ESCREVA O NOME DA PLANILHA
    --> """))

tabela = pd.read_excel(Excel + '.xlsx')
Produto = ''
rand = str(randint(1,100))
salvo = f'{Excel + rand}.xlsx'
Arquivo = salvo.split('\\')
if len(tabela) == 0:
    tabela['Produto'] = ""
    tabela['Estoque'] = ""
    tabela['Valor de Compra'] = ""
    tabela['Valor de Venda'] = ""
    tabela['Quantidade Vendida'] = ""
while Produto != '@_@':
    print(f'''
            ATENÇÃO!!
    QUANDO FINALIZADO O ARQUIVO 
    SERÁ SALVO COM {Arquivo[-1]} NO FINAL
    ''')
    cont = len(tabela['Produto'])
    print(tabela)
    print('Para finalizar os cadastros digite @_@ como nome de produto')
    Produto = str(input('Produto: '))
    if Produto != '@_@':
        Estoque = int(input('Estoque (Digitar apenas numeros): '))
        ValorCompra = float(input('Valor da Compra.(utilize ".") R$:'))
        ValorVenda = float(input('Valor de venda. (utilize ".") R$:'))
        QuantidadeVendida = int(input('Quantidade Vendida (Digitar apenas numeros): '))
        tabela.loc[cont,'Produto'] = Produto
        tabela.loc[cont,'Estoque'] = Estoque
        tabela.loc[cont,'Valor de Compra'] = ValorCompra
        tabela.loc[cont,'Valor de Venda'] = ValorVenda
        tabela.loc[cont,'Quantidade Vendida'] = QuantidadeVendida
        cont = cont + 1
        print(tabela)
        time.sleep(2)
        print('--/--'*4 +'Proximo Produto'+ '--/--'*4 )
    else:
        print('--->Calculando Tabela e Finalizando<---')
        tabela = tabela.dropna(how='all', axis=1)
        tabela.to_excel(Excel + rand + '.xlsx', index= False)
        time.sleep(2)
        print(tabela)
        print(f'''
{'-/-'*5} Programa finalizado {'-/-'*5}
Salvo em: {salvo}''')
        

        