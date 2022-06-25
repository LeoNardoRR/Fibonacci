#Nome: Leonardo Rodrigues Ribeiro
#RA: 0040972123041

'''

Um matemático está precisando de um programa que calcule a SEQUÊNCIA DE FIBONACCI. 
SEQUÊNCIA DE FIBONACCI é uma sucessão de números que, misteriosamente, aparece em muitos fenômenos da 
natureza. Descrita no final do século 12 pelo italiano Leonardo Fibonacci, ela é infinita e começa com 0 e 1. Os 
números seguintes são sempre a soma dos dois números anteriores. Portanto, depois de 0 e 1, vêm 1, 2, 3, 5, 8, 13, 21, 
34… 
O matemático está solicitando a você que faça um programa que gerará uma sequência de Fibonacci num arquivo 
Excel, bastando informar a quantidade de elementos que o usuário deseja que o programa gere. 
Então, o seu programa deverá: 
1) Primeiramente solicitar ao usuário que digite a quantidade de elementos (n) que deseja gerar. 
2) Preencher a sequência de Fibonacci num arquivo Excel que deverá ser criado pelo programa. Os dois primeiros 
elementos (0 e 1) deverão ser gravados nas células A1 e B1, e na sequência das colunas e linha 1, os demais elementos 
até atingir a quantidade de n elementos a serem gravados no arquivo. Veja o exemplo abaixo, onde foi escolhido a 
quantidade de 7 elementos: 

'''

from openpyxl import Workbook
def SolicitaUsuario():
      print ('Olá, bem vind(o)a!')
      nome = str(input('Digite seu nome?\nNome: '))
      print ('É um prazer te conhecer', nome ,'!')

SolicitaUsuario()

def Fibonacci():

    print('-'*30)
    print('Sequencia de Fibonacci')
    print('-'*30)   
    elemento = int(input('Agora, digite a quantidade de elementos que deseja gerar: '))

    termo1 = 0
    termo2 = 1
    lista = [termo1, termo2]
    print('~'*30)
    print('{} -> {}'.format(termo1, termo2), end='')
    cont = 3
    while cont <= elemento:
        termo3 = termo1 + termo2
        lista.append(termo3)
        print('-> {}'.format(termo3), end='')
        termo1 = termo2
        termo2 = termo3
        cont +=1
    print(' -> Sequência Finalizada')
    return lista

def CalculaFibo(lista):
    cont = 0
    for i in lista:
        cont += i
    return cont


def Gravar(lista):
    wb = Workbook()
    planilha_excel = wb.active
    planilha_excel.append(lista)

    total = CalculaFibo(lista)
    planilha_excel['A3'] = total

    wb.save('Fibonacci.xlsx')
    print('documento salvo com sucesso!')

lista = Fibonacci()
Gravar(lista)
