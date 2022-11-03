import pandas as pd
import pyautogui as pg
import time
from datetime import datetime as dt
import os
from tkinter.filedialog import askopenfilename

print("Iniciando os trabalhos:")
origem = ''
origem = askopenfilename()

if origem == '':
    print("\nNenhum arquivo foi selecionado:\n")
    os.system('PAUSE')
    os._exit(0)

if origem[-5:] != '.xlsx':
    print("\nFormato de arquivo não suportado:\n")
    os.system('PAUSE')
    os._exit(0)


type = {'CPF': str,
        'Nome Completo': str,
        'Endereço de e-mail': str,
        'Endereço de e-mail': str,
        'Telefone Fixo': str,
        'Telefone Celular': str} #Tipos para importação do excel
obg = ['CPF',
       'Nome completo',
       'Data de Nascimento',
       'Data de formatura do treinamento:',
       'Carga Horária']  #Valores de existência obrigatória
D = ['Data de Nascimento','Data de formatura do treinamento:']  #Datas
name = ['CPF',
        'Nome completo',
        'Data de Nascimento',
        'Data de formatura do treinamento:',
        'Endereço de e-mail',
        'Endereço de e-mail',
        'Telefone Fixo',
        'Telefone Celular'] #Índices
o = False
divergente = []

x = pd.read_excel(origem, engine = 'openpyxl', dtype = type)
inst = int(input("Digite o numero de instrutores cadastrados:\n"))
n = len(x['Nome completo'])


for i in D: #Formatação das datas
    x[i] = pd.to_datetime(x[i], errors='coerce')
    x[i] = x[i].dt.strftime('%Y%m%d')

for i in range(0,n): #Teste de Existência
     for j in obg:
        if str(x[j][i]) == 'nan':
            print(j, 'inexistente para', x['Nome completo'][i],'\n')
            if j in D and o == False:
                if int(x[j][i]) > int(dt.today().strftime('%Y%m%d')):
                    print("O treinamento de", x['Nome completo'][i],"ainda não foi concluído. Favor rever os dados\n")
                    o = True

     if int(x['Data de formatura do treinamento:'][i]) + 10000 < int(dt.today().strftime('%Y%m%d')):
        print("O teinamento de", x['Nome completo'][i], "está vencido. Favor rever os dados.\n")
        o = True


     if len(x['CPF'][i]) != 11:
         #print("CPF divergente para os seguintes brigadistas:",x['Nome completo'][i])
         divergente.append(str(x['Nome completo'][i]))
         o = True

if o:
    if len(divergente) != 0:
        print("\nCPF divergente para os seguintes brigadistas:\n")
        for l in divergente:
            print(l)
    os.system('PAUSE')
    os._exit(0)

os.system('PAUSE')

for i in D: #Formatação das datas
    x[i] = pd.to_datetime(x[i], errors='coerce')
    x[i] = x[i].dt.strftime('%d%m%Y')

# Começo do Preenchimento
pg.hotkey('alt','tab')
for i in range(0,n):
    pg.press('tab', presses = 10 + inst)
    pg.press('enter')

    for j in name:
        pg.press('tab')
        if str(x[j][i]) != 'nan':
            pg.write(str(x[j][i]))

    nivel = ['4h','8h','24h']
    if x['Carga Horária'][i] in nivel:
        pg.press('tab')
        t = nivel.index(x['Carga Horária'][i])
        pg.press('down', presses = t + 1)
    else:
        os._exit(0)

    pg.press('tab', presses = 2)
    pg.press('enter')

    t = 2
    time.sleep(t)
    pg.press('enter')
    time.sleep(t)
    pg.press('enter')

pg.hotkey('alt','tab')
print()
print('\nAutopreenchimento realizado!\n')
