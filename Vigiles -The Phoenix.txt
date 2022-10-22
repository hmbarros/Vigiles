import pandas as pd
import pyautogui as pg
import time
from datetime import datetime as dt
import os
from tkinter.filedialog import askopenfilename

def main(location):
    data, location = get_archive(location)
    header_check(data, location)
    existencia_check(data, location)
    CPF_check(data, location)
    date_check(data,location)
    workload_check(data, location)
    replication_check(data, location)
    auto_fill(data)

def get_archive(arquivo):
    origem = arquivo
    if arquivo == 0:
        origem = askopenfilename()
    if origem[-5:] != '.xlsx':
        print("\nFormato não suportado:\n")
        main(0)
    else:
        x = pd.read_excel(origem, engine = 'openpyxl', dtype = str)
        return x, origem

def header_check(x, origem):
    error = []
    header = ["Endereço de e-mail:",
              "Nome Completo:",
              "Data de Nascimento:",
              "CPF:",
              "Telefone Celular:",
              "Telefone Fixo:",
              "Data de Formatura do Treinamento:",
              "Carga Horária:"]
    for i in header:
        try:
            x[i]
        except:
            error.append(i)
    if len(error) != 0:
        print("Divergência nos seguintes itens do cabeçalho:")
        for i in error:
            print(i)
        print("\nCorrija o cabeçalho e aperte enter.")
        os.system('PAUSE')
        print()
        main(origem)
    else:
        print("----------------------------------------")
        print("Cabeçalho: Válido!")
        print("----------------------------------------")

def existencia_check(x, origem):
    obrigatorio = ["CPF:",
                   "Nome Completo:",
                   "Data de Nascimento:",
                   "Data de Formatura do Treinamento:",
                   "Carga Horária:"]
    
    n = 0 #Contador de erro
    for j in obrigatorio:
        error = []
        for i in range(0,len(x["Nome Completo:"])):
            if str(x[j][i]) == "nan":
                error.append(i)
        if len(error) != 0:
            print(str(j)[:-1],"inexistente para os seguintes brigadistas:")
            for i in error:
                print(x["Nome Completo:"][i]) 
            n += 1
            print("----------------------------------------")
    if n != 0:
        print("\nCorrija e pressione'ENTER'")
        os.system('PAUSE')
        print()
        main(origem)
    else:
        print("Existência de Dados Obrigatórios: Válido!")
        print("----------------------------------------")

def CPF_check(x, origem):
    error = []
    count = 0
    for CPF in x["CPF:"]:
        if len(str(CPF)) != 11:
            error.append(count)
        else:
            a = 0
            for i in range(1,10):
                a = a + i*int(CPF[i-1])
            resto = a%11
            if resto == 10:
                resto = 0
            if resto != int(CPF[-2]):
                error.append(count)
            else:
                a = 0
                for i in range(0,10):
                    a = a + i*int(CPF[i])
                resto = a%11
                if resto == 10:
                    resto = 0
                if resto != int(CPF[-1]):
                    error.append(count)
        count += 1
    if len(error) != 0:
        print("CPF divergente para os seguintes brigadistas:")
        for i in error:
            print(x["Nome Completo:"][i])
        print("Corrija os 'CPF's, salve o arquivo e tente novamente.")
        os.system('PAUSE')
        print()
        main(origem)
    else:
        print("Dados de CPF: Válido!")
        print("----------------------------------------")

def date_check(x, origem):
    error = []
    D = ['Data de Nascimento:','Data de Formatura do Treinamento:']
    hoje = dt.today().strftime('%Y%m%d')
    for i in D: #Formatação das datas
        x[i] = pd.to_datetime(x[i], errors='coerce')
        x[i] = x[i].dt.strftime('%Y%m%d')
    for i in range(0,len(x['Data de Nascimento:'])):
        if int(hoje) - int(x['Data de Nascimento:'][i]) < 180000:
            error.append(i)
    if len(error) != 0:
        print("\nData de nascimento divergente para os seguintes brigadistas:")
        for i in error:
            print(x["Nome Completo:"][i])
        print("Corrija as datas de nacimento, salve o arquivo e tente novamente.")
        os.system('PAUSE')
        main(origem)
        print()
    else:
        print("Data de Nascimento: Válido!")
        print("----------------------------------------")
    
    error = []
    for i in range(0,len(x['Data de Formatura do Treinamento:'])):
        if 0 < (int(hoje) - int(x['Data de Formatura do Treinamento:'][i])) > 10000:
            error.append(i)
    if len(error) != 0:
        print("\nData de Formatura do Treinamento divergente para os seguintes brigadistas:")
        for i in error:
            print(x["Nome Completo:"][i])
        print("Corrija as Datas de Formatura, salve o arquivo e tente novamente.")
        os.system('PAUSE')
        print()
        main(origem)
    else:
        print("Data de Formatura do Treinamento: Válido!")
        print("----------------------------------------")        

def workload_check(x, origem):
    error = []
    for i in range(0,len(x["Nome Completo:"])):
        if x['Carga Horária:'][i] in ["4h","8h","24h"]:
            pass
        else:
            error.append(i)
    if len(error) != 0:
        print("Carga Horária divergente para os seguintes brigadistas:")
        for i in error:
            print(x["Carga Horária:"][i])
        print("Corrija os dados de Carga Horária, salve o arquivo e tente novamente.")
        os.system('PAUSE')
        main(origem)
        print()
    else:
        print("Carga Horária: Válido!")
        print("----------------------------------------")

def replication_check(x, origem):
    duplicate = {}
    header = ["Endereço de e-mail:",
          "Nome Completo:",
          "CPF:",
          "Telefone Celular:"]
    for i in header:
        singular = []
        error = []
        for j in x[i]:
            if j in singular and str(j) != "nan":
                error.append(j)
            else:
                singular.append(j)
        if len(error) != 0:
            duplicate[i] = error
    if duplicate != {}:
        "----------------------------------------"
        print("Os seguinte dados estão duplicados:", end = "\n\n")
        for i in duplicate:
            print(i)
            for j in duplicate[i]:
                print(j)
            print()
        
        print("Corrija os dados duplicados, salve o arquivo e tente novamente.")
        print("----------------------------------------")
        os.system('PAUSE')
        main(origem)
        print()

    else:
        pass

def auto_fill(x):
    D = ['Data de Nascimento:','Data de Formatura do Treinamento:']
    for i in D: #Formatação das datas
        x[i] = pd.to_datetime(x[i], errors='coerce')
        x[i] = x[i].dt.strftime('%d%m%Y')

    print("ATENÇÃO!!!!")
    print("A PRÓXIMA ETAPA PRECISA QUE:")
    print("O SITE DO BOMBEIRO ESTEJA ABERTO NA PRIMEIRA ABA DO 'ALT' + 'TAB'")
    print("ESTEJA CADASTRADO EXATAMENTE UM INSTRUTOR")
    print("NENHUM BOTÃO ESTEJA SELECIONADO")
    print("CERTIFIQUE-SE QUE OS ITENS ACIMA FORAM CUMPRIDOS E DIGITE 'ENTER'")
    os.system('PAUSE')
    name = ['CPF:',
        'Nome Completo:',
        'Data de Nascimento:',
        'Data de Formatura do Treinamento:',
        'Endereço de e-mail:',
        'Endereço de e-mail:',
        'Telefone Fixo:',
        'Telefone Celular:',
        "Carga Horária:"]

    pg.hotkey('alt','tab')
    for i in range(0,len(x["Nome Completo:"])):
        for _ in range(0,11):
            pg.press('tab')
        pg.press('enter')

        for j in name:
            pg.press('tab')
            if j == "Carga Horária:":
                t = ['4h','8h','24h'].index(x['Carga Horária:'][i])
                pg.press('down', presses = t + 1)
                pg.press('tab')                 
            elif str(x[j][i]) != 'nan':
                pg.write(str(x[j][i]))

        pg.press('tab'  )
        pg.press('enter')

        t = 2
        time.sleep(t)   
        pg.press('enter')
        time.sleep(t)
        pg.press('enter')


    pg.hotkey('alt','tab')

main(0)
