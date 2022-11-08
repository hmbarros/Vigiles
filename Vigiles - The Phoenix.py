import pandas as pd
import pyautogui as pg
import time
from datetime import datetime as dt
import os
from tkinter.filedialog import askopenfilename
import re

def main(location):
    data, location = get_archive(location)
    header_check(data, location)
    existencia_check(data, location)
    email_check(data, location)
    CPF_check(data, location)
    date_check(data, location)
    CEL_check(data, location)
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
        print("----------------------------------------")
        print("Divergência nos seguintes itens do cabeçalho:\n")
        for i in error:
            print(i)
        print("\nCorrija o cabeçalho, salve o arquivo e aperte enter.")
        print("----------------------------------------")
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
            print(str(j)[:-1],"inexistente para os seguintes brigadistas:\n")
            for i in error:
                print(x["Nome Completo:"][i]) 
            n += 1

    if n != 0:
        print("\nCorrija os dados inexistentes, salve o arquivo e pressione'ENTER'")
        os.system('PAUSE')
        print()
        main(origem)
    else:
        print("Existência de Dados Obrigatórios: Válido!")
        print("----------------------------------------")

def CPF_check(x, origem):
    count = 0
    error = []
    for CPF in x["CPF:"]:
        conf = re.match("^(\d{3})\.?(\d{3})\.?(\d{3})-?(\d{2})$", CPF)
        data = ""
        if conf:
            for i in conf.groups():
                data += i
            #print("Conferindo:", data)
            a = 0
            for i in range(1,10):
                a = a + i*int(data[i-1])
            resto = a%11
            if resto == 10:
                resto = 0
            if resto != int(data[-2]):
                error.append(count)
            else:
                a = 0
                for i in range(0,10):
                    a = a + i*int(data[i])
                resto = a%11
                if resto == 10:
                    resto = 0
                if resto != int(data[-1]):
                    error.append(count)
        else:
            error.append(count)
        count += 1

    if len(error) > 0:
        print("CPF divergente para os seguintes brigadistas:\n")

        for i in error:
            print(x["Nome Completo:"][i])

        print()
        print("Corrija os 'CPF's, salve o arquivo e tente novamente.")
        os.system('PAUSE')
        print()
        main(origem)

    else:
        print("Dados de CPF: Válido!")
        print("----------------------------------------")

def email_check(x, origem):
    count = 0
    error = []
    for email in x["Endereço de e-mail:"]:
        if str(email) == "nan":
            pass
        else:
            conf = re.match("^[a-zA-Z0-9.!#$%&'*+\/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$", email)
            if conf:
                pass
                #print("conferido", email)
            else:
                error.append(count)
        count += 1
    
    if len(error) != 0:
        print("E-mail divergente para os seguintes brigadistas:\n")

        for i in error:
            print(x["Nome Completo:"][i])

        print("\nCaso o(a) brigadista não possua e-mail deixe o campo em branco.")
        print("Corrija os 'E-mail's, salve o arquivo e tente novamente.")
        print("-----------------------------------------------")
        os.system('PAUSE')
        print()
        main(origem)

    else:
        print("Dados de E-mail: Válido!")
        print("----------------------------------------")

def date_check(x, origem):
    hoje = dt.today()#.strftime('%Y%m%d')
    D = ['Data de Nascimento:','Data de Formatura do Treinamento:']
    
    for i in D: #Formatação das datas
        x[i] = pd.to_datetime(x[i], errors='coerce')
        error = []

        if i == 'Data de Nascimento:':
            count = 0
            majority = []
            for data in x[i]:
                try:
                    days = (hoje-data).days
                    if days < 6570:
                        majority.append(count)                        
                except:
                    error.append(count)
                count += 1

            if len(error) > 0:
                print("Data de Nascimento inválida para os seguintes brigadistas:\n")
                for j in error:
                    print(x["Nome Completo:"][j])
                print("--------------------------------------------------------------")
            if len(majority) > 0:
                print("Os seguintes brigadistas são menores de idade e devem ser removidos da lista:\n")
                for j in majority:
                    print(x["Nome Completo:"][j])
                print("--------------------------------------------------------------")
            if len(error) > 0 or len(majority) > 0:
                print("Corrija as 'Datas de Nascimento', salve o arquivo e tente novamente.")
                print("-----------------------------------------------")
                os.system('PAUSE')
                print()
                main(origem)
            else:
                print("Datas de Nascimento: Válidas!")
                print("-----------------------------------------------")
        else:
            count = 0
            overdue = []
            for data in x[i]:
                try:
                    days = (hoje-data).days
                    if days > 365:
                        overdue.append(count)
                    if days < 0:
                        overdue.append(count)
                except:
                    error.append(count)
                count += 1

            if len(error) > 0:
                print("Data de Formatura do Treinamento inválida para os seguintes brigadistas:\n")
                for j in error:
                    print(x["Nome Completo:"][j])
                print("--------------------------------------------------------------")
            if len(overdue) > 0:
                print("Treinamento vencido para os seguintes brigadistas:\n")
                for j in overdue:
                    print(x["Nome Completo:"][j])
                print("--------------------------------------------------------------")
            if len(error) > 0 or len(overdue) > 0:
                print("Corrija as 'Datas de Formatura do Treinamento', salve o arquivo e tente novamente.")
                print("-----------------------------------------------")
                os.system('PAUSE')
                print()
                main(origem)
            else:
                print("Datas de Formatura do Treinamento: Válidas!")
                print("-----------------------------------------------")

def CEL_check(x, origem):
    error = []
    count = 0
    for TEL in x["Telefone Fixo:"]:
        num = 0
        for i in str(TEL):
            try:
                j = int(i)
                num += 1
            except:
                pass
        if num == 10 or num == 0:
            pass
        else:
            error.append(count)
        count += 1

    if len(error) != 0:
        print("-----------------------------------------------")
        print("Telefone Fixo divergente para os seguintes brigadistas:\n")

        for i in error:
            print(x["Nome Completo:"][i])

        
        print("\nCaso o brigadista não possua telefone deixe o campo em branco")
        print("Corrija os 'Telefones Fixos', salve o arquivo e tente novamente.")
        print("-----------------------------------------------")
        os.system('PAUSE')
        print()
        main(origem)
    
    else:
        print("Telefones Fixos: Válidos!")
        print("-----------------------------------------------")

    count = 0
    error = []
    for CEL in x["Telefone Celular:"]:
        num = 0
        for i in str(CEL):
            try:
                j = int(i)
                num += 1
            except:
                pass
        if num == 11 or num == 0:
            pass
        else:
            error.append(count)
        count += 1

    if len(error) != 0:
        print("Telefone Celular divergente para os seguintes brigadistas:\n")

        for i in error:
            print(x["Nome Completo:"][i])

        print("\nCaso o brigadista não possua Telefone Celular deixe o campo em branco")
        print("Corrija os 'Telefones Celulares', salve o arquivo e tente novamente.")
        print("-----------------------------------------------")
        os.system('PAUSE')
        print()
        main(origem)

    else:
        print("Telefones Celulares: Válidos!")
        print("-----------------------------------------------")

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
            print(i, "\n")
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

    print("ATENÇÃO!!!!\n")
    print("A PRÓXIMA ETAPA PRECISA QUE:")
    print("O SITE DO BOMBEIRO ESTEJA ABERTO NA PRIMEIRA ABA DO 'ALT' + 'TAB'")
    print("ESTEJA CADASTRADO EXATAMENTE UM INSTRUTOR")
    print("NENHUM BOTÃO ESTEJA SELECIONADO")
    print("CERTIFIQUE-SE QUE OS ITENS ACIMA FORAM CUMPRIDOS E DIGITE 'ENTER'\n")
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

if __name__ == "__main__":
    main(0)
