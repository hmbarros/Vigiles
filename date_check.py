from datetime import datetime as dt
import re
import pandas as pd


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
                #os.system('PAUSE')
                #print()
                #main(origem)
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
                #os.system('PAUSE')
                #print()
                #main(origem)
            else:
                print("Datas de Formatura do Treinamento: Válidas!")
                print("-----------------------------------------------")
            

location = "E:/Master Safety/Bombeiro/Modelo Padrão Lista Treinamento Bombeiro - Centro Paula Souza.xlsx"
x = pd.read_excel(location, engine = 'openpyxl', dtype = str)
date_check(x, location)