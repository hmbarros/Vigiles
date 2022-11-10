import re
import os
from tkinter.filedialog import askopenfilename
import pandas as pd

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
        #print()
        #main(origem)
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
        #print()
        #main(origem)

    else:
        print("Telefones Celulares: Válidos!")
        print("-----------------------------------------------")
        

location = "E:/Master Safety/Bombeiro/Modelo Padrão Lista Treinamento Bombeiro - Centro Paula Souza.xlsx"
x = pd.read_excel(location, engine = 'openpyxl', dtype = str)
CEL_check(x, location)