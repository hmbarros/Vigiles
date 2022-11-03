import re
import os
from tkinter.filedialog import askopenfilename
import pandas as pd


def CPF_check(x):
    count = 0
    error = []
    for CPF in x["CPF:"]:
        conf = re.match("^(\d{3})\.?(\d{3})\.?(\d{3})-?(\d{2})$", CPF)
        data = ""
        if conf:
            for i in conf.groups():
                data += i
            print("Conferindo:", data)
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

    if len(error) != 0:
        print("-----------------------------------------------")
        print("CPF divergente para os seguintes brigadistas:")

        for i in error:
            print(x["Nome Completo:"][i])

        print("-----------------------------------------------")
        print("Corrija os 'CPF's, salve o arquivo e tente novamente.")
        os.system('PAUSE')

x = pd.read_excel("E:/Master Safety/Bombeiro/Modelo Padrão Lista Treinamento Bombeiro (1).xlsx", engine = 'openpyxl', dtype = str)
CPF_check(x)
