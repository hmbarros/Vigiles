import re
import pandas as pd

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

        print("\nCaso o Brigadista não possua e-mail deixe o campo em branco.")
        print("Corrija os 'E-mail's, salve o arquivo e tente novamente.")
        print("-----------------------------------------------")
    
    else:
        print("Dados de E-mail: Válido!")
        print("----------------------------------------")

location = "E:/Master Safety/Bombeiro/Modelo Padrão Lista Treinamento Bombeiro - Centro Paula Souza.xlsx"
x = pd.read_excel(location, engine = 'openpyxl', dtype = str)

email_check(x,location)