from datetime import datetime

# Data de referência

data = 30/09/2023
data_referencia = datetime.strptime(str(data), '%d/%m/%Y')

# Data atual
data_atual = datetime.now()

# Verificar se já passou um ano
if (data_atual - data_referencia).days >= 365:
    print("Já passou um ano desde 01/01/2023.")
else:
    print("Ainda não passou um ano desde 08/06/2023.")