import win32com.client as win32
import pandas as pd
from datetime import date

#Leitura do Arquivo
table = "contas.xlsx"
content = pd.read_excel(table, sheet_name="Contas")
content["Data de vencimento"] = pd.to_datetime(content["Data de vencimento"])
html_conteudo =''
total=0 
for index, row in content.iterrows():
    if row['Data de vencimento'] == pd.to_datetime(date.today()):
        print(f"Conta: {row['Conta']} no valor: {row['Valor']} vence hoje!!")
        html_conteudo += "".join(f"<p>Conta: {row['Conta']} no valor: {row['Valor']} vence hoje!!</p>")
        total += row['Valor']



#Criando integração com o outlook
outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

#Configurar informações do email

email.To = "youremail@hotmail.com"
email.Subject = "Contas a vencer"
email.HTMLBody = f"""
<h3>As seguintes Contas vão vencer hoje:</h3>
{html_conteudo}
<p>Total: {total}</p>
"""

try:
    email.Send()
    print("Email enviado")
except ValueError as e:
    print(e)