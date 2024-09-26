import os 
import pandas as pd
import win32com.client as win32 
from datetime import datetime


caminho = "bases/" 
arquivos = os.listdir(caminho) 

tabela_consolidada = pd.DataFrame() 

for nome in arquivos:
    tabela_venda = pd.read_csv(os.path.join(caminho, nome))

    # tratamento de dados 

    # formatando data
    tabela_venda["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_venda["Data de Venda"], unit="d")

    tabela_consolidada = pd.concat([tabela_consolidada, tabela_venda]) 


# ordenando datas
tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda") 

# ordenando os indices
tabela_consolidada = tabela_consolidada.reset_index(drop=True) 

# Gerando planilha em excel
tabela_consolidada.to_excel("Vendas.xlsx", index=False) 

outlook = win32.Dispatch('outlook.application') 
email = outlook.CreateItem(0)
email.To = "rafael.alves@labcor.com.br"
data_atual = datetime.today().strftime("%d/%m/%Y") 
email.Subject = f"Relatório de vendas - {data_atual}" 
email.body = f"""
Prezados,

Segue em anexo o relatório de Vendas de {data_atual} atualizado. 

Qualquer dúvida estou a disposição.

Atenciosamente,
Rafael Alves 
"""

caminho_codigo = os.getcwd()
anexo = os.path.join(caminho_codigo, "Vendas.xlsx") 

email.Attachments.Add(anexo)

email.Send()

