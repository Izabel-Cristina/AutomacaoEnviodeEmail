import win32com.client as win32
import pandas as pd
import datetime as dt

#Lendo arquivo#
tabela = pd.read_excel('Contas a Receber.xlsx')
display(tabela)
tabela.info()

#Verificar a data de hoje#
hoje = dt.datetime.now()
print(hoje)

#Coletar dados necessários#
tabela_devedores = tabela.loc[tabela['Status']=='Em aberto']
display(tabela_devedores)
tabela_devedores = tabela_devedores.loc[tabela_devedores['Data Prevista para pagamento']<hoje]
display(tabela_devedores)

#Configurar email#
#Integralizar Python com Outlook#
outlook = win32.Dispatch('Outlook.Application')
emissor = outlook.session.Accounts['email@gmail.com']
mensagem = outlook.CreatItem(0)
mensagem.Display()
mensagem.To = ''
mensagem.Subject = 'Pagamento em aberto'
mensagem.Body= '''
Bom dia, 
Ainda não identificamos o pagamento do boleto, se já foi realizado o pagamento por gentileza nos envie o comprovante de pagameto.
'''
mensagem._oleobj_.Invoke(*(64209, 0,8,0, emissor))
mensagem.Save()
mensagem.Send()

#Criar uma lista#
dados = tabela_devedores[['Valor em aberto', 'Data Prevista para pagamento', 'Email', 'NF'].values.tolist()]

#Enviando email para clientes#
for dado in dados:
    destinatario = dados[2]
    valor = dados[0]
    prazo = dados[1]
    prazo = prazo.str("%d/%m/%y")
    nf = dados[3]
    mensagem = outlook.CreatItem(0)
    mensagem.Display()
    mensagem.To = ''
    mensagem.Subject = 'Pagamento em aberto'
    mensagem.Body= f'''
    Bom dia, 
    Ainda não identificamos o pagamento da nota {nf},que venceu no dia {prazo}, no valor de R$ {valor } 
    se já foi realizado o pagamento por gentileza nos envie o comprovante de pagameto.
    
    '''
    mensagem._oleobj_.Invoke(*(64209, 0,8,0, emissor))
    mensagem.Save()
    mensagem.Send()
    

