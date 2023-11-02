import pandas as pd
import smtplib
import email.message

# importa base de dados
tabela_vendas = pd.read_excel('vendas.xlsx')

# exibe base de dados
pd.set_option('display.max_columns', None) # Não limita um valor maximo de colunas
print(tabela_vendas)

print('-' * 50)
# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum() # filtra a tabela, agrupa lojas que se repetem na lista e soma valor final
print(faturamento)

print('-' * 50)
# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# Ticket médio por cada loja (faturamento/quantidade de vendas)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)

# Envia um e-mail com relatório
def enviar_email():
    corpo_email = """ 
    <p>Prezado amigo, segue o relatório de vendas por loja.</p>
    <p>Faturamento:</p>
    <p>Quantidade Vendida:</p>
    <p>Ticket médio de produtos por loja:</p>
    <p>Qualquer dúvida estou á disposição.</p>
    <p>Att,</p>
    <p>Python</p>
    """
    msg = email.message.Message()
    msg['Subject'] = "Assunto"
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'],)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')