import pandas as pd
import win32com.client as win32
email = 'email@example.com'
assunto = 'Resultados de vendas'
try:
    tabela_teste = pd.read_excel('Vendas.xlsx')
    pd.set_option('display.max_columns', None)
    faturamento = tabela_teste[['ID Loja', 'Valor Final']].groupby('ID Loja').sum('Valor Final')
    quantidade = tabela_teste[['ID Loja', 'Quantidade']].groupby('ID Loja').sum('Quantidade')
    tkm = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
    tkm = tkm.rename(columns={0: 'TKM'})
    #abaixo script para enviar email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = assunto
    mail.HTMLbody = f'''
    <p>Prezados, segue abaixo o relatório de vendas,</p>
    <p>Faturamento por loja:</p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
    <p>Quantidade de vendas por loja:</p>
    {quantidade.to_html()}
    <p>Ticket Médio:</p>
    {tkm.to_html(formatters={'TKM': 'R${:,.2f}'.format})}

    <p>Att,</p>
    <p>Weslley</p>
    '''
    mail.Send() #Display() caso queira montar o e-mail mas não enviar diretamente (ideal para consultar se está correto), caso queira enviar, trocar por Send().
except FileNotFoundError:
    print('Arquivo não encontrado')
except KeyError:
    print('Arquivo com chaves inválidas')
else:
    print('E-mail enviado com sucesso')
