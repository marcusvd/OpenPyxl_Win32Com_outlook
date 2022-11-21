from pywin.mfc.object import Object

from plan_data import Data, EmailSender, ClientObj

data = Data('clientes.xlsx')
table = data.load_plan().active
numberRows = table['D2':'D10']

subject = 'No Stop TI - Assistência técnica mensal de computadores OUTUBRO.'
body = """Boa tarde, segue em anexo a nfe da assistência técnica mensal,
prestada durante o período do mês de OUTUBRO 2022. Muito Obrigado!"""

for n in range(len(numberRows)):
    nn = n+2

    client1 = ClientObj(table['D' + str(nn)].value, subject, body, table['E' + str(nn)].value)
    email = EmailSender(client1.client_obj_create().to, client1.client_obj_create().subject,
                        client1.client_obj_create().html_body, client1.client_obj_create().attach)
    email.sender()

