"""
Projeto de Automção de envio de E-Mail

By Carlos Henrique Barros Silva Campos
"""
import win32com.client as win32
import time

# integração com outlook
outlook = win32.Dispatch('outlook.application')

# gerar e-mail
email = outlook.CreateItem(0)


def menu():
    print('=' * 40)
    print('*         Automação de E-mail          *')
    print('*                Menu:                 *')
    print('=' * 40)
    print('    (1 - Mandar apenas um E-mail.')
    print('    (2 - Usar uma lista de E-mail. ')
    print('    (3 - Sair.')
    opc = int(input('Informe a opção:'))
    if opc == 1:
        pessoa = input('Para:')
        titulo = input('Título:')
        msg = input('Msg:')
        configuraEmail(pessoa, titulo, msg)
    elif opc == 2:
        lista_de_email()
    elif opc == 3:
        exit(0)
    else:
        print('Opção incorreta!')
        menu()


def lista_de_email():
    lista = input('Lista de E-mail:')
    with open(lista) as lst:
        lista = lst.read()
        email.to = lista
        titulo = input('Título:')
        msg = input('Msg:')
        email.Subject = titulo
        email.HTMLBody = f"""<p>{msg}</p>"""


def configuraEmail(Email, titulo_msg, corpo_msg):
    email.to = f"{Email};"
    email.Subject = titulo_msg
    email.HTMLBody = f"""<p>{corpo_msg}</p>"""

    # anexo = "C://Users/Root/Desktop/Orçamento.xlsx"
    # email.Attachments.Add(anexo)


if __name__ == '__main__':
    menu()
    print('E-mail enviado com sucesso!')
    email.Send()
