"""
Projeto de Automção de envio de E-Mail

By Carlos Henrique Barros Silva Campos
"""
import win32com.client as win32

print('=' * 40)
print('*                                      *')
print('*         Automação de E-mail          *')
print('*                                      *')
print('=' * 40)


# integração com outlook
outlook = win32.Dispatch('outlook.application')

# gerar e-mail
email = outlook.CreateItem(0)


# Configurar e-mail

email.to = "carlos.barros@aedb.br"
email.Subject = "Teste de e-mail automção"
email.HTMLBody = """
<p>Testando envio de E-mail em Python</p>

<p>Olá Carlos Henrique tudo bem... </p>

"""

print('E-mail enviado com sucesso!')
email.Send()