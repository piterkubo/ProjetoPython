import pandas as pd
import win32com.client as win32
import pyautogui as auto
import pyperclip as pyper
import time

auto.hotkey('winleft','r')
time.sleep(1)

caminho = r'C:\Base_clientes\NIVER.xlsx'

pyper.copy(caminho)

auto.hotkey('ctrl','v')
time.sleep(1)



auto.press('enter')
time.sleep(5)

auto.hotkey('alt', 'F4')
time.sleep(1)

auto.press('enter')
time.sleep(1)


df = pd.read_excel(r'C:\Base_clientes\NIVER.xlsx')

dados_funcionarios = df.query('mes_admissao == mes_hoje and dia_admissao == dia_hoje')[['funcionarios', 'data_admissao' ,'tempo_casa']]


if dados_funcionarios['funcionarios'].empty:
        print('Dados vazio !!!')
        exit()



else:


        print(dados_funcionarios)



        dados_funcionarios.to_excel(r'C:\Base_clientes\tabela_atualizada.xlsx')

        outlook = win32.Dispatch('outlook.application')

        email = outlook.CreateItem(0)

        email.To = 'cristiane@guitta.com.br;karolyne@guitta.com.br'
        email.Bcc ='piterkubo@guitta.com.br'

        email.Subject = 'Relação de tempo de casa dos funcionários'
        email.HTMLBody = f'''

        <p> Prezados ! </p>

        <p> Segue no anexo, planilha dos aniversariantes de tempo de casa.</p>

        <p> Qualquer esclarecimento, por gentileza entrar em contato.

        <p> Atenciosamente </3>

        '''
        anexo = (r'C:\Base_clientes\tabela_atualizada.xlsx')

        email.Attachments.Add(anexo)

        email.Send()
