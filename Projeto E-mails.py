#!/usr/bin/env python
# coding: utf-8

# # Código para envio de e-mails automático
# 
# ### Descrição
# 
# Todo dia, você, a equipe ou até mesmo um programa, gera um report diferente para cada área da empresa:
# - Financeiro
# - Logística
# - Manutenção
# - Marketing
# - Operações
# - Produção
# - Vendas
# 
# Cada um desses reports deve ser enviado por e-mail para o Gerente de cada Área.
# 
# Crie um programa que faça isso automaticamente. A relação de Gerentes (com seus respectivos e-mails) e áreas está no arquivo 'Enviar E-mails.xlsx'.
# 

# In[1]:


import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
import pandas as pd

gerentes_df = pd.read_excel('Enviar E-mails.xlsx')


for i, email in enumerate(gerentes_df['E-mail']):
    #Pegando da linha i (que vai dar pelo enumerate) e coluna 'Gerente'
    gerente = gerentes_df.loc[i, 'Gerente']
    area = gerentes_df.loc[i, 'Relatório']
    
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.CC = 'seuemail@gmail.com'
    mail.Subject = "Relatório de {}".format(area)
    #Três aspas para escrever o texto em mais de uma linha
    mail.Body = f'''
    Prezado, {gerente}. Segue anexo do relatório de {area} atualizado.
    Qualquer dúvida, estou à disposição.
    Abraços,
    
    '''

    # Anexos (pode colocar quantos quiser, mas preencha o seu caminho):
    attachment  = r'C:\Users\Área de Trabalho\#Python\Projetos Consultoria\Projeto E-mails\{}.xlsx'.format(area)
    mail.Attachments.Add(attachment)

    mail.Send()

