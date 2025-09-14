import win32com.client as win32
from openpyxl import load_workbook 
import os

outlook = win32.Dispatch('Outlook.Application')
caminho_arquivo = r'./dados/ListaEmail.xlsx'

planilha_aberta = load_workbook(filename=caminho_arquivo)

sheet_selecionada = planilha_aberta['Dados']

for linha in range(2, len(sheet_selecionada['A']) + 1):

    nome = sheet_selecionada['A%s' % linha].value
    nome_completo = sheet_selecionada['B%s' % linha].value
    email = sheet_selecionada['C%s' % linha].value

    emailOutlook = outlook.CreateItem(0)
    emailOutlook.To = email
    emailOutlook.Subject = f'Lista de vendas {nome_completo}'

    emailOutlook.HTMLBody = f"""
    <p>Boa noite <b>{nome}</b>. </p>
    <p>Segue o relatório com suas vendas</p>
    <p>Atenciosamente Thamires Sarges</p>
    """

    pasta_relatorios = r'./relatorios/'

    nome_arquivo = nome_completo.replace(' ', '_') + '.xlsx'
    anexoEmail = os.path.abspath(os.path.join(pasta_relatorios, nome_arquivo))

    emailOutlook.Attachments.Add(anexoEmail)

    emailOutlook.save() 