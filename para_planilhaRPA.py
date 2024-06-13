import os
import re
import openpyxl
import pandas as pd
import PyPDF2
import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta


def extrair_informacoes_pdf(pdf_path):

    info = {
        'Responsável': 'Alexandre Rego',
        'Data Lançamento': datetime.now().strftime('%d/%m/%Y'),
        'Carteira': 'AIRPROMO',
        'Sistema SAP': 'R/3',
        'Cód. Fornecedor': None,
        'Nome Fornecedor': None,
        'CNPJ Fornecedor': None,
        'Tipo': None,
        'NF': None,
        'Série': None,
        'Emissão': None,
        'Valor Total': None,
        'Vencimento': None,
        'Centro': None,
        'Pedido': None
    }

    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        num_pages = len(pdf_reader.pages)

        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            page_text = page.extract_text()
            
            match_data_emissao = re.search(
                r'DATA DA EMISSÃO: (\d{2}/\d{2}/\d{4})', page_text)
            if match_data_emissao:
                info['Emissão'] = match_data_emissao.group(1)

            match_valor_total = re.search(
                r'VALOR TOTAL DOS PRODUTOS: (R\$\s*\d+\.\d+,\d+\.\d+)', page_text)
            if match_valor_total:
                info['Valor Total'] = match_valor_total.group(1)

            match_serie = re.search(r'SÉRIE: (\d+)', page_text)
            if match_serie:
                info['Série'] = match_serie.group(1)

    return info

diretorio_pdf = r'C:\Users\alexa\Desktop\Interfaceairp'

planilha_excel = 'sua_planilha.xlsx'

if not os.path.exists(planilha_excel):
    df = pd.DataFrame(columns=['Responsável', 'Data Lançamento', 'Carteira', 'Sistema SAP', 'Cód. Fornecedor', 'Nome Fornecedor',
                      'CNPJ Fornecedor', 'Tipo', 'NF', 'Série', 'Emissão', 'Valor Total', 'Vencimento', 'Centro', 'Pedido'])
else:
    df = pd.read_excel(planilha_excel)

for filename in os.listdir(diretorio_pdf):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(diretorio_pdf, filename)
        info = extrair_informacoes_pdf(pdf_path)

        termos_titulo = filename.split(' - ')
        if len(termos_titulo) >= 3:
            info['Cód. Fornecedor'] = termos_titulo[-1].replace('.pdf', '')
            info['Tipo'] = termos_titulo[0]
            info['NF'] = termos_titulo[1]

        df = pd.concat([df, pd.DataFrame([info])], ignore_index=True)

df.to_excel(planilha_excel, index=False)

def calcular_vencimento(data_emissao):
    try:
        data_emissao = datetime.strptime(data_emissao, '%d/%m/%Y')
        data_vencimento = data_emissao + timedelta(days=85)
        return data_vencimento.strftime('%d/%m/%Y')
    except Exception as e:
        print(f"Erro ao calcular a data de vencimento: {str(e)}")
        return None

def extract_info_from_pdf(pdf_file_path):
    try:
        with open(pdf_file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            date_pattern = r'\d{2}/\d{2}/\d{4}'
            date_match = re.search(date_pattern, text)
            date_found = date_match.group() if date_match else None

            order_number_pattern = r'PEDIDO (\d+)'
            order_number_match = re.search(order_number_pattern, text)
            order_number = order_number_match.group(
                1) if order_number_match else None

            valor_total_pattern = r'([\d.,]+)\s+VALOR TOTAL DA NOTA'
            valor_total_match = re.search(valor_total_pattern, text)
            valor_total = valor_total_match.group(
                1) if valor_total_match else None

            cnpj_1_pattern = r'9060019859 (\d{2}[.-]?\d{3}[.-]?\d{3}[/-]?\d{4}[.-]?\d{2})'
            cnpj_1_match = re.search(cnpj_1_pattern, text)
            cnpj_1 = cnpj_1_match.group(1) if cnpj_1_match else None

            cnpj_2_pattern = r'ESTADUALBOTICARIO PRODUTOS DE BELEZA LTDA (\d{2}[.-]?\d{3}[.-]?\d{3}[/-]?\d{4}[.-]?\d{2})'
            cnpj_2_match = re.search(cnpj_2_pattern, text)
            cnpj_2 = cnpj_2_match.group(1) if cnpj_2_match else None

            nome_forn_pattern = r'SÉRIE: 1RECEBEMOS DE ([A-Z\s]+?) EIRELI'
            nome_forn_match = re.search(nome_forn_pattern, text)
            nome_forn = nome_forn_match.group(1) if nome_forn_match else None

            vencimento = calcular_vencimento(date_found)

            return date_found, order_number, valor_total, cnpj_1, cnpj_2, nome_forn, vencimento
    except Exception as e:
        print(f"Erro ao processar {pdf_file_path}: {str(e)}")
        return None, None

pdf_files = []
dates = []
order_numbers = []
valor_total_list = []
cnpj_list = []
nome_forn_list = []
cnpj2_list = []
vencimento_list = []

for filename in os.listdir(diretorio_pdf):
    if filename.endswith('.pdf'):
        pdf_file_path = os.path.join(diretorio_pdf, filename)
        pdf_files.append(filename)
        date, order_number, valor_total, cnpj_1, cnpj_2, nome_forn, vencimento = extract_info_from_pdf(
            pdf_file_path)
        dates.append(date)
        order_numbers.append(order_number)
        valor_total_list.append(valor_total)
        cnpj_list.append(cnpj_1)
        cnpj2_list.append(cnpj_2)
        nome_forn_list.append(nome_forn)
        vencimento_list.append(vencimento)

pdf_df = pd.DataFrame({'Nome do PDF': pdf_files, 'Data de Emissão da NF': dates, 'Número do Pedido': order_numbers,
                      'Valor Total': valor_total_list, 'cnpj_1': cnpj_list, 'nome_forn': nome_forn_list, 'Centro': cnpj2_list, 'vencimento': vencimento_list})

existing_excel_df = pd.read_excel(planilha_excel)

existing_excel_df['Emissão'] = pdf_df['Data de Emissão da NF']
existing_excel_df['Pedido'] = pdf_df['Número do Pedido']
existing_excel_df['Valor Total'] = pdf_df['Valor Total']
existing_excel_df['CNPJ Fornecedor'] = pdf_df['cnpj_1']
existing_excel_df['Nome Fornecedor'] = pdf_df['nome_forn']
existing_excel_df['Centro'] = pdf_df['Centro']
existing_excel_df['Vencimento'] = pdf_df['vencimento']

existing_excel_df.to_excel(planilha_excel, index=False)

def exibir_tabela():

    janela = tk.Tk()
    janela.title("Tabela de Dados")

    tree = ttk.Treeview(janela, columns=list(
        existing_excel_df.columns), show="headings")
    for col in existing_excel_df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    for i, row in existing_excel_df.iterrows():
        tree.insert("", "end", values=list(row))

    tree.pack()

    janela.mainloop()

exibir_tabela()
