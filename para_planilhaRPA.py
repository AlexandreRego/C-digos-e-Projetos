import os
import re
import openpyxl
import pandas as pd
import PyPDF2
import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta


def extrair_informacoes_pdf(pdf_path):
    # Inicialize um dicionário para armazenar informações extraídas

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

    # Abra o PDF usando PyPDF2
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        num_pages = len(pdf_reader.pages)

        # Percorra todas as páginas do PDF
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            page_text = page.extract_text()

            # Use expressões regulares para encontrar informações específicas e atualize o dicionário
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


# Diretório onde os arquivos PDF estão localizados
diretorio_pdf = r'C:\Users\alexa\Desktop\Interfaceairp'

# Nome da planilha Excel
planilha_excel = 'sua_planilha.xlsx'

# Verifique se a planilha já existe ou crie uma nova
if not os.path.exists(planilha_excel):
    # Crie um DataFrame vazio com as colunas apropriadas
    df = pd.DataFrame(columns=['Responsável', 'Data Lançamento', 'Carteira', 'Sistema SAP', 'Cód. Fornecedor', 'Nome Fornecedor',
                      'CNPJ Fornecedor', 'Tipo', 'NF', 'Série', 'Emissão', 'Valor Total', 'Vencimento', 'Centro', 'Pedido'])
else:
    # Se a planilha existir, carregue-a para o DataFrame
    df = pd.read_excel(planilha_excel)

# Percorra todos os arquivos PDF na pasta
for filename in os.listdir(diretorio_pdf):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(diretorio_pdf, filename)
        info = extrair_informacoes_pdf(pdf_path)

        # Extrair termos do título do PDF
        termos_titulo = filename.split(' - ')
        if len(termos_titulo) >= 3:
            info['Cód. Fornecedor'] = termos_titulo[-1].replace('.pdf', '')
            info['Tipo'] = termos_titulo[0]
            info['NF'] = termos_titulo[1]

        # Adicione o dicionário de informações ao DataFrame
        df = pd.concat([df, pd.DataFrame([info])], ignore_index=True)

# Salve o DataFrame no arquivo Excel
df.to_excel(planilha_excel, index=False)


def calcular_vencimento(data_emissao):
    try:
        data_emissao = datetime.strptime(data_emissao, '%d/%m/%Y')
        data_vencimento = data_emissao + timedelta(days=85)
        return data_vencimento.strftime('%d/%m/%Y')
    except Exception as e:
        print(f"Erro ao calcular a data de vencimento: {str(e)}")
        return None


# Função para extrair informações de um arquivo PDF
def extract_info_from_pdf(pdf_file_path):
    try:
        with open(pdf_file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()

            # Use expressões regulares para encontrar a data de emissão
            date_pattern = r'\d{2}/\d{2}/\d{4}'
            date_match = re.search(date_pattern, text)
            date_found = date_match.group() if date_match else None

            # Use expressões regulares para encontrar o número do pedido
            order_number_pattern = r'PEDIDO (\d+)'
            order_number_match = re.search(order_number_pattern, text)
            order_number = order_number_match.group(
                1) if order_number_match else None

            # Use expressões regulares para encontrar números com vírgula e ponto antes de "VALOR TOTAL DA NOTA"
            valor_total_pattern = r'([\d.,]+)\s+VALOR TOTAL DA NOTA'
            valor_total_match = re.search(valor_total_pattern, text)
            valor_total = valor_total_match.group(
                1) if valor_total_match else None

            # Use expressões regulares para encontrar o número do CNPJ Forn
            cnpj_1_pattern = r'9060019859 (\d{2}[.-]?\d{3}[.-]?\d{3}[/-]?\d{4}[.-]?\d{2})'
            cnpj_1_match = re.search(cnpj_1_pattern, text)
            cnpj_1 = cnpj_1_match.group(1) if cnpj_1_match else None

            # Use expressões regulares para encontrar o número do CNPJ Tomador
            cnpj_2_pattern = r'ESTADUALBOTICARIO PRODUTOS DE BELEZA LTDA (\d{2}[.-]?\d{3}[.-]?\d{3}[/-]?\d{4}[.-]?\d{2})'
            cnpj_2_match = re.search(cnpj_2_pattern, text)
            cnpj_2 = cnpj_2_match.group(1) if cnpj_2_match else None

            # Use expressões regulares para encontrar o nome do fornecedor
            nome_forn_pattern = r'SÉRIE: 1RECEBEMOS DE ([A-Z\s]+?) EIRELI'
            nome_forn_match = re.search(nome_forn_pattern, text)
            nome_forn = nome_forn_match.group(1) if nome_forn_match else None

            # Calcular a data de vencimento
            vencimento = calcular_vencimento(date_found)

            return date_found, order_number, valor_total, cnpj_1, cnpj_2, nome_forn, vencimento
    except Exception as e:
        print(f"Erro ao processar {pdf_file_path}: {str(e)}")
        return None, None


# Listas para armazenar os dados
pdf_files = []
dates = []
order_numbers = []
valor_total_list = []
cnpj_list = []
nome_forn_list = []
cnpj2_list = []
vencimento_list = []

# Loop pelos arquivos no diretório
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

# Criar um DataFrame com as informações
pdf_df = pd.DataFrame({'Nome do PDF': pdf_files, 'Data de Emissão da NF': dates, 'Número do Pedido': order_numbers,
                      'Valor Total': valor_total_list, 'cnpj_1': cnpj_list, 'nome_forn': nome_forn_list, 'Centro': cnpj2_list, 'vencimento': vencimento_list})


# Carregar a planilha existente
existing_excel_df = pd.read_excel(planilha_excel)

# Adicionar as colunas 'Emissão' e 'Pedido' à planilha existente e preencher com os dados extraídos dos PDFs
existing_excel_df['Emissão'] = pdf_df['Data de Emissão da NF']
existing_excel_df['Pedido'] = pdf_df['Número do Pedido']
existing_excel_df['Valor Total'] = pdf_df['Valor Total']
existing_excel_df['CNPJ Fornecedor'] = pdf_df['cnpj_1']
existing_excel_df['Nome Fornecedor'] = pdf_df['nome_forn']
existing_excel_df['Centro'] = pdf_df['Centro']
existing_excel_df['Vencimento'] = pdf_df['vencimento']


# Salvar a planilha atualizada
existing_excel_df.to_excel(planilha_excel, index=False)

# Função para exibir a tabela


def exibir_tabela():
    # Crie uma janela
    janela = tk.Tk()
    janela.title("Tabela de Dados")

    # Crie uma árvore (TreeView) para exibir a tabela
    tree = ttk.Treeview(janela, columns=list(
        existing_excel_df.columns), show="headings")
    for col in existing_excel_df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)

    # Adicione os dados ao TreeView
    for i, row in existing_excel_df.iterrows():
        tree.insert("", "end", values=list(row))

    tree.pack()

    janela.mainloop()


# Chame a função para exibir a tabela
exibir_tabela()
