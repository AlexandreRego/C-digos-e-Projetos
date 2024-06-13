import tkinter as tk
from tkinter import filedialog
from tkinter import PhotoImage
import shutil
import PyPDF2
import PyPDF2 as pyf
import pathlib
from tkinter import PhotoImage
from pathlib import Path
import os
import re

numero_de_paginas = 0
arquivos_renomeados_enumerados = False

def contar_paginas(file_path):
    global numero_de_paginas
    try:
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            numero_de_paginas = len(pdf_reader.pages)
            status_label.config(
                text=f"O arquivo que você enviou possui {numero_de_paginas} página(s).")
    except Exception as e:
        status_label.config(
            text="Erro ao contar as páginas do arquivo: " + str(e))


def obter_numero_pedido(text):
    order_number_pattern_45 = r'PEDIDO 45(\d+)'
    order_number_pattern_47 = r'PEDIDO 47(\d+)'

    if re.search(order_number_pattern_45, text):
        return "111319"
    elif re.search(order_number_pattern_47, text):
        return "5698"
    else:
        return "N/A"  

def renomear_arquivos_enumerados(pdf_dir):
    for i, NF in enumerate(os.listdir(pdf_dir)):
        if NF.endswith('.pdf'):
            num_pagina = i
            pdf_path = os.path.join(pdf_dir, NF)
            extracted_info = extract_info_from_pdf(pdf_path)
            order_number = obter_numero_pedido(extracted_info)
            new_filename = f"DANFE - {num_pagina + 1} - {order_number}.pdf"
            new_pdf_path = os.path.join(pdf_dir, new_filename)
            os.rename(pdf_path, new_pdf_path)

def renomear_arquivos_de_acordo_com_NF(pdf_dir):
    for NF in os.listdir(pdf_dir):
        if NF.endswith('.pdf'):
            pdf_path = os.path.join(pdf_dir, NF)
            extracted_info = extract_info_from_pdf(pdf_path)
            order_number = obter_numero_pedido(extracted_info)
            new_filename = extracted_info.split()[9:10]
            new_filename = "_".join(new_filename).replace(
                '000.0', '').replace('.', '')
            new_pdf_path = os.path.join(
                pdf_dir, f"DANFE - {new_filename} - {order_number}.pdf")
            os.rename(pdf_path, new_pdf_path)

def extract_info_from_pdf(pdf_path):
    pdf_reader = pyf.PdfReader(pdf_path)
    text = ""

    for page in pdf_reader.pages:
        text += page.extract_text()

    return text

def fazer_upload():

    file_paths = filedialog.askopenfilenames()
    if file_paths:
        global numero_de_paginas, arquivos_renomeados_enumerados
        numero_de_paginas = 0
        arquivos_renomeados_enumerados = False

        for file_path in file_paths:
    
            shutil.copy(file_path, file_path.split('/')[-1])
            status_label.config(text="Arquivos enviados com sucesso: " +
                                ", ".join(file_path.split('/')[-1] for file_path in file_paths))
            contar_paginas(file_path)


def separar_paginas_pdf(diretorio_pdf):
    if os.path.exists(diretorio_pdf):
        pdf_files = [f for f in os.listdir(
            diretorio_pdf) if f.endswith('.pdf')]

        for pdf_file in pdf_files:
            pdf_path = os.path.join(diretorio_pdf, pdf_file)

            with open(pdf_path, 'rb') as pdf_input_file:
                pdf_reader = pyf.PdfReader(pdf_input_file)

                for page_num in range(len(pdf_reader.pages)):
                    output_pdf = pyf.PdfWriter()
                    output_pdf.add_page(pdf_reader.pages[page_num])

                    nome_saida = f"{os.path.splitext(pdf_file)[0]}_pagina_{page_num + 1}.pdf"
                    caminho_saida = os.path.join(diretorio_pdf, nome_saida)

                    with open(caminho_saida, 'wb') as output_file:
                        output_pdf.write(output_file)

            os.remove(pdf_path)

        status_label.config(
            text="Os arquivos PDF foram divididos em páginas individuais. O original foi excluído.")
    else:
        status_label.config(
            text="O diretório especificado não foi encontrado.")


root = tk.Tk()
root.title("Janela de Upload de Arquivos")

root.attributes('-fullscreen', True)

def sair_tela_cheia(event):
    root.attributes('-fullscreen', False)

root.bind("<Escape>", sair_tela_cheia)

root.configure(bg='#040404')

image = PhotoImage(
    file=r"C:\Users\alexa\Desktop\Interfaceairp\BOTS.png")

largura_imagem = 10
altura_imagem = 5

border_color = "#040404"  # Cor da borda
border_width = 0  # Largura da borda
image_label = tk.Label(root, image=image, bd=border_width, relief="solid")
image_label.pack()


highlight_color = root.cget('bg')
status_label = tk.Label(root, text="Bem-vindo! Facilities e Varejo - Adequação de NF para RPA (AIR PROMO) ",
                        font=("Arial", 18), fg="white", bg=highlight_color)
status_label.pack(pady=5)

upload_button = tk.Button(
    root, text="Faça o Upload do(os) Arquivo(os)", command=fazer_upload)
upload_button.pack(pady=10)
upload_button.configure(bg='#6e98b9')

separar_paginas_button = tk.Button(
    root, text="Separar Páginas do PDF caso estejam em um só arquivo", command=lambda: separar_paginas_pdf(pdf_dir))
separar_paginas_button.pack(pady=10)
separar_paginas_button.configure(bg='#6e98b9')

rename_enumerated_button = tk.Button(
    root, text="Renomear os arquivos ordinalmente para prosseguir", command=lambda: renomear_arquivos_enumerados(pdf_dir))
rename_enumerated_button.pack(pady=10)
rename_enumerated_button.configure(bg='#6e98b9')

rename_nf_button = tk.Button(
    root, text="Renomear os arquivos para o formato adequado ao RPA", command=lambda: renomear_arquivos_de_acordo_com_NF(pdf_dir))
rename_nf_button.pack(pady=10)
rename_nf_button.configure(bg='#6e98b9')

pdf_dir = r'C:\Users\alexa\Desktop\Interfaceairp'

root.mainloop()
