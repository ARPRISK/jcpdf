import os
os.environ["TK_SILENCE_DEPRECATION"] = "1"
import pdfrw
import pandas as pd
from unidecode import unidecode
from tkinter import Tk, Button, Label, filedialog, messagebox
from datetime import datetime

def limpar_nome(nome):
    return unidecode(nome)

def extrair_dados_pdf(pdf_file):
    campos_preenchidos = False
    dados = {'Campo': [], 'Valor': []}
    pdf = pdfrw.PdfReader(pdf_file)
    for page in pdf.pages:
        if '/Annots' in page:
            annotations = page['/Annots']
            for annotation in annotations:
                subtype = annotation.get('/Subtype')
                if subtype == '/Widget':
                    field_name = annotation.get('/T')
                    if field_name:
                        field_name = limpar_nome(field_name[1:-1])  
                        if '/V' in annotation:
                            field_value = annotation['/V']
                            if isinstance(field_value, pdfrw.objects.pdfstring.PdfString):
                                field_value = field_value.decode()
                            elif isinstance(field_value, pdfrw.objects.pdfdict.PdfDict):
                                if '/AS' in field_value:
                                    field_value = field_value['/AS'].decode()
                                elif '/V' in field_value:
                                    field_value = field_value['/V'].decode()
                            campos_preenchidos = True
                        else:
                            field_value = ''
                        dados['Campo'].append(field_name)
                        dados['Valor'].append(field_value)
    if not campos_preenchidos:
        messagebox.showwarning("Aviso", "Nenhum campo foi preenchido. Preencha pelo menos um campo.")
    else:
        return dados

def adicionar_a_nova_aba(planilha_existente, dados_extraidos):
    with pd.ExcelWriter(planilha_existente, engine='openpyxl', mode='a') as writer:
        df = pd.DataFrame(dados_extraidos)
        df.to_excel(writer, sheet_name='Nova_Aba', index=False)

def selecionar_pdf_e_salvar_em_nova_aba():
    caminho_pdf = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if caminho_pdf:
        planilha_existente = filedialog.askopenfilename(filetypes=[("Planilha Excel", "*.xlsx")], title="Selecione a Planilha Existente")
        if planilha_existente:
            dados_extraidos = extrair_dados_pdf(caminho_pdf)
            if dados_extraidos:
                adicionar_a_nova_aba(planilha_existente, dados_extraidos)
                messagebox.showinfo("Sucesso", f"Dados adicionados com sucesso à planilha existente.")

def selecionar_pdf_e_salvar_em_nova_planilha():
    caminho_pdf = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if caminho_pdf:
        novo_nome = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Planilha Excel", "*.xlsx")], title="Salvar Planilha Como")
        if novo_nome:
            dados_extraidos = extrair_dados_pdf(caminho_pdf)
            if dados_extraidos:
                df = pd.DataFrame(dados_extraidos)
                df.to_excel(novo_nome, index=False)
                messagebox.showinfo("Sucesso", f"Dados salvos com sucesso em '{novo_nome}'")

root = Tk()
root.title("Adicionar Dados de PDF")

label = Label(root, text="Selecione uma opção abaixo para adicionar os dados do PDF:")
label.pack()

btn_nova_aba = Button(root, text="Adicionar a uma Nova Aba em uma Planilha Existente", command=selecionar_pdf_e_salvar_em_nova_aba)
btn_nova_aba.pack(pady=10)

btn_nova_planilha = Button(root, text="Salvar em uma Nova Planilha", command=selecionar_pdf_e_salvar_em_nova_planilha)
btn_nova_planilha.pack(pady=10)

root.mainloop()
