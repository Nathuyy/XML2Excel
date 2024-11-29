import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import xml.etree.ElementTree as ET
from process_xml import process_xml_identificacao

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Converta XML para Excel")
        self.root.geometry("400x300")
        
        self.title_label = tk.Label(root, text="XML para Excel", font=("Arial", 16))
        self.title_label.pack(pady=10)
        
        self.select_button = tk.Button(root, text="Selecionar Pasta", command=self.select_folder)
        self.select_button.pack(pady=10)
        
        self.folder_label = tk.Label(root, text="Nenhuma pasta selecionada", fg="gray", wraplength=400)
        self.folder_label.pack(pady=5)
        
        self.type_label = tk.Label(root, text="Selecione o tipo de Excel", font=("Arial", 12))
        self.type_label.pack(pady=10)
        
        self.button_identificacao = tk.Button(root, text="NF-e Identificação", command=lambda: self.convert_to_excel("NF-e Identificação"))
        self.button_identificacao.pack(pady=5)
        
        self.button_identificacao_pagamento = tk.Button(root, text="NF-e Identificação e Pagamento", command=lambda: self.convert_to_excel("NF-e Identificação e Pagamento"))
        self.button_identificacao_pagamento.pack(pady=5)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path = folder_path
            self.folder_label.config(text=f"Pasta selecionada: {folder_path}")
        else:
            self.folder_label.config(text="Nenhuma pasta selecionada", fg="gray")

    def convert_to_excel(self, tipo):
        if hasattr(self, 'folder_path'):
            xml_files = [f for f in os.listdir(self.folder_path) if f.endswith('.xml')]
            
            if not xml_files:
                messagebox.showwarning("Aviso", "A pasta selecionada não contém arquivos XML.")
                return
            
            if tipo == "NF-e Identificação":
                columns = [
                    "Período", "Chave NF-e", "Situação", "Natureza Operação", "Indicador Forma Pagamento",
                    "Modelo", "Série", "Número Documento", "Data Emissão Documento", "Tipo Operação",
                    "Finalidade Emissão", "Indicador Operação Consumidor", "Indicador Presença Comprador",
                    "CNPJ/CPF Emitente", "Inscrição Estadual Emitente", "Nome Emitente", "UF Emitente",
                    "CNPJ/CPF Destinatário", "Inscrição Estadual Destinatário", "Nome Destinatário",
                    "UF Destinatário", "SUFRAMA", "CNPJ/CPF Transportador", "Nome Transportador",
                    "Vlr Total NF-e", "Vlr Total Produto", "Vlr Total ICMS", "Vlr Total ICMS ST",
                    "Vlr Total Frete", "Vlr Total Seguro", "Vlr Total Desconto Produto",
                    "Vlr Total IPI", "Vlr Total PIS", "Vlr Total Cofins"
                ]
                data = [process_xml_identificacao(os.path.join(self.folder_path, xml_file)) for xml_file in xml_files]
            elif tipo == "NF-e Identificação e Pagamento":
                columns = [
                    "Período", "Chave NF-e", "Situação", "Natureza Operação", "Indicador Forma Pagamento",
                    "Modelo", "Série", "Número Documento", "Data Emissão Documento", "Tipo Operação",
                    "Finalidade Emissão", "Indicador Operação Consumidor", "Indicador Presença Comprador", 
                    "CNPJ/CPF Emitente", "Inscrição Estadual Emitente", "Nome Emitente", "UF Emitente", 
                    "CNPJ/CPF Destinatário", "Inscrição Estadual Destinatário", "Nome Destinatário", 
                    "UF Destinatário", "SUFRAMA", "CNPJ/CPF Transportador", "Nome Transportador", 
                    "Vlr Total NF-e", "Vlr Total Produto", "Vlr Total Serviços", "Vlr Total ICMS", 
                    "Vlr Total ICMS ST", "Vlr Total Frete", "Vlr Total Seguro", "Vlr Total Desconto Produto", 
                    "Vlr Total Desconto Incondicional Serviços", "Vlr Total Desconto Condicional Serviços", 
                    "Vlr Total Outras Despesas", "Vlr Total II", "Vlr Total IPI", "Vlr Total PIS", 
                    "Vlr Total Cofins", "Vlr Total ISSQN", "Vlr Total PIS Serviços", "Vlr Total Cofins Serviços", 
                    "Vlr Total Retenção ISSQN", "Forma Pagamento", "Meio Pagamento", "Vlr Pagamento", 
                    "Tipo Integração Pagamento", "CNPJ Instituição Pagamento", 
                    "Bandeira Operadora Cartão Crédito/Débito", 
                    "Número Autorização Operação Cartão Crédito/Débito", "Vlr Troco"
                ]
                data = [process_xml_identificacao(os.path.join(self.folder_path, xml_file)) for xml_file in xml_files]
            
            if data:
                df = pd.DataFrame(data, columns=columns)
                excel_path = os.path.join(self.folder_path, f"relatorio_{tipo.replace(' ', '_').lower()}.xlsx")
                df.to_excel(excel_path, index=False)
                messagebox.showinfo("Sucesso", f"Arquivo Excel gerado com sucesso em {excel_path}")
            else:
                messagebox.showwarning("Aviso", "Nenhuma informação foi extraída dos arquivos XML.")


if __name__ == "__main__":
    root = tk.Tk()
    window = MainWindow(root)
    root.mainloop()