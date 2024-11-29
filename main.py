import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from process_xml import process_xml_identificacao

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Converta XML para Excel")
        self.root.geometry("800x800")
        
        # Elementos da interface
        self.title_label = tk.Label(root, text="XML para Excel", font=("Arial", 16))
        self.title_label.pack(pady=10)
        
        self.select_button = tk.Button(root, text="Selecionar Pasta", command=self.select_folder)
        self.select_button.pack(pady=10)
        
        self.folder_label = tk.Label(root, text="Nenhuma pasta selecionada", fg="gray", wraplength=400)
        self.folder_label.pack(pady=5)
        
        self.type_label = tk.Label(root, text="Selecione o tipo de Excel", font=("Arial", 12))
        self.type_label.pack(pady=10)
        
        self.button_identificacao = tk.Button(
            root, text="300 - Identificação da NF-e-NFC-e", 
            command=lambda: self.convert_to_excel("NF-e Identificação")
        )
        self.button_identificacao.pack(pady=5)
        
        self.button_pagamento = tk.Button(
            root, text="301 - Identificação e Informações de Pagamento", 
            command=lambda: self.convert_to_excel("NF-e Identificação e Pagamento")
        )
        self.button_pagamento.pack(pady=5)
        
        self.button_referencia = tk.Button(
            root, text="302 - Identificação e Documento Fiscal Referenciado", 
            command=lambda: self.convert_to_excel("NF-e Referenciado")
        )
        self.button_referencia.pack(pady=5)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path = folder_path
            self.folder_label.config(text=f"Pasta selecionada: {folder_path}", fg="black")
        else:
            self.folder_label.config(text="Nenhuma pasta selecionada", fg="gray")

    def convert_to_excel(self, tipo):
        if not hasattr(self, 'folder_path'):
            messagebox.showwarning("Aviso", "Por favor, selecione uma pasta primeiro.")
            return

        xml_files = [f for f in os.listdir(self.folder_path) if f.endswith('.xml')]
        if not xml_files:
            messagebox.showwarning("Aviso", "A pasta selecionada não contém arquivos XML.")
            return

        options = {
            "NF-e Identificação": {
                "columns": [
                    "Período", "Chave NF-e", "Situação", "Natureza Operação", "Indicador Forma Pagamento",
                    "Modelo", "Série", "Número Documento", "Data Emissão Documento", "Tipo Operação",
                    "Finalidade Emissão", "Indicador Operação Consumidor", "Indicador Presença Comprador",
                    "CNPJ/CPF Emitente", "Inscrição Estadual Emitente", "Nome Emitente", "UF Emitente",
                    "CNPJ/CPF Destinatário", "Inscrição Estadual Destinatário", "Nome Destinatário",
                    "UF Destinatário", "SUFRAMA", "CNPJ/CPF Transportador", "Nome Transportador",
                    "Vlr Total NF-e", "Vlr Total Produto", "Vlr Total ICMS", "Vlr Total ICMS ST",
                    "Vlr Total Frete", "Vlr Total Seguro", "Vlr Total Desconto Produto",
                    "Vlr Total IPI", "Vlr Total PIS", "Vlr Total Cofins"
                ],
                "process_func": process_xml_identificacao
            },
            "NF-e Identificação e Pagamento": {
                "columns": [
                    "Período", "Chave NF-e", "Situação", "Natureza Operação", "Indicador Forma Pagamento",
                    "Modelo", "Série", "Número Documento", "Data Emissão Documento", "Tipo Operação",
                    "Finalidade Emissão", "Indicador Operação Consumidor", "Indicador Presença Comprador",
                    "CNPJ/CPF Emitente", "Inscrição Estadual Emitente", "Nome Emitente", "UF Emitente",
                    "CNPJ/CPF Destinatário", "Inscrição Estadual Destinatário", "Nome Destinatário",
                    "UF Destinatário", "SUFRAMA", "CNPJ/CPF Transportador", "Nome Transportador",
                    "Vlr Total NF-e", "Vlr Total Produto", "Vlr Total Serviços", "Vlr Total ICMS",
                    "Vlr Total ICMS ST", "Vlr Total Frete", "Vlr Total Seguro", "Vlr Total Desconto Produto",
                    "Vlr Total IPI", "Vlr Total PIS", "Vlr Total Cofins", "Forma Pagamento", "Meio Pagamento",
                    "Vlr Pagamento", "CNPJ Instituição Pagamento", "Bandeira Operadora Cartão"
                ],
                "process_func": process_xml_identificacao
            }
        }

        if tipo not in options:
            messagebox.showerror("Erro", "Tipo de conversão não suportado.")
            return

        config = options[tipo]
        data = [config["process_func"](os.path.join(self.folder_path, xml_file)) for xml_file in xml_files]

        if data:
            df = pd.DataFrame(data, columns=config["columns"])
            excel_path = os.path.join(self.folder_path, f"relatorio_{tipo.replace(' ', '_').lower()}.xlsx")
            df.to_excel(excel_path, index=False)
            messagebox.showinfo("Sucesso", f"Arquivo Excel gerado com sucesso em {excel_path}")
        else:
            messagebox.showwarning("Aviso", "Nenhuma informação foi extraída dos arquivos XML.")


if __name__ == "__main__":
    root = tk.Tk()
    window = MainWindow(root)
    root.mainloop()
