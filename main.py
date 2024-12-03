import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from process_xml import process_xml_identificacao, process_xml_307

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Converta XML para Excel")
        self.root.geometry("800x800")
        
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
            root, text="307 - Detalhamento de Produtos e Serviços e Tributos - Completo", 
            command=lambda: self.convert_to_excel("Detalhamento de Produtos e Serviços e Tributos - Completo")
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
            },
            "Detalhamento de Produtos e Serviços e Tributos - Completo": {
                "columns": [
                    "Período", "Chave NF-e", "Situação", "Data Situação", "Natureza Operação", "Modelo", "Série","Número Documento", "Data Emissão Documento","Tipo Operação", "Finalidade Emissão", 
                    "Indicador Operação Consumidor", "Indicador Presença Comprador", "Tipo Emissão", "CNPJ/CPF Emitente", "Inscrição Estadual Emitente", "Nome Emitente", "Logradouro Emitente","Número Emitente", 
                    "Bairro Emitente", "Município Emitente", "UF Emitente","CEP Emitente", "Fone Emitente","Regime Tributário Emitente","CNPJ/CPF Destinatário","Nome Destinatário","Logradouro Destinatário","Número Destinatário",
                    "Bairro Destinatário","Município Destinatário","UF Destinatário", "Indicador Contribuinte Destinatário","Vlr Total Base Cálculo ICMS","Vlr Total ICMS","Vlr Total ICMS Desonerado",
                    "Vlr Total Base Cálculo ICMS ST", "Vlr Total ICMS ST","Vlr Total Produto","Vlr Total Frete","Vlr Total Seguro","Vlr Total Desconto","Vlr Total II","Vlr Total IPI",
                    "Vlr Total PIS", "Vlr Total Cofins","Vlr Total Outras Despesas","Vlr Total NF-e","Vlr Total Aproximado Tributos", "Modalidade Frete", "Informação Adicional Contribuinte", 
                    "Número Item", "Código Item", "EAN", "Descrição Item","Código Benefício", "NCM", "Ex TIPI","CFOP", "Qtde","Vlr Unitário", "Vlr Total Produtos", "EAN Tributável", "CST ICMS",
                    "Modalidade Base Cálculo ICMS", "Percentual Redução ICMS", "Alíquota ICMS","Vlr ICMS Operação Diferimento", "Vlr Base Cálculo Efetiva","Alíquota ICMS Efetiva",
                    "Vlr ICMS Efetivo", "Vlr ICMS Substituto", "Vlr Base Cálculo FCP", "CST PIS", "Vlr Base Cálculo PIS","Alíquota PIS","Vlr PIS","CST Cofins","Vlr Base Cálculo Cofins","Alíquota Cofins", "Vlr Cofins"
                ],
                "process_func": process_xml_307
            }
        }

        if tipo not in options:
            messagebox.showerror("Erro", "Tipo de conversão não suportado.")
            return

        config = options[tipo]
        processed_data = []

        for xml_file in xml_files:
            try:
                results = config["process_func"](os.path.join(self.folder_path, xml_file))
                
                if isinstance(results, list):
                    processed_data.extend(results)
                elif isinstance(results, dict):  
                    processed_data.append(list(results.values()))
                else:
                    raise ValueError("Formato de dados inesperado.")
            except Exception as e:
                messagebox.showwarning("Erro", f"Erro ao processar {xml_file}: {e}")
                continue

        if processed_data:
            try:
                df = pd.DataFrame(processed_data, columns=config["columns"])
                excel_path = os.path.join(self.folder_path, f"relatorio_{tipo.replace(' ', '_').lower()}.xlsx")
                df.to_excel(excel_path, index=False)
                messagebox.showinfo("Sucesso", f"Arquivo Excel gerado com sucesso em {excel_path}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar o Excel: {e}")
        else:
            messagebox.showwarning("Aviso", "Nenhuma informação foi extraída dos arquivos XML.")



if __name__ == "__main__":
    root = tk.Tk()
    window = MainWindow(root)
    root.mainloop()
