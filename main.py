import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import xml.etree.ElementTree as ET


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Converta XML para Excel")
        self.root.geometry("400x200")
        
        self.title_label = tk.Label(root, text="XML para Excel", font=("Arial", 16))
        self.title_label.pack(pady=10)
        
        self.select_button = tk.Button(root, text="Selecionar Pasta", command=self.select_folder)
        self.select_button.pack(pady=10)
        
        self.convert_button = tk.Button(root, text="Converter para Excel", command=self.convert_to_excel)
        self.convert_button.pack(pady=10)
        
        self.folder_label = tk.Label(root, text="Nenhuma pasta selecionada", fg="gray", wraplength=400)
        self.folder_label.pack(pady=5)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path = folder_path
            self.folder_label.config(text=f"Pasta selecionada: {folder_path}")
        else:
            self.folder_label.config(text="Nenhuma pasta selecionada", fg="gray")

    def convert_to_excel(self):
        if hasattr(self, 'folder_path'):
            xml_files = [f for f in os.listdir(self.folder_path) if f.endswith('.xml')]
            
            if not xml_files:
                messagebox.showwarning("Aviso", "A pasta selecionada não contém arquivos XML.")
                return
            
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

            data = []

            for xml_file in xml_files:
                file_path = os.path.join(self.folder_path, xml_file)
                xml_data = self.process_xml(file_path)
                if xml_data:
                    data.append(xml_data)

            if data:
                df = pd.DataFrame(data, columns=columns)
                excel_path = os.path.join(self.folder_path, "relatorio_nfe.xlsx")
                df.to_excel(excel_path, index=False)
                messagebox.showinfo("Sucesso", f"Arquivo Excel gerado com sucesso em {excel_path}")
            else:
                messagebox.showwarning("Aviso", "Nenhuma informação foi extraída dos arquivos XML.")

    def process_xml(self, xml_file):
        """Processa o XML e extrai as informações desejadas"""
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()

            # Definindo o namespace do XML
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

            data = {
                'Período': self.get_text_from_element(root, './/nfe:ide/nfe:dhEmi', ns),
                'Chave NF-e': self.get_infNFe_id(root, ns),
                'Situação': self.get_text_from_element(root, './/nfe:protNFe/nfe:infProt/nfe:xMotivo', ns),
                'Natureza Operação': self.get_text_from_element(root, './/nfe:ide/nfe:natOp', ns),
                'Indicador Forma Pagamento': self.get_text_from_element(root, './/nfe:ide/nfe:indPag', ns),
                'Modelo': self.get_text_from_element(root, './/nfe:ide/nfe:mod', ns),
                'Série': self.get_text_from_element(root, './/nfe:ide/nfe:serie', ns),
                'Número Documento': self.get_text_from_element(root, './/nfe:ide/nfe:nNF', ns),
                'Data Emissão Documento': self.get_text_from_element(root, './/nfe:ide/nfe:dhEmi', ns),
                'Tipo Operação': self.get_text_from_element(root, './/nfe:ide/nfe:tpNF', ns),
                'CNPJ/CPF Emitente': self.get_text_from_element(root, './/nfe:emit/nfe:CNPJ', ns),
                'Nome Emitente': self.get_text_from_element(root, './/nfe:emit/nfe:xNome', ns),
                'UF Emitente': self.get_text_from_element(root, './/nfe:emit/nfe:enderEmit/nfe:UF', ns),
                'CNPJ/CPF Destinatário': self.get_text_from_element(root, './/nfe:dest/nfe:CPF', ns),
                'Nome Destinatário': self.get_text_from_element(root, './/nfe:dest/nfe:xNome', ns),
                'UF Destinatário': self.get_text_from_element(root, './/nfe:dest/nfe:enderDest/nfe:UF', ns),
                'Vlr Total NF-e': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vNF', ns)
            }

            return data
        except Exception as e:
            print(f"Erro ao processar o arquivo {xml_file}: {str(e)}")
            return None

    def get_infNFe_id(self, root, ns):
        """Obtém o valor do atributo Id da tag <infNFe>"""
        inf_nfe = root.find('.//nfe:infNFe', ns)
        if inf_nfe is not None:
            inf_nfe_id = inf_nfe.get('Id')
            if inf_nfe_id and inf_nfe_id.startswith("NFe"):
                return inf_nfe_id[3:]
            return inf_nfe_id
        return ""

    def get_text_from_element(self, root, xpath, ns):
        """Retorna o texto do elemento XML ou uma string vazia se não encontrado"""
        element = root.find(xpath, ns)
        return element.text if element is not None else ""


if __name__ == "__main__":
    root = tk.Tk()
    window = MainWindow(root)
    root.mainloop()
