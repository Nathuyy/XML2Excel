import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import xml.etree.ElementTree as ET

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
                data = [self.process_xml_identificacao(os.path.join(self.folder_path, xml_file)) for xml_file in xml_files]
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
                data = [self.process_xml_identificacao_pagamento(os.path.join(self.folder_path, xml_file)) for xml_file in xml_files]
            
            if data:
                df = pd.DataFrame(data, columns=columns)
                excel_path = os.path.join(self.folder_path, f"relatorio_{tipo.replace(' ', '_').lower()}.xlsx")
                df.to_excel(excel_path, index=False)
                messagebox.showinfo("Sucesso", f"Arquivo Excel gerado com sucesso em {excel_path}")
            else:
                messagebox.showwarning("Aviso", "Nenhuma informação foi extraída dos arquivos XML.")

    def process_xml_identificacao(self, xml_file):
        """Processa XML para NF-e Identificação"""
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
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
                'Vlr Total NF-e': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vNF', ns),
                'Finalidade Emissão': self.get_text_from_element(root, './/nfe:ide/nfe:finNFe', ns),
                'Indicador Operação Consumidor': self.get_text_from_element(root, './/nfe:ide/nfe:indFinal', ns),
                'Indicador Presença Comprador': self.get_text_from_element(root, './/nfe:ide/nfe:indPres', ns),
                'Inscrição Estadual Emitente': self.get_text_from_element(root, './/nfe:emit/nfe:IE', ns),
                'Inscrição Estadual Destinatário': self.get_text_from_element(root, './/nfe:dest/nfe:IE', ns),
                # 'SUFRAMA': self.get_text_from_element(root, './/nfe:dest/nfe:IE', ns),
                # 'CNPJ/CPF Transportador': self.get_text_from_element(root, './/nfe:dest/nfe:IE', ns),
                # Nome Transportador:
                'Vlr Total Produto': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vProd', ns),
                'Vlr Total ICMS ST': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vST', ns),
                'Vlr Total ICMS': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vICMS', ns),
                'Vlr Total Frete': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vFrete', ns),
                'Vlr Total Seguro': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vSeg', ns),
                'Vlr Total Desconto Produto': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vDesc', ns),
                # 'Vlr Total Desconto Incondicional Serviços': ,
                # 'Vlr Total Desconto Condicional Serviços': ,
                # 'Vlr Total Outras Despesas': ,
                'Vlr Total II': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vII', ns),
                'Vlr Total IPI': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vIPI', ns),
                'Vlr Total PIS': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vPIS', ns),
                'Vlr Total Cofins': self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vCOFINS', ns),
                'Meio Pagamento': self.get_text_from_element(root, './/nfe:pag/nfe:detPag/nfe:tPag', ns),
                'Vlr Pagamento': self.get_text_from_element(root, './/nfe:pag/nfe:detPag/nfe:vPag', ns),
                'Tipo Integração Pagamento': self.get_text_from_element(root, './/nfe:pag/nfe:detPag/nfe:card/nfe:tpIntegra', ns),
                'CNPJ Instituição Pagamento': self.get_text_from_element(root, './/nfe:pag/nfe:detPag/nfe:card/nfe:CNPJ', ns),
                'Bandeira Operadora Cartão Crédito/Débito': self.get_text_from_element(root, './/nfe:pag/nfe:detPag/nfe:card/nfe:tBand', ns),
                'Número Autorização Operação Cartão Crédito/Débito': self.get_text_from_element(root, './/nfe:pag/nfe:detPag/nfe:card/nfe:cAut', ns)
                # 'Vlr Total ISSQN': 
                # 'Vlr Total PIS Serviços':
                # 'Vlr Total Cofins Serviços':
                # 'Vlr Total Retenção ISSQN':
                
            }
            
            return data
        except Exception as e:
            print(f"Erro ao processar o arquivo {xml_file}: {str(e)}")
            return None

    def process_xml_identificacao_pagamento(self, xml_file):
        """Processa XML para NF-e Identificação e Pagamento"""
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            data = self.process_xml_identificacao(xml_file)
            # Adiciona campos relacionados ao pagamento, por exemplo:
            data['Vlr Total PIS'] = self.get_text_from_element(root, './/nfe:total/nfe:PIS/nfe:vPIS', ns)
            data['Vlr Total Cofins'] = self.get_text_from_element(root, './/nfe:total/nfe:COFINS/nfe:vCOFINS', ns)
            data['Vlr Troco'] = self.get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vTroco', ns)
            
            return data
        except Exception as e:
            print(f"Erro ao processar o arquivo {xml_file}: {str(e)}")
            return None

    def get_text_from_element(self, root, xpath, ns):
        """Auxiliar para obter texto de um elemento XML com XPath"""
        element = root.find(xpath, ns)
        return element.text if element is not None else ""
        
    def get_infNFe_id(self, root, ns):
        """Obtém a chave de identificação da NF-e"""
        infNFe = root.find('.//nfe:infNFe', ns)
        return infNFe.attrib.get('Id', '') if infNFe is not None else ''

if __name__ == "__main__":
    root = tk.Tk()
    window = MainWindow(root)
    root.mainloop()