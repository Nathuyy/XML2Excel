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
                    "Período", "Chave NF-e", "Situação", "Data Situação", "Natureza Operação", "Indicador Forma Pagamento", "Modelo", "Série", 
                    "Número Documento", "Data Emissão Documento", "Data Entrada/Saída", "Tipo Operação", "Finalidade Emissão", 
                    "Indicador Operação Consumidor", "Indicador Presença Comprador", "Tipo Emissão", "Data/Hora Contingência", 
                    "Justificativa Contingência", "CNPJ/CPF Emitente", "Inscrição Estadual Emitente", "Nome Emitente", "Nome Fantasia Emitente", 
                    "Logradouro Emitente", "Número Emitente", "Complemento Emitente", "Bairro Emitente", "Município Emitente", "UF Emitente", 
                    "CEP Emitente", "Fone Emitente", "Inscrição Estadual ST Emitente", "Inscrição Municipal Emitente", "CNAE Fiscal Emitente", 
                    "Regime Tributário Emitente", "CNPJ/CPF Destinatário", "Inscrição Estadual Destinatário", "ID Estrangeiro Destinatário", 
                    "Nome Destinatário", "SUFRAMA", "Logradouro Destinatário", "Número Destinatário", "Complemento Destinatário", 
                    "Bairro Destinatário", "Município Destinatário", "UF Destinatário", "CEP Destinatário", "País Destinatário", "Fone Destinatário", 
                    "Indicador Contribuinte Destinatário", "SUFRAMA Destinatário", "Inscrição Municipal Destinatário", "E-mail Destinatário", 
                    "CNPJ/CPF Retirada", "Logradouro Retirada", "Número Retirada", "Complemento Retirada", "Bairro Retirada", "Município Retirada", 
                    "UF Retirada", "Nome Expedidor", "CEP", "Código País", "Nome País", "Telefone", "E-mail", "Inscrição Estadual Expedidor", 
                    "CNPJ/CPF Entrega", "Logradouro Entrega", "Número Entrega", "Complemento Entrega", "Bairro Entrega", "Município Entrega", 
                    "UF Entrega", "Nome Entrega", "CEP Entrega", "Código País Entrega", "Nome País Entrega", "Telefone Entrega", "E-mail Entrega", 
                    "Inscrição Estadual Entrega", "Vlr Total Base Cálculo ICMS", "Vlr Total ICMS", "Vlr Total ICMS Desonerado", "Vlr Total Base Cálculo ICMS ST", 
                    "Vlr Total ICMS ST", "Vlr Total Produto", "Vlr Total Frete", "Vlr Total Seguro", "Vlr Total Desconto", "Vlr Total II", 
                    "Vlr Total IPI", "Vlr Total PIS", "Vlr Total Cofins", "Vlr Total Outras Despesas", "Vlr Total NF-e", "Vlr Total Aproximado Tributos", 
                    "Vlr Total Serviços", "Vlr Total Base Cálculo ISSQN", "Vlr Total ISSQN", "Vlr Total PIS Serviços", "Vlr Total Cofins Serviços", 
                    "Data Prestação Serviços", "Vlr Total Redução Base Serviços", "Vlr Total Outras Retenções Serviços", "Vlr Total Desconto Incondicional Serviços", 
                    "Vlr Total Desconto Condicional Serviços", "Vlr Total Retenção ISSQN", "Modalidade Frete", "CNPJ/CPF Transportador", 
                    "Nome Transportador", "Inscrição Estadual Transportador", "Endereço Transportador", "Município Transportador", "UF Transportador", 
                    "Vlr Serviço", "Vlr Base Retenção ICMS", "Alíquota Retenção", "Vlr ICMS Retido", "CFOP Retenção", "Código Município FG", 
                    "Placa Veículo", "UF Veículo", "RNTC Veículo", "Placa Reboque", "UF Reboque", "RNTC Reboque", "Vagão", "Balsa", 
                    "Qtde Volume Transportado", "Espécie Volume Transportado", "Marca Volume Transportado", "Peso Líquido", "Peso Bruto", 
                    "Informação Adicional Fisco", "Informação Adicional Contribuinte", "Número Item", "Código Item", "EAN", "Descrição Item", 
                    "NCM", "CEST", "Código Benefício", "Ex TIPI", "CFOP", "Unidade Comercial", "Qtde", "Vlr Unitário", "Vlr Total Produtos", 
                    "EAN Tributável", "Unidade Tributável", "Quantidade Tributável", "Vlr Frete Item", "Vlr Seguro Item", "Vlr Desconto Item", 
                    "Vlr Outas Despesas Item", "Número FCI", "Número Pedido Compra", "Item Pedido Compra", "Número DI", "Data DI", 
                    "Local Desembaraço", "UF Desembaraço", "Data Desembaraço", "Número Lote Produto", "Qtde Produto Lote", "Data Fabricação/Produção", 
                    "Data Validade", "Código Agregação", "Informação Adicional Produto", "Tipo Operação Venda", "Chassi Veículo", "Tipo Combustível", 
                    "Ano/Modelo Fabricação", "Ano Fabricação", "Tipo Veículo", "Espécie Veículo", "VIN - Chassi Remarcado", "Capacidade Lotação", 
                    "Restrição", "Origem CST ICMS", "CST ICMS", "Modalidade Base Cálculo ICMS", "Qtde Tributada", "Alíquota ad rem ICMS", 
                    "Vlr ICMS Próprio", "Qtde Tributada Retenção", "Alíquota ad rem ICMS Retenção", "Vlr ICMS Retido", "Percentual Redução adrem", 
                    "Motivo Redução adrem", "Qtde Tributada Diferida", "Alíquota ad rem ICMS Diferido", "Vlr ICMS Monofásico Diferido", 
                    "Vlr ICMS Monofásico Operação", "Qtde Tributada Retida Anteriormente", "Alíquota ad rem Retido Anteriormente", 
                    "Vlr ICMS Retido Anteriormente", "Percentual Redução ICMS", "Vlr Base Cálculo ICMS", "Alíquota ICMS", "Vlr ICMS Operação Diferimento", 
                    "Percentual Diferimento", "Vlr ICMS Diferido", "Vlr ICMS", "Percentual Redução Base Efetiva", "Vlr Base Cálculo Efetiva", 
                    "Alíquota ICMS Efetiva", "Vlr ICMS Efetivo", "Vlr ICMS Substituto", "Vlr Base Cálculo ICMS Destinatário", "Vlr ICMS Destinatário", 
                    "Modalidade Base Cálculo ICMS ST", "Percentual MVA", "Percentual Redução Base ICMS ST", "Vlr Base Cálculo ICMS ST", 
                    "Alíquota ICMS ST", "Vlr ICMS ST", "Vlr Base Cálculo ICMS ST Retido", "Vlr ICMS ST Retido", "Vlr ICMS Desonerado", 
                    "Motivo Desoneração ICMS", "Percentual Base Cálculo Operação Própria", "UF Devido ICMS ST", "Alíquota Crédito Simples", 
                    "Vlr Crédito Simples", "Vlr Base Cálculo FCP", "Alíquota FCP", "Vlr FCP", "Vlr Base Cálculo FCP ST", "Alíquota FCP ST", 
                    "Vlr FCP ST", "Vlr Base Cálculo FCP Retido Anteriormente ST", "Alíquota FCP Retido Anteriormente ST", "Vlr FCP Retido Anteriormente ST", 
                    "Vlr Base Cálculo ICMS UF Destino", "Vlr Base Cálculo FCP UF Destino", "Percentual ICMS Relativo FCP UF Destino", 
                    "Alíquota Interna UF Destino", "Alíquota Interestadual UF Envolvidas", "Percentual Provisório Partilha ICMS Interestadual", 
                    "Vlr ICMS Relativo FCP UF Destino", "Vlr ICMS Interestadual UF Destino", "Vlr ICMS Interestadual UF Remetente", 
                    "Vlr Total ICMS Relativo FCP UF Destino", "Vlr Total ICMS Interestadual UF Destino", "Vlr Total ICMS Interestadual UF Remetente", 
                    "Vlr Total FCP Retido Anteriormente ST", "Classe Enquadramento IPI", "CST IPI", "Vlr Base Cálculo IPI", "Alíquota IPI", 
                    "Vlr IPI", "Vlr Base Cálculo PIS", "Alíquota PIS", "Vlr PIS", "Vlr Base Cálculo COFINS", "Alíquota COFINS", "Vlr COFINS", 
                    "Base Cálculo ISSQN", "Alíquota ISSQN", "Vlr ISSQN", "Vlr Retenção ISSQN", "Código Receita", "Observações", "Tag NFe"
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
