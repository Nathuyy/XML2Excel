# Conversor de XML para Excel

Este é um aplicativo simples em Python usando a biblioteca Tkinter, que permite converter arquivos XML em planilhas Excel. O programa processa arquivos XML com informações de notas fiscais eletrônicas (NF-e) e gera relatórios em formato Excel de acordo com os tipos de informações selecionados.

## Funcionalidades

- Selecione uma pasta contendo arquivos XML.
- Converta arquivos XML em planilhas Excel para diferentes tipos de informações de NF-e, como:
  - Identificação da NF-e
  - Identificação e Informações de Pagamento
  - Detalhamento de Produtos e Serviços e Tributos
  - Identificação da NF-e e Documento Fiscal Referenciado
- O resultado é um arquivo Excel gerado dentro da mesma pasta contendo os dados extraídos do XML.

## Tecnologias Utilizadas

- **Python 3.x**
- **Tkinter** para a interface gráfica
- **Pandas** para manipulação e exportação de dados para Excel
- **xml.etree.ElementTree** para processamento específico dos arquivos XML

## Como Usar

1. Clone o repositório ou baixe os arquivos.
2. Instale as dependências necessárias:
3. Execute o arquivo `main.py` para abrir a interface gráfica:
   ```bash
   python main.py
4. Selecione uma pasta contendo arquivos XML.
5. Escolha o tipo de conversão desejado.
6. O arquivo Excel será gerado na mesma pasta.
