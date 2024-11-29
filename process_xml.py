import xml.etree.ElementTree as ET

def get_text_from_element(root, xpath, ns):
    """Extrai texto de um elemento XML com base no xpath."""
    element = root.find(xpath, ns)
    return element.text if element is not None else None

def get_infNFe_id(root, ns):
    """Extrai a chave NF-e a partir do atributo Id no elemento infNFe."""
    inf_nfe = root.find('.//nfe:infNFe', ns)
    return inf_nfe.attrib['Id'][3:] if inf_nfe is not None else None

def process_xml_identificacao(xml_file):
    """Processa XML para NF-e Identificação"""
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        
        data = {
            'Período': get_text_from_element(root, './/nfe:ide/nfe:dhEmi', ns),
            'Chave NF-e': get_infNFe_id(root, ns),
            'Situação': get_text_from_element(root, './/nfe:protNFe/nfe:infProt/nfe:xMotivo', ns),
            'Natureza Operação': get_text_from_element(root, './/nfe:ide/nfe:natOp', ns),
            'Indicador Forma Pagamento': get_text_from_element(root, './/nfe:ide/nfe:indPag', ns),
            'Modelo': get_text_from_element(root, './/nfe:ide/nfe:mod', ns),
            'Série': get_text_from_element(root, './/nfe:ide/nfe:serie', ns),
            'Número Documento': get_text_from_element(root, './/nfe:ide/nfe:nNF', ns),
            'Data Emissão Documento': get_text_from_element(root, './/nfe:ide/nfe:dhEmi', ns),
            'Tipo Operação': get_text_from_element(root, './/nfe:ide/nfe:tpNF', ns),
            'CNPJ/CPF Emitente': get_text_from_element(root, './/nfe:emit/nfe:CNPJ', ns),
            'Nome Emitente': get_text_from_element(root, './/nfe:emit/nfe:xNome', ns),
            'UF Emitente': get_text_from_element(root, './/nfe:emit/nfe:enderEmit/nfe:UF', ns),
            'CNPJ/CPF Destinatário': get_text_from_element(root, './/nfe:dest/nfe:CPF', ns),
            'Nome Destinatário': get_text_from_element(root, './/nfe:dest/nfe:xNome', ns),
            'UF Destinatário': get_text_from_element(root, './/nfe:dest/nfe:enderDest/nfe:UF', ns),
            'Vlr Total NF-e': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vNF', ns),
            'Finalidade Emissão': get_text_from_element(root, './/nfe:ide/nfe:finNFe', ns),
            'Indicador Operação Consumidor': get_text_from_element(root, './/nfe:ide/nfe:indFinal', ns),
            'Indicador Presença Comprador': get_text_from_element(root, './/nfe:ide/nfe:indPres', ns),
            'Inscrição Estadual Emitente': get_text_from_element(root, './/nfe:emit/nfe:IE', ns),
            'Inscrição Estadual Destinatário': get_text_from_element(root, './/nfe:dest/nfe:IE', ns),
            'Vlr Total Produto': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vProd', ns),
            'Vlr Total ICMS ST': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vST', ns),
            'Vlr Total ICMS': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vICMS', ns),
            'Vlr Total Frete': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vFrete', ns),
            'Vlr Total Seguro': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vSeg', ns),
            'Vlr Total Desconto Produto': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vDesc', ns),
            'Vlr Total II': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vII', ns),
            'Vlr Total IPI': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vIPI', ns),
            'Vlr Total PIS': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vPIS', ns),
            'Vlr Total Cofins': get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vCOFINS', ns),
        }
        return data
    except Exception as e:
        print(f"Erro ao processar o arquivo {xml_file}: {str(e)}")
        return None