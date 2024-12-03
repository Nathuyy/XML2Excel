import xml.etree.ElementTree as ET

def get_text_from_element(root, xpath, ns):
    element = root.find(xpath, ns)
    return element.text if element is not None else None

def get_infNFe_id(root, ns):
    inf_nfe = root.find('.//nfe:infNFe', ns)
    return inf_nfe.attrib['Id'][3:] if inf_nfe is not None else None

def process_xml_identificacao(xml_file):
    """Processa XML para NF-e Identificação"""
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        
        
        data = [{
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
        }]
        return data
    except Exception as e:
        print(f"Erro ao processar o arquivo {xml_file}: {str(e)}")
        return None


def process_xml_307(xml_file):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

        produtos = []

        periodo = get_text_from_element(root, './/nfe:ide/nfe:dhEmi', ns)
        chavenfe = get_infNFe_id(root, ns)
        situacao = get_text_from_element(root, './/nfe:protNFe/nfe:infProt/nfe:xMotivo', ns)
        natureza = get_text_from_element(root, './/nfe:ide/nfe:natOp', ns)
        modelo = get_text_from_element(root, './/nfe:ide/nfe:mod', ns)
        serie = get_text_from_element(root, './/nfe:ide/nfe:serie', ns)
        numero = get_text_from_element(root, './/nfe:ide/nfe:nNF', ns)
        data_doc = get_text_from_element(root, './/nfe:ide/nfe:dhEmi', ns)
        tipo = get_text_from_element(root, './/nfe:ide/nfe:tpNF', ns)
        finalidade = get_text_from_element(root, './/nfe:ide/nfe:finNFe', ns)
        indicador = get_text_from_element(root, './/nfe:ide/nfe:indFinal', ns)
        indicador_presenca = get_text_from_element(root, './/nfe:ide/nfe:indPres', ns)
        tipo = get_text_from_element(root, './/nfe:ide/nfe:tpEmis', ns)
        cnpj =  get_text_from_element(root, './/nfe:emit/nfe:CNPJ', ns)
        inscricao = get_text_from_element(root, './/nfe:emit/nfe:IE', ns)
        nome_emit = get_text_from_element(root, './/nfe:emit/nfe:xNome', ns)
        logradouro = get_text_from_element(root, './/nfe:emit/nfe:enderEmit/nfe:xLgr', ns)
        numero_emit =  get_text_from_element(root, './/nfe:emit/nfe:enderEmit/nfe:nro', ns)
        bairro = get_text_from_element(root, './/nfe:emit/nfe:enderEmit/nfe:xBairro', ns)
        municipio = get_text_from_element(root, './/nfe:emit/nfe:enderEmit/nfe:xMun', ns)
        uf_emit = get_text_from_element(root, './/nfe:emit/nfe:enderEmit/nfe:UF', ns)
        cep_emit =  get_text_from_element(root, './/nfe:emit/nfe:enderEmit/nfe:CEP', ns)
        fone_emit = get_text_from_element(root, './/nfe:emit/nfe:enderEmit/nfe:fone', ns)
        regime_emit = get_text_from_element(root, './/nfe:emit/nfe:CRT', ns)
        cnpj_dest =  get_text_from_element(root, './/nfe:dest/nfe:CPF', ns)
        nome_dest = get_text_from_element(root, './/nfe:dest/nfe:xNome', ns)
        logradouro_dest = get_text_from_element(root, './/nfe:dest/nfe:enderDest/nfe:xLgr', ns)
        numero_dest = get_text_from_element(root, './/nfe:dest/nfe:enderDest/nfe:nro', ns)
        bairro_dest = get_text_from_element(root, './/nfe:dest/nfe:enderDest/nfe:xBairro', ns)
        municipio_dest = get_text_from_element(root, './/nfe:dest/nfe:enderDest/nfe:cMun', ns)
        uf_dest = get_text_from_element(root, './/nfe:dest/nfe:enderDest/nfe:UF', ns)
        indicador_dest =  get_text_from_element(root, './/nfe:dest/nfe:indIEDest', ns)
        valor_base_icms = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vBC', ns)
        valor_icms = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vICMS', ns)
        valor_icms_desonerado = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vICMSDeson', ns)
        valor_bc_st_icms = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vBCST', ns)
        icms_st = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vST', ns)
        vt_prod = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vProd', ns)
        vt_frete = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vFrete', ns)
        vt_seguro = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vSeg', ns)
        vt_desc = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vDesc', ns)
        vt_II = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vII', ns)
        vt_ipi = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vIPI', ns)
        vt_pis = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vPIS', ns)
        vt_cofins = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vCOFINS', ns)
        vt_despesas = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vOutro', ns)
        vt_nfe = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vNF', ns)
        vt_tributos = get_text_from_element(root, './/nfe:total/nfe:ICMSTot/nfe:vTotTrib', ns)
        modalidade_frete = get_text_from_element(root, './/nfe:transp/nfe:modFrete', ns)
        info_contr = get_text_from_element(root, './/nfe:infAdic/nfe:infCpl', ns)

        for det in root.findall('.//nfe:det', ns):  
            nItem = det.get('nItem')
            prod = det.find('nfe:prod', ns)
            imposto = det.find('nfe:imposto', ns)
            icms = imposto.find('.//nfe:ICMS', ns) if imposto is not None else None
            icms60 = icms.find('.//nfe:ICMS60', ns) if icms is not None else None
            icms00 = icms.find('.//nfe:ICMS00', ns) if icms is not None else None
            pis = imposto.find('.//nfe:PIS', ns) if imposto is not None else None
            pis_aliq = pis.find('.//nfe:PISAliq', ns) if pis is not None else None
            cofins = imposto.find('.//nfe:COFINS/nfe:COFINSAliq', ns) if imposto is not None else None



            cProd = prod.find('nfe:cProd', ns) if prod is not None else None
            cEAN = prod.find('nfe:cEAN', ns) if prod is not None else None
            xProd = prod.find('nfe:xProd', ns) if prod is not None else None
            NCM = prod.find('nfe:NCM', ns) if prod is not None else None
            vProd = prod.find('nfe:vProd', ns) if prod is not None else None
            cBenef = prod.find('nfe:cBenef', ns) if prod is not None else None
            EXTIPI = prod.find('nfe:EXTIPI', ns) if prod is not None else None
            CFOP = prod.find('nfe:CFOP', ns) if prod is not None else None
            unidade = prod.find('nfe:uCom', ns) if prod is not None else None
            quant = prod.find('nfe:qCom', ns) if prod is not None else None
            valor_un = prod.find('nfe:vUnCom', ns) if prod is not None else None
            cEANTrib = prod.find('nfe:cEANTrib', ns) if prod is not None else None

            vTotTrib = imposto.find('.//nfe:vTotTrib', ns) if imposto is not None else None
            vICMSEfet = imposto.find('.//nfe:ICMS60/nfe:vICMSEfet', ns) if imposto is not None else None

            origem = icms.find('nfe:orig', ns).text if icms is not None and icms.find('nfe:orig', ns) is not None else ''
            cst = icms.find('nfe:CST', ns).text if icms is not None and icms.find('nfe:CST', ns) is not None else ''
            modBC = icms.find('nfe:modBC', ns).text if icms is not None and icms.find('nfe:modBC', ns) is not None else ''
            pICMS = icms.find('nfe:pICMS', ns).text if icms is not None and icms.find('nfe:pICMS', ns) is not None else ''
            vBC = icms.find('nfe:vBC', ns).text if icms is not None and icms.find('nfe:vBC', ns) is not None else ''
            pICMS = icms.find('nfe:pICMS', ns).text if icms is not None and icms.find('nfe:pICMS', ns) is not None else ''
            vICMS = icms.find('nfe:vICMS', ns).text if icms is not None and icms.find('nfe:vICMS', ns) is not None else ''

            vBCEfet = icms60.find('nfe:vBCEfet', ns).text if icms60 is not None and icms60.find('nfe:vBCEfet', ns) is not None else ''
            pICMSEfet = icms60.find('nfe:pICMSEfet', ns).text if icms60 is not None and icms60.find('nfe:pICMSEfet', ns) is not None else ''
            vICMSEfet = icms60.find('nfe:vICMSEfet', ns).text if icms60 is not None and icms60.find('nfe:vICMSEfet', ns) is not None else ''
            vICMSSubstituto = icms60.find('nfe:vICMSSubstituto', ns).text if icms60 is not None and icms60.find('nfe:vICMSSubstituto', ns) is not None else ''
            vBC = icms00.find('nfe:vBC', ns).text if icms00 is not None and icms00.find('nfe:vBC', ns) is not None else ''

            cst_pis = pis_aliq.find('nfe:CST', ns).text if pis_aliq is not None and pis_aliq.find('nfe:CST', ns) is not None else ''
            vbc_pis = pis_aliq.find('nfe:vBC', ns).text if pis_aliq is not None and pis_aliq.find('nfe:vBC', ns) is not None else ''
            pp_pis = pis_aliq.find('nfe:pPIS', ns).text if pis_aliq is not None and pis_aliq.find('nfe:pPIS', ns) is not None else ''
            vp_pis = pis_aliq.find('nfe:vPIS', ns).text if pis_aliq is not None and pis_aliq.find('nfe:vPIS', ns) is not None else ''

            cst_cofins = cofins.find('nfe:CST', ns).text if cofins is not None else None
            vbc_cofins = cofins.find('nfe:vBC', ns).text if cofins is not None else None
            pcofins = cofins.find('nfe:pCOFINS', ns).text if cofins is not None else None
            vcofins = cofins.find('nfe:vCOFINS', ns).text if cofins is not None else None


            produto_data = {
                'Período': periodo,
                'Chave NF-e': chavenfe,
                "Situação": situacao,
                "Data Situação": periodo,
                "Natureza Operação": natureza,
                "Modelo": modelo,
                "Série": serie,
                "Número Documento": numero,
                "Data Emissão Documento": data_doc,
                "Tipo Operação": tipo,
                "Finalidade Emissão": finalidade,
                "Indicador Operação Consumidor": indicador,
                "Indicador Presença Comprador": indicador_presenca,
                "Tipo Emissão": tipo,
                "CNPJ/CPF Emitente": cnpj,
                "Inscrição Estadual Emitente": inscricao, 
                "Nome Emitente": nome_emit,
                "Logradouro Emitente": logradouro,
                "Número Emitente": numero_emit,
                "Bairro Emitente": bairro,
                "Município Emitente": municipio,
                "UF Emitente": uf_emit,
                "CEP Emitente": cep_emit,
                "Fone Emitente": fone_emit, 
                "Regime Tributário Emitente": regime_emit,
                "CNPJ/CPF Destinatário": cnpj_dest,
                "Nome Destinatário": nome_dest,
                "Logradouro Destinatário": logradouro_dest,
                "Número Destinatário": numero_dest,
                "Bairro Destinatário": bairro_dest,
                "Município Destinatário": municipio_dest,
                "UF Destinatário": uf_dest,
                "Indicador Contribuinte Destinatário": indicador_dest,
                "Vlr Total Base Cálculo ICMS": valor_base_icms,
                "Vlr Total ICMS": valor_icms, 
                "Vlr Total ICMS Desonerado": valor_icms_desonerado,
                "Vlr Total Base Cálculo ICMS ST": valor_bc_st_icms,
                "Vlr Total ICMS ST": icms_st,
                "Vlr Total Produto": vt_prod,
                "Vlr Total Frete": vt_frete,
                "Vlr Total Seguro": vt_seguro,
                "Vlr Total Desconto": vt_desc,
                "Vlr Total II": vt_II,
                "Vlr Total IPI": vt_ipi,
                "Vlr Total PIS": vt_pis,
                "Vlr Total Cofins": vt_cofins,
                "Vlr Total Outras Despesas": vt_despesas,
                "Vlr Total NF-e": vt_nfe,
                "Vlr Total Aproximado Tributos": vt_tributos,
                "Modalidade Frete": modalidade_frete,
                "Informação Adicional Contribuinte": info_contr,
                'Número Item': nItem,
                'Código Item': cProd.text if cProd is not None else '',
                'EAN': cEAN.text if cEAN is not None else '',
                'Descrição Item': xProd.text if xProd is not None else '',
                'NCM': NCM.text if NCM is not None else '',
                "Código Benefício": cBenef.text if cBenef is not None else '',
                "Ex TIPI": EXTIPI.text if EXTIPI is not None else '',
                "CFOP": CFOP.text if CFOP is not None else '',
                "Unidade Comercial": unidade.text if unidade is not None else '',
                "Qtde": quant.text if quant is not None else '',
                "Vlr Unitário": valor_un.text if valor_un is not None else '',
                'Vlr Total Produtos': vProd.text if vProd is not None else '',
                "EAN Tributável":  cEANTrib.text if cEANTrib is not None else '',
                "Origem CST ICMS": origem if origem is not None else '',
                "CST ICMS": cst if cst is not None else '',
                "Modalidade Base Cálculo ICMS": modBC if cst is not None else '',
                "Percentual Redução ICMS": pICMS if pICMS is not None else '',
                "Vlr Base Cálculo ICMS": vBC if vBC is not None else '',
                "Alíquota ICMS":pICMS if pICMS is not None else '',
                "Vlr ICMS Operação Diferimento": vICMS if vICMS is not None else '',
                "Vlr Base Cálculo Efetiva": vBCEfet if vBCEfet is not None else '',
                "Alíquota ICMS Efetiva": pICMSEfet if pICMSEfet is not None else '',
                "Vlr ICMS Efetivo": vICMSEfet if vICMSEfet is not None else '',
                "Vlr ICMS Substituto": vICMSSubstituto if vICMSSubstituto is not None else '',
                "Vlr Base Cálculo FCP": vBC if vBC is not None else '',
                "CST PIS": cst_pis if cst_pis is not None else '',
                "Vlr Base Cálculo PIS": vbc_pis if vbc_pis is not None else '',
                "Alíquota PIS": pp_pis if pp_pis is not None else '',
                "Vlr PIS": vp_pis if vp_pis is not None else '',
                "CST Cofins": cst_cofins  if cst_cofins is not None else '',
                "Vlr Base Cálculo Cofins":vbc_cofins  if vbc_cofins is not None else '',
                "Alíquota Cofins": pcofins  if pcofins is not None else '',
                "Vlr Cofins":vcofins  if vcofins is not None else '',
            }

            produtos.append(produto_data)

        return produtos

    except Exception as e:
        print(f"Erro ao processar o arquivo {xml_file}: {str(e)}")
        return None