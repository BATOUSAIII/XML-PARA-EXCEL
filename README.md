
import xml.etree.ElementTree as ET
import openpyxl

# Função para ler o XML da nota fiscal e extrair os dados
def ler_xml_nfe(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    # Namespaces que podem ser usados no XML (ajuste conforme necessário)
    ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}
    
    # Dados do emitente
    emitente = root.find('.//ns:emit/ns:xNome', ns).text
    destinatario = root.find('.//ns:dest/ns:xNome', ns).text

    # Lista para armazenar os dados dos produtos
    produtos = []

    for item in root.findall('.//ns:det', ns):
        produto = {}
        produto['nome'] = item.find('.//ns:prod/ns:xProd', ns).text
        produto['quantidade'] = item.find('.//ns:prod/ns:qCom', ns).text
        produto['valor'] = item.find('.//ns:prod/ns:vProd', ns).text
        produto['desconto'] = item.find('.//ns:prod/ns:vDesc', ns).text if item.find('.//ns:prod/ns:vDesc', ns) is not None else "0.00"
        
        # Valor Líquido = Valor do Produto - Desconto
        produto['valor_liquido'] = str(float(produto['valor']) - float(produto['desconto']))
        produtos.append(produto)
    
    return emitente, destinatario, produtos

# Função para preencher o Excel com os dados da nota fiscal
def preencher_excel(emitente, destinatario, produtos, output_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nota Fiscal"
    
    # Cabeçalhos
    ws.append(["Emitente", "Destinatário", "Produto", "Quantidade", "Valor", "Valor Líquido", "Desconto"])
    
    # Preencher dados
    for produto in produtos:
        ws.append([emitente, destinatario, produto['nome'], produto['quantidade'], produto['valor'], produto['valor_liquido'], produto['desconto']])
    
    # Salvar arquivo Excel
    wb.save(output_file)

# Exemplo de uso
if __name__ == "__main__":
    xml_file = '.\XML\procNFE52240810866276000100550010016132961452897206.xml'
    output_file = 'nota_fiscal.xlsx'

    emitente, destinatario, produtos = ler_xml_nfe(xml_file)
    preencher_excel(emitente, destinatario, produtos, output_file)

    print(f"Arquivo Excel '{output_file}' gerado com sucesso!")
