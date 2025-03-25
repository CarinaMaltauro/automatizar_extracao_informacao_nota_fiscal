import os
import pandas as pd
import openpyxl
import xmltodict 

# Funções para abrir, ler e extrair informações da NF-e e NFS-e


def ler_xml_danfe(nota):
    with open(nota, 'rb') as arquivo: 
        # traduz arquivo xml para dicionário python
        documento = xmltodict.parse(arquivo) 
   
    # acessar infNFe dentro NFe dentro nfeProc, vide estrutura xml
    dic_notafiscal = documento['nfeProc']['NFe']['infNFe'] 

    # acessar informações dentro de infNFe
    # get foi colocado para o caso da nota não ter cpf ou cnpj, o que provocaria um erro; não tendo cpf, substituir pelo valor N/A
    valor_total = dic_notafiscal['total']['ICMSTot']['vNF']
    cnpj_vendeu = dic_notafiscal['emit'].get('CNPJ', 'N/A')
    nome_vendeu = dic_notafiscal['emit']['xNome']
    cpf_comprou = dic_notafiscal['dest'].get('CPF', 'N/A')
    nome_comprou = dic_notafiscal['dest']['xNome']
    produtos = dic_notafiscal['det']

    lista_produtos = []
    for produto in produtos:

        if 'vDesc' in produto['prod']:
            valor_desconto = produto['prod']['vDesc']
        else:
            valor_desconto = 0

        valor_produto = produto['prod']['vProd']
        nome_produto = produto['prod']['xProd']
        lista_produtos.append((nome_produto, valor_produto, valor_desconto))

    resposta = {
        'valor_total': [valor_total],
        'cnpj_vendeu': [cnpj_vendeu],
        'nome_vendeu': [nome_vendeu],
        'cpf_comprou': [cpf_comprou],
        'nome_comprou': [nome_comprou],
        'lista_produtos': [lista_produtos],
        'nota_fiscal':'NF-e/RJ'
    }
    return resposta


def ler_xml_servico(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    dic_notafiscal = documento['ConsultarNfseResposta']['ListaNfse']['CompNfse']['Nfse']['InfNfse']

    valor_total = dic_notafiscal['Servico']['Valores']['ValorServicos']
    cnpj_vendeu = dic_notafiscal['PrestadorServico']['IdentificacaoPrestador'].get('Cnpj', 'N/A')
    nome_vendeu = dic_notafiscal['PrestadorServico']['RazaoSocial']
    cpf_comprou = dic_notafiscal['TomadorServico']['IdentificacaoTomador']['CpfCnpj'].get('Cnpj', 'N/A')
    nome_comprou = dic_notafiscal['TomadorServico']['RazaoSocial']
    produtos = dic_notafiscal['Servico']['Discriminacao']

    resposta = {
        'valor_total': [valor_total],
        'cnpj_vendeu': [cnpj_vendeu],
        'nome_vendeu': [nome_vendeu],
        'cpf_comprou': [cpf_comprou],
        'nome_comprou': [nome_comprou],
        'lista_produtos': [produtos],
        'nota_fiscal': 'NFS-e/RJ'
    }
    return resposta




# Adicionar resultado NFs em um mesmo arquivo Excel

arquivos = os.listdir("NFs")

dataframes = []
for arquivo in arquivos:
     if 'xml' in arquivo:
         if 'DANFE' in arquivo:
             df = pd.DataFrame.from_dict(ler_xml_danfe(f"NFs/{arquivo}"))
         else:
             df = pd.DataFrame.from_dict(ler_xml_servico(f'NFs/{arquivo}'))
         dataframes.append(df)

# concatenar todos os dataframes em um só
df_final = pd.concat(dataframes, ignore_index=True)

df_final.to_excel('NFs.xlsx', index=False)
