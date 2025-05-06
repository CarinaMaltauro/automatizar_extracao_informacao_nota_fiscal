<h1 align="center"> Projeto - Automatizar Extração de Dados de Notas Fiscais  <img src="https://cdn-icons-png.freepik.com/256/14477/14477586.png?ga=GA1.1.763163565.1742925562&semt=ais_hybrid" width="25" height="25"> </h1>

<p align="center"> Carina R. P. M Dias</p>


## Objetivo

No Brasil, cada estado pode adotar padrões distintos para notas fiscais eletrônicas, o que afeta diretamente a estrutura dos arquivos XML. Além disso, a Nota Fiscal Eletrônica (NF-e) e a Nota Fiscal de Serviço Eletrônica (NFS-e) possuem formatos diferentes entre si.

Este projeto tem como objetivo automatizar a leitura e extração de dados de arquivos XML de NF-e e NFS-e do Estado do Rio de Janeiro. A estrutura do código foi desenvolvida de forma modular, permitindo a adição de novos padrões e layouts de outros estados e municípios com facilidade.

A automação realiza a leitura dos arquivos XML, identifica o tipo de nota e extrai informações relevantes, como:

Razão social e CNPJ do emitente

Data de emissão

Itens ou serviços descritos

Valores totais e impostos (ICMS, ISS, retenções)

Os dados extraídos são organizados e registrados em uma planilha Excel, facilitando:

O armazenamento e a consulta rápida de informações financeiras

A análise de custos, fornecedores e tributos

A escalabilidade do controle fiscal, mesmo com grandes volumes de documentos

Essa ferramenta é útil para escritórios contábeis, empresas e profissionais que lidam com grandes quantidades de notas fiscais e precisam otimizar o processo de análise e arquivamento.
