import csv
import re
import openpyxl

##Abrindo o .csv da apos conversao e pegando os dados. 
results = []
##MUDAR O NOME DO ARQUIVO, NO MOMENTO SÓ FUNCIONA PARA COPEL
##DEPOIS IREI DINAMIZAR
with open("2_520_70531.csv") as csvfile:
    reader = csv.reader(csvfile, quoting=csv.QUOTE_NONNUMERIC)
    for row in reader: 
        results.append(row)

#Tratando Unidade (in progress)
Unidade = results[1][results[0].index('UC')]
Unidade = Unidade.strip()

#Tratando CNPJ (in progress)
CNPJ = results[1][results[0].index('CNPJ')]
CNPJ = CNPJ.replace("CNPJ: ", "")
CNPJ = CNPJ.strip()

#Tratando ContaContrato (in progress)
ContaContrato = results[1][results[0].index('Conta_contrato')]

#Tratando a NF (in progress)
NF = results[1][results[0].index('Nota_fiscal')]
match = re.search(r'[0-9]{3,}', NF)
# Se houver correspondência
if match:
    # Extrai o valor
    NF = match.group(0)

#Tratando o Codigo de Barras, se houver (in progress)
CodigoDeBarras = results[1][results[0].index('Boleto_1')]
if CodigoDeBarras == "":
    CodigoDeBarras = results[1][results[0].index('Boleto_2')]
    if CodigoDeBarras == "":
      CodigoDeBarras = results[1][results[0].index('Boleto_3')]  

#Tratando Fisco
Fisco = results[1][results[0].index('Fisco')]
Fisco = Fisco.strip()

#Tratando Constante
Constante = results[1][results[0].index('Constante')]

#Tratamento de Leituras (in progress)
LeituraAtual = results[1][results[0].index('Leituras_1')]
LeituraAnterior = results[1][results[0].index('Leituras_1')]

match = re.search(r'[0-9,.]{3,5}\s+[0-9,.]{3,5}', LeituraAnterior)
# Se houver correspondência
if match:
    # Extrai o valor correspondente
    indice = match.group(0).split()
    LeituraAnterior = indice[0]
    LeituraAtual = indice[1]

#Tratando Chave de Acesso
ChaveDeAcesso = results[1][results[0].index('QR_code')]
match = re.search(r'Chave\s+de\s+Acesso\s+[0-9]+\s+[0-9]+\s+[0-9]+\s+[0-9]+\s+[0-9]+\s+[0-9]+\s+[0-9]+\s+[0-9]+\s+[0-9]+\s+[0-9]+\s+[0-9]+', ChaveDeAcesso)
# Se houver correspondência
if match:
    # Extrai o valor
    indice = match.group(0)
    ChaveDeAcesso = re.sub(r"Chave de Acesso ", "", indice)


#Tratando o Consumo (pego das leituras)
ConsumoFaturado = results[1][results[0].index('Consumo_2')]

indice = results[1][results[0].index('Tarifas_1')]
indice = indice.split()
BaseIcms = indice[1]
AliqIcms = indice[2]
Icms = indice[3]


indice = results[1][results[0].index('Tarifas_3')]
indice = indice.split()
BasePisCofins = indice[1]
AliqPis = indice[2]
Pis = indice[3]

indice = results[1][results[0].index('Tarifas_4')]
indice = indice.split()
AliqCofins = indice[1]
Cofins = indice[2]


Consumo = 0
ConsumoInjetado = 0
CIP = 0

TotalFatura = results[1][results[0].index('Total_fatura')]
for i in range(1, 16):
    variavel = f"Mais_{i}"
    indice = results[0].index(variavel)  # Obter o índice de "Mais_$i" em results[0]
    valor = results[1][indice] 

    match = re.search(r'ENERGIA\s+ELET\s+CONSUMO\s+kWh\s+[0-9.,]+', valor)
    # Se houver correspondência
    if match:
        # Extrai o valor correspondente
        pivot = match.group(0).split()
        Consumo = Consumo + int(pivot[4])


    match = re.search(r'ENERGIA\s+INJ.\s+OUC\s+MPT\s+TUSD\s+[0-9]{2}\/[0-9]{4}\s+kWh\s+[0-9.,-]+\s+[0-9.,]+\s+[0-9.,-]+\s+[0-9.,-]+\s+[0-9.,-]+', valor)
    # Se houver correspondência
    if match:
        # Extrai o valor correspondente
        pivot = match.group(0).split()
        ## COLOCAR A POSICAO CORRETA QUANDO TIVER INJ
        ConsumoInjetado = Consumo + int(pivot[4])

    
    match = re.search(r'CONT\s+ILUMIN\s+PUBLICA\s+MUNICIPIO\s+UN\s+[0-9]+\s+[0-9.,]+\s+[0-9.,]+', valor)
    # Se houver correspondência
    if match:
        # Extrai o valor correspondente
        pivot = match.group(0).split()
        CIP = pivot[7]

        

#Criando a Planilha
wb = openpyxl.Workbook()
planilha = wb.active

headers = ['Unidade', 'CNPJ', 'Conta Contrato', 'Nota Fiscal', 
           'CodigoDeBarras', 'Fisco', 'Constante', 'Leitura Anterior', 
           'Leitura Atual', 'Chave De Acesso', 'Consumo Faturado', 'Base de ICMS', 
           'Base de PIS/COFINS', 'Alíquota ICMS', 'ICMS', 'Alíquota PIS', 'PIS', 
           'Alíquota COFINS', 'COFINS', 'CIP', 'Consumo', 'Consumo Injetado','Total Fatura']
planilha.append(headers)

# Coloca os dados em cada linha
for row_index, row in enumerate(results):
    if row_index > 0:
        data = [Unidade, CNPJ, ContaContrato, NF, 
                CodigoDeBarras, Fisco, Constante, 
                LeituraAnterior, LeituraAtual, 
                ChaveDeAcesso, ConsumoFaturado, BaseIcms, BasePisCofins,
                AliqIcms, Icms, AliqPis, Pis, AliqCofins, Cofins, CIP, Consumo, 
                ConsumoInjetado, TotalFatura]
        planilha.append(data)

wb.save('dados_extraidos.csv')

print("Arquivo gerado com sucesso!")


