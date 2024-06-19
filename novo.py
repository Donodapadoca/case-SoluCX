
import openpyxl
import pandas as pd


dados = pd.read_excel('clientes1.xlsx', sheet_name='clientes')
dados1 = pd.read_excel('auto.xlsx',sheet_name='auto')

workbook = openpyxl.Workbook()
planilha = workbook.active
dadospestado = dados.groupby('state')

#pega a coluna monthly_income e "salva" em salario
salario = 'monthly_income'
#pega a coluna state e "salva" em estados
estados = dados['state']
#separa os estados do sudeste
estadossud = ['SP', 'ES', 'MG','RJ']

estsud_data = estados[estados.isin(estadossud)]
#agrupa e faz a media dos salarios dos estados do sudeste
salariopestado = dados.groupby(estsud_data)[salario].mean()

maiorpestado = dadospestado['monthly_income'].max()
menorpestado = dadospestado['monthly_income'].min()


planilha.cell(row = 3, column = 2).value = "Estado"
planilha.cell(row = 4, column = 2).value = "Salario"

numero_coluna = 2
# Colocando a string Estado e Média em frente aos estados do sudeste
est = planilha.cell(row=3, column=1)
med = planilha.cell(row=4, column=1)
est.value = "Estados"
med.value = "Média"


#organiza as medias e os estados em linha por linha e coluna por coluna( apenas os estados do Sudeste)
for estado, mediasalario in salariopestado.items():
    planilha.cell(row = 3, column = numero_coluna).value = estado
    planilha.cell(row = 4, column = numero_coluna).value = mediasalario
    numero_coluna += 2

numero_coluna = 2

#coloca strings em frente aos dados para melhor entendimento
est = planilha.cell(row=7,column=1)
maior = planilha.cell(row=8, column=1)
menor = planilha.cell(row=11, column=1)
est.value = "Estados"
maior.value = "Maior"
menor.value = "Menor"


#organiza os maiores salarios por estado
for estado, maioestado in maiorpestado.items():
    planilha.cell(row = 7, column=numero_coluna).value = estado
    planilha.cell(row = 8, column=numero_coluna).value = maioestado
    numero_coluna +=1

numero_coluna = 2
est = planilha.cell(row=11,column=1)
est.value = "Estados"

#organiza os menores salarios por estado
for estado, menorestado in menorpestado.items():
    planilha.cell(row= 11, column=numero_coluna).value = estado
    planilha.cell(row=12, column=numero_coluna).value = menorestado
    numero_coluna += 1


p60mais = (dados['age']>60).mean() * 100

pformat = f"{p60mais:.2f}%"

mais60 = planilha.cell(row=15, column=1)
mais60.value = "Clientes 60+"
planilha ['B15'].value = pformat

#combina os" dados da planilha clientes com a planilha auto 
#utiliza o "how='left'" para garantir que todas as informaçoes dos clientes sejam mantidas, mesmo aqueles que não possuem carro
dadoscombinados = dados.merge(dados1, how='left', on='id_cliente')





#encontrado os carros com a marca ford
dadosford = dadoscombinados.query("auto_brand== 'Ford' ")

#agrupando os carros com a marca ford em cada cidade x utilizando o id_cliente
cidadecarro = dadosford.groupby('city')['id_cliente'].size()

cmaiscarros = cidadecarro.idxmax()
mcontagem = cidadecarro.max()

cidade = planilha.cell(row=18,column=1)
qtd = planilha.cell(row=19, column=1)
cidade.value = "Cidade"
qtd.value = "Quantidade"

#imprimindo na planiha a cidade que contém mais carros da marca Ford
city = planilha.cell(row=18, column=2)
qts = planilha.cell(row=19, column=2)
city.value = cmaiscarros
qts.value = mcontagem


#procurar clientes pertencentes a cidades do estado do RS
dadosrs = dados.query("state == 'RS' ")

#limpar todas as cidades duplicadas que provavelmente vão existir e contar elas
cunicas = dadosrs.drop_duplicates(subset='city')
ncidades = cunicas.shape[0]

numdecidades = planilha.cell(row=22, column=1)
numdecidades.value = "Qtd de Cidades"

cidadesn = planilha.cell(row=22, column=2)
cidadesn.value = ncidades


ccolum = 4
cstate = 5





cunicas = []

for row in planilha.iter_rows(min_row = 2):
    cname = row[ccolum - 1].value
    sname = row[cstate - 1].value
    

    if sname == "RS" and cname not in cunicas:
        cunicas.append(cname)

    
nsheet = workbook.create_sheet("Cidades - RS")

row_index = 1

for cdds in cunicas:
    nsheet.cell(row=row_index, column = 1).value = cdds
    row_index +=1
    



print(cdds)




















workbook.save('Dadosfinal.xlsx')





