import pandas as pd

#abertura do arquivo

sitelist = pd.read_excel("SiteList.xlsx")
results = pd.read_excel("Results.xlsx")

#conversão para dicionário, criação de uma lista auxiliar

sitesBuscados = sitelist.to_dict()
resultadosTestes = results.to_dict()

dicTodosSites = resultadosTestes["Site Name"].values()
listaTodosSites = list(dicTodosSites)

# a utilização do dicionário aqui facilita a busca dos sites no documento. no resto do código, listas se tornam mais convenientes.

sitesEncontrados = [sitesBuscados["Site Name"][x] for x in sitesBuscados["Site Name"] if sitesBuscados["Site Name"][x] in dicTodosSites]
foraDosResultados = [sitesBuscados["Site Name"][x] for x in sitesBuscados["Site Name"] if not sitesBuscados["Site Name"][x] in dicTodosSites]

#transferindo os dados apropriados para uma lista para ordená-la de acordo com os estados

categoriasResultados = list(results.columns.values)
ordenarEstados = []
contadorAux = 0

for x in sitesEncontrados:
    ordenarEstadosAux =[]
    foundIndex = listaTodosSites.index(x)
    for y in categoriasResultados:
        ordenarEstadosAux.append(resultadosTestes[y][foundIndex])
    ordenarEstados.append(ordenarEstadosAux)
    contadorAux += 1

ordenarEstados.sort(key = lambda x : x[categoriasResultados.index("State")])

#jogando a lista finalizada de volta para o DataFrame para melhor visualização

tabelaFinal = pd.DataFrame(ordenarEstados, columns=categoriasResultados)

tabelaComparacao = [results, tabelaFinal]
dadosComparacao = []

#preparando a exibição das estatísticas pedidas

for x in tabelaComparacao:
    dadosComparacaoAux = []
    dadosComparacaoAux.append(x["Alerts"].value_counts()["Yes"])
    dadosComparacaoAux.append(x["Quality (0-10)"].value_counts()[0])
    dadosComparacaoAux.append(x["Mbps"][x["Mbps"] > 80].count())
    dadosComparacaoAux.append(x["Mbps"][x["Mbps"] < 10].count())
    dadosComparacao.append(dadosComparacaoAux)

dadosComparacao[0].insert(0, "Dados Completos")
dadosComparacao[1].insert(0, "Dados Selecionados")

dadosComparados = pd.DataFrame(dadosComparacao, columns=["Amostra", "Alertas Totais", "Qualidade zero", "Velocidade > 80 Mbps", "Velocidade < 10 Mbps"])
print (tabelaFinal, "\n")
print (dadosComparados, "\n")
print (len(foraDosResultados), " valores não foram encontrados na base de dados fornecida.")
