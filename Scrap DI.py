import openpyxl
import pandas as pd
import numpy as np

path = "C:\\Users\\Dell\\Desktop\\Estudos 2\\Finanças\\Derivativos\\PlanilhaDeJuros.xlsm"
wb_DI = openpyxl.load_workbook(path, data_only=True)
wb_DI_aba_DI= wb_DI["Taxas de DI"]
Taxas_DI =[]
Contratos_DI = []
Vencimentos_DI = []
for column in wb_DI_aba_DI.iter_cols():
    column_name = column[1].value
    if column_name == "Taxa de Fechamento":
        for cell in column:
            Taxas_DI.append(cell.value)
    column_name = column[1].value
    if column_name == "Contratos":
        for cell in column:
            Contratos_DI.append(cell.value)
    if column_name == "Vencimentos":
        for cell in column:
            Vencimentos_DI.append(cell.value)                

Taxas_DI.remove(None)
Taxas_DI.remove("Taxa de Fechamento")
Contratos_DI.remove(None)
Contratos_DI.remove("Contratos")
Vencimentos_DI.remove(None)
Vencimentos_DI.remove("Vencimentos")

wb_DI_aba_Datas_COPOM = wb_DI["Datas COPOM"]
Datas_COPOM=[]
for data in wb_DI_aba_Datas_COPOM:
    Datas = data[0].value
    if Datas == None:
        continue
    Datas_COPOM.append(Datas)

Datas_COPOM.remove("Calendario de reunioes do COPOM")
df_Datas_COPOM = pd.DataFrame(Datas_COPOM, columns = ["Datas"])

dict_DI = {"Datas": Vencimentos_DI, "Eventos": Contratos_DI,"Taxas Spot": Taxas_DI}
df_DI_Base = pd.DataFrame(dict_DI)
df_DI = df_DI_Base.append(df_Datas_COPOM, ignore_index = True)
df_DI.sort_values(by="Datas", ascending = True, inplace = True)

df_DI["Eventos"].fillna("Copom", inplace = True)
df_DI_2 = df_DI.reset_index(drop=True)

wb_DI_Parametrizacao = wb_DI["Parametrização"]

Data_Ref = []
for data in wb_DI_Parametrizacao:
    Data = data[1].value
    if Data == None:
        continue
    Data_Ref.append(Data)

Data_Ref.remove("Data de Referencia")

wb_DI_Feriados = wb_DI["Feriados"]

Feriados = []
for data in wb_DI_Feriados:
    Datas = data[0].value
    if Datas == None:
        continue
    Feriados.append(Datas)

Feriados.remove("Column1")

Feriados_np = np.array(Feriados, dtype = 'datetime64[D]')
Data_Ref_np = np.array(Data_Ref, dtype = 'datetime64[D]')

NDU_1 = np.busday_count(Data_Ref_np, df_DI_2["Datas"][0].date(), holidays = Feriados_np)

lista_indices = df_DI_2.index.values.tolist()

NDU_defacto = []
for indices in lista_indices:
    valor = indices
    NDU = np.busday_count(Data_Ref_np, df_DI_2["Datas"][valor].date(), holidays = Feriados_np)
    NDU_defacto.append(NDU)

df_DI_2.insert(2, "NDU", NDU_defacto, allow_duplicates=True)

Taxa_Selic_Over=[]
for data in wb_DI_Parametrizacao:
    Selic = data[0].value
    if Selic == None:
        continue
    Taxa_Selic_Over.append(Selic)

Taxa_Selic_Over.remove("Taxa Selic Over")

Taxa_Selic_Over_2 = [100*x for x in Taxa_Selic_Over]

Selic_Hoje_dict = {"Datas": Data_Ref, "Eventos": "Selic Hoje", "NDU": 1, "Taxas Spot": Taxa_Selic_Over_2}
Selic_Hoje_DataFrame = pd.DataFrame(Selic_Hoje_dict)
df_DI_3 = df_DI_2.append(Selic_Hoje_DataFrame, ignore_index = True)
df_DI_3.sort_values(by="Datas", ascending = True, inplace = True)
df_DI_4 = df_DI_3.reset_index(drop=True)

indices_copom_removidos = []
for indices in lista_indices:
    if df_DI_4.loc[indices, "NDU"] <= 0:
        indices_copom_removidos.append(indices)

df_DI_5 = df_DI_4.drop(index=indices_copom_removidos)
df_DI_6 = df_DI_5.reset_index(drop=True)

df_DI_6["Taxas Spot"] = df_DI_6["Taxas Spot"].fillna(0)

indices_interpolados = df_DI_6.index[df_DI_6["Taxas Spot"].eq(0)].tolist()

indices_taxas_anteriores = [x - 1 for x in indices_interpolados]
indices_taxas_posteriores = [x + 1 for x in indices_interpolados]

taxa_anterior = df_DI_6.loc[indices_taxas_anteriores, "Taxas Spot"]
taxa_posterior = df_DI_6.loc[indices_taxas_posteriores, "Taxas Spot"]

NDU_anterior = df_DI_6.loc[indices_taxas_anteriores, "NDU"]
NDU_posterior = df_DI_6.loc[indices_taxas_posteriores, "NDU"]
NDU_hoje = df_DI_6.loc[indices_interpolados, "NDU"]

for (indices,taxaTMenos1,taxaTMais1,NDUMenos1,NDUMais1,NDUHoje) in zip(indices_interpolados,taxa_anterior,taxa_posterior,NDU_anterior,NDU_posterior,NDU_hoje):
    taxa_interpolada = (pow(pow((1+taxaTMenos1/100),(NDUMenos1/252))*pow(pow((1+taxaTMais1/100),(NDUMais1/252))/pow((1+taxaTMenos1/100),(NDUMenos1/252)),((NDUHoje-NDUMenos1)/(NDUMais1-NDUMenos1))),(252/NDUHoje))-1)*100
    df_DI_6.loc[indices, "Taxas Spot"] = taxa_interpolada

def Criador_De_Colunas(n):
    lista_zeros = [0]*n
    return lista_zeros

lista_indices_2 = df_DI_6.index.values.tolist()

Taxas_Termo_DI = Criador_De_Colunas(len(lista_indices_2))
df_DI_6.insert(4, "Taxas Termo", Taxas_Termo_DI, allow_duplicates=True)

lista_indices_2_termo = df_DI_6.index.values.tolist()
lista_indices_2_termo.remove(0)

lista_indices_2_termo_anterior = [x - 1 for x in lista_indices_2_termo]
taxa_inicio_DI = df_DI_6.loc[lista_indices_2_termo_anterior, "Taxas Spot"]
taxa_inicio_DI_lista = taxa_inicio_DI.values.tolist()
taxa_fim_DI = df_DI_6.loc[lista_indices_2_termo, "Taxas Spot"]
taxa_fim_DI_lista = taxa_fim_DI.values.tolist()

NDU_inicio_DI = df_DI_6.loc[lista_indices_2_termo_anterior, "NDU"]
NDU_inicio_DI_lista = NDU_inicio_DI.values.tolist()
NDU_fim_DI = df_DI_6.loc[lista_indices_2_termo, "NDU"]
NDU_fim_DI_lista = NDU_fim_DI.values.tolist()

for indices in lista_indices_2:
    if indices == 0:
        df_DI_6.loc[0, "Taxas Termo"] = Taxa_Selic_Over_2
        for (indices_termo,taxa_inicio_DI,taxa_fim_DI,NDU_inicio_DI,NDU_fim_DI) in zip(lista_indices_2_termo,taxa_inicio_DI_lista,taxa_fim_DI_lista,NDU_inicio_DI_lista,NDU_fim_DI_lista):
            taxa_termo = (pow(pow((1+taxa_fim_DI/100),(NDU_fim_DI/252))/pow((1+taxa_inicio_DI/100),(NDU_inicio_DI/252)),(252/(NDU_fim_DI - NDU_inicio_DI))) - 1)*100
            df_DI_6.loc[indices_termo,"Taxas Termo"] = taxa_termo

print(df_DI_6)

Precificacao_Pol_Monetaria = Criador_De_Colunas(len(lista_indices_2))
df_DI_6.insert(5, "Precificacao_Pol_Monetaria", Precificacao_Pol_Monetaria, allow_duplicates=True)

termo_fim = df_DI_6.loc[lista_indices_2_termo, "Taxas Termo"]
termo_inicio = df_DI_6.loc[lista_indices_2_termo_anterior, "Taxas Termo"]

for indices in lista_indices_2:
    if indices == 0:
        df_DI_6.loc[0, "Precificacao_Pol_Monetaria"] = 0
        for (indice,termo_fim,termo_inicio) in zip(lista_indices_2_termo, termo_fim, termo_inicio):
            Precificacao = round((termo_fim - termo_inicio)*100,2)
            df_DI_6.loc[indice,"Precificacao_Pol_Monetaria"] = Precificacao

pd.set_option('display.max_columns', None)