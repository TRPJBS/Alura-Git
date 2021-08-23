#!python3

# Consolidação Base Segmentos

# Este arquivo .py deve estar dentro da MESMA pasta no computador local que as planilhas que ele recebe
# Este arquivo tem de insumo as planilhas: ERP Centro de Custo, ERP Vendas, ERP Veículo e Tabelas_Apoio
# Este arquivo gera uma planilha chamada BaseSegmentos.{data de geração}.Python.xlsx dentro da mesma pasta local
# Este script deve rodar antes do script de geração da BaseFrotas
# Tempo de execução: Aproximadamente 2.44 minutos

# Última atualização: 12.08.2021
##//////////////////////////////////////////////////////////////////////////////////////////////////////////

import pandas as pd
import re
import numpy as np
from datetime import date
import time
import os

start_time = time.time()

##//////////////////////////////////////////////////////////////////////////////////////////////////////////
# Leitura de objetos .xlsx

ERP_centro_de_custo = pd.read_excel("ERP_centrodecusto.xlsx")
ERP_vendas = pd.read_excel("ERP_vendas.xlsx", header=[1])
ERP_veiculo = pd.read_excel("ERP_veiculo.xlsx", header=[1])
cadastro_familia = pd.read_excel("Tabelas_Apoio.xlsx", sheet_name="CadastroFamilia")
base_outros = pd.read_excel("Tabelas_Apoio.xlsx", sheet_name="BaseOutros")
segmento_filiais = pd.read_excel("Tabelas_Apoio.xlsx", sheet_name="SegmentoFiliais")
placas_execoes = pd.read_excel("Tabelas_Apoio.xlsx", sheet_name="Placas Exceções")
base_segmentos = ERP_centro_de_custo.copy()

##//////////////////////////////////////////////////////////////////////////////////////////////////////////
## FUNUÇÕES ##

# Classifica o nível hierárquico da conta contábil de acordo com número de caracteres
def nivel(value):
    valuestr = str(value)
    length = len(valuestr)
    if length == 3:
        return 2
    elif length == 5:
        return 3
    elif length == 7:
        return 4
    elif length == 9:
        return 5
    elif length == 12:
        return 6   
    
# Função PROCV que retorna "-" para valores encontrados (esse retorno pode ser alterado no terceiro argumento)
def PROCV(Valor_Busca,Valor_Referencia,Valor_Retorno, Erro:str = ''):
    Verdadeiro = Valor_Retorno.loc[Valor_Referencia == Valor_Busca]
    if Verdadeiro.empty:
        return "-" if Erro == '' else Erro
    else:
        return Verdadeiro.tolist()[0]
    
# Encontra placa em modelo antigo e mercosul. 
# Se não encontrar a placa no nome, procura na lista de placas com exceção dentro da tabela de apoio
# Nessa tabela de placas com exceção, incluir somente placas onde o centro de custo não contém a placa
def placa(value):
    string2 = value.lower()
    match = re.search('[a-z]{3} ?(-|)? ?([0-9]{4}|[0-9][a-z][0-9]{2})', string2)
    if match:
        strmatch = match.group(0).replace("-","").replace(" ", "").replace("", "").upper()
        return strmatch
    else:
        return PROCV(value,placas_execoes["Nome"], placas_execoes["Placa"])

# Formatação de Conta Contábil
def Reduzido(value):
    valuestr = str(value)
    length = len(valuestr)
    if length < 9:
        return "-"
    else:
        fatia1 = valuestr[:1]
        fatia2 = valuestr[1:3]
        fatia3 = valuestr[3:5]
        fatia4 = valuestr[5:7]
        fatia5 = valuestr[7:9]
        return ("{}.{}.{}.{}.{}".format(fatia1, fatia2, fatia3, fatia4, fatia5))

# PROCV para encontrar Região
def Region(value):
    valuestr = str(value)
    if valuestr == "":
        return 1
    else:
        return PROCV(value, segmento_filiais["Numero Centro Custo"], segmento_filiais["Regional"])

##////////////////////////////////////////////////////////////////////////////////////////////////////////////
#CRIAÇÃO DE COLUNAS

base_segmentos["Placa"] = base_segmentos["Nome"].apply(placa).fillna("-")
base_segmentos["Nível"] = base_segmentos["Número Conta de Custo"].apply(nivel)
base_segmentos["Reduzido"] = base_segmentos["Número Conta de Custo"].apply(Reduzido)
base_segmentos["Regional"] = base_segmentos["Reduzido"].apply(Region)
base_segmentos["Status[PlanVendas]"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_vendas["Placa"], ERP_vendas["Status"]))
base_segmentos["Status[Cadastro]"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Status"])).fillna("-")
base_segmentos["Ativo[Cadastro]"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Ativo"])).fillna("-")
base_segmentos["Tipo Veículo"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Nome"])).fillna("-")
base_segmentos["Família"] = base_segmentos["Tipo Veículo"].apply(PROCV, args = (cadastro_familia["Nome"], cadastro_familia["Família"]))
base_segmentos["Início Outros"] = base_segmentos["Placa"].apply(PROCV, args = (base_outros["Placa"], base_outros["Início Outros"]))
base_segmentos["Status[Outros]"] = base_segmentos["Placa"].apply(PROCV, args = (base_outros["Placa"], base_outros["Status[Outros]"]))
base_segmentos["Status[Outros]Detalhe"] = base_segmentos["Placa"].apply(PROCV, args = (base_outros["Placa"], base_outros["Status[Outros]Detalhe"]))
base_segmentos["Contrato"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Contrato de manutenção"])).fillna("-")
base_segmentos["Supervisor"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Supervisor Placa"])).fillna("-")
base_segmentos["Núm.Car"]= base_segmentos["Número Conta de Custo"].astype(str).apply(len)
base_segmentos["Procv"] = base_segmentos["Código"].apply(PROCV, args = (segmento_filiais["Cód. Centro Custo"], segmento_filiais["Nome Área"]))
base_segmentos["Data Início de Operação[cadastro]"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Inicio Operação"])).fillna("-")
base_segmentos["Data Fim de Operação[cadastro]"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Fim Operação"])).fillna("-")
base_segmentos["Data Fim da Operação[Planvenda]"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_vendas["Placa"], ERP_vendas["Disponiblização"]))
base_segmentos["Ano Modelo"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Modelo"])).fillna("-")

#Quantas vezes a placa aparece na tabela
#Para registros que não têm placa (placa = "-") a função retorna a contagem de "-" na tabela
#Substituição da contagem de "-" por 0
base_segmentos["Duplicidade"] = base_segmentos.groupby("Placa")["Placa"].transform("count")
base_segmentos["Duplicidade"] = np.where(base_segmentos["Duplicidade"]>10,0,base_segmentos["Duplicidade"])

#Procv de placa na dbveículo erp foi com sucesso?
base_segmentos["procv placa veiculo"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Placa"])).fillna("-")


#SIM para cavalos ativos, em operação e que não constam no relatório de veículos do ERP
#NÃO para todos outros
base_segmentos["Status[Processo Novo]"]  = np.where(np.logical_and(base_segmentos["Ativo[Cadastro]"]=="1-Sim", np.logical_and(base_segmentos["Status[Cadastro]"]=="1-Em Operação",np.logical_and(base_segmentos["Família"]=="Cavalo",base_segmentos["procv placa veiculo"] == "-"))),"Sim","Não")

#Coluna duplicada somente para manter ordem de colunas originais
base_segmentos["Ccusto"] = base_segmentos["Código"]

#Compara nome do relatório centro de custo ERP com o mesmo na tabela Base Segmentos Filiais
condicoes = [(base_segmentos["Procv"]==base_segmentos["Nome"]),(base_segmentos["Procv"]=="-")]
resultados = ["igual", "-"]
base_segmentos["Validação[Nome Área c/ Base]"] = np.select(condicoes, resultados, default="diferente")

# Se CC Reduzido é vazio, busca Segmento pelo código ERP na tabela Base Segmentos Filiais
# Se CC Reduzido não é vazio, busca Segmento pelo CC Reduzido na tabela Base Segmentos Filiais
Segmento = []
index = 0
for i in base_segmentos["Reduzido"]:
    if i == "-":
        Segmento.append(PROCV(base_segmentos["Código"][index], segmento_filiais["Cód. Centro Custo"], segmento_filiais["Segmento"]))
    else:
        Segmento.append(PROCV(i, segmento_filiais["Numero Centro Custo"], segmento_filiais["Segmento"]))
    index+=1

# Se CC Reduzido é vazio, busca Filial pelo código ERP na tabela Base Segmentos Filiais
# Se CC Reduzido não é vazio, busca Filial pelo CC Reduzido na tabela Base Segmentos Filiais
Filial = []
index = 0
for i in base_segmentos["Reduzido"]:
    if i == "":
        Filial.append(PROCV(base_segmentos["Código"][index], segmento_filiais["Cód. Centro Custo"], segmento_filiais["Filial"]))
    else:
        Filial.append(PROCV(i, segmento_filiais["Numero Centro Custo"], segmento_filiais["Filial"]))
    index+=1
    
#Criação de colunas Segmento e Filial utilizando resultado das buscas acima
base_segmentos["Segmento"] = Segmento
base_segmentos["Filial"] = Filial

##///////////////////////////////////////////////////////////////////////////////////////////////////////////////
##CRIAÇÃO DE NOVAS COLUNAS PARA FACILITAR PROCESSOS DE VALIDAÇÃO


#Colunas serão utilizadas durante a criação da BaseFrota
base_segmentos["Codigo Tipo Veiculo"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Nome.1"])).fillna("-")
base_segmentos["Ano Fabricação"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Fabricação"])).fillna("-")
base_segmentos["procv placa veiculo"] = base_segmentos["Placa"].apply(PROCV, args = (ERP_veiculo["Placa"], ERP_veiculo["Placa"])).fillna("-")

#Verifica se veículo apresenta status na base outros E no relatório de veículos do ERP
base_segmentos["Status Vendas e Outros?"] = np.where(np.logical_and(base_segmentos["Status[Outros]"]!="-", base_segmentos["Status[PlanVendas]"]!="-"),"Sim","Não")

#Avalia se registro nível 9 não consta na base segmentos filiais
base_segmentos["Atualização BaseSF"] = np.where(np.logical_and(base_segmentos["Núm.Car"]==9, base_segmentos["Procv"]=="-"),"Sim","Não")

#Verifica se existe PLACA duplicada (sim, não, -)
condicoes2 = [(base_segmentos["Duplicidade"]==1),(base_segmentos["Duplicidade"]>10)]
resultados2 = ["Não", "-"]
base_segmentos["Placa Duplicada?"] = np.select(condicoes2, resultados2, default="Sim")

##///////////////////////////////////////////////////////////////////////////////////////////////////////////////
#Organizar colunas para ficar igual a base antiga

base_segmentos = base_segmentos[["Código", "Nome", "Número Conta de Custo", "Ativo", "Placa",
             "Nível", "Reduzido", "Regional", "Segmento", "Filial", "Status[Cadastro]",
             "Status[PlanVendas]", "Status[Processo Novo]", "Início Outros", "Status[Outros]",
             "Status[Outros]Detalhe", "Contrato", "Tipo Veículo", "Família", "Supervisor", 
             "Ativo[Cadastro]", "Núm.Car", "Procv", "Validação[Nome Área c/ Base]", "Ccusto",
             "Duplicidade", "Data Início de Operação[cadastro]", "Data Fim de Operação[cadastro]",
             "Data Fim da Operação[Planvenda]", "Ano Modelo", "Placa Duplicada?", 
             "Status Vendas e Outros?", "Atualização BaseSF", "Codigo Tipo Veiculo", "Ano Fabricação",
             "procv placa veiculo"]]

##///////////////////////////////////////////////////////////////////////////////////////////////////////////////
#Exportando DBCC para excel (base segmentos)

nome_arquivo = "BaseSegmentos."+str(date.today())+".Python.xlsx"
base_segmentos.to_excel(nome_arquivo, index=False)

#Tempo de execução impresso no final do run
tempoexec = ((time.time()-start_time)/60)
print("\nConsolidação realizada com sucesso!\n")
print("O Arquivo '", nome_arquivo, "' foi gerado na seguinte pasta: \n\n", os.getcwd())
print("\nTempo de execução: ", "{:.2f}".format(tempoexec), "minutos")
fim = input("Pressione 'Enter' para finalizar!")

## Alteração para git