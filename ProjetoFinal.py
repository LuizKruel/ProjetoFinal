"""
ProjetoFinal
Descrição: programa para extrair um arquivo csv da Web (Portal de Dados Abertos do TCE-RS)
Autor: Luiz Kruel
Data: 19/09/2022
Versão: 0.0.1
"""

# importa pacotes

import requests
import pandas as pd
from openpyxl import Workbook, load_workbook


# abre site TCE
def catch(site):
    dados = requests.get(site)
    return dados

# grava arquivo com as informações da variável "dados"
def save(dados):
    with open('balancete.csv','wb') as arquivo:
        for texto in dados.iter_content():
            arquivo.write(texto)
        arquivo.close()

# usa Pandas para ler o arquivo "balancete.csv" e Openpyxl para transformar em .xlsx
def xlsx(dados):
    balancete = pd.read_csv('balancete.csv')
    balancete.to_excel("balancete.xlsx", sheet_name="despesas", index = False)
    novo_balancete = load_workbook(filename = 'balancete.xlsx')
    novo_balancete.save("Novo Balancete.xlsx")

# módulo principal
def main():
    site = "http://dados.tce.rs.gov.br/dados/municipal/balancete-despesa/2022.csv" 
    dados = catch(site)
    save(dados)
    xlsx(dados)
    print("Balancetes salvos com sucesso!")
    input("Tecle Enter para Finalizar.")

if __name__ == "__main__":
    main()
