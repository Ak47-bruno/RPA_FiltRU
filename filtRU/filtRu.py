from datetime import datetime, timedelta
from time import strptime
from tkinter import filedialog, messagebox, ttk
import re
from tkinter.ttk import Progressbar
import openpyxl
from tkinter import *
from pathlib import Path

import csv
import pandas as pd
# Cria a coluna
global srcfile


srcfile = openpyxl.load_workbook("C:/Users/HatchBack/Downloads/bruto.xlsx", read_only=False, keep_vba=False)

# excel = srcfile.get_sheet_by_name("Chamados_Infraestrutura_Created")
global path
excel = srcfile.active
excel.title = "Planilha1"
excel.insert_cols(15)
soma = 0
global quantidade
quantidade = 0
colaborador = ""

for row in excel.iter_rows():
    soma += 1
    colaborador = excel[f'N{soma}'].value
    excel['O1'] = str("Cargo")
    data_planilha = excel[f'F{soma}'].value
    data_anterior = str(datetime.now() - timedelta(days=3))
    data_local = str(datetime.now())
    data_local = data_local[:-16]
    data_planilha = data_planilha[:-26]
    data_anterior = data_anterior[:-16]

    # Verifica estado Aprovados e põe para Aberto.
    if excel[f'I{soma}'].value == "Aprovado":
        excel[f'I{soma}'] = str("Aberto")

    if excel[f'I{soma}'].value == "merged" and data_local == data_planilha:
        excel[f'I{soma}'] = str("Em verificação de Eficácia")
    elif excel[f'I{soma}'].value == "merged" and data_anterior >= data_planilha:
        excel[f'I{soma}'] = str("Fechado Segundo Nivel")
    elif excel[f'I{soma}'].value == "merged" and data_anterior < data_planilha:
        excel[f'I{soma}'] = str("Em verificação de Eficácia")

    if colaborador == "RV Instalações":
        excel[f'O{soma}'] = str("RV Instalações")
    elif colaborador == "Alex Figueiredo Ferreira" or colaborador == "Jessika Pontes":
        excel[f'O{soma}'] = str("Operações")
        excel[f'H{soma}'] = str("Infraestrutura")
    elif colaborador == "Adriano Cunha N" or colaborador == "Adriano Cunha" or colaborador == "Bryan Lima N" \
            or colaborador == "Bryan Lima" or colaborador == "Diego Souza N" or colaborador == "Diego Souza" \
            or colaborador == "Glauco Guimaraes N" or colaborador == "Glauco Guimaraes" \
            or colaborador == "Helenio  Fernandes" or colaborador == "Marcelo Alexandre Calacina Guimarães" \
            or colaborador == "Nathanael Carvalho Projetos" or colaborador == "Paollo Gomes N" \
            or colaborador == "Paollo Gomes" or colaborador == "Thiago Lima" or colaborador == "William Silva N" \
            or colaborador == "William Silva" or colaborador == "Richard Augusto Santos Sousa" \
            or colaborador == "Gustavo de Lima Bessa N" or colaborador == "Gustavo de Lima Bessa" \
            or colaborador == "Lucas de Melo" or colaborador == "Lucas de Melo N" \
            or colaborador == "Marcelo Guimarães":
        excel[f'O{soma}'] = str("Projetos")
    else:
        excel[f'O{soma}'] = str("Operações")


srcfile.save('C:/Users/HatchBack/Downloads/BaseInfra.xlsx')


#Comando para gerar o executavel com janela
# pyinstaller.exe --onefile --windowed --icon=ico.ico filtroRU.py

#Comando para gerar o executavel com script
# pyinstaller --onefile filtroRU.py
