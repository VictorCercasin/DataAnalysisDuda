import pandas as pd
import numpy as np
import os
from pathlib import Path
from dataclasses import dataclass
from openpyxl import load_workbook
from typing import Any, Literal, TypeAlias
from openpyxl.utils import get_column_letter, column_index_from_string





Dia: TypeAlias = Literal[0, 7, 14]
Concentracao: TypeAlias = Literal[0, 25, 50, 75, 100, -1]
Sistema: TypeAlias = Literal['AI', 'A', 'BI', 'B', 'CN', 'CP']

@dataclass
class Amostra:
    fungo: str
    sistema: Sistema
    dia: Dia
    concentracao: Concentracao
    analise: str
    valor: Any




class Extrator:
    def __init__(self, file_path: str, sheet: str, fungo: str = ''):
        print("Iniciando extrator")
        self.file_path = file_path
        self.fungo = fungo if len(fungo) > 0 else self.file_path.stem
        self._amostras: list[Amostra] = []
        self.erro = ""
        if os.path.exists(file_path):
            self.wb = load_workbook(file_path, data_only=True)
        else:
            print(f"ERRO - Caminho {file_path} não existe")
            self.erro = "ERRO - Caminho {file_path} não existe"
            return
        self.ws= self.wb[sheet]

    def extrair_ufc(self):
        return []


    def extrair_base_seca(self):
        if len (self.erro) > 0:
            return []
        # AI
        self.extrair_celula(start_col_letter='g', start_row=3, dia=14, sistema='AI', analise='Massa Fungica', concentracao=0)
        self.extrair_celula(start_col_letter='g', start_row=5, dia=14, sistema='AI', analise='Massa Fungica', concentracao=25)
        self.extrair_celula(start_col_letter='g', start_row=7, dia=14, sistema='AI', analise='Massa Fungica', concentracao=50)
        self.extrair_celula(start_col_letter='g', start_row=9, dia=14, sistema='AI', analise='Massa Fungica', concentracao=75)
        self.extrair_celula(start_col_letter='g', start_row=11, dia=14, sistema='AI', analise='Massa Fungica', concentracao=100)

        # A
        self.extrair_celula(start_col_letter='g', start_row=13, dia=14, sistema='A', analise='Massa Fungica', concentracao=0)
        self.extrair_celula(start_col_letter='g', start_row=15, dia=14, sistema='A', analise='Massa Fungica', concentracao=25)
        self.extrair_celula(start_col_letter='g', start_row=17, dia=14, sistema='A', analise='Massa Fungica', concentracao=50)
        self.extrair_celula(start_col_letter='g', start_row=19, dia=14, sistema='A', analise='Massa Fungica', concentracao=75)
        self.extrair_celula(start_col_letter='g', start_row=21, dia=14, sistema='A', analise='Massa Fungica', concentracao=100)

        # BI
        self.extrair_celula(start_col_letter='g', start_row=23, dia=14, sistema='BI', analise='Massa Fungica', concentracao=0)
        self.extrair_celula(start_col_letter='g', start_row=25, dia=14, sistema='BI', analise='Massa Fungica', concentracao=25)
        self.extrair_celula(start_col_letter='g', start_row=27, dia=14, sistema='BI', analise='Massa Fungica', concentracao=50)
        self.extrair_celula(start_col_letter='g', start_row=29, dia=14, sistema='BI', analise='Massa Fungica', concentracao=75)
        self.extrair_celula(start_col_letter='g', start_row=31, dia=14, sistema='BI', analise='Massa Fungica', concentracao=100)

        # B
        self.extrair_celula(start_col_letter='g', start_row=33, dia=14, sistema='B', analise='Massa Fungica', concentracao=0)
        self.extrair_celula(start_col_letter='g', start_row=35, dia=14, sistema='B', analise='Massa Fungica', concentracao=25)
        self.extrair_celula(start_col_letter='g', start_row=37, dia=14, sistema='B', analise='Massa Fungica', concentracao=50)
        self.extrair_celula(start_col_letter='g', start_row=39, dia=14, sistema='B', analise='Massa Fungica', concentracao=75)
        self.extrair_celula(start_col_letter='g', start_row=41, dia=14, sistema='B', analise='Massa Fungica', concentracao=100)


    def extrair_analise_macroscopica(self, dia: Dia):
        if len (self.erro) > 0:
            return []
        # Sementes germinadas
        self.extrair_coluna(start_col_letter='d', start_row=4, dia=dia, sistema='AI', analise='Sementes Germinadas')
        self.extrair_coluna(start_col_letter='d', start_row=9, dia=dia, sistema='A', analise='Sementes Germinadas')
        self.extrair_coluna(start_col_letter='d', start_row=14, dia=dia, sistema='BI', analise='Sementes Germinadas')
        self.extrair_coluna(start_col_letter='d', start_row=19, dia=dia, sistema='B', analise='Sementes Germinadas')

        self.extrair_celula(start_col_letter='d', start_row=24, dia=dia, sistema='CN', analise='Sementes Germinadas', concentracao=-1)
        self.extrair_celula(start_col_letter='d', start_row=25, dia=dia, sistema='CP', analise='Sementes Germinadas', concentracao=-1)

        # TG
        self.extrair_coluna(start_col_letter='e', start_row=4, dia=dia, sistema='AI', analise='TG')
        self.extrair_coluna(start_col_letter='e', start_row=9, dia=dia, sistema='A', analise='TG')
        self.extrair_coluna(start_col_letter='e', start_row=14, dia=dia, sistema='BI', analise='TG')
        self.extrair_coluna(start_col_letter='e', start_row=19, dia=dia, sistema='B', analise='TG')

        self.extrair_celula(start_col_letter='e', start_row=24, dia=dia, sistema='CN', analise='TG', concentracao=-1)
        self.extrair_celula(start_col_letter='e', start_row=25, dia=dia, sistema='CP', analise='TG', concentracao=-1)

        # GRS
        self.extrair_coluna(start_col_letter='f', start_row=4, dia=dia, sistema='AI', analise='GRS')
        self.extrair_coluna(start_col_letter='f', start_row=9, dia=dia, sistema='A', analise='GRS')
        self.extrair_coluna(start_col_letter='f', start_row=14, dia=dia, sistema='BI', analise='GRS')
        self.extrair_coluna(start_col_letter='f', start_row=19, dia=dia, sistema='B', analise='GRS')

        self.extrair_celula(start_col_letter='f', start_row=24, dia=dia, sistema='CN', analise='GRS', concentracao=-1)
        self.extrair_celula(start_col_letter='f', start_row=25, dia=dia, sistema='CP', analise='GRS', concentracao=-1)

        # IG
        self.extrair_coluna(start_col_letter='g', start_row=4, dia=dia, sistema='AI', analise='IG')
        self.extrair_coluna(start_col_letter='g', start_row=9, dia=dia, sistema='A', analise='IG')
        self.extrair_coluna(start_col_letter='g', start_row=14, dia=dia, sistema='BI', analise='IG')
        self.extrair_coluna(start_col_letter='g', start_row=19, dia=dia, sistema='B', analise='IG')

        self.extrair_celula(start_col_letter='g', start_row=24, dia=dia, sistema='CN', analise='IG', concentracao=-1)
        self.extrair_celula(start_col_letter='g', start_row=25, dia=dia, sistema='CP', analise='IG', concentracao=-1)

        # CRR
        self.extrair_coluna(start_col_letter='h', start_row=4, dia=dia, sistema='AI', analise='CRR')
        self.extrair_coluna(start_col_letter='h', start_row=9, dia=dia, sistema='A', analise='CRR')
        self.extrair_coluna(start_col_letter='h', start_row=14, dia=dia, sistema='BI', analise='CRR')
        self.extrair_coluna(start_col_letter='h', start_row=19, dia=dia, sistema='B', analise='CRR')

        self.extrair_celula(start_col_letter='h', start_row=24, dia=dia, sistema='CN', analise='CRR', concentracao=-1)
        self.extrair_celula(start_col_letter='h', start_row=25, dia=dia, sistema='CP', analise='CRR', concentracao=-1)

        # IGN
        self.extrair_coluna(start_col_letter='i', start_row=4, dia=dia, sistema='AI', analise='IGN')
        self.extrair_coluna(start_col_letter='i', start_row=9, dia=dia, sistema='A', analise='IGN')
        self.extrair_coluna(start_col_letter='i', start_row=14, dia=dia, sistema='BI', analise='IGN')
        self.extrair_coluna(start_col_letter='i', start_row=19, dia=dia, sistema='B', analise='IGN')

        self.extrair_celula(start_col_letter='i', start_row=24, dia=dia, sistema='CN', analise='IGN', concentracao=-1)
        self.extrair_celula(start_col_letter='i', start_row=25, dia=dia, sistema='CP', analise='IGN', concentracao=-1)

        # IER
        self.extrair_coluna(start_col_letter='j', start_row=4, dia=dia, sistema='AI', analise='IER')
        self.extrair_coluna(start_col_letter='j', start_row=9, dia=dia, sistema='A', analise='IER')
        self.extrair_coluna(start_col_letter='j', start_row=14, dia=dia, sistema='BI', analise='IER')
        self.extrair_coluna(start_col_letter='j', start_row=19, dia=dia, sistema='B', analise='IER')

        self.extrair_celula(start_col_letter='j', start_row=24, dia=dia, sistema='CN', analise='IER', concentracao=-1)
        self.extrair_celula(start_col_letter='j', start_row=25, dia=dia, sistema='CP', analise='IER', concentracao=-1)




        for a in self._amostras:
            if a.valor == '#DIV/0!':
                a.valor = 0




    def extrair_metais_toxicos(self):
        if len (self.erro) > 0:
            return []
        # Apenas as concentrações do zinco ultrapassam os mínimos de sensibilidade do equipamento

        # Dia 0
        self.extrair_coluna(start_col_letter='i', start_row=5, dia=0, sistema='AI', analise='Zn')
        self.extrair_coluna(start_col_letter='i', start_row=10, dia=0, sistema='A', analise='Zn')
        self.extrair_coluna(start_col_letter='i', start_row=15, dia=0, sistema='BI', analise='Zn')
        self.extrair_coluna(start_col_letter='i', start_row=20, dia=0, sistema='B', analise='Zn')

        # Dia 14
        self.extrair_coluna(start_col_letter='o', start_row=5, dia=14, sistema='AI', analise='Zn')
        self.extrair_coluna(start_col_letter='o', start_row=10, dia=14, sistema='A', analise='Zn')
        self.extrair_coluna(start_col_letter='o', start_row=15, dia=14, sistema='BI', analise='Zn')
        self.extrair_coluna(start_col_letter='o', start_row=20, dia=14, sistema='B', analise='Zn')



    
    

    def extrair_fisicoquimico(self):
        if len (self.erro) > 0:
            return []
        # Análise pH
        self.extrair_quadrante(start_col_letter='d', start_row=5, sistema='AI', analise='pH')
        self.extrair_quadrante(start_col_letter='d', start_row=11, sistema='A', analise='pH')
        self.extrair_quadrante(start_col_letter='d', start_row=17, sistema='BI', analise='pH')
        self.extrair_quadrante(start_col_letter='d', start_row=23, sistema='B', analise='pH')

        # Análise turbidez 
        self.extrair_quadrante(start_col_letter='h', start_row=5, sistema='AI', analise='Turbidez')
        self.extrair_quadrante(start_col_letter='h', start_row=11, sistema='A', analise='Turbidez')
        self.extrair_quadrante(start_col_letter='h', start_row=17, sistema='BI', analise='Turbidez')
        self.extrair_quadrante(start_col_letter='h', start_row=23, sistema='B', analise='Turbidez')

        # Análise Condutividade
        self.extrair_quadrante(start_col_letter='l', start_row=5, sistema='AI', analise='Condutividade')
        self.extrair_quadrante(start_col_letter='l', start_row=11, sistema='A', analise='Condutividade')
        self.extrair_quadrante(start_col_letter='l', start_row=17, sistema='BI', analise='Condutividade')
        self.extrair_quadrante(start_col_letter='l', start_row=23, sistema='B', analise='Condutividade')

        # Análise Cor
        self.extrair_quadrante(start_col_letter='p', start_row=5, sistema='AI', analise='Cor')
        self.extrair_quadrante(start_col_letter='p', start_row=11, sistema='A', analise='Cor')
        self.extrair_quadrante(start_col_letter='p', start_row=17, sistema='BI', analise='Cor')
        self.extrair_quadrante(start_col_letter='p', start_row=23, sistema='B', analise='Cor')

        # Análise DQO
        self.extrair_quadrante(start_col_letter='t', start_row=5, sistema='AI', analise='DQO')
        self.extrair_quadrante(start_col_letter='t', start_row=11, sistema='A', analise='DQO')
        self.extrair_quadrante(start_col_letter='t', start_row=17, sistema='BI', analise='DQO')
        self.extrair_quadrante(start_col_letter='t', start_row=23, sistema='B', analise='DQO')


    def extrair_celula(self, start_col_letter: str, start_row: int,dia: int, sistema: str, analise: str, concentracao: int):
        ncols = 1
        nrows = 1

        print(f"Extraindo célula. Sistema: {sistema}. Análise: {analise}. Fungo: {self.fungo}. Concentração: {concentracao}", )
        
        ws = self.ws
        # Convert start column (e.g., "D") to a number
        start_col = column_index_from_string(start_col_letter)

        # Compute end column and row
        end_col = start_col + ncols - 1
        end_row = start_row + nrows - 1

        # Convert back to Excel letters
        end_col_letter = get_column_letter(end_col)

        # Build Excel-style range (e.g., "D5:F9")
        cell_range = f"{start_col_letter}{start_row}:{end_col_letter}{end_row}"

        # Extract data from that range

        

        for i, row in enumerate(ws[cell_range]):
            for j, cell in enumerate(row):
                
                self.add_amostra(sistema=sistema, dia=dia, concentracao=concentracao, analise=analise, valor=cell.value)
    
    
    def extrair_coluna(self, start_col_letter: str, start_row: int,dia: int, sistema: str, analise: str):
        ncols = 1
        nrows = 5

        print(f"Extraindo coluna. Sistema: {sistema}. Análise: {analise}. Fungo: {self.fungo}", )
        
        ws = self.ws
        # Convert start column (e.g., "D") to a number
        start_col = column_index_from_string(start_col_letter)

        # Compute end column and row
        end_col = start_col + ncols - 1
        end_row = start_row + nrows - 1

        # Convert back to Excel letters
        end_col_letter = get_column_letter(end_col)

        # Build Excel-style range (e.g., "D5:F9")
        cell_range = f"{start_col_letter}{start_row}:{end_col_letter}{end_row}"

        # Extract data from that range

        concentracao_map = [0, 25, 50, 75, 100]
        

        for i, row in enumerate(ws[cell_range]):
            for j, cell in enumerate(row):
                
                self.add_amostra(sistema=sistema, dia=dia, concentracao=concentracao_map[i], analise=analise, valor=cell.value)


    def extrair_quadrante(self, start_col_letter: str, start_row: int, sistema: str, analise: str):
        ncols = 3
        nrows = 5

        print(f"Extraindo quadrante. Sistema: {sistema}. Análise: {analise}. Fungo: {self.fungo}", )
        
        ws = self.ws
        # Convert start column (e.g., "D") to a number
        start_col = column_index_from_string(start_col_letter)

        # Compute end column and row
        end_col = start_col + ncols - 1
        end_row = start_row + nrows - 1

        # Convert back to Excel letters
        end_col_letter = get_column_letter(end_col)

        # Build Excel-style range (e.g., "D5:F9")
        cell_range = f"{start_col_letter}{start_row}:{end_col_letter}{end_row}"

        # Extract data from that range

        concentracao_map = [0, 25, 50, 75, 100]
        dia_map = [0, 7, 14]
        

        for i, row in enumerate(ws[cell_range]):
            for j, cell in enumerate(row):
                
                self.add_amostra(sistema=sistema, dia=dia_map[j], concentracao=concentracao_map[i], analise=analise, valor=cell.value)




    def add_amostra(self, sistema: str, dia: str, concentracao: int, analise: str, valor: Any):
        if valor is None or valor == "":
            valor = np.nan
        fungo = self.fungo
        amostra = Amostra(fungo=fungo, sistema=sistema, dia=dia, concentracao=concentracao, analise=analise, valor = valor)
        self._amostras.append(amostra)




