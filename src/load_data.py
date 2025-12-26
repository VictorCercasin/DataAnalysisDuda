import pandas as pd
import numpy as np
import os
from pathlib import Path
from dataclasses import dataclass
from openpyxl import load_workbook
from typing import Any, Literal
from openpyxl.utils import get_column_letter, column_index_from_string


@dataclass
class Amostra:
    fungo: str
    sistema: Literal['AI', 'A', 'BI', 'B', 'CN', 'CP']
    dia: Literal[0, 7, 14]
    concentracao: Literal[0, 25, 50, 75, 100, -1]
    analise: str
    valor: Any
    
    def __post_init__(self):
        valid_concentrations = {0, 25, 50, 75, 100, -1}
        if self.concentracao not in valid_concentrations:
            raise ValueError(f"concentracao deve ser um dos valores a seguir: {valid_concentrations}")
        valid_days = {0, 7, 14}
        if self.dia not in valid_days:
            raise ValueError(f"dia deve ser um dos valores a seguir: {valid_days}")
        valid_systems = {'AI', 'A', 'BI', 'B', 'CN', 'CP'}
        if self.sistema not in valid_systems:
            raise ValueError(f"sistema deve ser um dos valores a seguir: {valid_systems}")



class Extrator:
    def __init__(self, file_path: str, sheet: str, fungo: str = ''):
        print("Iniciando extrator")
        self.file_path = file_path
        self.fungo = fungo if len(fungo) > 0 else self.file_path.stem
        self._amostras: list[Amostra] = []
        if os.path.exists(file_path):
            self.wb = load_workbook(file_path, data_only=True)
        else:
            print(f"ERRO - Caminho {file_path} não existe")
            return []
        self.ws= self.wb[sheet]


    def extrair_analise_macroscopica(self):
        # self.extrair_coluna(start_col_letter='d', start_row=4, dia=0, sistema='AI', analise='Sementes Germinadas')
        # self.extrair_coluna(start_col_letter='d', start_row=9, dia=0, sistema='A', analise='Sementes Germinadas')
        # self.extrair_coluna(start_col_letter='d', start_row=14, dia=0, sistema='BI', analise='Sementes Germinadas')
        # self.extrair_coluna(start_col_letter='d', start_row=19, dia=0, sistema='B', analise='Sementes Germinadas')

        # self.extrair_celula(start_col_letter='d', start_row=24, dia=0, sistema='CN', analise='Sementes Germinadas', concentracao=-1)
        # self.extrair_celula(start_col_letter='d', start_row=25, dia=0, sistema='CP', analise='Sementes Germinadas', concentracao=-1)
        self.extrair_celula(start_col_letter='d', start_row=12, dia=0, sistema='A', analise='Sementes Germinadas', concentracao=75)

        # self._amostras = [
        #     a for a in self._amostras
        #     if a.valor != '#DIV/0!'
        # ]
    def extrair_metais_toxicos(self):

        # Apenas as concentrações do zinco ultrapassam os mínimos de sensibilidade do equipamento

        # Dia 0
        # self.extrair_coluna(start_col_letter='d', start_row=5, dia=0, sistema='AI', analise='Cd')
        # self.extrair_coluna(start_col_letter='d', start_row=10, dia=0, sistema='A', analise='Cd')
        # self.extrair_coluna(start_col_letter='d', start_row=15, dia=0, sistema='BI', analise='Cd')
        # self.extrair_coluna(start_col_letter='d', start_row=20, dia=0, sistema='B', analise='Cd')

        # self.extrair_coluna(start_col_letter='e', start_row=5, dia=0, sistema='AI', analise='Cr')
        # self.extrair_coluna(start_col_letter='e', start_row=10, dia=0, sistema='A', analise='Cr')
        # self.extrair_coluna(start_col_letter='e', start_row=15, dia=0, sistema='BI', analise='Cr')
        # self.extrair_coluna(start_col_letter='e', start_row=20, dia=0, sistema='B', analise='Cr')

        # self.extrair_coluna(start_col_letter='f', start_row=5, dia=0, sistema='AI', analise='Cu')
        # self.extrair_coluna(start_col_letter='f', start_row=10, dia=0, sistema='A', analise='Cu')
        # self.extrair_coluna(start_col_letter='f', start_row=15, dia=0, sistema='BI', analise='Cu')
        # self.extrair_coluna(start_col_letter='f', start_row=20, dia=0, sistema='B', analise='Cu')

        # self.extrair_coluna(start_col_letter='g', start_row=5, dia=0, sistema='AI', analise='Ni')
        # self.extrair_coluna(start_col_letter='g', start_row=10, dia=0, sistema='A', analise='Ni')
        # self.extrair_coluna(start_col_letter='g', start_row=15, dia=0, sistema='BI', analise='Ni')
        # self.extrair_coluna(start_col_letter='g', start_row=20, dia=0, sistema='B', analise='Ni')

        # self.extrair_coluna(start_col_letter='h', start_row=5, dia=0, sistema='AI', analise='Pb')
        # self.extrair_coluna(start_col_letter='h', start_row=10, dia=0, sistema='A', analise='Pb')
        # self.extrair_coluna(start_col_letter='h', start_row=15, dia=0, sistema='BI', analise='Pb')
        # self.extrair_coluna(start_col_letter='h', start_row=20, dia=0, sistema='B', analise='Pb')

        self.extrair_coluna(start_col_letter='i', start_row=5, dia=0, sistema='AI', analise='Zn')
        self.extrair_coluna(start_col_letter='i', start_row=10, dia=0, sistema='A', analise='Zn')
        self.extrair_coluna(start_col_letter='i', start_row=15, dia=0, sistema='BI', analise='Zn')
        self.extrair_coluna(start_col_letter='i', start_row=20, dia=0, sistema='B', analise='Zn')



        # Dia 14
        # self.extrair_coluna(start_col_letter='j', start_row=5, dia=14, sistema='AI', analise='Cd')
        # self.extrair_coluna(start_col_letter='j', start_row=10, dia=14, sistema='A', analise='Cd')
        # self.extrair_coluna(start_col_letter='j', start_row=15, dia=14, sistema='BI', analise='Cd')
        # self.extrair_coluna(start_col_letter='j', start_row=20, dia=14, sistema='B', analise='Cd')

        # self.extrair_coluna(start_col_letter='k', start_row=5, dia=14, sistema='AI', analise='Cr')
        # self.extrair_coluna(start_col_letter='k', start_row=10, dia=14, sistema='A', analise='Cr')
        # self.extrair_coluna(start_col_letter='k', start_row=15, dia=14, sistema='BI', analise='Cr')
        # self.extrair_coluna(start_col_letter='k', start_row=20, dia=14, sistema='B', analise='Cr')

        # self.extrair_coluna(start_col_letter='l', start_row=5, dia=14, sistema='AI', analise='Cu')
        # self.extrair_coluna(start_col_letter='l', start_row=10, dia=14, sistema='A', analise='Cu')
        # self.extrair_coluna(start_col_letter='l', start_row=15, dia=14, sistema='BI', analise='Cu')
        # self.extrair_coluna(start_col_letter='l', start_row=20, dia=14, sistema='B', analise='Cu')

        # self.extrair_coluna(start_col_letter='m', start_row=5, dia=14, sistema='AI', analise='Ni')
        # self.extrair_coluna(start_col_letter='m', start_row=10, dia=14, sistema='A', analise='Ni')
        # self.extrair_coluna(start_col_letter='m', start_row=15, dia=14, sistema='BI', analise='Ni')
        # self.extrair_coluna(start_col_letter='m', start_row=20, dia=14, sistema='B', analise='Ni')

        # self.extrair_coluna(start_col_letter='n', start_row=5, dia=14, sistema='AI', analise='Pb')
        # self.extrair_coluna(start_col_letter='n', start_row=10, dia=14, sistema='A', analise='Pb')
        # self.extrair_coluna(start_col_letter='n', start_row=15, dia=14, sistema='BI', analise='Pb')
        # self.extrair_coluna(start_col_letter='n', start_row=20, dia=14, sistema='B', analise='Pb')

        self.extrair_coluna(start_col_letter='o', start_row=5, dia=14, sistema='AI', analise='Zn')
        self.extrair_coluna(start_col_letter='o', start_row=10, dia=14, sistema='A', analise='Zn')
        self.extrair_coluna(start_col_letter='o', start_row=15, dia=14, sistema='BI', analise='Zn')
        self.extrair_coluna(start_col_letter='o', start_row=20, dia=14, sistema='B', analise='Zn')



    def extrair_celula(self, start_col_letter: str, start_row: int,dia: int, sistema: str, analise: str, concentracao: int):
        ncols = 1
        nrows = 1

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
    

    def extrair_fisicoquimico(self):
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




