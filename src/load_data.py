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
    sistema: Literal['AI','A','BI','B']
    dia: Literal[0, 7, 14]
    concentracao: Literal[0, 25, 50, 75, 100]
    analise: str
    valor: Any
    
    def __post_init__(self):
        valid_concentrations = {0, 25, 50, 75, 100}
        if self.concentracao not in valid_concentrations:
            raise ValueError(f"concentracao deve ser um dos valores a seguir: {valid_concentrations}")
        valid_days = {0, 7, 14}
        if self.dia not in valid_days:
            raise ValueError(f"dia deve ser um dos valores a seguir: {valid_days}")
        valid_systems = {'AI','A','BI','B'}
        if self.sistema not in valid_systems:
            raise ValueError(f"sistema deve ser um dos valores a seguir: {valid_systems}")



class Extrator:
    def __init__(self, file_path: str, sheet: str):
        print("Iniciando extrator")
        self.file_path = file_path
        self.wb = load_workbook(file_path, data_only=True)
        self.fungo = self.file_path.stem
        self.ws= self.wb[sheet]
        self._amostras: list[Amostra] = []


    def extrair(self):
        # self.add_amostra("AI", 7, 25, "pH", 3.5)
        print("Iniciando extração")

        print("Extraíndo Fisico Químico")
        self.extrair_fisicoquimico()

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

        print(f"Extraindo quadrante. Sistema: {sistema}. Análise: {analise}", )
        
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




