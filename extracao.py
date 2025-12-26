import pandas as pd
import numpy as np
import os
from dataclasses import asdict
from pathlib import Path
from openpyxl import load_workbook
from src.load_data import  Extrator


FILE_PATH_CANDIDA_FISICOQUIMICO = Path('data/etapa_2/Candida/Fisico-Quimico.xlsx')
FILE_PATH_ASPERGILLIUS_FISICOQUIMICO = Path('data/etapa_2/Aspergillius/Fisico-Quimico.xlsx')
FILE_PATH_SACCHAROMYCES_FISICOQUIMICO = Path('data/etapa_2/Saccharomyces/Fisico-Quimico.xlsx')
FILE_PATH_PENICILLIUM_FISICOQUIMICO = Path('data/etapa_2/Penicillium/Fisico-Quimico.xlsx')

FILE_PATH_CANDIDA_METAIS_TOXICOS = Path('data/etapa_2/Análise Metais.xlsx')
FILE_PATH_ASPERGILLIUS_METAIS_TOXICOS = Path('data/etapa_2/Análise Metais.xlsx')
FILE_PATH_SACCHAROMYCES_METAIS_TOXICOS = Path('data/etapa_2/Análise Metais.xlsx')
FILE_PATH_PENICILLIUM_METAIS_TOXICOS = Path('data/etapa_2/Análise Metais.xlsx')


FILE_PATH_SACCHAROMYCES_ANALISE_MACROSCOPICA = Path('data/etapa_2/Saccharomyces/Análise Macroscópica.xlsx')

def main():
    todas_amostras = []

    # # Extração fisicoquímico
    # extrator_candida_fisicoquimoco = Extrator(file_path=FILE_PATH_CANDIDA_FISICOQUIMICO, sheet='PROCESSAMENTO', fungo="Candida")
    # extrator_candida_fisicoquimoco.extrair_fisicoquimico()
    # todas_amostras = todas_amostras + extrator_candida_fisicoquimoco._amostras

    # extrator_aspergillus_fisicoquimoco = Extrator(file_path=FILE_PATH_ASPERGILLIUS_FISICOQUIMICO, sheet='PROCESSAMENTO', fungo="Aspergillius")
    # extrator_aspergillus_fisicoquimoco.extrair_fisicoquimico()
    # todas_amostras = todas_amostras + extrator_aspergillus_fisicoquimoco._amostras

    # extrator_saccharomyces_fisicoquimoco = Extrator(file_path=FILE_PATH_SACCHAROMYCES_FISICOQUIMICO, sheet='PROCESSAMENTO', fungo="Saccharomyces")
    # extrator_saccharomyces_fisicoquimoco.extrair_fisicoquimico()
    # todas_amostras = todas_amostras + extrator_saccharomyces_fisicoquimoco._amostras

    # extrator_saccharomyces_fisicoquimoco = Extrator(file_path=FILE_PATH_PENICILLIUM_FISICOQUIMICO, sheet='PROCESSAMENTO', fungo="Penicillium")
    # extrator_saccharomyces_fisicoquimoco.extrair_fisicoquimico()
    # todas_amostras = todas_amostras + extrator_saccharomyces_fisicoquimoco._amostras

    # # Extração metais tóxicos

    # extrator_candida_fisicoquimoco = Extrator(file_path=FILE_PATH_CANDIDA_METAIS_TOXICOS, sheet='Candida', fungo='Candida')
    # extrator_candida_fisicoquimoco.extrair_metais_toxicos()
    # todas_amostras = todas_amostras + extrator_candida_fisicoquimoco._amostras

    # extrator_candida_fisicoquimoco = Extrator(file_path=FILE_PATH_ASPERGILLIUS_METAIS_TOXICOS, sheet='Aspergillius', fungo='Aspergillius')
    # extrator_candida_fisicoquimoco.extrair_metais_toxicos()
    # todas_amostras = todas_amostras + extrator_candida_fisicoquimoco._amostras

    # extrator_candida_fisicoquimoco = Extrator(file_path=FILE_PATH_SACCHAROMYCES_METAIS_TOXICOS, sheet='Scharomicys', fungo='Saccharomyces')
    # extrator_candida_fisicoquimoco.extrair_metais_toxicos()
    # todas_amostras = todas_amostras + extrator_candida_fisicoquimoco._amostras

    # extrator_penicillium_fisicoquimoco = Extrator(file_path=FILE_PATH_SACCHAROMYCES_METAIS_TOXICOS, sheet='Penicillium', fungo='Penicillium')
    # extrator_penicillium_fisicoquimoco.extrair_metais_toxicos()
    # todas_amostras = todas_amostras + extrator_penicillium_fisicoquimoco._amostras
    # todas_amostras = todas_amostras + extrator_penicillium_fisicoquimoco._amostras


    # Extração análise macroscópica
    extrator_saccharomyces_analise_macroscopica = Extrator(file_path=FILE_PATH_SACCHAROMYCES_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 0', fungo="Saccharomyces")
    extrator_saccharomyces_analise_macroscopica.extrair_analise_macroscopica()
    todas_amostras = todas_amostras + extrator_saccharomyces_analise_macroscopica._amostras
    df = pd.DataFrame([asdict(a) for a in todas_amostras])

    # Ensure output directory exists
    output_dir = Path("./output")
    output_dir.mkdir(parents=True, exist_ok=True)

    # Save the file
    output_path = output_dir / "structured_data.xlsx"
    df.to_excel(output_path, index=False)





# Example usage:
if __name__ == "__main__":
    main()