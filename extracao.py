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


FILE_PATH_CANDIDA_ANALISE_MACROSCOPICA = Path('data/etapa_2/Candida/Análise Macroscópica.xlsx')
FILE_PATH_PENICILLIUM_ANALISE_MACROSCOPICA = Path('data/etapa_2/Penicillium/Análise Macroscópica.xlsx')
FILE_PATH_ASPERGILLIUS_ANALISE_MACROSCOPICA = Path('data/etapa_2/Aspergillius/Análise Macroscópica.xlsx')
FILE_PATH_SACCHAROMYCES_ANALISE_MACROSCOPICA = Path('data/etapa_2/Saccharomyces/Análise Macroscópica.xlsx')


FILE_PATH_CANDIDA_BASE_SECA = Path('data/etapa_2/Candida/Base Seca.xlsx')
FILE_PATH_PENICILLIUM_BASE_SECA = Path('data/etapa_2/Penicillium/Base Seca.xlsx')
FILE_PATH_SACCHAROMYCES_BASE_SECA = Path('data/etapa_2/Saccharomyces/Base Seca.xlsx')
FILE_PATH_ASPERGILLIUS_BASE_SECA = Path('data/etapa_2/Aspergillius/Base Seca.xlsx')


FILE_PATH_CANDIDA_UFC = Path('data/etapa_2/Candida/Calculo de UFC.xlsx')

def main():
    todas_amostras = []

    # Extração fisicoquímico
    extrator_candida_fisicoquimoco = Extrator(file_path=FILE_PATH_CANDIDA_FISICOQUIMICO, sheet='PROCESSAMENTO', fungo="Candida")
    extrator_candida_fisicoquimoco.extrair_fisicoquimico()
    todas_amostras = todas_amostras + extrator_candida_fisicoquimoco._amostras

    extrator_aspergillus_fisicoquimoco = Extrator(file_path=FILE_PATH_ASPERGILLIUS_FISICOQUIMICO, sheet='PROCESSAMENTO', fungo="Aspergillius")
    extrator_aspergillus_fisicoquimoco.extrair_fisicoquimico()
    todas_amostras = todas_amostras + extrator_aspergillus_fisicoquimoco._amostras

    extrator_saccharomyces_fisicoquimoco = Extrator(file_path=FILE_PATH_SACCHAROMYCES_FISICOQUIMICO, sheet='PROCESSAMENTO', fungo="Saccharomyces")
    extrator_saccharomyces_fisicoquimoco.extrair_fisicoquimico()
    todas_amostras = todas_amostras + extrator_saccharomyces_fisicoquimoco._amostras

    extrator_saccharomyces_fisicoquimoco = Extrator(file_path=FILE_PATH_PENICILLIUM_FISICOQUIMICO, sheet='PROCESSAMENTO', fungo="Penicillium")
    extrator_saccharomyces_fisicoquimoco.extrair_fisicoquimico()
    todas_amostras = todas_amostras + extrator_saccharomyces_fisicoquimoco._amostras

    # Extração metais tóxicos

    extrator_candida_fisicoquimoco = Extrator(file_path=FILE_PATH_CANDIDA_METAIS_TOXICOS, sheet='Candida', fungo='Candida')
    extrator_candida_fisicoquimoco.extrair_metais_toxicos()
    todas_amostras = todas_amostras + extrator_candida_fisicoquimoco._amostras

    extrator_candida_fisicoquimoco = Extrator(file_path=FILE_PATH_ASPERGILLIUS_METAIS_TOXICOS, sheet='Aspergillius', fungo='Aspergillius')
    extrator_candida_fisicoquimoco.extrair_metais_toxicos()
    todas_amostras = todas_amostras + extrator_candida_fisicoquimoco._amostras

    extrator_candida_fisicoquimoco = Extrator(file_path=FILE_PATH_SACCHAROMYCES_METAIS_TOXICOS, sheet='Scharomicys', fungo='Saccharomyces')
    extrator_candida_fisicoquimoco.extrair_metais_toxicos()
    todas_amostras = todas_amostras + extrator_candida_fisicoquimoco._amostras

    extrator_penicillium_fisicoquimoco = Extrator(file_path=FILE_PATH_SACCHAROMYCES_METAIS_TOXICOS, sheet='Penicillium', fungo='Penicillium')
    extrator_penicillium_fisicoquimoco.extrair_metais_toxicos()
    todas_amostras = todas_amostras + extrator_penicillium_fisicoquimoco._amostras
    todas_amostras = todas_amostras + extrator_penicillium_fisicoquimoco._amostras


    # Extração análise macroscópica
    # Saccharomyces
    extrator_saccharomyces_analise_macroscopica = Extrator(file_path=FILE_PATH_SACCHAROMYCES_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 0', fungo="Saccharomyces")
    extrator_saccharomyces_analise_macroscopica.extrair_analise_macroscopica(dia=0)
    todas_amostras = todas_amostras + extrator_saccharomyces_analise_macroscopica._amostras

    extrator_saccharomyces_analise_macroscopica = Extrator(file_path=FILE_PATH_SACCHAROMYCES_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 7', fungo="Saccharomyces")
    extrator_saccharomyces_analise_macroscopica.extrair_analise_macroscopica(dia=7)
    todas_amostras = todas_amostras + extrator_saccharomyces_analise_macroscopica._amostras

    extrator_saccharomyces_analise_macroscopica = Extrator(file_path=FILE_PATH_SACCHAROMYCES_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 14', fungo="Saccharomyces")
    extrator_saccharomyces_analise_macroscopica.extrair_analise_macroscopica(dia=14)
    todas_amostras = todas_amostras + extrator_saccharomyces_analise_macroscopica._amostras
    
    # Candida
    extrator_candida_analise_macroscopica = Extrator(file_path=FILE_PATH_CANDIDA_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 0', fungo="Candida")
    extrator_candida_analise_macroscopica.extrair_analise_macroscopica(dia=0)
    todas_amostras = todas_amostras + extrator_candida_analise_macroscopica._amostras

    extrator_candida_analise_macroscopica = Extrator(file_path=FILE_PATH_CANDIDA_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 7', fungo="Candida")
    extrator_candida_analise_macroscopica.extrair_analise_macroscopica(dia=7)
    todas_amostras = todas_amostras + extrator_candida_analise_macroscopica._amostras

    extrator_candida_analise_macroscopica = Extrator(file_path=FILE_PATH_CANDIDA_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 14', fungo="Candida")
    extrator_candida_analise_macroscopica.extrair_analise_macroscopica(dia=14)
    todas_amostras = todas_amostras + extrator_candida_analise_macroscopica._amostras

    # Penicillium
    extrator_penicillium_analise_macroscopica = Extrator(file_path=FILE_PATH_PENICILLIUM_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 0', fungo="Penicillium")
    extrator_penicillium_analise_macroscopica.extrair_analise_macroscopica(dia=0)
    todas_amostras = todas_amostras + extrator_penicillium_analise_macroscopica._amostras

    extrator_penicillium_analise_macroscopica = Extrator(file_path=FILE_PATH_PENICILLIUM_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 7', fungo="Penicillium")
    extrator_penicillium_analise_macroscopica.extrair_analise_macroscopica(dia=7)
    todas_amostras = todas_amostras + extrator_penicillium_analise_macroscopica._amostras

    extrator_penicillium_analise_macroscopica = Extrator(file_path=FILE_PATH_PENICILLIUM_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 14', fungo="Penicillium")
    extrator_penicillium_analise_macroscopica.extrair_analise_macroscopica(dia=14)
    todas_amostras = todas_amostras + extrator_penicillium_analise_macroscopica._amostras

    # Aspergillius
    extrator_aspergillius_analise_macroscopica = Extrator(file_path=FILE_PATH_ASPERGILLIUS_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 0', fungo="Aspergillius")
    extrator_aspergillius_analise_macroscopica.extrair_analise_macroscopica(dia=0)
    todas_amostras = todas_amostras + extrator_aspergillius_analise_macroscopica._amostras

    extrator_aspergillius_analise_macroscopica = Extrator(file_path=FILE_PATH_ASPERGILLIUS_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 7', fungo="Aspergillius")
    extrator_aspergillius_analise_macroscopica.extrair_analise_macroscopica(dia=7)
    todas_amostras = todas_amostras + extrator_aspergillius_analise_macroscopica._amostras

    extrator_aspergillius_analise_macroscopica = Extrator(file_path=FILE_PATH_ASPERGILLIUS_ANALISE_MACROSCOPICA, sheet='Dados prontos Dia 14', fungo="Aspergillius")
    extrator_aspergillius_analise_macroscopica.extrair_analise_macroscopica(dia=14)
    todas_amostras = todas_amostras + extrator_aspergillius_analise_macroscopica._amostras

    # Base Seca
    extrator_Candida_base_seca = Extrator(file_path=FILE_PATH_CANDIDA_BASE_SECA, sheet='Dia 14', fungo="Candida")
    extrator_Candida_base_seca.extrair_base_seca()
    todas_amostras = todas_amostras + extrator_Candida_base_seca._amostras

    extrator_Penicillium_base_seca = Extrator(file_path=FILE_PATH_PENICILLIUM_BASE_SECA, sheet='Dia 14', fungo="Penicillium")
    extrator_Penicillium_base_seca.extrair_base_seca()
    todas_amostras = todas_amostras + extrator_Penicillium_base_seca._amostras

    extrator_Saccharomyces_base_seca = Extrator(file_path=FILE_PATH_SACCHAROMYCES_BASE_SECA, sheet='Dia 14', fungo="Saccharomyces")
    extrator_Saccharomyces_base_seca.extrair_base_seca()
    todas_amostras = todas_amostras + extrator_Saccharomyces_base_seca._amostras

    extrator_aspergillius_base_seca = Extrator(file_path=FILE_PATH_ASPERGILLIUS_BASE_SECA, sheet='Dia 14', fungo="Aspergillius")
    extrator_aspergillius_base_seca.extrair_base_seca()
    todas_amostras = todas_amostras + extrator_aspergillius_base_seca._amostras


    # # ufc
    # extrator_aspergillius_ufc = Extrator(file_path=FILE_PATH_CANDIDA_UFC, sheet='Dia 14', fungo="Aspergillius")
    # extrator_aspergillius_ufc.extrair_ufc()
    # todas_amostras = todas_amostras + extrator_aspergillius_ufc._amostras

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