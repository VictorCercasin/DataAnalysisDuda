import pandas as pd
import numpy as np
import os
from dataclasses import asdict
from pathlib import Path
from openpyxl import load_workbook
from src.load_data import  Extrator


FILE_PATH_CANDIDA = Path('data/etapa_2/Candida.xlsx')
FILE_PATH_ASPERGILLUS = Path('data/etapa_2/Aspergillus.xlsx')
FILE_PATH_SACCHAROMYCES = Path('data/etapa_2/Saccharomyces.xlsx')

def main():

    extrator_candida = Extrator(file_path=FILE_PATH_CANDIDA, sheet='PROCESSAMENTO')
    extrator_candida.extrair()

    extrator_aspergillus = Extrator(file_path=FILE_PATH_ASPERGILLUS, sheet='PROCESSAMENTO')
    extrator_aspergillus.extrair()

    extrator_saccharomyces = Extrator(file_path=FILE_PATH_SACCHAROMYCES, sheet='PROCESSAMENTO')
    extrator_saccharomyces.extrair()

    todas_amostras = (
        extrator_candida._amostras
        + extrator_aspergillus._amostras
        + extrator_saccharomyces._amostras
    )


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