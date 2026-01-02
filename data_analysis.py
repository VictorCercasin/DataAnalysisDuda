import pandas as pd
import numpy as np
import os
from pathlib import Path
import matplotlib.pyplot as plt
import seaborn as sns


FILE_PATH = Path('output/structured_data.xlsx')

def main():
    df = pd.read_excel(FILE_PATH)


    exemplo(df)    

    

def exemplo(df: pd.DataFrame):
    analise = "pH"
    fungo = "Aspergillius"
    concentracao = 25


    df_agrupado = df[(df["fungo"] == fungo) & (df["analise"].str.contains(analise, case=False)) & (df["concentracao"] == concentracao)]

    # Plot progression by sistema
    plt.figure(figsize=(10, 6))
    sns.lineplot(
        data=df_agrupado,
        x="dia",
        y="valor",
        hue="sistema",
        marker="o",
        # errorbar=None 
    )

    plt.title(f"Progressão do(a) {analise} ao longo dos dias ({fungo}), concentração: {concentracao}%")
    plt.xlabel("Dias")
    plt.ylabel(analise)
    plt.legend(title="Sistema")
    plt.grid(True)
    plt.show()
    pass



if __name__ == "__main__":
    main()