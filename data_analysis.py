import pandas as pd
import numpy as np
import os
from pathlib import Path
import matplotlib.pyplot as plt
import seaborn as sns


FILE_PATH = Path('output/structured_data.xlsx')

def main():
    df = pd.read_excel(FILE_PATH)
    heatmap_fungo_metal(df, sistema='A', dia=0)


def heatmap_fungo_metal(df: pd.DataFrame, sistema: str = "", dia: int = -1, fungo: str = ""):
    """
    Creates a heatmap where: rows = concentration, columns = metal, cell = mean valor
    for a given fungus.
    """
    metais = ['Cd', 'Cr', 'Ni', 'Pb', 'Zn']

    cond = df["analise"].isin(metais)

    if len(sistema) > 0:
        cond &= (df["sistema"] == sistema)

    if dia != -1:
        cond &= (df["dia"] == dia)

    if len(fungo) > 0:
        cond &= df["fungo"].str.contains(fungo, case=False)

    df_f = df[cond]

    # pivot = concentration (rows), analise (columns), average value inside
    tabela = df_f.pivot_table(
        index="concentracao",
        columns="analise",
        values="valor",
        aggfunc="mean"
    )

    plt.figure(figsize=(10,6))
    ax = sns.heatmap(
        tabela,
        annot=True,
        cmap="viridis",
        fmt=".2f"
    )

    # colorbar label
    cbar = ax.collections[0].colorbar
    cbar.set_label("mg/L")

    plt.title(f"{f'Sistema {sistema}' if len(sistema) > 0 else ''} {f', fungo {fungo}' if len(fungo) > 0 else ''} {f'Dia {dia}' if dia != -1 else ''}")
    plt.xlabel("Metal")
    plt.ylabel("Concentração (%)")
    plt.show()

    

    

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