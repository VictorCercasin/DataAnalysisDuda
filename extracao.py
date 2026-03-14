from dataclasses import asdict
from pathlib import Path

import numpy as np
import pandas as pd

from src.load_data import Extrator


BASE_DIR = Path("data/etapa_2")
OUTPUT_DIR = Path("output")
OUTPUT_FILE = OUTPUT_DIR / "structured_data.xlsx"

FUNGI = {
    "candida": {
        "label": "Candida",
        "dir": BASE_DIR / "Candida",
        "metals_sheet": "Candida",
    },
    "aspergillius": {
        "label": "Aspergillius",
        "dir": BASE_DIR / "Aspergillius",
        "metals_sheet": "Aspergillius",
    },
    "saccharomyces": {
        "label": "Saccharomyces",
        "dir": BASE_DIR / "Saccharomyces",
        "metals_sheet": "Scharomicys",  # kept as-is because that seems to be the Excel sheet name
    },
    "penicillium": {
        "label": "Penicillium",
        "dir": BASE_DIR / "Penicillium",
        "metals_sheet": "Penicillium",
    },
}

TOXIC_METALS_FILE = BASE_DIR / "Análise Metais.xlsx"


def append_samples(all_samples: list, extractor: Extrator) -> None:
    all_samples.extend(extractor._amostras)


def extract_physicochemical(all_samples: list) -> None:
    for fungus in FUNGI.values():
        extractor = Extrator(
            file_path=fungus["dir"] / "Fisico-Quimico.xlsx",
            sheet="PROCESSAMENTO",
            fungo=fungus["label"],
        )
        extractor.extrair_fisicoquimico()
        append_samples(all_samples, extractor)


def extract_toxic_metals(all_samples: list) -> None:
    for fungus in FUNGI.values():
        extractor = Extrator(
            file_path=TOXIC_METALS_FILE,
            sheet=fungus["metals_sheet"],
            fungo=fungus["label"],
        )
        extractor.extrair_metais_toxicos()
        append_samples(all_samples, extractor)


def extract_macroscopic_analysis(all_samples: list) -> None:
    days = [0, 7, 14]

    for fungus in FUNGI.values():
        file_path = fungus["dir"] / "Análise Macroscópica.xlsx"

        for day in days:
            extractor = Extrator(
                file_path=file_path,
                sheet=f"Dados prontos Dia {day}",
                fungo=fungus["label"],
            )
            extractor.extrair_analise_macroscopica(dia=day)
            append_samples(all_samples, extractor)


def extract_dry_basis(all_samples: list) -> None:
    for fungus in FUNGI.values():
        extractor = Extrator(
            file_path=fungus["dir"] / "Base Seca.xlsx",
            sheet="Dia 14",
            fungo=fungus["label"],
        )
        extractor.extrair_base_seca()
        append_samples(all_samples, extractor)


def clean_valor_column(df: pd.DataFrame) -> pd.DataFrame:
    if "valor" not in df.columns:
        return df

    value_series = df["valor"].astype(str).str.strip()

    # Fix malformed decimal text like "8,,98" -> "8,98"
    value_series = value_series.str.replace(",,", ",", regex=False)

    # Convert Brazilian decimal comma to dot
    value_series = value_series.str.replace(",", ".", regex=False)

    # Normalize obvious invalid tokens
    invalid_tokens = {"", "nan", "none", "-", "—", "na", "n/a", "nd"}
    value_series_lower = value_series.str.lower()
    value_series = value_series.where(
        ~value_series_lower.isin(invalid_tokens),
        other=np.nan,
    )

    df["valor"] = pd.to_numeric(value_series, errors="coerce")
    df = df.dropna(subset=["valor"]).reset_index(drop=True)

    return df


def build_dataframe(all_samples: list) -> pd.DataFrame:
    df = pd.DataFrame([asdict(sample) for sample in all_samples])
    return clean_valor_column(df)


def save_dataframe(df: pd.DataFrame) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df.to_excel(OUTPUT_FILE, index=False)


def main() -> None:
    all_samples = []

    extract_physicochemical(all_samples)
    extract_toxic_metals(all_samples)
    extract_macroscopic_analysis(all_samples)
    extract_dry_basis(all_samples)

    df = build_dataframe(all_samples)
    save_dataframe(df)


if __name__ == "__main__":
    main()