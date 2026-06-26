from __future__ import annotations

import argparse
from pathlib import Path
import sys

import pandas as pd


DEFAULT_INPUT_FILE = r"C:\Users\ricar\Desktop\DIR CYBER SECURITY.xlsx"
DEFAULT_DIRETORIA = "DIR CYBER SECURITY"
REQUIRED_COLUMNS = ["Diretoria", "Indicador", "Requisito"]
DEFAULT_PILLAR_COUNT = 6


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Le uma planilha Excel e lista os indicadores e requisitos "
            "existentes para apoiar a configuracao do PILLAR_CONFIG."
        )
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        default=DEFAULT_INPUT_FILE,
        help="Caminho do arquivo Excel de entrada.",
    )
    parser.add_argument(
        "-o",
        "--output",
        help=(
            "Arquivo de saida com a sugestao de configuracao. "
            "Se omitido, salva ao lado do script como pillar_config_sugerido.py."
        ),
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="Nome ou indice da aba a ser lida. Padrao: primeira aba.",
    )
    parser.add_argument(
        "--diretoria",
        default=DEFAULT_DIRETORIA,
        help=(
            "Valor da coluna Diretoria usado no filtro. "
            "Use vazio para listar todas as diretorias."
        ),
    )
    parser.add_argument(
        "--pilares",
        type=int,
        default=DEFAULT_PILLAR_COUNT,
        help=f"Quantidade de pilares na sugestao de PILLAR_CONFIG. Padrao: {DEFAULT_PILLAR_COUNT}.",
    )
    return parser


def normalize_sheet_arg(sheet: str) -> str | int:
    try:
        return int(sheet)
    except ValueError:
        return sheet


def validate_columns(dataframe: pd.DataFrame) -> None:
    missing_columns = [
        column for column in REQUIRED_COLUMNS if column not in dataframe.columns
    ]
    if missing_columns:
        available_columns = ", ".join(str(column) for column in dataframe.columns)
        missing = ", ".join(missing_columns)
        raise ValueError(
            f"Colunas obrigatorias nao encontradas: {missing}. "
            f"Colunas disponiveis: {available_columns}"
        )


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def default_output_path() -> Path:
    return Path.cwd() / "pillar_config_sugerido.py"


def build_inventory(dataframe: pd.DataFrame, diretoria: str) -> pd.DataFrame:
    working = dataframe.copy()

    if diretoria:
        working = working[
            working["Diretoria"].map(clean_text).eq(diretoria)
        ].copy()

    working["Indicador"] = working["Indicador"].map(clean_text)
    working["Requisito"] = working["Requisito"].map(clean_text)

    inventory = (
        working[["Indicador", "Requisito"]]
        .drop_duplicates()
        .sort_values(["Indicador", "Requisito"])
        .reset_index(drop=True)
    )
    return inventory


def quote(value: str) -> str:
    return repr(value)


def pillar_weights(pillar_count: int) -> list[float]:
    if pillar_count <= 0:
        raise ValueError("A quantidade de pilares deve ser maior que zero.")

    base_weight = round(100 / pillar_count, 2)
    weights = [base_weight for _ in range(pillar_count)]
    weights[-1] = round(100 - sum(weights[:-1]), 2)
    return weights


def render_indicators(inventory: pd.DataFrame, indentation: str) -> list[str]:
    lines = []

    for indicator, group in inventory.groupby("Indicador", sort=True):
        requirements = group["Requisito"].tolist()
        requirement_weight = round(100 / len(requirements), 4) if requirements else 0

        lines.extend(
            [
                f"{indentation}{{",
                f'{indentation}    "Indicador": {quote(indicator)},',
                f'{indentation}    "Peso Indicador %": 100,',
                f'{indentation}    "Requisitos": [',
            ]
        )

        for requirement in requirements:
            lines.extend(
                [
                    f"{indentation}        {{",
                    f'{indentation}            "Requisito": {quote(requirement)},',
                    f'{indentation}            "Peso Requisito %": {requirement_weight},',
                    f"{indentation}        }},",
                ]
            )

        lines.extend(
            [
                f"{indentation}    ],",
                f"{indentation}}},",
            ]
        )

    return lines


def render_pillar_config(inventory: pd.DataFrame, pillar_count: int) -> str:
    weights = pillar_weights(pillar_count)
    lines = ["PILLAR_CONFIG = ["]

    for index, weight in enumerate(weights, start=1):
        lines.extend(
            [
                "    {",
                f'        "Pilar": "Pilar {index} - Ajustar nome",',
                f'        "Peso Pilar %": {weight},',
                '        "Indicadores": [',
            ]
        )

        if index == 1:
            lines.extend(render_indicators(inventory, "            "))

        lines.extend(
            [
                "        ],",
                "    },",
            ]
        )

    lines.extend(
        [
            "]",
            "",
        ]
    )
    return "\n".join(lines)


def list_indicators_requirements(
    input_file: str,
    output_file: str | None,
    sheet: str,
    diretoria: str,
    pilares: int,
) -> Path:
    input_path = Path(input_file).expanduser()
    if not input_path.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {input_path}")

    dataframe = pd.read_excel(input_path, sheet_name=normalize_sheet_arg(sheet))
    validate_columns(dataframe)

    inventory = build_inventory(dataframe, diretoria)
    output_path = Path(output_file).expanduser() if output_file else default_output_path()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(render_pillar_config(inventory, pilares), encoding="utf-8")

    print(f"Diretoria filtrada: {diretoria or 'Todas'}")
    print(f"Pilares na sugestao: {pilares}")
    print(f"Indicadores encontrados: {inventory['Indicador'].nunique()}")
    print(f"Requisitos encontrados: {len(inventory)}")
    print()
    print(inventory.to_string(index=False))
    print()
    print(f"Sugestao de PILLAR_CONFIG salva em: {output_path}")
    return output_path


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        list_indicators_requirements(
            input_file=args.input_file,
            output_file=args.output,
            sheet=args.sheet,
            diretoria=args.diretoria,
            pilares=args.pilares,
        )
    except Exception as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
