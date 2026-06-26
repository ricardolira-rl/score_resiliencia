from __future__ import annotations

import argparse
from pathlib import Path
import sys

import pandas as pd


DEFAULT_DIRETORIA = "DIR CYBER SECURITY"
SUMMARY_COLUMNS = ["Release Train", "Squad", "Indicador", "Requisito"]
REQUIRED_COLUMNS = ["Diretoria", "Status", "Quantidade", *SUMMARY_COLUMNS]
STATUS_COLUMNS = ["PASSED", "FAILED", "Percentual obtido"]
PILLAR_CONFIG = [
    {
        "Pilar": "Pilar 1 - Ajustar nome",
        "Peso Pilar %": 16.67,
        "Indicadores": [
            {
                "Indicador": "Testes em tempo de desenvolvimento (Sonar, TaaC, Hopper)",
                "Peso Indicador %": 100,
                "Requisitos": [
                    {
                        "Requisito": "SonarQube - bugs",
                        "Peso Requisito %": 50,
                    },
                    {
                        "Requisito": "Taac e Hopper - Reuso e Casos de Teste",
                        "Peso Requisito %": 50,
                    },
                ],
            },
        ],
    },
    {
        "Pilar": "Pilar 2 - Ajustar nome",
        "Peso Pilar %": 16.67,
        "Indicadores": [],
    },
    {
        "Pilar": "Pilar 3 - Ajustar nome",
        "Peso Pilar %": 16.67,
        "Indicadores": [],
    },
    {
        "Pilar": "Pilar 4 - Ajustar nome",
        "Peso Pilar %": 16.67,
        "Indicadores": [],
    },
    {
        "Pilar": "Pilar 5 - Ajustar nome",
        "Peso Pilar %": 16.67,
        "Indicadores": [],
    },
    {
        "Pilar": "Pilar 6 - Ajustar nome",
        "Peso Pilar %": 16.65,
        "Indicadores": [],
    },
]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Le uma ou mais planilhas Excel e filtra pelos arquivos cujo nome "
            "sem extensao seja igual a 'DIR CYBER SECURITY'."
        )
    )
    parser.add_argument(
        "input_files",
        nargs="*",
        default=[r"C:\Users\ricar\Desktop\DIR CYBER SECURITY.xlsx"],
        help="Caminho de uma ou mais planilhas Excel de entrada.",
    )
    parser.add_argument(
        "-o",
        "--output",
        help=(
            "Caminho do arquivo Excel de saida. "
            "Se omitido, salva ao lado do arquivo original com sufixo '_filtrado'."
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
            "Valor usado para filtrar pelo nome do arquivo sem extensao. "
            f"Padrao: {DEFAULT_DIRETORIA!r}."
        ),
    )
    return parser


def normalize_sheet_arg(sheet: str) -> str | int:
    try:
        return int(sheet)
    except ValueError:
        return sheet


def default_output_path(input_paths: list[Path]) -> Path:
    if len(input_paths) == 1:
        input_path = input_paths[0]
        return input_path.with_name(f"{input_path.stem}_filtrado.xlsx")
    return Path.cwd() / "score_resiliencia_filtrado.xlsx"


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


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).replace("\xa0", " ").split()).upper()


def filter_by_diretoria(dataframe: pd.DataFrame, diretoria: str) -> pd.DataFrame:
    if not diretoria:
        return dataframe.copy()

    target = normalize_text(diretoria)
    return dataframe[
        dataframe["Diretoria Arquivo"].map(normalize_text).eq(target)
    ].copy()


def format_diretoria_diagnostics(dataframe: pd.DataFrame) -> str:
    file_values = (
        dataframe["Diretoria Arquivo"]
        .fillna("")
        .astype(str)
        .str.strip()
        .value_counts()
        .head(20)
    )
    if file_values.empty:
        return "Nenhum valor encontrado no nome do arquivo."
    return (
        "Diretorias inferidas pelo nome do arquivo:\n"
        f"{file_values.to_string() if not file_values.empty else '(vazio)'}"
    )


def read_spreadsheets(input_files: list[str], sheet: str) -> tuple[pd.DataFrame, list[Path]]:
    input_paths = [Path(input_file).expanduser() for input_file in input_files]
    dataframes = []

    for input_path in input_paths:
        if not input_path.exists():
            raise FileNotFoundError(f"Arquivo nao encontrado: {input_path}")

        dataframe = pd.read_excel(input_path, sheet_name=normalize_sheet_arg(sheet))
        validate_columns(dataframe)
        dataframe["Arquivo Origem"] = input_path.name
        dataframe["Diretoria Arquivo"] = input_path.stem
        dataframes.append(dataframe)

    if not dataframes:
        raise ValueError("Informe pelo menos uma planilha de entrada.")

    return pd.concat(dataframes, ignore_index=True), input_paths


def build_summary(dataframe: pd.DataFrame) -> pd.DataFrame:
    working = dataframe.copy()
    working["Status"] = working["Status"].astype(str).str.strip()
    working["Status Resumo"] = working["Status"].str.upper().map(
        {
            "PASSED": "PASSED",
            "FAILED": "FAILED",
            "PERCENTUAL OBTIDO": "Percentual obtido",
        }
    )
    working["Quantidade"] = pd.to_numeric(working["Quantidade"], errors="coerce").fillna(0)

    summary = (
        working[working["Status Resumo"].isin(STATUS_COLUMNS)]
        .pivot_table(
            index=SUMMARY_COLUMNS,
            columns="Status Resumo",
            values="Quantidade",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    for status in STATUS_COLUMNS:
        if status not in summary.columns:
            summary[status] = 0

    summary = summary[[*SUMMARY_COLUMNS, *STATUS_COLUMNS]]
    summary["TOTAL"] = summary["PASSED"] + summary["FAILED"]
    summary = summary.sort_values(SUMMARY_COLUMNS).reset_index(drop=True)
    return summary


def build_pillar_config() -> pd.DataFrame:
    rows = []
    for pillar in PILLAR_CONFIG:
        for indicator in pillar.get("Indicadores", []):
            for requirement in indicator.get("Requisitos", []):
                rows.append(
                    {
                        "Pilar": pillar.get("Pilar"),
                        "Peso Pilar %": pillar.get("Peso Pilar %", 0),
                        "Indicador": indicator.get("Indicador"),
                        "Peso Indicador %": indicator.get("Peso Indicador %", 0),
                        "Requisito": requirement.get("Requisito"),
                        "Peso Requisito %": requirement.get("Peso Requisito %", 0),
                    }
                )

    config = pd.DataFrame(rows)
    required_config_columns = [
        "Pilar",
        "Peso Pilar %",
        "Indicador",
        "Peso Indicador %",
        "Requisito",
        "Peso Requisito %",
    ]
    missing_columns = [
        column for column in required_config_columns if column not in config.columns
    ]
    if missing_columns:
        missing = ", ".join(missing_columns)
        raise ValueError(f"Colunas ausentes em PILLAR_CONFIG: {missing}")

    config = config[required_config_columns].copy()
    for column in ["Peso Pilar %", "Peso Indicador %", "Peso Requisito %"]:
        config[column] = pd.to_numeric(config[column], errors="coerce").fillna(0)
    return config


def build_pillar_view(summary: pd.DataFrame) -> pd.DataFrame:
    config = build_pillar_config()
    detailed = summary.merge(
        config,
        on=["Indicador", "Requisito"],
        how="left",
    )
    detailed["Pilar"] = detailed["Pilar"].fillna("Nao configurado")
    for column in ["Peso Pilar %", "Peso Indicador %", "Peso Requisito %"]:
        detailed[column] = detailed[column].fillna(0)

    detailed["Resultado %"] = 0.0
    binary_mask = detailed["TOTAL"].gt(0)
    detailed.loc[binary_mask, "Resultado %"] = (
        detailed.loc[binary_mask, "PASSED"] / detailed.loc[binary_mask, "TOTAL"] * 100
    )

    percentual_mask = detailed["Percentual obtido"].gt(0)
    detailed.loc[percentual_mask, "Resultado %"] = detailed.loc[
        percentual_mask, "Percentual obtido"
    ]
    detailed["Pontuacao Requisito"] = (
        detailed["Resultado %"] * detailed["Peso Requisito %"] / 100
    )

    indicator_view = (
        detailed.groupby(
            [
                "Release Train",
                "Squad",
                "Pilar",
                "Peso Pilar %",
                "Indicador",
                "Peso Indicador %",
            ],
            dropna=False,
        )
        .agg(
            {
                "Peso Requisito %": "sum",
                "PASSED": "sum",
                "FAILED": "sum",
                "Percentual obtido": "sum",
                "TOTAL": "sum",
                "Pontuacao Requisito": "sum",
            }
        )
        .reset_index()
        .rename(
            columns={
                "Peso Requisito %": "Peso Requisitos %",
                "Pontuacao Requisito": "Resultado Indicador %",
            }
        )
    )
    indicator_view["Pontuacao Indicador no Pilar"] = (
        indicator_view["Resultado Indicador %"]
        * indicator_view["Peso Indicador %"]
        / 100
    )

    view = (
        indicator_view.groupby(
            ["Release Train", "Squad", "Pilar", "Peso Pilar %"],
            dropna=False,
        )
        .agg(
            {
                "PASSED": "sum",
                "FAILED": "sum",
                "Percentual obtido": "sum",
                "TOTAL": "sum",
                "Pontuacao Indicador no Pilar": "sum",
            }
        )
        .reset_index()
        .sort_values(["Release Train", "Squad", "Pilar"])
    )
    view = view.rename(
        columns={"Pontuacao Indicador no Pilar": "Resultado Pilar %"}
    )
    view["Pontuacao Pilar Final"] = view["Resultado Pilar %"] * view["Peso Pilar %"] / 100
    indicator_view = indicator_view.sort_values(
        ["Release Train", "Squad", "Pilar", "Indicador"]
    ).reset_index(drop=True)
    return detailed, indicator_view, view


def filter_spreadsheet(
    input_files: list[str],
    output_file: str | None,
    sheet: str,
    diretoria: str,
) -> Path:
    dataframe, input_paths = read_spreadsheets(input_files, sheet)
    output_path = Path(output_file).expanduser() if output_file else default_output_path(input_paths)

    filtered = filter_by_diretoria(dataframe, diretoria)
    if filtered.empty:
        print(
            "Aviso: nenhuma linha encontrada para a diretoria "
            f"{diretoria!r}. Diretorias disponiveis:"
        )
        print(format_diretoria_diagnostics(dataframe))

    summary = build_summary(filtered)
    summary_with_pillars, indicator_view, pillar_view = build_pillar_view(summary)
    pillar_config = build_pillar_config()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        filtered.to_excel(writer, sheet_name="Dados Filtrados", index=False)
        summary_with_pillars.to_excel(writer, sheet_name="Resumo", index=False)
        indicator_view.to_excel(writer, sheet_name="Visao Indicadores", index=False)
        pillar_view.to_excel(writer, sheet_name="Visao Pilares", index=False)
        pillar_config.to_excel(writer, sheet_name="Config Pilares", index=False)

    print(f"Arquivos lidos: {len(input_paths)}")
    print(f"Linhas encontradas: {len(filtered)}")
    print(f"Linhas no resumo: {len(summary)}")
    print(f"Linhas na visao de pilares: {len(pillar_view)}")
    print(f"Arquivo salvo em: {output_path}")
    return output_path


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        filter_spreadsheet(
            input_files=args.input_files,
            output_file=args.output,
            sheet=args.sheet,
            diretoria=args.diretoria,
        )
    except Exception as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
