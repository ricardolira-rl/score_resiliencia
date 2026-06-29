from __future__ import annotations

import argparse
from pathlib import Path
import sys

import pandas as pd


SQUAD_NAO_IDENTIFICADA = "SQUAD NAOIDENTIFICADA"
REQUIRED_COLUMNS = ["Squad", "Release Train", "Sigla"]

# Preencha este dicionario com as regras reais.
# Exemplo:
# SIGLA_DESTINO_MAP = {
#     "ABC": {
#         "Squad": "NOME DA SQUAD ABC",
#         "Release Train": "NOME DA RT ABC",
#     },
#     "XYZ": {
#         "Squad": "NOME DA SQUAD XYZ",
#         "Release Train": "NOME DA RT XYZ",
#     },
# }
SIGLA_DESTINO_MAP = {
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Atualiza linhas com Squad igual a 'SQUAD NAOIDENTIFICADA' usando "
            "a coluna Sigla e o dicionario interno SIGLA_DESTINO_MAP."
        )
    )
    parser.add_argument(
        "input_files",
        nargs="+",
        help="Caminho de uma ou mais planilhas Excel de entrada.",
    )
    parser.add_argument(
        "-o",
        "--output",
        help=(
            "Arquivo de saida. Use apenas quando informar uma unica planilha "
            "de entrada."
        ),
    )
    parser.add_argument(
        "--output-dir",
        help=(
            "Diretorio de saida para uma ou mais planilhas. Se omitido, salva "
            "ao lado de cada arquivo original."
        ),
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="Nome ou indice da aba a ser lida. Padrao: primeira aba.",
    )
    return parser


def normalize_sheet_arg(sheet: str) -> str | int:
    try:
        return int(sheet)
    except ValueError:
        return sheet


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).replace("\xa0", " ").split()).upper()


def default_output_path(input_path: Path, output_dir: str | None) -> Path:
    file_name = f"{input_path.stem}_squads_atualizadas.xlsx"
    if output_dir:
        return Path(output_dir).expanduser() / file_name
    return input_path.with_name(file_name)


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


def normalized_sigla_map() -> dict[str, dict[str, str]]:
    normalized_map = {}
    for sigla, destino in SIGLA_DESTINO_MAP.items():
        normalized_sigla = normalize_text(sigla)
        squad = str(destino.get("Squad", "")).strip()
        release_train = str(destino.get("Release Train", "")).strip()

        if normalized_sigla and squad and release_train:
            normalized_map[normalized_sigla] = {
                "Squad": squad,
                "Release Train": release_train,
            }

    return normalized_map


def update_squad_by_sigla(dataframe: pd.DataFrame) -> tuple[pd.DataFrame, int, list[str]]:
    validate_columns(dataframe)

    sigla_map = normalized_sigla_map()
    updated = dataframe.copy()

    unidentified_mask = updated["Squad"].map(normalize_text).eq(SQUAD_NAO_IDENTIFICADA)
    mapped_destinations = updated["Sigla"].map(lambda value: sigla_map.get(normalize_text(value)))
    update_mask = unidentified_mask & mapped_destinations.notna()

    updated.loc[update_mask, "Squad"] = mapped_destinations.loc[
        update_mask
    ].map(lambda destino: destino["Squad"])
    updated.loc[update_mask, "Release Train"] = mapped_destinations.loc[
        update_mask
    ].map(lambda destino: destino["Release Train"])

    unmatched_siglas = (
        updated.loc[unidentified_mask & mapped_destinations.isna(), "Sigla"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .sort_values()
        .tolist()
    )

    return updated, int(update_mask.sum()), unmatched_siglas


def process_file(
    input_file: str,
    output_file: str | None,
    output_dir: str | None,
    sheet: str,
) -> Path:
    input_path = Path(input_file).expanduser()
    if not input_path.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {input_path}")

    if output_file and output_dir:
        raise ValueError("Use --output ou --output-dir, nao ambos.")

    output_path = (
        Path(output_file).expanduser()
        if output_file
        else default_output_path(input_path, output_dir)
    )

    dataframe = pd.read_excel(input_path, sheet_name=normalize_sheet_arg(sheet))
    updated, updated_rows, unmatched_siglas = update_squad_by_sigla(dataframe)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    updated.to_excel(output_path, index=False)

    print(f"Arquivo lido: {input_path}")
    print(f"Linhas atualizadas: {updated_rows}")
    if unmatched_siglas:
        print("Siglas sem mapeamento para Squad e Release Train:")
        for sigla in unmatched_siglas:
            print(f"- {sigla}")
    print(f"Arquivo salvo em: {output_path}")
    print()

    return output_path


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.output and len(args.input_files) != 1:
        print("Erro: use --output apenas com uma unica planilha de entrada.", file=sys.stderr)
        return 1

    try:
        for input_file in args.input_files:
            process_file(
                input_file=input_file,
                output_file=args.output,
                output_dir=args.output_dir,
                sheet=args.sheet,
            )
    except Exception as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
