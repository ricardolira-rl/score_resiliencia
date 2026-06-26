# Score Resiliencia

Script Python para ler uma planilha Excel, filtrar os registros da diretoria
`DIR CYBER SECURITY` e gerar uma nova planilha com visoes de resumo.

## O que o script gera

- `Dados Filtrados`: registros filtrados pela coluna `Diretoria`.
- `Resumo`: contagem/soma por `Release Train`, `Squad`, `Indicador` e `Requisito`.
- `Visao Indicadores`: consolidacao ponderada por indicador.
- `Visao Pilares`: consolidacao ponderada por pilar.
- `Config Pilares`: configuracao de pilares, indicadores, requisitos e pesos usada no calculo.

## Como executar

```powershell
python filter_dir_cyber_security.py "C:\Users\ricar\Desktop\DIR CYBER SECURITY.xlsx"
```

Para escolher o arquivo de saida:

```powershell
python filter_dir_cyber_security.py "C:\Users\ricar\Desktop\DIR CYBER SECURITY.xlsx" -o "C:\Users\ricar\Documents\Score Resiliencia\DIR CYBER SECURITY_filtrado.xlsx"
```

## Configuracao dos pesos

Os pesos ficam no codigo, na variavel `PILLAR_CONFIG`, em tres niveis:

- `Peso Pilar %`
- `Peso Indicador %`
- `Peso Requisito %`

Exemplo:

```python
PILLAR_CONFIG = [
    {
        "Pilar": "Desenvolvimento Seguro",
        "Peso Pilar %": 100,
        "Indicadores": [
            {
                "Indicador": "Testes em tempo de desenvolvimento (Sonar, TaaC, Hopper)",
                "Peso Indicador %": 100,
                "Requisitos": [
                    {"Requisito": "SonarQube - bugs", "Peso Requisito %": 50},
                    {"Requisito": "Taac e Hopper - Reuso e Casos de Teste", "Peso Requisito %": 50},
                ],
            },
        ],
    },
]
```

## Dependencias

```powershell
pip install pandas openpyxl
```
