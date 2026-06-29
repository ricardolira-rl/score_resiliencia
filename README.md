# Score Resiliencia

Scripts Python para ler uma planilha Excel do Score de Resiliencia, identificar
a diretoria pelo nome do arquivo sem extensao e gerar visoes consolidadas por
`Release Train`, `Squad`, `Indicador`, `Requisito` e `Pilar`.

Por padrao, os scripts tambem filtram somente linhas em que a coluna
`Comunidade/Supt` seja igual a `CYBER SECURITY`.

## Scripts e arquivos

### `atualizar_squad_por_sigla.py`

Script de pre-processamento.

Ele le uma ou mais planilhas originais, procura linhas onde a coluna `Squad`
esteja igual a `SQUAD NAOIDENTIFICADA` e usa a coluna `Sigla` para substituir
pelo nome correto da Squad e da Release Train configurados na variavel interna
`SIGLA_DESTINO_MAP`.

Entrada:

- planilha Excel original;
- coluna `Squad`;
- coluna `Release Train`;
- coluna `Sigla`.

Saida:

- arquivo `.xlsx` com sufixo `_squads_atualizadas.xlsx`.

### `filter_dir_cyber_security.py`

Script principal do projeto.

Ele le uma ou mais planilhas originais, identifica a diretoria pelo nome do
arquivo sem extensao e gera uma nova planilha Excel com abas de dados
filtrados, resumo, visao por indicador, visao por pilar e configuracao usada no
calculo.

Entrada:

- planilha Excel original;
- coluna `Diretoria`;
- coluna `Comunidade/Supt`;
- coluna `Release Train`;
- coluna `Squad`;
- coluna `Indicador`;
- coluna `Requisito`;
- coluna `Status`;
- coluna `Quantidade`.

Saida:

- arquivo `.xlsx` filtrado e calculado.

### `listar_indicadores_requisitos.py`

Script auxiliar para preparar a configuracao dos pilares.

Ele varre uma ou mais planilhas originais, identifica os pares unicos de `Indicador` e
`Requisito`, mostra uma amostra no terminal e gera um arquivo com uma sugestao
inicial de `PILLAR_CONFIG` com 6 pilares.

Entrada:

- planilha Excel original.

Saida:

- lista de indicadores e requisitos no terminal;
- arquivo `pillar_config_sugerido.py` com 6 pilares.

### `pillar_config_sugerido.py`

Arquivo gerado pelo script `listar_indicadores_requisitos.py`.

Ele nao executa o calculo final sozinho. A funcao dele e servir como base para
copiar ou adaptar a variavel `PILLAR_CONFIG` dentro do script
`filter_dir_cyber_security.py`.

Entrada:

- gerado automaticamente a partir da planilha original.

Saida:

- bloco Python sugerido para configurar 6 pilares, indicadores, requisitos e
  pesos.

## 1. Instalar dependencias

Execute no PowerShell:

```powershell
pip install pandas openpyxl
```

## 2. Atualizar Squads nao identificadas

Antes dos demais scripts, configure o dicionario `SIGLA_DESTINO_MAP` dentro do
arquivo `atualizar_squad_por_sigla.py`.

Exemplo:

```python
SIGLA_DESTINO_MAP = {
    "ABC": {
        "Squad": "NOME DA SQUAD ABC",
        "Release Train": "NOME DA RT ABC",
    },
    "XYZ": {
        "Squad": "NOME DA SQUAD XYZ",
        "Release Train": "NOME DA RT XYZ",
    },
}
```

Depois execute:

```powershell
python atualizar_squad_por_sigla.py "C:\caminho\DIR CYBER SECURITY.xlsx" "C:\caminho\DIR TEC OPER CYBER SECURITY.xlsx"
```

O script gera novos arquivos ao lado dos originais:

```text
DIR CYBER SECURITY_squads_atualizadas.xlsx
DIR TEC OPER CYBER SECURITY_squads_atualizadas.xlsx
```

Use esses arquivos atualizados nos proximos scripts.

Mesmo com o sufixo `_squads_atualizadas`, os demais scripts continuam
inferindo a diretoria pelo nome original do arquivo. Por exemplo,
`DIR CYBER SECURITY_squads_atualizadas.xlsx` continua sendo tratado como
`DIR CYBER SECURITY`.

## 3. Rodar o inventario de indicadores e requisitos

Antes de configurar os pilares, rode o script de inventario:

```powershell
python listar_indicadores_requisitos.py "C:\Users\ricar\Desktop\DIR CYBER SECURITY_squads_atualizadas.xlsx"
```

Para ler duas planilhas na mesma execucao, informe os dois caminhos:

```powershell
python listar_indicadores_requisitos.py "C:\caminho\primeira_planilha.xlsx" "C:\caminho\segunda_planilha.xlsx"
```

Por padrao, ele processa todos os arquivos informados. Cada arquivo vira uma
diretoria usando o nome do arquivo sem extensao.

Por padrao, ele tambem filtra somente a comunidade `CYBER SECURITY` na coluna
`Comunidade/Supt`.

Para filtrar uma diretoria especifica, informe o nome do arquivo sem extensao:

```powershell
python listar_indicadores_requisitos.py "C:\caminho\segunda_planilha.xlsx" --diretoria "DIR TEC OPER CYBER SECURITY"
```

Para listar todas as comunidades, sem aplicar esse filtro:

```powershell
python listar_indicadores_requisitos.py "C:\caminho\segunda_planilha.xlsx" --comunidade ""
```

Saida esperada no terminal:

```text
Diretoria filtrada: Todas
Comunidade filtrada: CYBER SECURITY
Indicadores encontrados: 1
Requisitos encontrados: 2
```

Esse comando executa o script `listar_indicadores_requisitos.py`.

Ele tambem gera o arquivo:

```text
pillar_config_sugerido.py
```

Use esse arquivo como base para configurar os 6 pilares, indicadores,
requisitos e pesos no script principal.

## 4. Configurar pilares e pesos no script principal

Abra o arquivo `filter_dir_cyber_security.py` e ajuste a variavel
`PILLAR_CONFIG`.

A configuracao possui 6 pilares e tres niveis de peso:

- `Peso Pilar %`: peso do pilar na visao final.
- `Peso Indicador %`: peso do indicador dentro do pilar.
- `Peso Requisito %`: peso do requisito dentro do indicador.

Exemplo:

```python
PILLAR_CONFIG = [
    {
        "Pilar": "Pilar 1 - Ajustar nome",
        "Peso Pilar %": 16.67,
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
```

## 5. Gerar a planilha final

Execute:

```powershell
python filter_dir_cyber_security.py "C:\Users\ricar\Desktop\DIR CYBER SECURITY_squads_atualizadas.xlsx"
```

Para consolidar duas planilhas em uma unica saida:

```powershell
python filter_dir_cyber_security.py "C:\caminho\primeira_squads_atualizadas.xlsx" "C:\caminho\segunda_squads_atualizadas.xlsx" -o "C:\Users\ricar\Documents\Score Resiliencia\score_resiliencia_filtrado.xlsx"
```

Para escolher o local e nome do arquivo de saida:

```powershell
python filter_dir_cyber_security.py "C:\Users\ricar\Desktop\DIR CYBER SECURITY_squads_atualizadas.xlsx" -o "C:\Users\ricar\Documents\Score Resiliencia\DIR CYBER SECURITY_filtrado.xlsx"
```

Esse comando executa o script `filter_dir_cyber_security.py`.

Se o resultado final ficar vazio, o script imprime no terminal as diretorias
inferidas pelo nome dos arquivos lidos.

A diretoria considerada pelo filtro e sempre o nome do arquivo sem extensao.
Por exemplo, um arquivo chamado `DIR CYBER SECURITY.xlsx` sera tratado como
diretoria `DIR CYBER SECURITY`, mesmo que a coluna `Diretoria` tenha outro
valor.

Outro exemplo: um arquivo chamado `DIR TEC OPER CYBER SECURITY.xlsx` sera
tratado como diretoria `DIR TEC OPER CYBER SECURITY`.

Além disso, a saída final usa apenas linhas em que `Comunidade/Supt` seja
`CYBER SECURITY`, a menos que voce rode com `--comunidade ""`.

## Colunas obrigatorias

As planilhas podem ter colunas extras ou estruturas diferentes, mas precisam
conter estas colunas:

- `Diretoria`
- `Comunidade/Supt`
- `Release Train`
- `Squad`
- `Sigla`
- `Indicador`
- `Requisito`
- `Status`
- `Quantidade`

O filtro de diretoria usa a coluna `Diretoria Arquivo`, criada
automaticamente a partir do nome do arquivo sem extensao. Esse filtro ignora
diferencas de maiusculas/minusculas, espacos repetidos e espacos invisiveis.

O filtro de comunidade usa a coluna `Comunidade/Supt` e tambem ignora
diferencas de maiusculas/minusculas, espacos repetidos e espacos invisiveis.

## 6. Abas geradas no Excel

A planilha final possui as abas:

- `Dados Filtrados`: linhas dos arquivos filtrados pela diretoria informada.
- `Resumo`: consolidacao por `Release Train`, `Squad`, `Indicador` e
  `Requisito`, com `PASSED`, `FAILED`, `Percentual obtido`, pesos e pontuacao.
- `Visao Indicadores`: consolidacao ponderada por indicador.
- `Visao Pilares`: consolidacao ponderada por pilar.
- `Config Pilares`: configuracao de pesos usada na execucao.

## 7. Regras principais do calculo

- Para status `PASSED` e `FAILED`, o script soma a coluna `Quantidade`.
- Para status `Percentual obtido`, o script usa a soma da coluna `Quantidade`
  como resultado percentual.
- A pontuacao do requisito usa `Resultado % * Peso Requisito %`.
- A pontuacao do indicador usa a soma dos requisitos ponderada pelo
  `Peso Indicador %`.
- A pontuacao final do pilar usa o resultado do pilar ponderado pelo
  `Peso Pilar %`.

## Fluxo recomendado

1. Instale as dependencias.
2. Edite `SIGLA_DESTINO_MAP` no `atualizar_squad_por_sigla.py`.
3. Rode `atualizar_squad_por_sigla.py` para corrigir `SQUAD NAOIDENTIFICADA`
   com Squad e Release Train.
4. Rode `listar_indicadores_requisitos.py` para descobrir os indicadores e
   requisitos existentes na planilha.
5. Use o arquivo `pillar_config_sugerido.py`, com 6 pilares, como referencia.
6. Edite o `PILLAR_CONFIG` no `filter_dir_cyber_security.py`.
7. Rode `filter_dir_cyber_security.py` para gerar a planilha final.
