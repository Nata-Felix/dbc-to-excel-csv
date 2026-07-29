# Relatório de integridade dos dados

Data da verificação: 29 de julho de 2026.

## Resultado

O arquivo `MENTBR26.DBC` foi descompactado e comparado com suas saídas CSV,
Excel e Parquet. Foram verificados **2.180 registros, 62 campos e 135.160
células por formato**.

| Formato | Registros | Campos | Células comparadas | Divergências |
|---|---:|---:|---:|---:|
| CSV | 2.180 | 62 | 135.160 | **0** |
| Excel (.xlsx) | 2.180 | 62 | 135.160 | **0** |
| Parquet | 2.180 | 62 | 135.160 | **0** |

Cabeçalhos e ordem dos campos também coincidiram em todos os formatos. A
planilha Excel não apresentou erros de fórmula.

## Metodologia

1. O DBC foi descompactado para DBF pelo conversor.
2. O DBF foi descompactado novamente por um programa de referência independente
   baseado no BLAST original. Os dois DBFs ficaram binariamente idênticos.
3. CSV e XLSX foram lidos por rotinas independentes e comparados célula por célula.
4. Parquet foi lido independentemente com PyArrow 25.0.0 e comparado da mesma forma.
5. Valores foram normalizados apenas quanto à representação: espaços de
   preenchimento DBF, datas `AAAAMMDD`/`AAAA-MM-DD` e campos vazios.

O SHA-256 lógico normalizado do DBF, CSV, XLSX e Parquet foi o mesmo:

```text
4e0709af594cc714d00dc330e6e5ce8521cdef627316e1c9fd58a24598451743
```

## Hashes dos arquivos analisados

```text
DBC     268A6D2C3E66DC5E8B589F1BB83CD3DED6DA9A27B5720D1B00E399A776A90DE5
DBF     61F4AECF6DE3629A2651FF8057DE802389639EC80288629E013B2A3E41D4D9C9
CSV     2CEDCAAD493640539C1EF99EDB7B54FB4D5070C39421097C7DF594C8D38C3806
XLSX    8ABAF11244A33B83CE7E13C8EEDF1E68B75C8EFC46C187B588004043BFAC353A
Parquet 91BC8A05D7FA3FE636BF5A9A43933AFBD9A823459BFE445E468D5A835A47B326
```

Conclusão: **não foi detectada perda de informação nos arquivos testados**.
