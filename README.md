<p align="center">
  <img src="assets/solppe-logo-horizontal.png" alt="SOLPPE" width="360">
</p>

<h1 align="center">DBC to Excel/CSV</h1>

<p align="center">
  Conversor gratuito para transformar arquivos DBC do DATASUS em Excel, CSV, Parquet, DBF ou JSONL no Windows.
</p>

## Download

Baixe somente o arquivo **`DBC-to-Excel-CSV.exe`** na
[versão mais recente](https://github.com/Nata-Felix/dbc-to-excel-csv/releases/latest).

Não é necessário instalar Python, Java, Microsoft Excel ou qualquer outro componente.
O programa é portátil, funciona no Windows 10/11 de 64 bits e não exige permissão de administrador.

> O executável ainda não possui assinatura digital. Por isso, o Windows SmartScreen pode exibir
> um aviso na primeira abertura. Confirme sempre que o arquivo foi baixado desta página oficial.

## Como usar

1. Abra `DBC-to-Excel-CSV.exe`.
2. Selecione um arquivo `.dbc` ou uma pasta com vários arquivos.
3. Confirme a pasta de saída sugerida como `convertidos`.
4. Marque os formatos desejados e clique em **Converter arquivos**.

CSV, Excel e Parquet já aparecem selecionados. O arquivo DBC original nunca é modificado.

## Formatos

- **CSV:** UTF-8 com ponto e vírgula, pronto para o Excel em português.
- **Excel (.xlsx):** planilha real com cabeçalho, filtros, datas e tipos preservados.
- **Parquet:** formato colunar compacto, indicado para análise de dados.
- **DBF:** arquivo descompactado compatível com ferramentas legadas.
- **JSONL:** um registro JSON por linha.

## Fidelidade dos dados

O projeto foi validado com um DBC real do DATASUS contendo **2.180 registros,
62 campos e 135.160 células**. CSV, Excel e Parquet foram comparados célula por
célula com a estrutura DBF descompactada, resultando em **zero divergências**.

Consulte o [relatório de integridade](RELATORIO_INTEGRIDADE.md) para conhecer a metodologia e os hashes.

## Compilar no Windows

Requisitos:

- Windows x64;
- LLVM-MinGW no `PATH`;
- PowerShell 5.1 ou superior.

Execute:

```powershell
./build-windows.ps1
```

O script baixa a biblioteca oficial DuckDB 1.5.5, confere seu SHA-256 e gera o
executável portátil em `dist/DBC-to-Excel-CSV.exe`.

## Licença

O código-fonte é distribuído sob a licença MIT. Consulte
[avisos de terceiros](THIRD_PARTY_NOTICES.md) e [uso da marca](TRADEMARKS.md).

## Gostou do projeto?

⭐ **Deixe uma estrela no repositório.** Isso ajuda outras pessoas a encontrarem a ferramenta e incentiva novas melhorias.
