# Textos escapados no SpreadsheetML

Este ensaio estreito foi acrescentado após encontrar uma perda de texto na leitura de shared strings pelo openpyxl 3.1.5: o texto literal `x005F_`, sem sublinhado inicial, podia desaparecer. Isso poderia produzir uma comparação falsamente igual. O complemento do leitor precisa conservar o valor interpretado do texto, além de usar a biblioteca central.

O XLSX do ensaio é original e mínimo. A coluna A armazena as sequências em uma shared string table (SST); a coluna B armazena o mesmo conteúdo como inline string. Um LibreOffice real converte o arquivo para CSV UTF-8, e um parser CSV da biblioteca padrão verifica os valores observados. São **14 casos, 28 células**, incluindo dois limites entre trechos de rich text. Não há fórmulas, macros, referências externas ou dados pessoais no arquivo.

| Texto dentro do elemento XML `t` | Valor observado em SST e inline |
| --- | --- |
| `x005F_` | `x005F_`, intacto |
| `_x005F_` | Um sublinhado |
| `_x0041_` | `A` |
| `_x005F_x0041_` | Texto literal `_x0041_`, sem uma segunda decodificação |
| `_xD83D__xDE00_` | `😀`, um único ponto de código U+1F600 |
| `_x000A_` | Quebra de linha U+000A |
| `_x005F_x000A_` | Texto literal `_x000A_` |
| `_x000D_` | Retorno de carro U+000D |

Outros casos verificam caixa dos dígitos hexadecimais e duas sequências consecutivas. Em rich text, dois elementos `t` contendo `_x004` e `1_` produzem o texto literal `_x0041_`, não `A`. O mesmo ocorre com `_x000` e `D_`: permanece `_x000D_`, sem virar retorno de carro. Portanto, decodificar depois de juntar todos os trechos alteraria valores válidos.

A [documentação Microsoft de ST_Xstring](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bd0aa042-434a-4ca7-b25f-4e1fd25a954d) descreve o escape de retorno de carro e o uso do escape de sublinhado para proteger o começo de uma sequência. A evidência de aplicação deste repositório é do **LibreOffice**, não de uma execução do Microsoft Excel. Não se deve ampliar esses casos para uma alegação de compatibilidade integral com todo o formato.

## Conferir os registros

```sh
python tools/xstring_rehearsal.py --check-fixture
```

O modo offline confere os hashes dos arquivos registrados, o conteúdo lógico dos membros do XLSX original, os valores do CSV realmente exportado e o resultado do leitor atual. Ele retorna `external_consumer_rerun: false`; não inicia o LibreOffice.

Para repetir o consumidor externo em POSIX com um diretório de saída novo:

```sh
mkdir -p local
python tools/xstring_rehearsal.py \
  --output local/novo-ensaio-textos \
  --libreoffice /usr/bin/libreoffice
```

O perfil é temporário e exclusivo, e os argumentos do [filtro CSV oficial](https://help.libreoffice.org/latest/en-US/text/shared/guide/csv_params.html) selecionam vírgula, aspas e UTF-8. Os artefatos e a versão efetivamente observada ficam em [fixtures/xstring](../fixtures/xstring). A geração lógica é conferida por membro do ZIP; não se exige igualdade de cabeçalhos ou bytes comprimidos entre versões de Python e zlib.
