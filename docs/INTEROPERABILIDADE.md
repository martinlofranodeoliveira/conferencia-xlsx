# Ensaio com LibreOffice real

O leitor foi exercitado com dois XLSX produzidos pelo **LibreOffice 26.2.6.3, 620(Build:3)** instalado no sistema. Os arquivos de origem são FODS originais, criados para esta ferramenta com pessoas, títulos e registros fictícios. Não são arquivos de uma operação comercial nem cópias dos testes do LibreOffice.

O consumidor externo foi executado por um comando explícito de reprodução, com perfil temporário exclusivo. O perfil existente do computador não foi utilizado. O processo terminou e o diretório do perfil temporário foi removido. Nenhuma conta ou serviço é necessário; os FODS não contêm macros nem referências externas. Isso descreve os dados e as ações do ensaio, sem afirmar que o sistema operacional bloqueou todas as possibilidades de rede do executável.

## Fluxo executado

1. Gerar `antes.fods` e `depois.fods` usando somente a biblioteca padrão Python. São três abas em cada arquivo.
2. Converter os dois arquivos com `libreoffice --headless --convert-to xlsx`, usando `-env:UserInstallation` apontado para um diretório criado por `TemporaryDirectory`.
3. Abrir os ZIPs e XMLs produzidos com um verificador independente, sem importar openpyxl ou o leitor do aplicativo. Conferir resultados conhecidos, tipos, fórmulas, abas e mesclagens.
4. Ler os mesmos bytes com `read_bytes` e os caminhos com `read_path`. Conferir a igualdade dos snapshots e os SHA-256 dos arquivos antes e depois da leitura.
5. Comparar os snapshots e exigir as **onze mudanças esperadas**, por categoria, aba e coordenada.
6. Gerar uma fixture derivada removendo somente o elemento `<v>` de `Inventário!C1` no conteúdo XML. Essa mutação é nossa e está identificada: não foi produzida diretamente pelo LibreOffice. O comparador deve registrar exatamente uma mudança de cache, de `value` para `absent`.

As fórmulas dos FODS **não recebem resultados em cache previamente preenchidos**. Na conversão real, o LibreOffice produziu os resultados conhecidos. Isso comprova esses casos do ensaio; a ferramenta de comparação continua sem calcular fórmulas.

| Caso no XLSX produzido | Antes | Depois | O que conferir |
| --- | --- | --- | --- |
| `Inventário!C1` | `SUM(A1:B1)`, cache 15 | Mesma fórmula, cache 17 | Mudança de cache, sem mudança de fórmula |
| `Inventário!D1` | `1+1`, cache 2 | `2*1`, cache 2 | Mudança de fórmula, sem mudança de cache |
| `Inventário!F1` | Número 1 | Texto `1` | Mudança de tipo literal |
| `Inventário!G1` | Booleano verdadeiro | Booleano falso | Tipo booleano real `t=b` |
| `Inventário!H1` | Data serial 46334 | Igual | Número armazenado; aparência não comparada |
| `Inventário!I1` | `1/0`, cache `#DIV/0!` | Igual | Cache tipado como erro |
| `Inventário!J1` | Fórmula `""`, elemento `<v/>` | Igual | Cache `empty`, sem inferir atualidade |
| `Inventário!K1` | Em branco | Em branco | Célula sem conteúdo fora do mapa |
| `Inventário!L1` | Texto `=SUM(1,2)` | Igual | Texto literal, sem execução |
| Estrutura | `Inventário`, `Legado`, `Auxiliar` | `Auxiliar`, `Inventário`, `Nova` | Remoção, adição e ordem das abas comuns |
| `Auxiliar` | Visível | Oculta | Mudança de visibilidade |
| Mesclagem | `A2:B2` | `A2:C2` | Uma remoção e uma adição de intervalo |

Há ainda texto com acentos, quebra de linha, `&`, `|` e `<valor>`. Os arquivos XLSX de demonstração são distribuídos em [samples](../src/conferencia_xlsx/samples); os FODS, o relatório da execução, a fixture sem cache e os hashes ficam em [fixtures/libreoffice](../fixtures/libreoffice).

## Conferir os registros sem executar LibreOffice

Após instalar as dependências, na raiz do código:

```sh
python tools/libreoffice_rehearsal.py --check-fixture
```

Esse modo confere os hashes registrados, repete o verificador XML independente e executa o leitor/comparador atual contra os fixtures. Seu resultado contém `external_consumer_rerun: false`: não é uma nova execução do programa externo. Os testes automatizados usam esse modo e podem rodar sem LibreOffice instalado.

## Repetir o consumidor externo

O caminho abaixo foi executado em Linux com `/usr/bin/libreoffice`. A reprodução externa usa grupos de processo POSIX para poder encerrar apenas o processo do ensaio em caso de timeout. Ela não procura nem encerra outras instâncias do LibreOffice.

```sh
mkdir -p local
python tools/libreoffice_rehearsal.py \
  --output local/novo-ensaio \
  --libreoffice /usr/bin/libreoffice
```

O diretório final precisa ser novo. O comando registra a versão observada, argumentos sem caminhos pessoais, oráculos, comparações e hashes. O XLSX é produzido pelo programa instalado, portanto versões diferentes do consumidor podem alterar os bytes ou a representação de algum recurso. Divergências no oráculo encerram o ensaio; não são transformadas em sucesso.

Este ensaio não executou Microsoft Excel, não certifica compatibilidade com todo o formato XLSX e não verifica aparência, gráficos ou equivalência matemática das fórmulas. Fora dos casos controlados aqui, a atualidade de qualquer resultado armazenado permanece **desconhecida**, mesmo quando há um valor ou uma propriedade de recálculo.

Fontes primárias: [parâmetros de inicialização do LibreOffice](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html), [fórmulas no openpyxl](https://openpyxl.readthedocs.io/en/stable/simple_formulae.html), [modo de leitura e dimensões](https://openpyxl.readthedocs.io/en/stable/optimized.html).
