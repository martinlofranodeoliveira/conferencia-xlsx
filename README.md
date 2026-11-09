# Conferência XLSX

Veja o que mudou entre duas entregas de uma planilha. A ferramenta compara células por aba e coordenada, separa fórmulas de seus resultados armazenados e gera relatórios locais, sem modificar os arquivos de entrada.

[Código e versões](https://github.com/martinlofranodeoliveira/conferencia-xlsx) · [LOFRANO TECNOLOGIA](https://lofranotecnologia.com.br/)

## Começar

Requisito: Python 3.11 ou mais recente. Na pasta do projeto:

```sh
python -m venv .venv
# Linux/macOS:
source .venv/bin/activate
# Windows PowerShell, em vez da linha acima:
# .venv\Scripts\Activate.ps1
python -m pip install .
conferencia-xlsx exemplo --saida exemplo-xlsx.zip
```

O exemplo contém duas planilhas fictícias e os relatórios correspondentes. Elas foram produzidas pelo LibreOffice a partir de FODS originais, com cálculos conferidos. Para trabalhar com seus arquivos:

```sh
conferencia-xlsx comparar antes.xlsx depois.xlsx --saida conferencia.zip
```

O ZIP contém `comparacao.json`, `comparacao.md`, `alteracoes.csv` e `manifesto.json`. **Seus arquivos de entrada não são incluídos.** O exemplo é a única opção que acrescenta as duas entradas fictícias. O manifesto informa SHA-256 e tamanho dos membros para conferir a integridade do pacote.

O destino precisa ser novo; arquivos existentes, diretórios e links não são substituídos. A publicação usa um temporário irmão e um link exclusivo. Em sistemas de arquivos que não permitem esse mecanismo, a operação falha sem tentar sobrescrever o destino.

## Saídas individuais e automação

```sh
conferencia-xlsx comparar antes.xlsx depois.xlsx --formato json --saida conferencia.json
conferencia-xlsx comparar antes.xlsx depois.xlsx --formato markdown --saida conferencia.md
conferencia-xlsx comparar antes.xlsx depois.xlsx --formato csv --saida alteracoes.csv
```

| Código de saída | Significado |
| --- | --- |
| 0 | Iguais no escopo conferido; ou exemplo criado |
| 1 | Diferenças encontradas; relatório completo foi publicado |
| 2 | Erro, formato não suportado ou limite; nenhum resultado parcial foi publicado |
| 130 | Interrompido pelo usuário |

Um código 1 é o resultado normal de uma comparação com alterações. O console mostra somente estado e contagens; os arquivos de relatório incluem os valores das células. Escolha uma pasta apropriada para guardá-los.

CSV contém somente a lista de alterações. Para o veredito, hashes e cobertura completa, use JSON, Markdown ou o bundle padrão.

## O que é comparado

- Sistema de datas 1900/1904.
- Abas adicionadas/removidas, ordem relativa das abas comuns e visibilidade.
- Intervalos mesclados.
- Literais com tipos distintos: texto, número, booleano, erro e data ISO armazenada.
- Texto exato de fórmulas comuns; fórmulas compartilhadas são expandidas pelo leitor.
- Estado e valor tipado do resultado armazenado de cada fórmula.

Números são lidos do XML como decimais exatos, sem conversão para float na comparação. `1`, `1.0` e `1e0` coincidem; valores distintos além da precisão usual de ponto flutuante continuam distintos. Datas seriais são comparadas como números, com o sistema de datas conferido separadamente. Datas ISO são comparadas pela representação armazenada. Formatos de exibição não fazem parte dessa interpretação.

Uma mudança literal → fórmula, mesmo que aparente o mesmo valor, é identificada. Fórmula igual pode ter resultado armazenado diferente; fórmula diferente pode ter o mesmo resultado. A ferramenta não considera expressões matematicamente equivalentes e **não calcula fórmulas**.

## Caches não comprovam recálculo

Ausência de `<v>`, elemento `<v/>` vazio e valor armazenado são três estados distintos. Elemento vazio pode ser apenas um marcador deixado pelo produtor. Todos são apresentados com atualidade **desconhecida**; flags de recálculo ou igualdade entre caches não comprovam que a planilha foi recalculada com os dados atuais. Essa fronteira segue a [documentação de fórmulas do openpyxl](https://openpyxl.readthedocs.io/en/stable/simple_formulae.html).

## Cobertura explícita

O resultado é **igual ou diferente no escopo conferido**, não uma declaração de igualdade integral da pasta de trabalho. `identical_bytes` informa separadamente se os snapshots tinham os mesmos bytes. Aparência, estilos, formatos numéricos, gráficos, imagens, tabelas estruturadas, validações, comentários, hyperlinks, nomes definidos, assinaturas e conteúdo de vínculos externos não são comparados.

Abas são associadas por nome exato e células por coordenada. Renomear abas ou inserir/mover linhas não acompanha identidade de registros. Abas adicionadas/removidas aparecem como mudanças estruturais com contagens, sem repetir todas as suas células como alterações individuais.

Arquivos aceitos seguem o perfil XLSX OOXML Transitional. Fórmulas matriciais, data tables de fórmula, chartsheets, macros, arquivos cifrados e OOXML Strict produzem diagnóstico explícito. Relações externas nunca são seguidas. O leitor usa `openpyxl` em modo somente leitura e inspeção XML complementar; não extrai o ZIP em disco nem salva os workbooks.

## Limites

Por entrada: 16 MiB compactados, 64 MiB expandidos, 2.048 partes, 100.000 células presentes, 500.000 nós XML e 128 abas. A soma de `maior linha × maior coluna` por aba é limitada a 2.000.000 posições para impedir expansão inesperada de grades esparsas. As coordenadas reais são usadas em vez de confiar no `<dimension>` declarado pelo produtor.

Há limite de 100.000 alterações, 1.000.000 de caracteres por texto e 64 MiB por relatório/conjunto do bundle. Ao atingir um limite, a operação falha; não trunca dados nem declara igualdade com base em um prefixo. Os limites de biblioteca ficam em `Limits`. Isso não é uma garantia de uso constante de RAM.

Tokens numéricos têm até 1.024 caracteres e expoente absoluto de até 1.000.000. Textos precisam ser Unicode bem formado. As sequências de escape OOXML são decodificadas em uma única passagem por trecho de texto, incluindo pares UTF-16 válidos. Controles presentes em strings por esses escapes são preservados e escapados na representação JSON usada pelos relatórios. Os [casos de texto conferidos no LibreOffice](docs/TEXTOS-OOXML.md) cobrem strings compartilhadas, inline e fronteiras de rich text.

## API Python

```python
from conferencia_xlsx.reader import read_path
from conferencia_xlsx.compare import compare_snapshots

antes = read_path("antes.xlsx")
depois = read_path("depois.xlsx")
resultado = compare_snapshots(antes, depois)
print(resultado["status"], resultado["summary"]["changes"])
```

Os snapshots são dataclasses imutáveis, com validação nas fronteiras. O SHA-256 refere-se aos bytes capturados durante a leitura; não promete que o caminho continue igual depois. O leitor verifica metadados do mesmo descriptor antes/depois e recusa fontes alteradas durante a captura. Não há rede, telemetria, execução de fórmulas ou leitura de outros arquivos a partir de vínculos.

## Desenvolver e verificar

```sh
python -m pip install -e .
python -m unittest discover -s tests -v
python -m pip install build
python -m build
```

O [contrato](CONTRACT.md) descreve os dados e as categorias. O [ensaio independente](docs/INTEROPERABILIDADE.md) registra a versão real do LibreOffice, as entradas sintéticas, cálculos e hashes, além da diferença entre conferir fixtures e executar novamente o consumidor externo.

```sh
python tools/libreoffice_rehearsal.py --check-fixture
python tools/xstring_rehearsal.py --check-fixture
python scripts/check_release.py
python scripts/check_installed.py --wheel dist/conferencia_xlsx-1.0.0-py3-none-any.whl
```

Os dois primeiros comandos repetem as conferências das fixtures sem iniciar o LibreOffice. Os dois últimos conferem os artefatos gerados por `python -m build`, inclusive a instalação do wheel fora do checkout.

Para contribuir, abra uma issue com objetivo e uma planilha sintética mínima; inclua o teste do comportamento alterado. Evite enviar planilhas pessoais ou operacionais em issues e pull requests.

## Licença

Código, exemplos e verificadores originais sob [MIT](LICENSE), por Martin Lofrano / LOFRANO TECNOLOGIA. openpyxl segue MIT e defusedxml segue PSF; as dependências mantêm seus próprios avisos. LibreOffice é um consumidor externo usado na validação, não parte redistribuída do aplicativo.
