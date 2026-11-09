# Contrato · Conferência XLSX

Biblioteca/CLI Python 3.11+, PT-BR, MIT. Compara dois snapshots XLSX locais por nome de aba e coordenada. Não altera nem salva as entradas. Saídas JSON, Markdown, CSV e ZIP de relatórios; sem interface web, rede, conta ou publicação automática. openpyxl 3.1.5 é o leitor central, em read_only=True/data_only=False/keep_links=False; defusedxml 0.7.1 é obrigatório. O XML OOXML complementa o leitor para números exatos, presença do cache e mesclagens. Não alegar igualdade integral de arquivos ou cálculo de fórmulas.

## Escopo e resultado

Comparar: sistema de datas 1900/1904; adição/remoção/ordem/visibilidade de abas; mesclagens; células literais tipadas; texto exato de fórmulas (shared formulas expandidas pelo openpyxl); estado/valor tipado dos resultados armazenados das fórmulas. Fórmulas matriciais/dataTable e chartsheets geram erro explícito antes de qualquer resultado. Macros/arquivos cifrados não são aceitos. Fórmulas comuns, inclusive funções desconhecidas pelo aplicativo, são texto, não avaliadas.

Fora de escopo: aparência/estilos/formatos numéricos, gráficos/imagens, tabelas estruturadas, validações, comentários, hyperlinks, nomes definidos, assinaturas e vínculos externos. Inventariar contagens de partes/relacionamentos ignorados quando possível; sempre declarar exclusões. Relações externas nunca são seguidas. Renomear abas ou inserir/mover linhas não é rastreado por identidade: gera diferenças em nomes/coordenadas.

Status completo `equal_in_scope` ou `different_in_scope`; `identical_bytes` separado por SHA-256. Erros/limites/recursos desconhecidos não geram relatório parcial nem igualdade. Resultados armazenados têm atualidade sempre `unknown`; <v> ausente, <v/> vazio e valor presente são três estados diferentes. Cache vazio não comprova resultado calculado em branco; propriedades de recálculo não comprovam atualização.

## Tipos compartilhados · src/conferencia_xlsx/models.py

Dataclasses frozen, listas imutáveis como tuple:

- `AuditError(code:str,message:str)` em errors.py; mensagens sanitizadas não repetem caminhos, valores nem exceções da biblioteca.
- `Limits(max_file_bytes=16MiB,max_expanded_bytes=64MiB,max_parts=2048,max_cells=100000,max_grid_slots=2000000,max_xml_nodes=500000,max_sheets=128,max_changes=100000,max_text_chars=1000000)`. Inteiros positivos, rejeitar bool; limites explícitos validados pelo leitor e comparador. Não precisa expor todos no CLI.
- `Value(kind:str,value:str|bool)` com kind string/number/boolean/error/date. Números são strings decimais canônicas exatas (sem float ou normalize arredondado). Não distinguir 1/1.0/1e0; preservar diferença além de IEEE754. Pode usar notação científica compacta para expoentes grandes. Datas `t=d` comparam representação ISO armazenada literalmente; datas seriais permanecem números e a época é comparada separadamente. Não comparar aparência renderizada.
- `Cache(state:str,value:Value|None)` state absent/empty/value; só value tem Value. Atualidade unknown no relatório.
- `Cell(coordinate:str,literal:Value|None,formula:str|None,cache:Cache|None)`. Fórmula implica literal=None, cache presente; literal implica fórmula/cache None. Células realmente em branco, inclusive só estilo, ficam fora do mapa; string vazia explícita permanece string vazia.
- `Sheet(name:str,state:str,cells:tuple[Cell,...],merges:tuple[str,...])`; cells ordenadas por linha/coluna, merges por intervalo normalizado determinístico. state visible/hidden/veryHidden.
- `Snapshot(sha256:str,size_bytes:int,epoch:str,sheets:tuple[Sheet,...],ignored_parts:tuple[tuple[str,int],...],warning_count:int=0)`.

## Leitor · reader.py

`read_bytes(data:bytes,limits:Limits=Limits())->Snapshot`, `read_path(path,limits=...)->Snapshot`.

Capturar entrada uma vez como bytes limitados. Em arquivo, usar descriptor de arquivo regular e fstat antes/depois da mesma API para detectar mudanças; rejeitar pipe/dispositivo, ler no máximo limite+1. O SHA refere-se aos bytes capturados, não promete que o caminho continue igual depois. Não extrair ZIP em disco.

Preflight ZIP: partes únicas/caminhos internos canônicos, sem criptografia, links ZIP ou formatos de compressão não suportados; CRC e limites de bytes reais conferidos. XML sem DTD/entidades/externos, com contagem de nós limitada. Resolver workbook/worksheets por relações do pacote; rejeitar referências internas ausentes, nomes/coords duplicados ou conteúdo celular não reconhecido. Relacionamentos externos são contados e nunca acessados. O formato é OOXML Transitional da família XLSX; Strict produz unsupported-format explícito.

Pré-analisar células realmente presentes, seus tipos, cache bruto e merges. Contar células e slots de grade derivados das coordenadas reais antes de iter_rows, para limitar a expansão de intervalos esparsos. Não confiar no <dimension> declarado; reset_dimensions e limites explícitos na iteração. Usar openpyxl para resolver strings e fórmulas comuns/compartilhadas. Números/valores de cache vêm de XML validado para não perder precisão ou confundir absent/empty. max_grid_slots soma maxRow*maxCol de cada aba; a rejeição é um limite de recurso, nunca comparação parcial. Não carregar desenhos/imagens em modo normal.

## Comparação · compare.py

`compare_snapshots(before:Snapshot,after:Snapshot,limits=Limits())->dict` JSON compatível e determinístico, sem relógio/caminhos absolutos.

Raiz: schema_version=1, status, identical_bytes, scope (included/excluded,coordinate_policy,formula_policy,cache_freshness='unknown'), before/after (sha256,size_bytes,epoch,sheet_names,ignored_parts,warning_count), summary (changes, por kind), changes lista.

Cada mudança `{kind,sheet,coordinate,before,after}`. kind: epoch, sheet_added, sheet_removed, sheet_order, sheet_state, merge_added, merge_removed, cell_added, cell_removed, literal, formula, cache. sheet/coordinate null quando não aplicáveis. before/after são valores JSON tipados ou resumos estruturais; fórmula renderizada como string, cache como {state,value,freshness:'unknown'}. Abas adicionadas/removidas têm mudança estrutural com contagem de células/mesclagens, sem expandir todas em mudanças redundantes. Em abas comuns, união de células: apenas uma mudança cell_added/removed, podendo ter fórmula e cache internos; comuns podem gerar formula/cache separadas quando pertinente. Alteração literal↔fórmula gera mudança formula com representação completa do conteúdo; não duplicar em literal. Fórmulas iguais podem ter cache diferente. Números iguais normalizados não mudam. Ordenar tipos estruturais em ordem fixa e depois abas por ordem do snapshot antes, acrescendo novas na ordem depois; células linha/coluna e categorias fixas. Não inferir correspondências por nome parecido/valor.

API recebe dataclasses de snapshots; validar invariantes ao entrar, não aceitar duplicatas/shape inválido ou forged tipo booleano/número inconsistente. Não mutar entradas. Ao atingir max_changes, erro antes de retornar resultado; não truncar.

## CLI/publicação · root

Comando `conferencia-xlsx comparar ANTES DEPOIS --saida ARQUIVO [--formato bundle|json|markdown|csv]`, bundle default. Exit0 iguais no escopo, exit1 diferenças completas, exit2 erro. Nenhum conteúdo de célula no stdout/stderr padrão: apenas status/contagens. Relatórios explícitos incluem dados selecionados e SHA; o bundle não inclui entradas. `conferencia-xlsx exemplo --saida demo.zip` inclui duas entradas sintéticas empacotadas e relatórios para reproduzir.

Relatórios são preparados integralmente, gravados em temporário irmão e publicados por link exclusivo sem sobrescrever destino existente (inclusive symlink); se hardlink não é suportado, falhar sem fallback destrutivo. Limpar só temporário próprio. Não garantir transação de origem, autenticidade ou durabilidade contra falta de energia. ZIP contém comparacao.json/comparacao.md/alteracoes.csv/manifesto.json (e antes.xlsx/depois.xlsx somente no exemplo). Manifesto contém versão, SHA e tamanhos dos membros, sem hash autorreferente. JSON ensure_ascii=False UTF-8/LF, Markdown escaping de HTML/pipes/links/linhas e CSV com proteção de prefixos de fórmula.

## Integração e distribuição

Dois FODS originais sintéticos convertidos por /usr/bin/libreoffice headless, perfil temporário exclusivo e sem arquivos pessoais. Conferir fórmulas/caches por XML independente e oráculos de valores conhecidos; inputs inalterados após leitura. Casos de tipos, fórmula igual com cache diferente, fórmula diferente com cache igual, aba e merge alterados. Fixture derivada sem <v> identifica cache ausente. Registrar versão real/flags/SHA, reproduzir sem copiar código/licenças alheios.

Fixtures originais publicadas em src/conferencia_xlsx/samples/antes.xlsx e depois.xlsx para demo instalada. Core unit tests podem usar openpyxl para casos adicionais e mutações OOXML explícitas, sem apresentar isso como consumidor independente. Wheel/sdist, execução instalada fora do checkout, CI Python3.11–3.14 e Windows3.14, release com hashes e publicação no GitHub. Perfil: somente PROJETOS-ABERTOS.md/EVIDENCIAS.md; README principal byte a byte preservado. Nenhum dado Meli/auth/serviço local é necessário.
