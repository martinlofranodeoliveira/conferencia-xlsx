# Validação da versão 1.0.0

A suíte local final tem **64 testes**, executados com Python 3.12.8. Os workflows repetem os testes, os checkers de fixtures, o build, a conferência dos arquivos distribuídos e a instalação do wheel fora do checkout em Python 3.11–3.14 no Linux e Python 3.14 no Windows. Um job Linux executa os dois ensaios com o LibreOffice instalado naquele ambiente.

## Casos de domínio e arquivos

Os testes conferem números além da precisão de float sob contexto Decimal reduzido, tipos literal/booleano/erro/data, datas seriais 59 e 60 distintas, sistema 1900/1904, fórmulas compartilhadas, fórmulas e caches alterados separadamente, adição/remoção/ordem/visibilidade de abas e mesclagens. Coordenadas implícitas e dimensões declaradas incorretamente são exercitadas. Relatórios limitados não podem retornar igualdade de um prefixo.

O leitor é conferido com entradas inválidas e limites de pacote, XML e grade, sem extração em disco. As entradas de arquivo são capturadas uma vez; testes e ensaios conferem os hashes antes/depois e alterações detectadas durante a captura. A CLI é exercitada com os códigos 0/1/2, destino existente, arquivos inválidos, exemplo completo e diagnóstico sanitizado.

JSON e CSV preservam os valores tipados e controles decodificados como escapes. Markdown é conferido para HTML, links, imagens, pipes e linhas provenientes de células. O CSV protege prefixos de fórmula em nomes de abas. Manifestos cobrem todos os membros do bundle, que não inclui entradas no comando de comparação. Falhas de publicação não usam sobrescrita como alternativa.

## Correções encontradas na revisão

Uma revisão independente identificou três falhas de fidelidade/formato, corrigidas antes da publicação:

1. A normalização de strings compartilhadas da dependência podia remover o literal `x005F_`. O leitor passou a usar os nós textuais validados, com decodificação ST_Xstring em uma passagem por nó. O LibreOffice real confirmou os valores em [14 casos e 28 células](TEXTOS-OOXML.md), incluindo pares UTF-16 e fronteiras de rich text.
2. Identificar XML apenas pelo sufixo permitia que uma parte declarada XML por Content Types ficasse fora dos controles. Agora overrides e defaults também identificam todas essas partes, sujeitas ao mesmo orçamento de nós e proibição de DTD/entidades.
3. Conferir o resultado de ZipExtFile contra o tamanho anunciado não comprovava a expansão real do fluxo. Agora o fluxo DEFLATE é validado com orçamento real, término completo, tamanho e CRC, preservando a verificação de cabeçalhos e sobreposição da biblioteca ZIP.

Uma confirmação dirigida posterior pontuou **os três achados como resolvidos**, lendo o código e executando os 64 testes e os dois checkers de fixtures. Esse veredito se limita à lista acima; não representa uma certificação global do formato ou uma garantia contra todos os arquivos possíveis.

Na revisão dos relatórios, a saída Markdown passou a incluir contagens observadas de partes fora do escopo e avisos do leitor. CSV é documentado como tabela de alterações; JSON/Markdown/bundle preservam veredito e cobertura. A verificação instalada confirmou o exemplo de onze mudanças e preservação de um destino já existente.

## Consumidores reais e reexecução

O [ensaio principal](INTEROPERABILIDADE.md) usa FODS originais sem caches preenchidos. O LibreOffice 26.2.6.3 gerou os XLSX e calculou os resultados conhecidos. O leitor/comparador confirmou onze alterações; uma fixture derivada sem `<v>` gerou uma mudança de cache ausente. A [prova de texto OOXML](TEXTOS-OOXML.md) usa XLSX sintético e a exportação CSV do programa externo como oráculo.

Os checkers `--check-fixture` verificam artefatos registrados e o código atual; não iniciam o consumidor externo. Os comandos completos com `--output` iniciam o LibreOffice em perfil temporário exclusivo. Os ensaios locais terminaram e removeram seus próprios perfis, sem usar contas ou arquivos pessoais.

Não foi executado Microsoft Excel. Aparência, cálculo de fórmulas e recursos excluídos no contrato não fazem parte destes resultados.
