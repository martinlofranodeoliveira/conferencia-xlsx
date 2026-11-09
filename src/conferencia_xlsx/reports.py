"""Relatórios locais: os dados aparecem somente nos arquivos pedidos."""
from __future__ import annotations

import csv
import hashlib
import html
import io
import json
import re
import zipfile

from .errors import AuditError

MAX_REPORT_BYTES = 64 * 1024 * 1024
KINDS = {
    "epoch": "Sistema de datas", "sheet_added": "Aba adicionada", "sheet_removed": "Aba removida",
    "sheet_order": "Ordem das abas", "sheet_state": "Visibilidade", "merge_added": "Mesclagem adicionada",
    "merge_removed": "Mesclagem removida", "cell_added": "Célula adicionada", "cell_removed": "Célula removida",
    "literal": "Valor ou tipo", "formula": "Fórmula", "cache": "Resultado armazenado",
}
SCOPE_LABELS = {
    "date_system": "sistema de datas", "worksheet_names_order_visibility": "nomes, ordem e visibilidade de abas",
    "merged_ranges": "mesclagens", "typed_stored_literals": "literais armazenados e seus tipos",
    "expanded_formula_text": "texto de fórmula expandida", "stored_formula_cache": "resultado armazenado de fórmula",
    "appearance_styles_and_number_formats": "aparência, estilos e formatos numéricos", "charts_and_images": "gráficos e imagens",
    "structured_tables": "tabelas estruturadas", "data_validations": "validações", "comments": "comentários",
    "hyperlinks": "hyperlinks", "defined_names": "nomes definidos", "digital_signatures": "assinaturas digitais",
    "external_link_targets": "conteúdo de vínculos externos", "mathematical_formula_equivalence": "equivalência matemática",
    "formula_recalculation": "recálculo de fórmulas",
}


def _bounded(data: bytes) -> bytes:
    if len(data) > MAX_REPORT_BYTES:
        raise AuditError("report-limit", "O relatório ultrapassa o limite de 64 MiB.")
    return data


def json_bytes(value: object) -> bytes:
    return _bounded((json.dumps(value, ensure_ascii=False, allow_nan=False, indent=2) + "\n").encode("utf-8"))


def markdown_text(value: object) -> str:
    text = html.escape(str(value), quote=True)
    text = re.sub(r"([\\`*_{}\[\]()#+.!|~])", r"\\\1", text)
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")


def _display(value: object) -> str:
    if value is None:
        return "—"
    return json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(",", ":"))


def markdown_bytes(report: dict) -> bytes:
    verdict = "Diferente no escopo conferido" if report["status"] == "different_in_scope" else "Igual no escopo conferido"
    lines = ["# Conferência XLSX", "", f"**{verdict}.** {report['summary']['changes']} alterações.", "",
             f"Arquivos idênticos em bytes: {'sim' if report['identical_bytes'] else 'não'}.", "",
             "A comparação usa nome da aba e coordenada. Inserir ou mover linhas não acompanha a identidade dos registros.", "",
             "## Entradas capturadas", "", "| Entrada | SHA-256 | Bytes | Sistema de datas |", "| --- | --- | ---: | --- |"]
    for key, label in [("before", "Antes"), ("after", "Depois")]:
        source = report[key]
        lines.append(f"| {label} | {markdown_text(source['sha256'])} | {source['size_bytes']} | {markdown_text(source['epoch'])} |")
    lines += ["", "## Alterações", "", "| Categoria | Aba | Célula | Antes | Depois |", "| --- | --- | --- | --- | --- |"]
    for change in report["changes"]:
        values = [KINDS.get(change["kind"], change["kind"]), change["sheet"] or "—", change["coordinate"] or "—",
                  _display(change["before"]), _display(change["after"])]
        lines.append("| " + " | ".join(markdown_text(v) for v in values) + " |")
    if not report["changes"]:
        lines.append("| Nenhuma alteração no escopo | — | — | — | — |")
    lines += ["", "## Como interpretar", "",
              "- Números são comparados como decimais armazenados exatos, sem arredondamento de ponto flutuante.",
              "- Fórmulas são texto; fórmulas compartilhadas são expandidas pelo leitor. Equivalência matemática não é avaliada.",
              "- Resultados armazenados têm atualidade **desconhecida**. Cache ausente, elemento vazio e valor presente são estados distintos.",
              "- Datas seriais continuam números; células ISO são comparadas pela representação armazenada. O sistema 1900/1904 é conferido separadamente.", "",
              "## Cobertura", "", "Comparados: " + markdown_text(", ".join(SCOPE_LABELS.get(k, k) for k in report["scope"]["included"])) + ".", "",
              "Fora de escopo: " + markdown_text(", ".join(SCOPE_LABELS.get(k, k) for k in report["scope"]["excluded"])) + ".", "",
              "| Observação do leitor | Antes | Depois |", "| --- | ---: | ---: |"]
    categories = sorted(set(report["before"]["ignored_parts"]) | set(report["after"]["ignored_parts"]))
    for category in categories:
        lines.append(f"| {markdown_text(category)} | {report['before']['ignored_parts'].get(category, 0)} | {report['after']['ignored_parts'].get(category, 0)} |")
    lines.append(f"| Avisos do leitor \u2014 somente contagem | {report['before']['warning_count']} | {report['after']['warning_count']} |")
    lines += ["", "As contagens inventariam recursos fora do escopo; não conferem seu conteúdo.", "",
              "O resultado não comprova igualdade visual, recálculo, integridade de negócio ou autenticidade. As entradas não foram salvas pelo aplicativo.", ""]
    return _bounded("\n".join(lines).encode("utf-8"))


def protect_csv(text: str) -> str:
    return "'" + text if re.sub(r"^[\s\ufeff]*", "", text).startswith(("=", "+", "-", "@")) else text


def csv_bytes(report: dict) -> bytes:
    output = io.StringIO(newline="")
    writer = csv.writer(output, lineterminator="\r\n")
    writer.writerow(["Categoria", "Aba", "Célula", "Antes", "Depois"])
    for change in report["changes"]:
        writer.writerow([protect_csv(str(v)) for v in [change["kind"], change["sheet"] or "", change["coordinate"] or "",
                                                       _display(change["before"]), _display(change["after"])]])
    return _bounded(("\ufeff" + output.getvalue()).encode("utf-8"))


def bundle_bytes(report: dict, *, samples: dict[str, bytes] | None = None) -> bytes:
    members = {"comparacao.json": json_bytes(report), "comparacao.md": markdown_bytes(report), "alteracoes.csv": csv_bytes(report)}
    if samples:
        if set(samples) != {"antes.xlsx", "depois.xlsx"}:
            raise AuditError("invalid-samples", "As amostras não são reconhecidas.")
        members.update(samples)
    if sum(map(len, members.values())) > MAX_REPORT_BYTES:
        raise AuditError("report-limit", "Os relatórios ultrapassam o limite combinado de 64 MiB.")
    manifest = {"schema_version": 1, "application": "conferencia-xlsx", "version": "1.0.0",
                "members": {name: {"sha256": hashlib.sha256(data).hexdigest(), "size_bytes": len(data)} for name, data in members.items()}}
    members["manifesto.json"] = json_bytes(manifest)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for name, data in members.items():
            item = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            item.compress_type = zipfile.ZIP_DEFLATED
            item.external_attr = 0o100600 << 16
            item.create_system = 3
            archive.writestr(item, data)
    return _bounded(output.getvalue())
