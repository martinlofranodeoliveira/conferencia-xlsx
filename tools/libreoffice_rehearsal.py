#!/usr/bin/env python3
"""Original synthetic FODS → real LibreOffice XLSX → independent/core checks.

The ordinary fixture check does not execute LibreOffice. An explicit --output
starts the external consumer with a newly-created private temporary profile.
"""
from __future__ import annotations

import argparse
from datetime import date, datetime, timezone
from decimal import Decimal
import hashlib
import json
import os
from pathlib import Path
import posixpath
import signal
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "fixtures" / "libreoffice"
SAMPLES = ROOT / "src" / "conferencia_xlsx" / "samples"
NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    "of": "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
    "dc": "http://purl.org/dc/elements/1.1/",
    "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
}
for _prefix, _uri in NS.items():
    ET.register_namespace(_prefix, _uri)
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
OD_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def q(prefix: str, name: str) -> str:
    return "{" + NS[prefix] + "}" + name


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def paragraph(parent: ET.Element, text: str) -> None:
    p = ET.SubElement(parent, q("text", "p"))
    lines = text.split("\n")
    p.text = lines[0]
    for line in lines[1:]:
        br = ET.SubElement(p, q("text", "line-break"))
        br.tail = line


def literal(row: ET.Element, kind: str, value: str, **attrs: str) -> None:
    cell = ET.SubElement(row, q("table", "table-cell"), attrs)
    cell.set(q("office", "value-type"), kind)
    if kind == "float":
        cell.set(q("office", "value"), value)
    elif kind == "boolean":
        cell.set(q("office", "boolean-value"), value)
    elif kind == "date":
        cell.set(q("office", "date-value"), value)
    paragraph(cell, value)


def formula(row: ET.Element, expression: str) -> None:
    # No office:value or cached text is supplied. LibreOffice calculates it.
    cell = ET.SubElement(row, q("table", "table-cell"))
    cell.set(q("table", "formula"), "of:=" + expression)
    ET.SubElement(cell, q("text", "p"))


def make_fods(after: bool) -> bytes:
    doc = ET.Element(q("office", "document"), {
        q("office", "version"): "1.3",
        q("office", "mimetype"): "application/vnd.oasis.opendocument.spreadsheet",
        "xmlns:of": NS["of"],
    })
    meta = ET.SubElement(doc, q("office", "meta"))
    ET.SubElement(meta, q("meta", "initial-creator")).text = "Equipe Exemplo"
    ET.SubElement(meta, q("dc", "creator")).text = "Equipe Exemplo"
    ET.SubElement(meta, q("dc", "title")).text = "Conferência sintética de planilhas"
    styles = ET.SubElement(doc, q("office", "automatic-styles"))
    ds = ET.SubElement(styles, q("number", "date-style"), {q("style", "name"): "date1"})
    for tag, text in [("year", None), ("text", "-"), ("month", None), ("text", "-"), ("day", None)]:
        item = ET.SubElement(ds, q("number", tag))
        if text is None:
            item.set(q("number", "style"), "long")
        else:
            item.text = text
    ET.SubElement(styles, q("style", "style"), {
        q("style", "name"): "date-cell", q("style", "family"): "table-cell",
        q("style", "data-style-name"): "date1",
    })
    bs = ET.SubElement(styles, q("number", "boolean-style"), {q("style", "name"): "boolean1"})
    ET.SubElement(bs, q("number", "boolean"))
    ET.SubElement(styles, q("style", "style"), {
        q("style", "name"): "boolean-cell", q("style", "family"): "table-cell",
        q("style", "data-style-name"): "boolean1",
    })
    hidden = ET.SubElement(styles, q("style", "style"), {
        q("style", "name"): "hidden-table", q("style", "family"): "table",
    })
    ET.SubElement(hidden, q("style", "table-properties"), {q("table", "display"): "false"})
    body = ET.SubElement(doc, q("office", "body"))
    spreadsheet = ET.SubElement(body, q("office", "spreadsheet"))

    def auxiliary(hidden_state: bool) -> None:
        table = ET.SubElement(spreadsheet, q("table", "table"), {q("table", "name"): "Auxiliar"})
        if hidden_state:
            table.set(q("table", "style-name"), "hidden-table")
        row = ET.SubElement(table, q("table", "table-row"))
        literal(row, "string", "Apoio fictício")

    if after:
        auxiliary(True)
    table = ET.SubElement(spreadsheet, q("table", "table"), {q("table", "name"): "Inventário"})
    row = ET.SubElement(table, q("table", "table-row"))
    literal(row, "float", "12" if after else "10")  # A1
    literal(row, "float", "5")                      # B1
    formula(row, "SUM([.A1:.B1])")                   # C1
    formula(row, "2*1" if after else "1+1")         # D1
    literal(row, "string", "1")                    # E1
    literal(row, "string" if after else "float", "1")  # F1
    literal(row, "boolean", "false" if after else "true", **{q("table", "style-name"): "boolean-cell"})  # G1
    literal(row, "date", "2026-11-08", **{q("table", "style-name"): "date-cell"})  # H1
    formula(row, "1/0")                             # I1: error cache
    formula(row, '""')                             # J1: empty <v>
    ET.SubElement(row, q("table", "table-cell"))    # K1: real blank
    literal(row, "string", "=SUM(1,2)")            # L1: formula-like literal
    literal(row, "string", "Linha 1 & revisão\nLinha 2 | <valor>")  # M1
    row2 = ET.SubElement(table, q("table", "table-row"))
    span = 3 if after else 2
    literal(row2, "string", "Resumo fictício", **{q("table", "number-columns-spanned"): str(span)})
    for _ in range(span - 1):
        ET.SubElement(row2, q("table", "covered-table-cell"))
    name = "Nova" if after else "Legado"
    another = ET.SubElement(spreadsheet, q("table", "table"), {q("table", "name"): name})
    another_row = ET.SubElement(another, q("table", "table-row"))
    literal(another_row, "string", "Registro fictício")
    if not after:
        auxiliary(False)
    return ET.tostring(doc, encoding="utf-8", xml_declaration=True) + b"\n"


def xml_oracle(data: bytes) -> dict:
    """Small independent oracle for the trusted original consumer fixtures.

    This intentionally does not import the application's reader or openpyxl.
    It is not advertised as a general-purpose XLSX parser.
    """
    from io import BytesIO
    with ZipFile(BytesIO(data)) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: posixpath.normpath(posixpath.join("xl", r.attrib["Target"]))
                   for r in rels if r.attrib.get("TargetMode") != "External"}
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            strings = ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared = ["".join(t.text or "" for t in si.iter("{" + MAIN + "}t")) for si in strings]
        order, sheets = [], {}
        props = wb.find("{" + MAIN + "}workbookPr")
        epoch = "1904" if props is not None and props.attrib.get("date1904") in ("1", "true") else "1900"
        for s in wb.find("{" + MAIN + "}sheets"):
            name = s.attrib["name"]
            order.append(name)
            root = ET.fromstring(z.read(targets[s.attrib["{" + OD_REL + "}id"]]))
            cells = {}
            for c in root.iter("{" + MAIN + "}c"):
                f, v = c.find("{" + MAIN + "}f"), c.find("{" + MAIN + "}v")
                kind = c.attrib.get("t", "n")
                if f is not None:
                    cells[c.attrib["r"]] = {
                        "formula": f.text or "", "cache_state": "absent" if v is None else "empty" if v.text is None else "value",
                        "type": kind, "value": None if v is None else v.text,
                    }
                elif kind == "s" and v is not None and v.text is not None:
                    cells[c.attrib["r"]] = {"type": "string", "value": shared[int(v.text)]}
                elif kind == "inlineStr":
                    cells[c.attrib["r"]] = {"type": "string", "value": "".join(t.text or "" for t in c.iter("{" + MAIN + "}t"))}
                elif v is not None and v.text is not None:
                    cells[c.attrib["r"]] = {"type": kind, "value": v.text}
            merge_parent = root.find("{" + MAIN + "}mergeCells")
            merges = [] if merge_parent is None else sorted(m.attrib["ref"] for m in merge_parent)
            sheets[name] = {"state": s.attrib.get("state", "visible"), "cells": cells, "merges": merges}
    return {"epoch": epoch, "sheet_order": order, "sheets": sheets}


def check_known_values(before: dict, after: dict) -> None:
    assert before["epoch"] == after["epoch"] == "1900"
    assert before["sheet_order"] == ["Inventário", "Legado", "Auxiliar"]
    assert after["sheet_order"] == ["Auxiliar", "Inventário", "Nova"]
    assert before["sheets"]["Auxiliar"]["state"] == "visible"
    assert after["sheets"]["Auxiliar"]["state"] == "hidden"
    b, a = before["sheets"]["Inventário"], after["sheets"]["Inventário"]
    assert b["merges"] == ["A2:B2"] and a["merges"] == ["A2:C2"]
    bc, ac = b["cells"], a["cells"]
    assert bc["A1"] == {"type": "n", "value": "10"}
    assert ac["A1"] == {"type": "n", "value": "12"}
    assert bc["C1"]["formula"] == ac["C1"]["formula"] == "SUM(A1:B1)"
    assert bc["C1"]["value"] == "15" and ac["C1"]["value"] == "17"
    assert bc["D1"]["formula"] == "1+1" and ac["D1"]["formula"] == "2*1"
    assert bc["D1"]["value"] == ac["D1"]["value"] == "2"
    assert bc["E1"] == ac["E1"] == {"type": "string", "value": "1"}
    assert bc["F1"] == {"type": "n", "value": "1"}
    assert ac["F1"] == {"type": "string", "value": "1"}
    assert bc["G1"] == {"type": "b", "value": "1"}
    assert ac["G1"] == {"type": "b", "value": "0"}
    serial = str((date(2026, 11, 8) - date(1899, 12, 30)).days)
    assert bc["H1"] == ac["H1"] == {"type": "n", "value": serial}
    assert bc["I1"]["type"] == ac["I1"]["type"] == "e"
    assert bc["I1"]["value"] == ac["I1"]["value"] == "#DIV/0!"
    assert bc["J1"]["formula"] == ac["J1"]["formula"] == '""'
    assert bc["J1"]["cache_state"] == ac["J1"]["cache_state"] == "empty"
    assert "K1" not in bc and "K1" not in ac
    assert bc["L1"] == ac["L1"] == {"type": "string", "value": "=SUM(1,2)"}
    assert bc["M1"] == ac["M1"] == {"type": "string", "value": "Linha 1 & revisão\nLinha 2 | <valor>"}


def without_cache(data: bytes) -> bytes:
    from io import BytesIO
    target = BytesIO()
    with ZipFile(BytesIO(data)) as original, ZipFile(target, "w") as changed:
        for info in original.infolist():
            payload = original.read(info)
            if info.filename == "xl/worksheets/sheet1.xml":
                root = ET.fromstring(payload)
                cell = root.find(".//{" + MAIN + "}c[@r='C1']")
                assert cell is not None
                value = cell.find("{" + MAIN + "}v")
                assert value is not None
                cell.remove(value)
                payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            changed.writestr(info, payload)
    return target.getvalue()


def check_core(before_data: bytes, after_data: bytes, missing_data: bytes) -> dict:
    sys.path.insert(0, str(ROOT / "src"))
    from conferencia_xlsx.reader import read_bytes
    from conferencia_xlsx.compare import compare_snapshots
    before, after, missing = [read_bytes(data) for data in (before_data, after_data, missing_data)]
    inv_before = next(s for s in before.sheets if s.name == "Inventário")
    inv_after = next(s for s in after.sheets if s.name == "Inventário")
    b, a = [{c.coordinate: c for c in s.cells} for s in (inv_before, inv_after)]
    assert b["C1"].formula == a["C1"].formula == "=SUM(A1:B1)"
    assert b["C1"].cache.state == a["C1"].cache.state == "value"
    assert Decimal(b["C1"].cache.value.value) == 15 and Decimal(a["C1"].cache.value.value) == 17
    assert b["D1"].formula == "=1+1" and a["D1"].formula == "=2*1"
    assert Decimal(b["D1"].cache.value.value) == Decimal(a["D1"].cache.value.value) == 2
    assert b["E1"].literal.kind == "string" and b["F1"].literal.kind == "number"
    assert a["F1"].literal.kind == "string"
    assert b["G1"].literal.value is True and a["G1"].literal.value is False
    assert b["H1"].literal.kind == "number"
    assert b["I1"].cache.value.kind == "error" and b["I1"].cache.value.value == "#DIV/0!"
    assert b["J1"].cache.state == "empty" and b["J1"].cache.value is None
    assert "K1" not in b
    assert b["L1"].literal.kind == "string" and b["L1"].literal.value == "=SUM(1,2)"
    report = compare_snapshots(before, after)
    assert report["status"] == "different_in_scope" and report["identical_bytes"] is False
    actual = {(x["kind"], x["sheet"], x["coordinate"]) for x in report["changes"]}
    expected = {
        ("sheet_added", "Nova", None), ("sheet_removed", "Legado", None),
        ("sheet_order", None, None), ("sheet_state", "Auxiliar", None),
        ("merge_added", "Inventário", "A2:C2"), ("merge_removed", "Inventário", "A2:B2"),
        ("literal", "Inventário", "A1"), ("cache", "Inventário", "C1"),
        ("formula", "Inventário", "D1"), ("literal", "Inventário", "F1"),
        ("literal", "Inventário", "G1"),
    }
    assert actual == expected, (actual, expected)
    assert len(report["changes"]) == len(expected) == 11
    missing_report = compare_snapshots(before, missing)
    assert len(missing_report["changes"]) == 1
    change = missing_report["changes"][0]
    assert (change["kind"], change["sheet"], change["coordinate"]) == ("cache", "Inventário", "C1")
    assert next(c for s in missing.sheets if s.name == "Inventário" for c in s.cells if c.coordinate == "C1").cache.state == "absent"
    assert report["scope"]["cache_freshness"] == "unknown"
    return {"status": "passed", "before_sheets": len(before.sheets), "after_sheets": len(after.sheets),
            "expected_changes": 11, "missing_cache_changes": 1, "source_hashes_match": before.sha256 == sha(before_data) and after.sha256 == sha(after_data),
            "cache_freshness": "unknown"}


def run_owned(args: list[str], timeout: float = 60) -> subprocess.CompletedProcess:
    process = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True)
    try:
        stdout, stderr = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        os.killpg(process.pid, signal.SIGTERM)
        try:
            process.communicate(timeout=5)
        except subprocess.TimeoutExpired:
            os.killpg(process.pid, signal.SIGKILL)
            process.communicate()
        raise RuntimeError("O consumidor externo excedeu o prazo do ensaio.") from None
    if process.returncode != 0:
        raise RuntimeError("O consumidor externo falhou no ensaio sintético.")
    return subprocess.CompletedProcess(args, process.returncode, stdout, stderr)


def run_consumer(output: Path, executable: Path) -> dict:
    if os.name != "posix":
        raise RuntimeError("A reprodução externa deste ensaio foi implementada para POSIX.")
    output.mkdir(parents=False, exist_ok=False)
    version = run_owned([str(executable), "--version"]).stdout.decode("utf-8").strip()
    source_hashes = {}
    for name, after in [("antes", False), ("depois", True)]:
        data = make_fods(after)
        (output / f"{name}.fods").write_bytes(data)
        source_hashes[name] = sha(data)
    with tempfile.TemporaryDirectory(prefix="conferencia-xlsx-lo-") as temporary:
        profile = Path(temporary) / "profile"
        # TemporaryDirectory is exclusive; LibreOffice creates this profile.
        args = [str(executable), "-env:UserInstallation=" + profile.as_uri(), "--headless",
                "--convert-to", "xlsx", "--outdir", str(output.resolve()),
                str((output / "antes.fods").resolve()), str((output / "depois.fods").resolve())]
        converted = run_owned(args)
    before_data, after_data = [(output / f"{name}.xlsx").read_bytes() for name in ("antes", "depois")]
    before_oracle, after_oracle = [xml_oracle(data) for data in (before_data, after_data)]
    check_known_values(before_oracle, after_oracle)
    missing_data = without_cache(before_data)
    (output / "sem-cache.xlsx").write_bytes(missing_data)
    input_hashes = [sha(before_data), sha(after_data)]
    core = check_core(before_data, after_data, missing_data)
    from conferencia_xlsx.reader import read_path, read_bytes
    assert read_path(output / "antes.xlsx") == read_bytes(before_data)
    assert read_path(output / "depois.xlsx") == read_bytes(after_data)
    core["read_path_matches_captured_bytes"] = True
    assert input_hashes == [sha((output / f"{name}.xlsx").read_bytes()) for name in ("antes", "depois")]
    report = {
        "schema_version": 1, "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "external_consumer_rerun": True, "consumer": {"name": "LibreOffice", "version": version,
            "arguments": ["-env:UserInstallation=<exclusive temporary profile>", "--headless", "--convert-to", "xlsx", "--outdir", "<owned output>", "antes.fods", "depois.fods"],
            "profile": "exclusive temporary profile; removed after the process exits", "returncode": converted.returncode},
        "authorship": "FODS e verificadores originais; dados inteiramente fictícios; fórmulas FODS sem valores em cache fornecidos.",
        "source_fods_sha256": source_hashes,
        "xlsx_sha256": dict(zip(("antes", "depois"), input_hashes)),
        "oracle": {"before": before_oracle, "after": after_oracle}, "core": core,
        "inputs_unchanged_after_read": True,
        "limitations": ["Ensaio sintético com LibreOffice, não Microsoft Excel.", "Formatos numéricos e aparência estão fora do escopo.", "Atualidade geral dos caches permanece desconhecida.", "Não usa conta, perfil existente ou arquivos pessoais; não instala nem consulta serviços."],
    }
    (output / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def check_fixture() -> dict:
    manifest = json.loads((FIXTURES / "manifest.json").read_text(encoding="utf-8"))
    for name, entry in manifest["files"].items():
        data = (ROOT / name).read_bytes()
        assert len(data) == entry["bytes"] and sha(data) == entry["sha256"], name
    before_data, after_data = [(SAMPLES / f"{name}.xlsx").read_bytes() for name in ("antes", "depois")]
    missing_data = (FIXTURES / "sem-cache.xlsx").read_bytes()
    recorded = json.loads((FIXTURES / "report.json").read_text(encoding="utf-8"))
    before_oracle, after_oracle = [xml_oracle(data) for data in (before_data, after_data)]
    assert recorded["oracle"] == {"before": before_oracle, "after": after_oracle}
    check_known_values(before_oracle, after_oracle)
    result = check_core(before_data, after_data, missing_data)
    return {**result, "recorded_hashes_verified": True, "external_consumer_rerun": False}


def main() -> int:
    parser = argparse.ArgumentParser(description="Ensaio original e isolado de interoperabilidade XLSX.")
    parser.add_argument("--check-fixture", action="store_true", help="Confere registros sem executar LibreOffice.")
    parser.add_argument("--output", type=Path, help="Novo diretório para uma execução real do consumidor.")
    parser.add_argument("--libreoffice", type=Path, default=Path("/usr/bin/libreoffice"))
    args = parser.parse_args()
    if args.check_fixture == bool(args.output):
        parser.error("Escolha --check-fixture ou --output, uma opção apenas.")
    try:
        result = check_fixture() if args.check_fixture else run_consumer(args.output, args.libreoffice)
    except Exception:
        print("O ensaio não foi concluído. Confira dependências, destino novo e contrato dos fixtures.", file=sys.stderr)
        return 2
    print(json.dumps(result if args.check_fixture else {"status": "passed", "external_consumer_rerun": True, "core": result["core"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
