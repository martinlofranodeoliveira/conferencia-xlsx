#!/usr/bin/env python3
"""Narrow real-consumer oracle for SpreadsheetML escaped string decoding."""
from __future__ import annotations

import argparse
import csv
from datetime import datetime, timezone
import hashlib
import importlib.util
from io import BytesIO, StringIO
import json
from pathlib import Path
import sys
import tempfile
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "fixtures/xstring"
SPEC = importlib.util.spec_from_file_location("owned_lo_rehearsal", ROOT / "tools/libreoffice_rehearsal.py")
OWNED = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(OWNED)
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CASES = (
    ("x005F_", "x005F_"), ("_x005F_", "_"), ("_x0041_", "A"),
    ("_x005F_x0041_", "_x0041_"), ("_xD83D__xDE00_", "😀"),
    ("_x000A_", "\n"), ("_x005F_x000A_", "_x000A_"),
    ("prefix_x005F_suffix", "prefix_suffix"), ("_x005f_x0041_", "_x0041_"),
    ("_x005F__x0041_", "_A"), ("_x000D_", "\r"), ("before_x0041_after", "beforeAafter"),
)
RICH_CASES = (
    (("_x004", "1_"), "_x0041_"),
    (("_x000", "D_"), "_x000D_"),
)
ALL_CASES = tuple(((raw,), expected) for raw, expected in CASES) + RICH_CASES


def make_workbook() -> bytes:
    """Original minimal XLSX: column A shared, column B inline, same raw text."""
    parts = {
        "[Content_Types].xml": b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
        "_rels/.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="Probe" sheetId="1" r:id="rId1"/></sheets></workbook>'.encode(),
        "xl/_rels/workbook.xml.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
    }
    strings, rows = [], []
    for index, (runs, _expected) in enumerate(ALL_CASES):
        content = f'<t>{escape(runs[0])}</t>' if len(runs) == 1 else ''.join(f'<r><t>{escape(run)}</t></r>' for run in runs)
        strings.append(f'<si>{content}</si>')
        rows.append(f'<row r="{index + 1}"><c r="A{index + 1}" t="s"><v>{index}</v></c>'
                    f'<c r="B{index + 1}" t="inlineStr"><is>{content}</is></c></row>')
    parts["xl/sharedStrings.xml"] = (f'<sst xmlns="{MAIN}" count="{len(ALL_CASES)}" uniqueCount="{len(ALL_CASES)}">'
        + ''.join(strings) + '</sst>').encode('utf-8')
    parts["xl/worksheets/sheet1.xml"] = (f'<worksheet xmlns="{MAIN}"><sheetData>'
        + ''.join(rows) + '</sheetData></worksheet>').encode('utf-8')
    output = BytesIO()
    with ZipFile(output, "w") as archive:
        for name, content in parts.items():
            info = ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, content)
    return output.getvalue()


def check_consumer_rows(csv_data: bytes) -> list[list[str]]:
    rows = list(csv.reader(StringIO(csv_data.decode("utf-8-sig"), newline="")))
    assert rows == [[expected, expected] for _runs, expected in ALL_CASES]
    return rows


def check_app(data: bytes) -> None:
    sys.path.insert(0, str(ROOT / "src"))
    from conferencia_xlsx.reader import read_bytes
    snapshot = read_bytes(data)
    assert len(snapshot.sheets) == 1 and snapshot.sheets[0].name == "Probe"
    cells = {c.coordinate: c for c in snapshot.sheets[0].cells}
    assert len(cells) == 2 * len(ALL_CASES)
    for index, (_runs, expected) in enumerate(ALL_CASES, 1):
        for column in ("A", "B"):
            cell = cells[f"{column}{index}"]
            assert cell.literal.kind == "string" and cell.literal.value == expected
            assert cell.formula is None and cell.cache is None


def run_consumer(output: Path, executable: Path) -> dict:
    output.mkdir(parents=False, exist_ok=False)
    original = make_workbook()
    (output / "probe.xlsx").write_bytes(original)
    version = OWNED.run_owned([str(executable), "--version"]).stdout.decode("utf-8").strip()
    with tempfile.TemporaryDirectory(prefix="conferencia-xstring-lo-") as temporary:
        args = [str(executable), "-env:UserInstallation=" + Path(temporary).as_uri(), "--headless",
                "--convert-to", "csv:Text - txt - csv (StarCalc):44,34,76,1",
                "--outdir", str(output.resolve()), str((output / "probe.xlsx").resolve())]
        converted = OWNED.run_owned(args)
    csv_data = (output / "probe.csv").read_bytes()
    rows = check_consumer_rows(csv_data)
    check_app(original)
    assert (output / "probe.xlsx").read_bytes() == original
    report = {
        "schema_version": 1, "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "external_consumer_rerun": True, "consumer": {"name": "LibreOffice", "version": version,
            "arguments": ["-env:UserInstallation=<exclusive temporary profile>", "--headless", "--convert-to",
                "csv:Text - txt - csv (StarCalc):44,34,76,1", "--outdir", "<owned output>", "probe.xlsx"],
            "returncode": converted.returncode, "profile": "exclusive temporary profile; removed after process exit"},
        "authorship": "XLSX mínimo original com texto em SST e inline; dados inteiramente sintéticos.",
        "cases": [{"xml_runs": list(runs), "shared_string_value": row[0], "inline_string_value": row[1],
                   "codepoints": [f"U+{ord(c):04X}" for c in row[0]]} for (runs, _), row in zip(ALL_CASES, rows)],
        "app_matches_consumer": True, "input_unchanged": True,
        "primary_sources": [
            "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bd0aa042-434a-4ca7-b25f-4e1fd25a954d",
            "https://help.libreoffice.org/latest/en-US/text/shared/guide/csv_params.html"],
        "limitations": ["Consumidor exercitado: LibreOffice, não Microsoft Excel.",
            "Casos de texto específicos; não certifica todo o formato XLSX."],
    }
    (output / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def check_fixture() -> dict:
    manifest = json.loads((FIXTURES / "manifest.json").read_text(encoding="utf-8"))
    for name, record in manifest["files"].items():
        data = (FIXTURES / name).read_bytes()
        assert len(data) == record["bytes"] and hashlib.sha256(data).hexdigest() == record["sha256"]
    data = (FIXTURES / "probe.xlsx").read_bytes()
    # Compare member bytes, not platform/compressor-dependent ZIP headers.
    with ZipFile(BytesIO(data)) as recorded, ZipFile(BytesIO(make_workbook())) as generated:
        assert recorded.namelist() == generated.namelist()
        for name in recorded.namelist():
            assert recorded.read(name) == generated.read(name)
    rows = check_consumer_rows((FIXTURES / "probe.csv").read_bytes())
    report = json.loads((FIXTURES / "report.json").read_text(encoding="utf-8"))
    assert len(report["cases"]) == len(ALL_CASES)
    for (runs, expected), row, recorded in zip(ALL_CASES, rows, report["cases"]):
        assert recorded["xml_runs"] == list(runs)
        assert recorded["shared_string_value"] == recorded["inline_string_value"] == expected == row[0] == row[1]
    check_app(data)
    return {"status": "passed", "cases": len(ALL_CASES), "cells": 2 * len(ALL_CASES),
            "app_matches_recorded_consumer": True, "recorded_hashes_verified": True, "external_consumer_rerun": False}


def main() -> int:
    parser = argparse.ArgumentParser(description="Ensaio estreito de textos escapados no SpreadsheetML.")
    parser.add_argument("--check-fixture", action="store_true")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--libreoffice", type=Path, default=Path("/usr/bin/libreoffice"))
    args = parser.parse_args()
    if args.check_fixture == bool(args.output):
        parser.error("Escolha --check-fixture ou --output.")
    try:
        result = check_fixture() if args.check_fixture else run_consumer(args.output, args.libreoffice)
    except Exception:
        print("O ensaio de texto não foi concluído. Confira dependências, destino novo e casos registrados.", file=sys.stderr)
        return 2
    print(json.dumps(result if args.check_fixture else {"status": "passed", "cases": len(result["cases"]), "external_consumer_rerun": True}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
