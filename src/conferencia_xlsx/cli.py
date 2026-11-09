"""Interface de linha de comando em português."""
from __future__ import annotations

import argparse
from importlib.resources import files
import sys

from .compare import compare_snapshots
from .errors import AuditError
from .publication import publish_bytes
from .reader import read_bytes, read_path
from .reports import bundle_bytes, csv_bytes, json_bytes, markdown_bytes


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="conferencia-xlsx", description="Confira conteúdo e estrutura selecionada de duas planilhas XLSX, sem modificá-las.")
    parser.add_argument("--version", action="version", version="conferencia-xlsx 1.0.0")
    commands = parser.add_subparsers(dest="command", required=True)
    compare = commands.add_parser("comparar", help="Comparar snapshots locais por aba e coordenada.")
    compare.add_argument("antes")
    compare.add_argument("depois")
    compare.add_argument("--saida", required=True, help="Arquivo novo; o destino existente não será sobrescrito.")
    compare.add_argument("--formato", choices=["bundle", "json", "markdown", "csv"], default="bundle")
    demo = commands.add_parser("exemplo", help="Criar um ZIP com entradas fictícias e relatórios de demonstração.")
    demo.add_argument("--saida", required=True)
    args = parser.parse_args(argv)
    try:
        samples = None
        if args.command == "exemplo":
            samples = {name: files("conferencia_xlsx").joinpath("samples", name).read_bytes() for name in ("antes.xlsx", "depois.xlsx")}
            before, after = read_bytes(samples["antes.xlsx"]), read_bytes(samples["depois.xlsx"])
        else:
            before, after = read_path(args.antes), read_path(args.depois)
        report = compare_snapshots(before, after)
        if args.command == "exemplo":
            data = bundle_bytes(report, samples=samples)
        else:
            renderer = {"json": json_bytes, "markdown": markdown_bytes, "csv": csv_bytes, "bundle": bundle_bytes}[args.formato]
            data = renderer(report)
        publish_bytes(args.saida, data)
        if args.command == "exemplo":
            print("Exemplo fictício publicado: duas entradas XLSX, relatórios e manifesto.")
            return 0
        if report["status"] == "different_in_scope":
            print(f"Comparação completa: {report['summary']['changes']} alterações no escopo conferido.")
            return 1
        print("Conteúdo e estrutura selecionada iguais. Consulte a cobertura do relatório.")
        return 0
    except AuditError as error:
        print(f"Erro [{error.code}]: {error.message}", file=sys.stderr)
        return 2
    except (OSError, ValueError, TypeError, KeyError, OverflowError):
        print("Erro [operation-failed]: Não foi possível concluir. Confira os arquivos e os limites documentados.", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("Operação interrompida.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
