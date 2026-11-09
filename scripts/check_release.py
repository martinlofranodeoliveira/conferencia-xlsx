"""Confere a distribuição produzida, os recursos e o RECORD do wheel."""
import base64
import csv
import hashlib
import io
from pathlib import Path, PurePosixPath
import tarfile
import tomllib
import zipfile


ROOT = Path(__file__).resolve().parents[1]
VERSION = tomllib.loads((ROOT / 'pyproject.toml').read_text(encoding='utf-8'))['project']['version']
NAME = 'conferencia_xlsx'


def main():
    package = ROOT / 'src' / NAME
    expected = {p.relative_to(ROOT / 'src').as_posix(): p.read_bytes() for p in package.rglob('*')
                if p.is_file() and p.suffix in ('.py', '.xlsx') and '__pycache__' not in p.parts}
    wheels = list((ROOT / 'dist').glob('*.whl'))
    sdists = list((ROOT / 'dist').glob('*.tar.gz'))
    assert len(wheels) == len(sdists) == 1, 'Esperado um wheel e um sdist.'
    with zipfile.ZipFile(wheels[0]) as archive:
        names = archive.namelist()
        assert len(names) == len(set(names)) and archive.testzip() is None
        assert set(n for n in names if n.startswith(NAME + '/')) == set(expected)
        for name, data in expected.items():
            assert archive.read(name) == data, name
        info = f'{NAME}-{VERSION}.dist-info/'
        allowed = {'METADATA', 'WHEEL', 'entry_points.txt', 'RECORD', 'licenses/LICENSE'}
        assert all(n in expected or (n.startswith(info) and n[len(info):] in allowed) for n in names)
        rows = list(csv.reader(io.StringIO(archive.read(info + 'RECORD').decode('utf-8'))))
        assert {r[0] for r in rows} == set(names) and len(rows) == len(names)
        for name, digest, size in rows:
            if name == info + 'RECORD':
                assert digest == size == ''
                continue
            data = archive.read(name)
            expected_digest = base64.urlsafe_b64encode(hashlib.sha256(data).digest()).rstrip(b'=').decode()
            assert digest == 'sha256=' + expected_digest and size == str(len(data))
        assert 'conferencia-xlsx = conferencia_xlsx.cli:main' in archive.read(info + 'entry_points.txt').decode()
    with tarfile.open(sdists[0], 'r:gz') as archive:
        members = archive.getmembers()
        prefix = f'{NAME}-{VERSION}/'
        paths = set()
        for item in members:
            path = PurePosixPath(item.name)
            assert not path.is_absolute() and '..' not in path.parts and not item.issym() and not item.islnk()
            assert item.name not in paths
            paths.add(item.name)
            assert item.name.startswith(prefix)
            relative = item.name[len(prefix):]
            assert not any(bad in path.parts for bad in ('.venv', 'local', '__pycache__', '.env', '.git'))
            if relative.startswith('src/') and relative[4:] in expected:
                assert archive.extractfile(item).read() == expected[relative[4:]]
        for name in expected:
            assert prefix + 'src/' + name in paths
        for name in ['README.md', 'LICENSE', 'CONTRACT.md', 'pyproject.toml', 'tools/libreoffice_rehearsal.py']:
            assert prefix + name in paths
    print(f'Distribuição conferida: {len(expected)} arquivos de execução, recursos XLSX, sdist e RECORD íntegros.')


if __name__ == '__main__':
    main()
