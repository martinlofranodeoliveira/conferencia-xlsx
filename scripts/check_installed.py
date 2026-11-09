"""Instala o wheel em ambiente novo e executa a CLI fora do checkout."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import venv
import zipfile


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--wheel', required=True)
    args = parser.parse_args()
    wheel = Path(args.wheel).resolve()
    env = os.environ.copy()
    env.pop('PYTHONPATH', None)
    env['PYTHONUTF8'] = '1'
    env['PYTHONNOUSERSITE'] = '1'
    with tempfile.TemporaryDirectory(prefix='conferencia-installed-') as directory:
        root = Path(directory)
        venv.EnvBuilder(with_pip=True, symlinks=os.name != 'nt').create(root / 'venv')
        python = root / 'venv' / ('Scripts/python.exe' if os.name == 'nt' else 'bin/python')

        def run(arguments, code=0):
            result = subprocess.run([str(python), *arguments], cwd=root, env=env, capture_output=True, text=True, encoding='utf-8', timeout=120)
            assert result.returncode == code, (result.stdout, result.stderr)
            return result

        run(['-m', 'pip', '--disable-pip-version-check', 'install', str(wheel)])
        result = run(['-I', '-c', 'import conferencia_xlsx; print(conferencia_xlsx.__file__)'])
        assert str(root / 'venv') in result.stdout
        run(['-I', '-m', 'conferencia_xlsx', 'exemplo', '--saida', 'demo.zip'])
        with zipfile.ZipFile(root / 'demo.zip') as archive:
            manifest = json.loads(archive.read('manifesto.json'))
            assert set(archive.namelist()) == set(manifest['members']) | {'manifesto.json'}
            for name, record in manifest['members'].items():
                data = archive.read(name)
                assert len(data) == record['size_bytes'] and hashlib.sha256(data).hexdigest() == record['sha256']
            for name in ['antes.xlsx', 'depois.xlsx']:
                (root / name).write_bytes(archive.read(name))
        run(['-I', '-m', 'conferencia_xlsx', 'comparar', 'antes.xlsx', 'depois.xlsx', '--saida', 'result.json', '--formato', 'json'], code=1)
        result = json.loads((root / 'result.json').read_bytes())
        assert result['summary']['changes'] == 11 and result['status'] == 'different_in_scope'
        original = (root / 'result.json').read_bytes()
        run(['-I', '-m', 'conferencia_xlsx', 'comparar', 'antes.xlsx', 'depois.xlsx', '--saida', 'result.json', '--formato', 'json'], code=2)
        assert (root / 'result.json').read_bytes() == original
        print('Wheel instalado fora do checkout: exemplo completo, 11 alterações e destino preservado.')


if __name__ == '__main__':
    main()
