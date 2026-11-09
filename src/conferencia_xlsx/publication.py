"""Publicação de um único arquivo preparado, sem sobrescrever o destino."""
from __future__ import annotations

import os
from pathlib import Path
import tempfile

from .errors import AuditError


def publish_bytes(destination: str | os.PathLike, data: bytes) -> None:
    if not isinstance(data, bytes):
        raise AuditError("invalid-output", "A saída precisa ser preparada antes da publicação.")
    path = Path(destination)
    temporary = None
    try:
        fd, temporary = tempfile.mkstemp(prefix=".conferencia-xlsx-", suffix=".tmp", dir=path.parent)
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        # link() fails for any existing directory entry, including a broken
        # symlink. No check-then-replace window and no destructive fallback.
        os.link(temporary, path, follow_symlinks=False)
    except FileExistsError:
        raise AuditError("output-exists", "O destino já existe. Escolha outro arquivo; nada foi sobrescrito.") from None
    except OSError:
        raise AuditError("output-failed", "Não foi possível publicar a saída. Use uma pasta gravável que permita links locais.") from None
    finally:
        if temporary is not None:
            try:
                os.unlink(temporary)
            except OSError:
                pass
