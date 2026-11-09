"""Conferência local de conteúdo e estrutura selecionada de XLSX."""
from .errors import AuditError
from .models import Limits, Value, Cache, Cell, Sheet, Snapshot

__version__ = '1.0.0'
__all__ = ['AuditError', 'Limits', 'Value', 'Cache', 'Cell', 'Sheet', 'Snapshot']
