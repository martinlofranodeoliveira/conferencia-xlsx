class AuditError(ValueError):
    """Erro público sanitizado; não inclui valores ou caminhos das entradas."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message
