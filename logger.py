import os
import traceback
from datetime import datetime

def log_erro(msg):
    """Grava erros em um arquivo de texto para auditoria e debug."""
    try:
        with open("log_erros.txt", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now()}: {msg}\n")
    except Exception:
        pass