import os
import sys

def resource_path(relative_path):
    """Função para acessar arquivos empacotados com o PyInstaller."""
    try:
        base_path = sys._MEIPASS  # Quando empacotado
    except AttributeError:
        base_path = os.path.abspath(".")  # Quando rodando como script
    return os.path.join(base_path, relative_path)