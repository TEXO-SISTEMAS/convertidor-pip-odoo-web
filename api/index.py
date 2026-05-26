import sys
from pathlib import Path

# Agregar raíz del proyecto y src/ al path
root = Path(__file__).parent.parent
sys.path.insert(0, str(root))
sys.path.insert(0, str(root / "src"))

from app import app
