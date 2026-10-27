import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1] / "skills" / "make-style-pptx" / "scripts"
sys.path.insert(0, str(SCRIPTS))
