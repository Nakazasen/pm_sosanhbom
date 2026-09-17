# Entrypoint for active SSBOM bundle
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from src.gui.app import main

if __name__ == '__main__':
    main()
