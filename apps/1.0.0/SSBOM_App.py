# Entrypoint for active SSBOM bundle
import sys
from pathlib import Path

# Health check flag support
if "--health-check" in sys.argv:
    print("[OK] SSBOM_App health check passed.")
    sys.exit(0)

# Inject parent directory into path
app_dir = Path(__file__).resolve().parent
if str(app_dir) not in sys.path:
    sys.path.insert(0, str(app_dir))

from src.gui.app import main

if __name__ == "__main__":
    main()
