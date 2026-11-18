"""Entry point: `python <skill folder>/oqa.py <command>` works from any folder, so the assistant never has to cd or set PYTHONPATH."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from oqa.cli import main  # noqa: E402

sys.exit(main())
