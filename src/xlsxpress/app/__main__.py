"""Enable `python -m xlsxpress.app`, the command the context menu entries run."""

from __future__ import annotations

import sys

from xlsxpress.app.main import main

sys.exit(main())
