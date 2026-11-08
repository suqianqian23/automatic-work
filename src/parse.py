import sys
from pathlib import Path


class Parse:
    def resource_root(self):
        if self._is_frozen():
            return Path(sys._MEIPASS)  # type: ignore[attr-defined]
        return Path(__file__).parent.parent

    def _is_frozen(self) -> bool:
        return getattr(sys, "frozen", False)
