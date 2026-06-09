from __future__ import annotations

import json
import threading
from datetime import datetime
from pathlib import Path

from .paths import ensure_dir


class JsonlLogger:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        ensure_dir(self.path.parent)
        self._lock = threading.Lock()

    def emit(self, event: str, **fields):
        payload = {"ts": datetime.now().isoformat(timespec="seconds"), "event": event, **fields}
        line = json.dumps(payload, ensure_ascii=False)
        with self._lock:
            with self.path.open("a", encoding="utf-8") as f:
                f.write(line + "\n")
        return payload

