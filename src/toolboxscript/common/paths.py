from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re


def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H%M%S")


def slugify(text: str) -> str:
    text = (text or "").strip().lower()
    text = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-")
    return text or "run"


def ensure_dir(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def new_run_dir(output_root: str | Path, script_name: str) -> Path:
    root = ensure_dir(output_root)
    base = root / f"{now_stamp()}_{slugify(script_name)}"
    run_dir = base
    idx = 2
    while run_dir.exists():
      run_dir = root / f"{base.name}-{idx}"
      idx += 1
    run_dir.mkdir(parents=True, exist_ok=False)
    for child in ("raw", "output", "raw/search", "raw/detail", "raw/meta"):
        (run_dir / child).mkdir(parents=True, exist_ok=True)
    return run_dir

