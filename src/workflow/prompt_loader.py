from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Any


PROMPTS_DIR = Path(__file__).resolve().parent / "prompts"


@lru_cache(maxsize=None)
def load_prompt(relative_path: str) -> str:
    path = PROMPTS_DIR / relative_path
    if not path.exists():
        raise FileNotFoundError(f"Prompt file not found: {path}")
    return path.read_text(encoding="utf-8").strip()


def render_prompt(relative_path: str, **kwargs: Any) -> str:
    template = load_prompt(relative_path)
    return template.format(**kwargs)
