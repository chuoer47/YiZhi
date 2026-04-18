from __future__ import annotations

import os

from dotenv import load_dotenv
from langchain_openai import ChatOpenAI

if __package__:
    from .workflow_config import DOTENV_PATH
else:
    from workflow_config import DOTENV_PATH


def load_env() -> None:
    load_dotenv(dotenv_path=DOTENV_PATH, override=False)


def init_llm() -> ChatOpenAI:
    api_key = os.getenv("LLM_API_KEY")
    model = os.getenv("LLM_MODEL")
    base_url = os.getenv("LLM_BASE_URL")
    if not api_key or not model or not base_url:
        missing = [
            key
            for key, value in (
                ("LLM_API_KEY", api_key),
                ("LLM_MODEL", model),
                ("LLM_BASE_URL", base_url),
            )
            if not value
        ]
        raise RuntimeError(f"Missing required LLM env vars: {', '.join(missing)}")
    return ChatOpenAI(model=model, api_key=api_key, temperature=0.1, base_url=base_url)
