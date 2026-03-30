"""
config.py — Centralised application configuration via environment variables.
All settings are loaded once and cached; downstream modules import `settings`.
"""

from functools import lru_cache
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application-wide settings resolved from .env or environment variables."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    # ── NVIDIA NIM ────────────────────────────────────────────────────────────
    nvidia_api_key: str = "nvapi-placeholder-replace-me"
    nvidia_base_url: str = "https://integrate.api.nvidia.com/v1"
    nvidia_embed_model: str = "nvidia/nv-embedqa-e5-v5"
    nvidia_llm_model: str = "meta/llama-3.1-70b-instruct"
    nvidia_rerank_model: str = "nvidia/nv-rerankqa-mistral-4b-v3"

    # ── ChromaDB ──────────────────────────────────────────────────────────────
    chroma_persist_dir: str = "./data/chroma_db"
    chroma_collection_name: str = "supply_chain_rag"

    # ── Retrieval ─────────────────────────────────────────────────────────────
    top_k_retrieval: int = 15       # candidates fetched before reranking
    top_k_rerank: int = 6           # final context chunks sent to LLM
    bm25_weight: float = 0.3        # weight for BM25 score in hybrid fusion
    vector_weight: float = 0.7      # weight for vector score in hybrid fusion

    # ── Application ───────────────────────────────────────────────────────────
    app_host: str = "0.0.0.0"
    app_port: int = 8000
    log_level: str = "INFO"
    evaluation_db_path: str = "./data/evaluation.db"


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    """Return a cached singleton Settings instance."""
    return Settings()


# Module-level alias for ergonomic imports: `from config import settings`
settings: Settings = get_settings()
