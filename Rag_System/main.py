"""
main.py — FastAPI application entry point.

Run with:
    uvicorn main:app --host 0.0.0.0 --port 8000 --reload
"""

from __future__ import annotations

import logging
import os
import sys
from contextlib import asynccontextmanager
from typing import AsyncGenerator

import uvicorn
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse

from api.routes import router
from config import settings

# ─────────────────────────────────────────────────────────────────────────────
# Logging setup
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=getattr(logging, settings.log_level.upper(), logging.INFO),
    format="%(asctime)s | %(levelname)-8s | %(name)s:%(lineno)d — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)],
)

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Startup / Shutdown lifecycle
# ─────────────────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI) -> AsyncGenerator[None, None]:
    """
    Application lifespan handler.
    - On startup : ensure data directories exist; warm up singletons.
    - On shutdown: nothing special needed (ChromaDB persists automatically).
    """
    # Ensure required directories
    os.makedirs("./data/chroma_db", exist_ok=True)
    os.makedirs(os.path.dirname(os.path.abspath(settings.evaluation_db_path)), exist_ok=True)

    # Warm up singleton services (lazy-init, but we trigger them early to
    # surface misconfiguration at startup rather than first request)
    try:
        from retrieval.vector_store import get_vector_store
        from retrieval.hybrid_retriever import get_hybrid_retriever
        from evaluation.logger import get_eval_logger

        vs  = get_vector_store()
        rtr = get_hybrid_retriever()
        evl = get_eval_logger()

        # If the vector store already has data, rebuild the BM25 index
        if vs.count > 0:
            logger.info(
                "Existing vector store found (%d chunks). Rebuilding BM25 index...",
                vs.count,
            )
            rtr.refresh_bm25()

        logger.info("All services initialised. Vector store: %d chunks.", vs.count)

    except Exception as exc:
        logger.warning(
            "Non-fatal startup warning (services will init on first request): %s", exc
        )

    logger.info(
        "🚀 Enterprise Supply Chain AI Assistant is ready on "
        "http://%s:%d  |  Docs: http://%s:%d/docs",
        settings.app_host, settings.app_port,
        settings.app_host, settings.app_port,
    )

    yield   # <-- application runs here

    logger.info("Shutting down.")


# ─────────────────────────────────────────────────────────────────────────────
# FastAPI app
# ─────────────────────────────────────────────────────────────────────────────

app = FastAPI(
    title="Enterprise Supply Chain AI Assistant",
    description=(
        "Production-grade RAG system for supply-chain data analysis.\n\n"
        "Upload Excel files (PO_Data, Vendor_Data, Inventory_Data) and query "
        "your data using natural language.\n\n"
        "Powered by NVIDIA NIM LLM + ChromaDB hybrid retrieval."
    ),
    version="1.0.0",
    lifespan=lifespan,
    docs_url="/docs",
    redoc_url="/redoc",
)

# ── CORS (allow all origins for local dev; restrict in production) ────────────
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Routes ────────────────────────────────────────────────────────────────────
app.include_router(router, prefix="/api/v1")


@app.get("/", include_in_schema=False)
async def root() -> RedirectResponse:
    """Redirect root to API docs."""
    return RedirectResponse(url="/docs")


# ─────────────────────────────────────────────────────────────────────────────
# Dev entry point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host=settings.app_host,
        port=settings.app_port,
        reload=True,
        log_level=settings.log_level.lower(),
    )
