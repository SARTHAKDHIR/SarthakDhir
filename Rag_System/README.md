#AI Assistant — RAG System

A production-grade Retrieval-Augmented Generation (RAG) system that lets you query Excel-based business data using natural language. Upload your data, ask questions, and get LLM-powered answers backed by hybrid vector + BM25 retrieval.

---

## Architecture

```
Excel Upload → Parsing & Chunking → ChromaDB (Vector Store)
                                  → BM25 Index (Keyword)
                                         ↓
User Query → Query Rewriting → Hybrid Retrieval → Reranking → LLM → Answer
```

**Retrieval pipeline:**
- Hybrid search: ChromaDB vector similarity (70%) + BM25 keyword (30%)
- Reranking via NVIDIA NIM reranker before passing context to LLM
- Query rewriting + sub-query decomposition for better recall

---

## Features

- 📤 **Excel ingestion** — Upload workbooks with PO, Vendor, and Inventory sheets; auto-parsed and embedded
- 🔍 **Natural language querying** — Ask anything about your data in plain English
- 🔄 **Query rewriting** — Automatic query expansion and sub-query generation
- 📊 **Business analytics** — Pre-computed KPIs injected into LLM context
- ⭐ **Evaluation logging** — Every query logged with latency breakdown and manual relevance rating
- 🐳 **Docker support** — One-command deployment with Docker Compose

---

## Tech Stack

| Layer | Technology |
|---|---|
| LLM & Embeddings | NVIDIA NIM (Llama 3.1 70B + NV-EmbedQA-E5) |
| Reranking | NVIDIA NIM (NV-RerankQA Mistral 4B) |
| Vector Store | ChromaDB |
| Keyword Search | BM25 (rank-bm25) |
| Backend API | FastAPI + Uvicorn |
| Frontend UI | Streamlit + Plotly |
| Data Processing | pandas, openpyxl |
| Config | Pydantic Settings |
| Deployment | Docker, Docker Compose |

---

## Project Structure

```
├── main.py                  # FastAPI app entry point
├── streamlit_app.py         # Streamlit UI
├── config.py                # Centralised settings via .env
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── .env.example             # Environment variable template
├── supply_chain_sample.xlsx # Sample input data (3 sheets)
├── dumy_data.py             # Script to regenerate sample data
│
├── api/
│   └── routes.py            # API route definitions
│
├── retrieval/
│   ├── vector_store.py      # ChromaDB wrapper
│   └── hybrid_retriever.py  # BM25 + vector hybrid retrieval + reranking
│
└── evaluation/
    └── logger.py            # Query logging and metrics to SQLite
```

---

## Setup

### Option 1 — Local

```bash
# 1. Clone the repo
git clone https://github.com/yourusername/supply-chain-rag.git
cd supply-chain-rag

# 2. Install dependencies
pip install -r requirements.txt

# 3. Configure environment
cp .env.example .env
# Add your NVIDIA API key to .env

# 4. Start the FastAPI backend
uvicorn main:app --host 0.0.0.0 --port 8000 --reload

# 5. Start the Streamlit UI (separate terminal)
streamlit run streamlit_app.py
```

Open `http://localhost:8501` for the UI — `http://localhost:8000/docs` for the API.

### Option 2 — Docker

```bash
cp .env.example .env
# Add your NVIDIA API key to .env

docker-compose up --build
```

---

## Environment Variables

Copy `.env.example` to `.env` and fill in:

```
NVIDIA_API_KEY=nvapi-your-key-here
NVIDIA_LLM_MODEL=meta/llama-3.1-70b-instruct
NVIDIA_EMBED_MODEL=nvidia/nv-embedqa-e5-v5
NVIDIA_RERANK_MODEL=nvidia/nv-rerankqa-mistral-4b-v3
```

Get your NVIDIA NIM API key at [build.nvidia.com](https://build.nvidia.com).

---

## Sample Data

`supply_chain_sample.xlsx` contains three sheets:
- **PO_Data** — 60 purchase orders with vendor, item, quantity, status, and delivery dates
- **Vendor_Data** — 12 vendors with reliability scores, lead times, and on-time delivery %
- **Inventory_Data** — 40 SKUs with stock levels, reorder points, and primary suppliers

Run `python dumy_data.py` to regenerate it.

---

## API Endpoints

| Method | Endpoint | Description |
|---|---|---|
| `POST` | `/api/v1/upload` | Upload and ingest an Excel file |
| `POST` | `/api/v1/query` | Run a RAG query |
| `GET` | `/api/v1/health` | Backend health + vector store stats |
| `GET` | `/api/v1/metrics` | Query logs and aggregate KPIs |
| `POST` | `/api/v1/metrics/rate` | Submit manual relevance rating |
