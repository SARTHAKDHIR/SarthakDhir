"""
streamlit_app.py — Streamlit UI for the Supply Chain AI Assistant.

Run with:
    streamlit run streamlit_app.py

Requires the FastAPI backend to be running at http://localhost:8000.
"""

from __future__ import annotations

import json
import time
from typing import Any, Dict

import plotly.express as px
import requests
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# Config
# ─────────────────────────────────────────────────────────────────────────────

API_BASE = "http://localhost:8000/api/v1"

st.set_page_config(
    page_title="Supply Chain AI Assistant",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ─────────────────────────────────────────────────────────────────────────────
# Helper utilities
# ─────────────────────────────────────────────────────────────────────────────

def _api_post(endpoint: str, **kwargs) -> Dict[str, Any]:
    try:
        resp = requests.post(f"{API_BASE}{endpoint}", timeout=120, **kwargs)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.ConnectionError:
        st.error("Cannot connect to the backend. Is FastAPI running on port 8000?")
        return {}
    except requests.HTTPError as exc:
        detail = exc.response.json().get("detail", str(exc))
        st.error(f"API error: {detail}")
        return {}


def _api_get(endpoint: str, params: Dict | None = None) -> Dict[str, Any]:
    try:
        resp = requests.get(f"{API_BASE}{endpoint}", params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.ConnectionError:
        st.error("Cannot connect to the backend. Is FastAPI running on port 8000?")
        return {}
    except requests.HTTPError as exc:
        detail = exc.response.json().get("detail", str(exc))
        st.error(f"API error: {detail}")
        return {}


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar — Navigation + Upload
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.image(
        "https://img.icons8.com/fluency/96/000000/factory.png",
        width=64,
    )
    st.title("Supply Chain AI")
    st.caption("Powered by NVIDIA NIM + ChromaDB")

    st.divider()

    page = st.radio(
        "Navigate",
        ["📤 Upload Data", "🔍 Query Assistant", "📊 Metrics & Logs"],
        label_visibility="collapsed",
    )

    st.divider()

    # Quick health check
    health = _api_get("/health")
    if health:
        docs = health.get("vector_store_docs", 0)
        bm25 = health.get("bm25_index_docs", 0)
        st.metric("Chunks in DB", docs)
        st.metric("BM25 Index", bm25)
        st.success("Backend: Online ✅")
    else:
        st.error("Backend: Offline ❌")


# ─────────────────────────────────────────────────────────────────────────────
# Page: Upload Data
# ─────────────────────────────────────────────────────────────────────────────

if page == "📤 Upload Data":
    st.header("📤 Upload Supply Chain Data")
    st.markdown(
        "Upload an Excel workbook with one or more of these sheets: "
        "**PO_Data**, **Vendor_Data**, **Inventory_Data**."
    )

    uploaded = st.file_uploader("Choose Excel file", type=["xlsx", "xls"])

    if uploaded and st.button("Ingest File", type="primary"):
        with st.spinner("Parsing, embedding, and indexing... this may take a moment."):
            result = _api_post(
                "/upload",
                files={"file": (uploaded.name, uploaded.getvalue(), "application/octet-stream")},
            )

        if result:
            st.success(f"✅ Ingestion complete in {result['duration_ms']:.0f}ms")

            col1, col2, col3 = st.columns(3)
            col1.metric("Sheets Parsed",    len(result.get("sheets_parsed", [])))
            col2.metric("Documents Built",  result.get("documents_built", 0))
            col3.metric("Chunks Stored",    result.get("chunks_stored", 0))

            st.markdown(f"**Sheets:** {', '.join(result.get('sheets_parsed', []))}")
            st.caption(f"Ingestion ID: `{result.get('ingestion_id', '')}`")

            # Show pre-computed analytics
            analytics = result.get("analytics", {})
            if analytics:
                st.subheader("Pre-computed Analytics")
                for sheet, kpis in analytics.items():
                    with st.expander(f"{sheet} KPIs", expanded=True):
                        st.json(kpis)


# ─────────────────────────────────────────────────────────────────────────────
# Page: Query Assistant
# ─────────────────────────────────────────────────────────────────────────────

elif page == "🔍 Query Assistant":
    st.header("🔍 Supply Chain Query Assistant")
    st.markdown("Ask any question about your PO, Vendor, or Inventory data.")

    # Query options
    with st.expander("⚙️ Query Options", expanded=False):
        col_a, col_b, col_c = st.columns(3)
        source_filter = col_a.selectbox(
            "Filter by Source",
            ["All", "PO_Data", "Vendor_Data", "Inventory_Data"],
        )
        top_k = col_b.slider("Top-K Retrieval", min_value=3, max_value=20, value=10)
        include_analytics = col_c.checkbox("Include Business Analytics", value=True)

    # Example queries
    examples = [
        "Which vendors have the most delays?",
        "What items are below reorder point?",
        "Show me all pending purchase orders.",
        "Which vendors have low reliability scores?",
        "What is our overall on-time delivery rate?",
    ]
    st.markdown("**Try an example query:**")
    example_cols = st.columns(len(examples))
    for i, ex in enumerate(examples):
        if example_cols[i].button(ex, key=f"ex_{i}", use_container_width=True):
            st.session_state["query_input"] = ex

    # Query input
    query_input = st.text_area(
        "Your question:",
        key="query_input",
        height=80,
        placeholder="E.g. Which vendors are performing worst in on-time delivery?",
    )

    if st.button("Ask", type="primary", disabled=not query_input.strip()):
        st.session_state["last_query"] = query_input

        payload = {
            "query":              query_input.strip(),
            "source_filter":      None if source_filter == "All" else source_filter,
            "top_k":              top_k,
            "include_analytics":  include_analytics,
        }

        with st.spinner("Running RAG pipeline..."):
            t0     = time.time()
            result = _api_post("/query", json=payload)
            wall   = (time.time() - t0) * 1000

        if result:
            # ── Answer ────────────────────────────────────────────────────────
            st.markdown("---")
            st.subheader("Answer")
            st.markdown(result.get("answer", "No answer returned."))

            # ── Latency strip ─────────────────────────────────────────────────
            lat = result.get("latency_ms", {})
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Retrieval",    f"{lat.get('retrieval_ms', 0):.0f} ms")
            c2.metric("Rerank",       f"{lat.get('rerank_ms', 0):.0f} ms")
            c3.metric("Generation",   f"{lat.get('generation_ms', 0):.0f} ms")
            c4.metric("Total",        f"{lat.get('total_ms', 0):.0f} ms")

            # ── Query rewriting transparency ──────────────────────────────────
            with st.expander("🔄 Query Rewriting", expanded=False):
                st.markdown(f"**Original:** {result.get('original_query', '')}")
                st.markdown(f"**Rewritten:** `{result.get('rewritten_query', '')}`")
                sub = result.get("sub_queries", [])
                if sub:
                    st.markdown("**Sub-queries:**")
                    for sq in sub:
                        st.markdown(f"  - {sq}")

            # ── Retrieved context chunks ───────────────────────────────────────
            chunks = result.get("retrieved_chunks", [])
            with st.expander(f"📄 Retrieved Chunks ({len(chunks)})", expanded=False):
                for i, chunk in enumerate(chunks, 1):
                    with st.container():
                        score = chunk.get("rerank_score") or chunk.get("rrf_score") or 0
                        st.markdown(
                            f"**[{i}] {chunk.get('source', 'N/A')}** "
                            f"| ID: `{chunk.get('id', '')}` "
                            f"| Score: `{score:.4f}`"
                        )
                        st.caption(chunk.get("text", "")[:300] + "...")
                        st.divider()

            # ── Business analytics ────────────────────────────────────────────
            analytics = result.get("analytics_used", {})
            if analytics:
                with st.expander("📊 Business Analytics Used", expanded=False):
                    st.json(analytics)

            # ── Manual relevance rating ───────────────────────────────────────
            st.markdown("---")
            st.markdown("**Rate this answer:**")
            rating_cols = st.columns(6)
            query_id = result.get("query_id", "")
            for score in range(6):
                label = ["⭐ 0", "⭐ 1", "⭐⭐ 2", "⭐⭐⭐ 3", "⭐⭐⭐⭐ 4", "⭐⭐⭐⭐⭐ 5"][score]
                if rating_cols[score].button(label, key=f"rate_{score}"):
                    _api_post(
                        f"/metrics/rate?query_id={query_id}&relevance={score}",
                    )
                    st.toast(f"Rated {score}/5 — thank you!", icon="✅")

            st.caption(f"Query ID: `{query_id}` | Model: `{result.get('model_used', '')}`")


# ─────────────────────────────────────────────────────────────────────────────
# Page: Metrics & Logs
# ─────────────────────────────────────────────────────────────────────────────

elif page == "📊 Metrics & Logs":
    st.header("📊 Pipeline Metrics & Evaluation Logs")

    metrics = _api_get("/metrics", params={"limit": 50})

    if metrics:
        agg = metrics.get("aggregate", {})
        qm  = agg.get("query_metrics", {})
        im  = agg.get("ingestion_metrics", {})

        # ── KPI strip ─────────────────────────────────────────────────────────
        st.subheader("Aggregate KPIs")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total Queries",     qm.get("total_queries",   0))
        k2.metric("Failures",          qm.get("total_failures",  0))
        k3.metric("Avg Latency",       f"{qm.get('avg_latency_ms') or 0:.0f} ms")
        k4.metric("Chunks in DB",      metrics.get("vector_store_count", 0))
        k5.metric("Docs Ingested",     im.get("total_docs",      0))

        # ── Latency breakdown chart ───────────────────────────────────────────
        if qm.get("total_queries", 0) > 0:
            st.subheader("Latency Breakdown")
            lat_data = {
                "Stage":      ["Retrieval", "Reranking", "Generation"],
                "Avg ms":     [
                    qm.get("avg_retrieval_ms",   0),
                    qm.get("avg_rerank_ms",      0),
                    qm.get("avg_generation_ms",  0),
                ],
            }
            fig = px.bar(
                lat_data, x="Stage", y="Avg ms",
                color="Stage", title="Average Latency per Pipeline Stage",
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            st.plotly_chart(fig, use_container_width=True)

        # ── Recent queries table ──────────────────────────────────────────────
        st.subheader("Recent Queries")
        recent = metrics.get("recent_queries", [])
        if recent:
            import pandas as pd
            df = pd.DataFrame(recent)[
                [
                    "timestamp", "original_query", "rewritten_query",
                    "total_latency_ms", "is_failure", "manual_relevance", "model_used",
                ]
            ].copy()
            df["total_latency_ms"] = df["total_latency_ms"].round(1)
            df["is_failure"]       = df["is_failure"].astype(bool)
            st.dataframe(df, use_container_width=True, height=350)
        else:
            st.info("No queries logged yet.")

        # ── Failure cases ─────────────────────────────────────────────────────
        failures = metrics.get("failure_cases", [])
        if failures:
            with st.expander(f"⚠️ Failure Cases ({len(failures)})", expanded=False):
                for f in failures:
                    st.markdown(
                        f"**Query:** {f.get('original_query', '')}  \n"
                        f"**Error:** {f.get('failure_reason', '')}  \n"
                        f"**Time:** {f.get('timestamp', '')}"
                    )
                    st.divider()
