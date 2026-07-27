# Hermes STR Performance Pipeline

An automated data pipeline that ingests weekly hotel STR (Smith Travel Research) reports, consolidates performance metrics, generates analytics, powers an interactive dashboard, and provides a conversational AI assistant — all from a single data source.

---

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the Pipeline

**Step 1 — Ingest reports & generate analytics:**
```bash
python process_reports.py
```
This scans the `Input Files/` folder (including inside `.zip` archives), extracts metrics from each weekly report, and produces:
- `STR_Master.xlsx` — consolidated database of all hotel metrics
- `performance_brief.txt` — portfolio summary with trends, alerts, and rankings

**Step 2 — Launch the dashboard:**
```bash
python -m streamlit run dashboard.py
```
Open `http://localhost:8501` in your browser to explore interactive charts, KPI cards, and property comparisons.

**Step 3 — Chat with your data:**
```bash
set GEMINI_API_KEY=your_api_key_here
python ask_str.py
```
Ask natural-language questions about hotel performance and get precise, data-backed answers.

---

## What Each File Does

| File | Purpose |
|------|---------|
| `process_reports.py` | Scans folders & zips, reads Excel reports, extracts 12 metrics, builds the master database, and writes the performance brief |
| `STR_Master.xlsx` | The central database — 14 columns (Hotel Code, Date, + 12 metrics) with one row per hotel per week |
| `performance_brief.txt` | Text summary covering portfolio momentum, historical trends, urgent alerts, and rankings |
| `dashboard.py` | Streamlit web dashboard with KPI cards, trend charts, heatmaps, and data export |
| `ask_str.py` | AI-powered Q&A assistant that answers questions grounded exclusively in the master database |
| `requirements.txt` | Python package dependencies |

---

## AI Architecture: How `ask_str.py` Works

### Model

**Google Gemini 2.0 Flash** (via the stable `v1` API endpoint using the `google-genai` Python SDK v1.15+).

Chosen for its mature, production-grade function-calling support, low latency, and cost efficiency — ideal for a tool-grounded conversational interface.

### Design Choice: Tool Calling (vs. Context Stuffing or RAG)

We evaluated three approaches for grounding the AI in hotel data:

| Approach | How It Works | Why We Didn't Choose It |
|----------|-------------|------------------------|
| **Context Stuffing** | Paste the entire dataset into the prompt | Fragile with growing data; wastes tokens; model may still hallucinate from noise |
| **RAG (Retrieval-Augmented Generation)** | Embed data into a vector store, retrieve relevant chunks | Over-engineered for structured tabular data; adds infrastructure complexity |
| **Tool Calling** ✅ | Model calls Python functions to query exact data on demand | **Selected** — precise, scalable, zero hallucination risk on numbers |

**Why Tool Calling wins here:**
- The data is **structured and tabular** (not free-text documents), so exact lookups outperform fuzzy retrieval
- The model never sees raw data in its context — it *must* call a function to get numbers, eliminating fabricated statistics
- Adding new query capabilities is as simple as writing a new Python function

### Available Tools

The AI has access to three data-retrieval functions:

| Function | What It Does | Example Use |
|----------|-------------|-------------|
| `get_latest_metrics(inn_code)` | Returns the most recent week's 12 metrics for a hotel | "How is HEZCN performing?" |
| `get_history(inn_code, metric, limit)` | Returns a time series of a specific metric | "Show me MSYHV's RGI trend" |
| `get_top_n(metric, n, ascending)` | Ranks all hotels by a metric | "Who are the top 3 performers?" |

### Known Limitation

**The AI cannot perform cross-metric calculations or custom aggregations** beyond what the three tool functions provide. For example, asking *"What is the average RGI Index across all hotels for July?"* requires a computation that none of the current tools support — the model will attempt to answer using available tools but may not produce the precise aggregate. Adding more specialized tool functions would address this for specific use cases.

---

## Sample Chat Transcripts

### Example 1: Checking a specific hotel

```
You: How is HEZCN doing right now?

  → [Tool Call] get_latest_metrics(inn_code="HEZCN")

Hermes AI: HEZCN's latest data is from August 31, 2025. Their 28-Day RGI
Index is 117.45, which means they're outperforming their competitive set
by about 17%. However, the 28-Day RGI % Change is -42.63%, indicating a
significant decline in relative performance compared to the prior period.
```

### Example 2: Ranking properties

```
You: Who are the top 3 performers right now?

  → [Tool Call] get_top_n(metric="RGI_28d_PctChg", n=3, ascending=False)

Hermes AI: Based on the latest 28-Day RGI % Change:
  1. JANGM  +54.82%  (Sep 28)
  2. LQCHA  +31.14%  (Sep 28)
  3. MSYHV  -19.91%  (Sep 28)
JANGM and LQCHA are showing strong upward momentum, while MSYHV is the
best of the remaining properties despite still being in negative territory.
```

### Example 3: Historical trend with follow-up

```
You: Show me MSYHV's RGI Index trend over the last 5 weeks.

  → [Tool Call] get_history(inn_code="MSYHV", metric="RGI_28d_Index", limit=5)

Hermes AI: MSYHV's 28-Day RGI Index over the last 5 weeks:
  Aug 31: 95.66
  Sep 07: 89.79
  Sep 14: 92.44
  Sep 21: 86.47
  Sep 28: 81.03
The trend shows a steady decline from 95.66 to 81.03 — the property has
dropped below the competitive set par line of 100.

You: Is that the worst in the portfolio?

  → [Tool Call] get_top_n(metric="RGI_28d_Index", n=1, ascending=True)

Hermes AI: No — JANTW actually has a lower 28-Day RGI Index at 43.18
(as of Aug 31), making it the weakest performer in the portfolio. MSYHV
at 81.03 is the second-lowest among the five properties.
```

---

## Project Structure

```
Final Output/
├── process_reports.py      ← Step 1: Run this first
├── STR_Master.xlsx         ← Generated database
├── performance_brief.txt   ← Generated analytics summary
├── dashboard.py            ← Step 2: Streamlit dashboard
├── ask_str.py              ← Step 3: AI chat assistant
├── requirements.txt        ← Dependencies
└── README.md               ← This file
```
