#!/usr/bin/env python3
"""
ask_str.py — Conversational AI CLI for STR Data
==================================================
A Gemini-powered conversational CLI that answers questions about hotel
performance data grounded exclusively in STR_Master.xlsx.

Uses function/tool calling so the model fetches exact data rather than
hallucinating. Out-of-scope questions are refused gracefully.

Usage:
    python ask_str.py                          # uses GEMINI_API_KEY env var
    python ask_str.py --api-key YOUR_KEY       # pass key directly
"""



import argparse
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

# ──────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).resolve().parent
MASTER_FILE = SCRIPT_DIR / "STR_Master.xlsx"

MODEL_NAME = "gemini-2.0-flash"

HOTEL_DISPLAY_NAMES = {
    "HEZCN": "Holiday Inn Express & Suites Natchez South",
    "JANGM": "Holiday Inn Express & Suites Jackson Downtown Coliseum",
    "JANTW": "JANTW Property",
    "LQCHA": "La Quinta Inn & Suites by Wyndham Chattanooga Downtown South",
    "MSYHV": "MSYHV Property",
}

SYSTEM_INSTRUCTION = """\
You are **Hermes AI**, a helpful hotel performance analyst for Hermes Hospitality.

**Your ONLY data source is STR_Master.xlsx**, accessed exclusively through the
tools provided to you. NEVER invent, guess, or extrapolate data. If you don't
have the data needed to answer a question, say so.

## Rules
1. **Always use tools** to look up data before answering any data question.
   Never quote numbers from memory or prior turns — always call a tool.
2. **Refuse out-of-scope questions** gracefully. You only know about hotel STR
   performance metrics (MPI, ARI, RGI indices and % changes). You cannot answer
   questions about weather, stock prices, food, directions, or anything outside
   hotel performance data.
3. **Be precise**: quote exact numbers, dates, and hotel codes from tool results.
4. **Be concise**: answer in 2-4 sentences unless the user asks for detail.
5. **Hotel codes**: HEZCN, JANGM, JANTW, LQCHA, MSYHV. If the user mentions
   a hotel by name (e.g. "La Quinta", "Natchez"), infer the code.
6. **Metric terminology**:
   - MPI = Market Penetration Index (occupancy vs comp set)
   - ARI = Average Rate Index (ADR vs comp set)
   - RGI = Revenue Generation Index (RevPAR vs comp set)
   - Index > 100 = outperforming comp set; < 100 = underperforming
   - % Change = week-over-week or 28-day change in the index
7. **Formatting**: Use plain text suitable for a terminal. Align numbers in
   tables using spaces. Use headers with dashes for visual structure.
"""


# ──────────────────────────────────────────────────────────────
# Data Layer (loads once at startup)
# ──────────────────────────────────────────────────────────────

_df: pd.DataFrame = pd.DataFrame()


def _load_data() -> pd.DataFrame:
    """Load and cache the master dataset."""
    global _df
    if _df.empty:
        _df = pd.read_excel(MASTER_FILE, engine="openpyxl")
        _df["Date"] = pd.to_datetime(_df["Date"])
        _df.sort_values(["Inn Code", "Date"], inplace=True)
        _df.reset_index(drop=True, inplace=True)
    return _df


# ──────────────────────────────────────────────────────────────
# Tool Functions (exposed to Gemini via function calling)
# ──────────────────────────────────────────────────────────────


def get_latest_metrics(inn_code: str) -> dict:
    """Get the most recent week's metrics for a specific hotel.

    Args:
        inn_code: The hotel's Inn Code (e.g. HEZCN, JANGM, JANTW, LQCHA, MSYHV).

    Returns:
        A dictionary containing the hotel code, reporting date, and all 12
        STR metrics (MPI/ARI/RGI indices and % changes for 7-day and 28-day
        periods). Returns an error message if the hotel is not found.
    """
    df = _load_data()
    code = inn_code.upper().strip()
    hotel_data = df[df["Inn Code"] == code]

    if hotel_data.empty:
        available = ", ".join(sorted(df["Inn Code"].unique()))
        return {"error": f"Hotel '{code}' not found. Available: {available}"}

    latest = hotel_data.sort_values("Date").iloc[-1]
    result = {"inn_code": code, "date": latest["Date"].strftime("%Y-%m-%d")}
    for col in df.columns:
        if col not in ("Inn Code", "Date"):
            val = latest[col]
            result[col] = round(float(val), 2) if pd.notna(val) else None
    return result


def get_history(inn_code: str, metric: str, limit: int = 10) -> dict:
    """Get the historical time series for a specific metric and hotel.

    Args:
        inn_code: The hotel's Inn Code (e.g. HEZCN, JANGM, JANTW, LQCHA, MSYHV).
        metric: The metric column name. Valid options:
            MPI_7d_PctChg, ARI_7d_PctChg, RGI_7d_PctChg,
            MPI_28d_PctChg, ARI_28d_PctChg, RGI_28d_PctChg,
            MPI_7d_Index, ARI_7d_Index, RGI_7d_Index,
            MPI_28d_Index, ARI_28d_Index, RGI_28d_Index.
        limit: Maximum number of most recent data points to return (default 10).

    Returns:
        A dictionary with the hotel code, metric name, and a list of
        {date, value} entries sorted chronologically. Returns an error
        message if the hotel or metric is not found.
    """
    df = _load_data()
    code = inn_code.upper().strip()
    hotel_data = df[df["Inn Code"] == code]

    if hotel_data.empty:
        available = ", ".join(sorted(df["Inn Code"].unique()))
        return {"error": f"Hotel '{code}' not found. Available: {available}"}

    if metric not in df.columns:
        valid = [c for c in df.columns if c not in ("Inn Code", "Date")]
        return {"error": f"Invalid metric '{metric}'. Valid: {', '.join(valid)}"}

    hotel_data = hotel_data.sort_values("Date").tail(limit)
    history = []
    for _, row in hotel_data.iterrows():
        val = row[metric]
        history.append({
            "date": row["Date"].strftime("%Y-%m-%d"),
            "value": round(float(val), 2) if pd.notna(val) else None,
        })

    return {"inn_code": code, "metric": metric, "data": history}


def get_top_n(metric: str, n: int = 3, ascending: bool = False) -> dict:
    """Rank hotels by a specific metric using their most recent data.

    Args:
        metric: The metric column name to rank by. Valid options:
            MPI_7d_PctChg, ARI_7d_PctChg, RGI_7d_PctChg,
            MPI_28d_PctChg, ARI_28d_PctChg, RGI_28d_PctChg,
            MPI_7d_Index, ARI_7d_Index, RGI_7d_Index,
            MPI_28d_Index, ARI_28d_Index, RGI_28d_Index.
        n: Number of top (or bottom) hotels to return (default 3).
        ascending: If False (default), returns top performers (highest values).
            If True, returns bottom performers (lowest values).

    Returns:
        A dictionary with the metric name, ranking direction, and a list of
        {inn_code, date, value} entries. Returns an error if the metric
        is not found.
    """
    df = _load_data()

    if metric not in df.columns:
        valid = [c for c in df.columns if c not in ("Inn Code", "Date")]
        return {"error": f"Invalid metric '{metric}'. Valid: {', '.join(valid)}"}

    # Use each hotel's latest row
    latest = df.sort_values("Date").groupby("Inn Code").tail(1)
    ranked = latest.dropna(subset=[metric]).sort_values(metric, ascending=ascending)
    top = ranked.head(n)

    results = []
    for _, row in top.iterrows():
        results.append({
            "inn_code": row["Inn Code"],
            "date": row["Date"].strftime("%Y-%m-%d"),
            "value": round(float(row[metric]), 2),
        })

    return {
        "metric": metric,
        "direction": "bottom" if ascending else "top",
        "rankings": results,
    }


# ──────────────────────────────────────────────────────────────
# CLI Interface
# ──────────────────────────────────────────────────────────────

WELCOME_BANNER = """
╔══════════════════════════════════════════════════════════════╗
║   HERMES AI — STR Performance Assistant                     ║
║                                                              ║
║   Ask me anything about your hotel portfolio's STR metrics.  ║
║   Type 'quit' or 'exit' to end the session.                  ║
╚══════════════════════════════════════════════════════════════╝

  Available hotels: HEZCN · JANGM · JANTW · LQCHA · MSYHV
  Metrics: MPI, ARI, RGI (7-Day & 28-Day, Index & % Change)

"""


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Conversational AI CLI for STR hotel performance data"
    )
    parser.add_argument(
        "--api-key",
        type=str,
        default=None,
        help="Gemini API key (or set GEMINI_API_KEY env var)",
    )
    parser.add_argument(
        "--model",
        type=str,
        default=MODEL_NAME,
        help=f"Gemini model name (default: {MODEL_NAME})",
    )
    args = parser.parse_args()

    # Resolve API key
    api_key = args.api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        print("Error: No API key provided.")
        print("  Set GEMINI_API_KEY environment variable, or pass --api-key")
        sys.exit(1)

    # Late import so missing google-genai gives a clear error
    try:
        from google import genai
        from google.genai import types
    except ImportError:
        print("Error: google-genai package not installed.")
        print("  pip install google-genai")
        sys.exit(1)

    # Pre-load data to catch issues early
    df = _load_data()
    print(f"  Loaded {len(df)} records across {df['Inn Code'].nunique()} hotels")
    print(f"  Date range: {df['Date'].min().strftime('%Y-%m-%d')} to {df['Date'].max().strftime('%Y-%m-%d')}")

    # Initialize Gemini client using the stable v1 API version
    client = genai.Client(api_key=api_key)

    tools = [get_latest_metrics, get_history, get_top_n]

    chat = client.chats.create(
        model=args.model,
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_INSTRUCTION,
            tools=tools,
            temperature=0.2,  # low temperature for factual answers
        ),
    )

    print(WELCOME_BANNER)

    # Conversation loop
    while True:
        try:
            user_input = input("You: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n\nGoodbye!")
            break

        if not user_input:
            continue
        if user_input.lower() in ("quit", "exit", "q", "bye"):
            print("\nGoodbye! 👋")
            break

        try:
            response = chat.send_message(user_input)
            answer = response.text if response.text else "(No response generated)"

            print(f"\nHermes AI: {answer}\n")

        except Exception as e:
            print(f"\n⚠ Error: {e}\n")
            print("  (The conversation context is preserved — try another question)\n")


if __name__ == "__main__":
    main()
