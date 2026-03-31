"""Chain of Spreadsheet (CoS) — QA pipeline for SpreadsheetLLM.

Paper Section 4: Three-stage methodology:
  Stage 1: Table identification from compressed encoding (Section 4.1)
  Stage 2: Response generation from uncompressed table (Section 4.2)
  Table splitting for large tables (Appendix M.2, Algorithm 2)

Supports: anthropic, openai, or placeholder LLM backends.
"""

from __future__ import annotations

import json
import logging
import os
import re
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)

# --- Prompt templates from Appendix L.3 ---

QA_STAGE1_PROMPT = """INSTRUCTION:
Given an input that is a string denoting data of cells in a table. The input \
table contains many tuples, describing the cells with content in the spreadsheet. \
Each tuple consists of two elements separated by a '|': the cell content and the \
cell address/region, like (Year|A1), ( |A1) or (IntNum|A1:B3). The content in \
some cells such as '#,##0'/'d-mmm-yy'/'H:mm:ss', etc., represents the CELL DATA \
FORMATS of Excel. The content in some cells such as 'IntNum'/'DateData'/'EmailData', \
etc., represents a category of data with the same format and similar semantics. \
For example, 'IntNum' represents integer type data, and 'ScientificNum' represents \
scientific notation type data. 'A1:B3' represents a region in a spreadsheet, from \
the first row to the third row and from column A to column B. Some cells with empty \
content in the spreadsheet are not entered. How many tables are there in the \
spreadsheet? Below is a question about one certain table in this spreadsheet. I need \
you to determine in which table the answer to the following question can be found, \
and return the RANGE of the ONE table you choose, LIKE ['range': 'A1:F9']. \
DON'T ADD OTHER WORDS OR EXPLANATION.

INPUT:
{encoded_sheet}

QUESTION:
{question}"""

QA_STAGE2_PROMPT = """INSTRUCTION:
Given an input that is a string denoting data of cells in a table and a question \
about this table. The answer to the question can be found in the table. The input \
table includes many pairs, and each pair consists of a cell address and the text in \
that cell with a ',' in between, like 'A1,Year'. Cells are separated by '|' like \
'A1,Year|A2,Profit'. The text can be empty so the cell data is like 'A1, |A2,Profit'. \
The cells are organized in row-major order. The answer to the input question is \
contained in the input table and can be represented by cell address. I need you to \
find the cell address of the answer in the given table based on the given question \
description, and return the cell ADDRESS of the answer like '[B3]' or \
'[SUM(A2:A10)]'. DON'T ADD ANY OTHER WORDS.

INPUT:
{encoded_table}

QUESTION:
{question}"""


# =============================================================================
# LLM Backend
# =============================================================================


def _call_llm(prompt: str, provider: str = "placeholder") -> str:
    """Call an LLM with the given prompt.

    Supported providers:
      - "anthropic": Uses ANTHROPIC_API_KEY env var (Claude)
      - "openai": Uses OPENAI_API_KEY env var (GPT-4)
      - "placeholder": Returns a demo response (no API needed)
    """
    if provider == "anthropic":
        return _call_anthropic(prompt)
    if provider == "openai":
        return _call_openai(prompt)
    return _call_placeholder(prompt)


def _call_anthropic(prompt: str) -> str:
    """Call Claude via the Anthropic SDK."""
    try:
        import anthropic
    except ImportError:
        logger.error("anthropic package not installed. Run: pip install anthropic")
        return ""

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        logger.error("ANTHROPIC_API_KEY not set.")
        return ""

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=512,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


def _call_openai(prompt: str) -> str:
    """Call GPT-4 via the OpenAI SDK."""
    try:
        import openai
    except ImportError:
        logger.error("openai package not installed. Run: pip install openai")
        return ""

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        logger.error("OPENAI_API_KEY not set.")
        return ""

    client = openai.OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=512,
    )
    return response.choices[0].message.content or ""


def _call_placeholder(prompt: str) -> str:
    """Return a placeholder response for demo/testing."""
    logger.warning("Using placeholder LLM — no real inference.")
    if "determine in which table" in prompt:
        return "['range': 'A1:F9']"
    if "find the cell address" in prompt:
        return "[B3]"
    return "[placeholder response]"


def _detect_provider() -> str:
    """Auto-detect available LLM provider from env vars."""
    if os.environ.get("ANTHROPIC_API_KEY"):
        return "anthropic"
    if os.environ.get("OPENAI_API_KEY"):
        return "openai"
    return "placeholder"


# =============================================================================
# CoS Pipeline
# =============================================================================


def identify_table(
    encoding: Dict[str, Any],
    query: str,
    provider: Optional[str] = None,
) -> Optional[str]:
    """CoS Stage 1: Identify the relevant table range for a query.

    Uses the compressed encoding + LLM to determine which table
    contains the answer.

    Returns: Table range string like 'A1:F9', or None.
    """
    if provider is None:
        provider = _detect_provider()

    sheet_name = _find_relevant_sheet(encoding, query)
    if not sheet_name:
        logger.warning("No relevant sheet found for query.")
        return None

    sheet_data = encoding["sheets"][sheet_name]
    encoded = json.dumps(sheet_data, ensure_ascii=False)

    prompt = QA_STAGE1_PROMPT.format(encoded_sheet=encoded, question=query)
    response = _call_llm(prompt, provider)

    # Parse range from response
    match = re.search(r"'([A-Z]+\d+:[A-Z]+\d+)'", response)
    if match:
        return match.group(1)

    logger.warning("Could not parse table range from: %s", response)
    return None


def generate_response(
    table_data: Dict[str, Any],
    query: str,
    provider: Optional[str] = None,
) -> str:
    """CoS Stage 2: Generate an answer from the identified table.

    Uses uncompressed table encoding + LLM to find the cell address.

    Returns: Answer string like '[B3]'.
    """
    if provider is None:
        provider = _detect_provider()

    encoded = json.dumps(table_data, ensure_ascii=False)
    prompt = QA_STAGE2_PROMPT.format(encoded_table=encoded, question=query)
    return _call_llm(prompt, provider)


def table_split_qa(
    sheet_data: Dict[str, Any],
    table_range: str,
    query: str,
    token_limit: int = 4096,
    provider: Optional[str] = None,
) -> str:
    """Handle QA for large tables by splitting into chunks.

    Implements Algorithm 2 from Appendix M.2:
    1. If table fits in token_limit, query directly.
    2. Otherwise, detect header, split body into chunks
       where header + chunk < limit.
    3. Query each chunk, aggregate answers.
    """
    if provider is None:
        provider = _detect_provider()

    table_tokens = len(json.dumps(sheet_data, ensure_ascii=False))

    if table_tokens <= token_limit:
        return generate_response(sheet_data, query, provider)

    logger.info("Table too large (%d tokens), splitting.", table_tokens)

    # Extract cells as a flat list for chunking
    cells = sheet_data.get("cells", {})
    all_items = list(cells.items())

    if not all_items:
        return generate_response(sheet_data, query, provider)

    # Approximate header as first ~10% of items
    header_size = max(1, len(all_items) // 10)
    header_items = all_items[:header_size]
    body_items = all_items[header_size:]

    header_data = {"cells": dict(header_items)}
    header_tokens = len(json.dumps(header_data, ensure_ascii=False))
    chunk_budget = token_limit - header_tokens

    if chunk_budget <= 0:
        # Header alone exceeds limit — just query full
        return generate_response(sheet_data, query, provider)

    # Split body into chunks
    answers: List[str] = []
    chunk: List[Any] = []
    chunk_tokens = 0

    for item in body_items:
        item_tokens = len(json.dumps(dict([item]), ensure_ascii=False))
        if chunk_tokens + item_tokens > chunk_budget and chunk:
            # Process this chunk
            chunk_data = {"cells": {**dict(header_items), **dict(chunk)}}
            answer = generate_response(chunk_data, query, provider)
            answers.append(answer)
            chunk = []
            chunk_tokens = 0

        chunk.append(item)
        chunk_tokens += item_tokens

    # Process remaining chunk
    if chunk:
        chunk_data = {"cells": {**dict(header_items), **dict(chunk)}}
        answer = generate_response(chunk_data, query, provider)
        answers.append(answer)

    if len(answers) == 1:
        return answers[0]

    return f"Aggregated from {len(answers)} chunks: " + " | ".join(answers)


def _find_relevant_sheet(
    encoding: Dict[str, Any],
    query: str,
) -> Optional[str]:
    """Find the most relevant sheet via keyword matching."""
    query_tokens = {t.lower() for t in query.split()}
    best_score = 0
    best_sheet = None

    for sheet_name, sheet_data in encoding.get("sheets", {}).items():
        score = 0
        for value in sheet_data.get("cells", {}):
            lower_val = str(value).lower()
            if any(token in lower_val for token in query_tokens):
                score += 1
        if score > best_score:
            best_score = score
            best_sheet = sheet_name

    return best_sheet
