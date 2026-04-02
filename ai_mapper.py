"""
AI-powered bank statement extractor using OpenAI.
Sends the full document content to GPT and asks it to extract all transactions directly.
Supports PDF and XLSX/XLSM files.
"""

from __future__ import annotations
import json
from pathlib import Path
import openpyxl

try:
    import pdfplumber
    _PDF_AVAILABLE = True
except ImportError:
    _PDF_AVAILABLE = False

try:
    from openai import OpenAI
    _OPENAI_AVAILABLE = True
except ImportError:
    _OPENAI_AVAILABLE = False


# ─────────────────────────── CONTENT READERS ─────────────────────────────────

def _read_pdf_text(filepath: Path) -> str:
    """Extract all text from a PDF, page by page."""
    if not _PDF_AVAILABLE:
        raise RuntimeError("pdfplumber não instalado. Execute: pip install pdfplumber")
    pages = []
    with pdfplumber.open(str(filepath)) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            if text.strip():
                pages.append(f"--- Página {i} ---\n{text}")
    return "\n\n".join(pages)


def _read_xlsx_text(filepath: Path) -> str:
    """Convert xlsx rows to a readable text block."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.worksheets[0]
    lines = []
    for row in ws.iter_rows(values_only=True):
        cells = [str(c) if c is not None else "" for c in row]
        line = " | ".join(cells).strip(" |")
        if line:
            lines.append(line)
    wb.close()
    return "\n".join(lines)


# ─────────────────────────── AI EXTRACTION ───────────────────────────────────

_SYSTEM = (
    "You are a financial data extraction assistant. "
    "You receive bank statement text and return structured transaction data as JSON."
)

_PROMPT = """\
Below is the content of a bank statement. Extract ALL financial transactions.

Rules:
- Skip balance rows (\"Saldo Anterior\", \"Total\", \"Saldo Invest\", etc.)
- If a date is missing for a row, inherit it from the previous transaction
- Income / credits → type = "RECEITAS", value must be positive
- Expenses / debits → type = "DESPESAS", value must be negative
- Merge multi-line descriptions into a single string
- Use Brazilian date format DD/MM/YYYY

Return ONLY this JSON (no explanation):
{{
  "transactions": [
    {{"date": "DD/MM/YYYY", "description": "...", "value": 0.00, "type": "RECEITAS|DESPESAS"}},
    ...
  ]
}}

Bank statement:
{content}"""


def _call_openai(content: str, api_key: str) -> list[dict]:
    if not _OPENAI_AVAILABLE:
        raise RuntimeError("openai não instalado. Execute: pip install openai")

    client = OpenAI(api_key=api_key)

    # Trim content if very long (keep ~12000 chars to stay within token limits)
    if len(content) > 12000:
        content = content[:12000] + "\n... [truncated]"

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": _SYSTEM},
            {"role": "user", "content": _PROMPT.format(content=content)},
        ],
        temperature=0,
        response_format={"type": "json_object"},
    )

    data = json.loads(response.choices[0].message.content)
    return data.get("transactions", [])


# ─────────────────────────── MAIN ENTRY ──────────────────────────────────────

def extract_with_ai(filepath: Path, api_key: str) -> list[dict]:
    """
    Read a bank file, ask OpenAI to extract all transactions, return normalized rows.
    Each row: {sheet, data, historico, valor}
    """
    from engine import _parse_date, _parse_value  # late import avoids circular

    suffix = filepath.suffix.lower()
    if suffix == ".pdf":
        content = _read_pdf_text(filepath)
    elif suffix in (".xlsx", ".xlsm", ".xls"):
        content = _read_xlsx_text(filepath)
    else:
        raise ValueError(f"Formato não suportado: {suffix}")

    if not content.strip():
        raise ValueError("Nenhum texto encontrado no arquivo.")

    transactions = _call_openai(content, api_key)

    if not transactions:
        raise ValueError("IA não encontrou transações no arquivo.")

    result = []
    for tx in transactions:
        parsed_date = _parse_date(str(tx.get("date", "")))
        parsed_val  = _parse_value(tx.get("value"))
        desc        = str(tx.get("description", "")).strip()
        tx_type     = str(tx.get("type", "")).upper().strip()

        if parsed_date is None or parsed_val is None or not desc:
            continue

        # Respect AI type field; fall back to value sign
        if tx_type == "RECEITAS":
            target = "RECEITAS"
            parsed_val = abs(parsed_val)
        elif tx_type == "DESPESAS":
            target = "DESPESAS"
            parsed_val = -abs(parsed_val)
        else:
            target = "DESPESAS" if parsed_val < 0 else "RECEITAS"

        result.append({
            "sheet":     target,
            "data":      parsed_date,
            "historico": desc,
            "valor":     parsed_val,
        })

    return result


# ─────────────────────────── CLASSIFICATION ─────────────────────────────────

_CLASSIFY_SYSTEM = (
    "You are a financial transaction classifier for a Brazilian company. "
    "Return only valid JSON, no explanation."
)


def classify_transactions(
    rows: list[dict],
    categories_by_sheet: dict,
    memory_rules: dict,
    api_key: str,
) -> list[dict]:
    """
    Classify all transactions in a single OpenAI call.
    categories_by_sheet = {"DESPESAS": [...], "RECEITAS": [...]}
    memory_rules = {"DESPESAS": {"desc": "Cat"}, "RECEITAS": {"desc": "Cat"}}
    Adds "classificacao" key to each row (empty string if no match).
    """
    if not _OPENAI_AVAILABLE or not rows:
        for row in rows:
            row.setdefault("classificacao", "")
        return rows

    despesas_cats = categories_by_sheet.get("DESPESAS", [])
    receitas_cats  = categories_by_sheet.get("RECEITAS", [])

    if not despesas_cats and not receitas_cats:
        for row in rows:
            row.setdefault("classificacao", "")
        return rows

    despesas_mem = memory_rules.get("DESPESAS", {})
    receitas_mem = memory_rules.get("RECEITAS", {})
    
    mem_lines = []
    if despesas_mem:
        mem_lines.append("Historically past DESPESAS (Use these if description matches closely):")
        for k, v in despesas_mem.items():
            mem_lines.append(f'- "{k}" -> "{v}"')
    if receitas_mem:
        mem_lines.append("Historically past RECEITAS (Use these if description matches closely):")
        for k, v in receitas_mem.items():
            mem_lines.append(f'- "{k}" -> "{v}"')
            
    mem_prompt = "\n".join(mem_lines) + "\n\n" if mem_lines else ""

    tx_lines = [
        f'{i}: [{row["sheet"]}] {row["historico"]}'
        for i, row in enumerate(rows)
    ]

    prompt = (
        "Classify each transaction below into one category from the provided lists.\n\n"
        f"{mem_prompt}"
        f"DESPESAS categories:\n{json.dumps(despesas_cats, ensure_ascii=False)}\n\n"
        f"RECEITAS categories:\n{json.dumps(receitas_cats, ensure_ascii=False)}\n\n"
        "Transactions (format: \"index: [TYPE] description\"):\n"
        + "\n".join(tx_lines)
        + "\n\nRules:\n"
        "- DESPESAS transactions → choose from DESPESAS list only\n"
        "- RECEITAS transactions → choose from RECEITAS list only\n"
        '- If no category fits, use ""\n'
        "- Return ONLY this JSON:\n"
        '{"classifications": [{"index": 0, "classificacao": "Category Name"}, ...]}'
    )

    try:
        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": _CLASSIFY_SYSTEM},
                {"role": "user",   "content": prompt},
            ],
            temperature=0,
            response_format={"type": "json_object"},
        )
        data = json.loads(response.choices[0].message.content)
        index_map = {
            item["index"]: item.get("classificacao", "")
            for item in data.get("classifications", [])
        }

        valid_despesas = set(despesas_cats)
        valid_receitas  = set(receitas_cats)

        for i, row in enumerate(rows):
            cat = index_map.get(i, "")
            valid_set = valid_despesas if row["sheet"] == "DESPESAS" else valid_receitas
            if cat and valid_set and cat not in valid_set:
                cat = ""
            row["classificacao"] = cat

    except Exception as exc:
        print(f"[CLASSIFY WARNING] {exc}")
        for row in rows:
            row.setdefault("classificacao", "")

    return rows
