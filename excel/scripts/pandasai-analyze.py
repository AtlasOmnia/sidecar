#!/usr/bin/env python3
"""Run optional PandasAI analysis against CSV/XLSX/Parquet data.

PandasAI executes LLM-generated Python. Use only with trusted prompts/data and,
for sensitive material, a local model endpoint or the official Docker sandbox.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any

# PandasAI otherwise enables anonymized Scarf usage analytics by default.
os.environ.setdefault("SCARF_NO_ANALYTICS", "true")

import pandas as pd
import pandasai as pai
from pandasai_litellm.litellm import LiteLLM


def load_frame(path: Path, sheet: str | None) -> Any:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        frame = pd.read_csv(path)
    elif suffix in {".xlsx", ".xlsm"}:
        frame = pd.read_excel(path, sheet_name=sheet or 0)
    elif suffix in {".parquet", ".pq"}:
        frame = pd.read_parquet(path)
    else:
        raise ValueError("Supported inputs: .csv, .xlsx, .xlsm, .parquet, .pq")
    return pai.DataFrame(frame)


def serialize_response(response: Any) -> dict[str, Any]:
    value = getattr(response, "value", response)
    if isinstance(value, pd.DataFrame):
        return {"type": "dataframe", "value": value.to_dict(orient="records")}
    if isinstance(value, pd.Series):
        return {"type": "series", "value": value.to_dict()}
    if isinstance(value, Path):
        return {"type": "path", "value": str(value)}
    if isinstance(value, (str, int, float, bool)) or value is None:
        return {"type": type(value).__name__, "value": value}
    return {"type": type(value).__name__, "value": str(value)}


def main() -> int:
    parser = argparse.ArgumentParser(description="Analyze trusted tabular data with PandasAI 3.")
    parser.add_argument("input", type=Path, help="CSV, XLSX/XLSM, or Parquet file")
    parser.add_argument("query", help="Natural-language analysis request")
    parser.add_argument("--sheet", help="Excel worksheet name (defaults to first sheet)")
    parser.add_argument("--model", default=os.getenv("PANDASAI_MODEL"), help="LiteLLM model identifier")
    parser.add_argument("--api-base", default=os.getenv("PANDASAI_API_BASE"), help="OpenAI-compatible endpoint")
    parser.add_argument("--api-key", default=os.getenv("PANDASAI_API_KEY", "lm-studio"), help=argparse.SUPPRESS)
    parser.add_argument("--temperature", type=float, default=0.0)
    args = parser.parse_args()

    if not args.input.is_file():
        parser.error(f"Input file does not exist: {args.input}")
    if not args.model:
        parser.error("Set --model or PANDASAI_MODEL")

    llm_kwargs: dict[str, Any] = {"api_key": args.api_key, "temperature": args.temperature}
    if args.api_base:
        llm_kwargs["api_base"] = args.api_base

    try:
        pai.config.set({"llm": LiteLLM(model=args.model, **llm_kwargs)})
        result = load_frame(args.input, args.sheet).chat(args.query)
        print(json.dumps({"ok": True, **serialize_response(result)}, ensure_ascii=False, default=str))
        return 0
    except Exception as exc:
        print(
            json.dumps(
                {"ok": False, "error": type(exc).__name__, "message": str(exc)},
                ensure_ascii=False,
            ),
            file=sys.stderr,
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
