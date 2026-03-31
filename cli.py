"""CLI entry point for SpreadsheetLLM encoder.

Usage:
    python cli.py encode input.xlsx [-o output.json] [-k 2]
    python cli.py encode input.xlsx --vanilla [-o output.txt]
    python cli.py qa input.xlsx "What is the total revenue?" [--provider anthropic]
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys


def _cmd_encode(args: argparse.Namespace) -> None:
    """Handle the 'encode' subcommand."""
    from spreadsheet_llm.encoder import encode_spreadsheet
    from spreadsheet_llm.vanilla import vanilla_encode

    if args.vanilla:
        out = args.output or os.path.splitext(args.excel_file)[0] + "_vanilla.txt"
        result = vanilla_encode(args.excel_file, out)
        if result:
            first_sheet = next(iter(result))
            token_count = len(result[first_sheet])
            print(f"Vanilla encoding saved to {out}")
            print(f"  Sheet '{first_sheet}': {token_count} chars")
    else:
        out = args.output or os.path.splitext(args.excel_file)[0] + "_encoded.json"
        result = encode_spreadsheet(args.excel_file, k=args.k, output_path=out)
        if result:
            metrics = result.get("compression_metrics", {}).get("overall", {})
            print(f"SheetCompressor encoding saved to {out}")
            print(f"  Original tokens:  {metrics.get('original_tokens', 0):,}")
            print(f"  Final tokens:     {metrics.get('final_tokens', 0):,}")
            print(f"  Anchor ratio:     {metrics.get('anchor_ratio', 0):.2f}x")
            print(f"  Index ratio:      {metrics.get('inverted_index_ratio', 0):.2f}x")
            print(f"  Format ratio:     {metrics.get('format_ratio', 0):.2f}x")
            print(f"  Overall ratio:    {metrics.get('overall_ratio', 0):.2f}x")


def _cmd_qa(args: argparse.Namespace) -> None:
    """Handle the 'qa' subcommand."""
    from spreadsheet_llm.encoder import encode_spreadsheet
    from spreadsheet_llm.cos import identify_table, generate_response, table_split_qa

    print(f"Encoding {args.excel_file}...")
    encoding = encode_spreadsheet(args.excel_file, k=2)
    if not encoding:
        print("Error: could not encode file.", file=sys.stderr)
        sys.exit(1)

    print(f"Identifying relevant table for: {args.question}")
    table_range = identify_table(encoding, args.question, provider=args.provider)
    if table_range:
        print(f"  Table range: {table_range}")
    else:
        print("  Could not identify table — using first sheet.")

    # For now, use the first sheet's full encoding
    first_sheet = next(iter(encoding["sheets"]))
    sheet_data = encoding["sheets"][first_sheet]

    if table_range:
        answer = table_split_qa(
            sheet_data, table_range, args.question, provider=args.provider,
        )
    else:
        answer = generate_response(sheet_data, args.question, provider=args.provider)

    print(f"\nAnswer: {answer}")


def main() -> None:
    """CLI main entry point."""
    parser = argparse.ArgumentParser(
        description="SpreadsheetLLM — Encode spreadsheets for LLMs",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Enable verbose logging",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # --- encode ---
    encode_parser = subparsers.add_parser("encode", help="Encode a spreadsheet")
    encode_parser.add_argument("excel_file", help="Path to .xlsx file")
    encode_parser.add_argument("-o", "--output", help="Output file path")
    encode_parser.add_argument(
        "-k", type=int, default=2,
        help="Neighborhood distance for anchors (default: 2)",
    )
    encode_parser.add_argument(
        "--vanilla", action="store_true",
        help="Use vanilla encoding instead of SheetCompressor",
    )

    # --- qa ---
    qa_parser = subparsers.add_parser("qa", help="Ask a question about a spreadsheet")
    qa_parser.add_argument("excel_file", help="Path to .xlsx file")
    qa_parser.add_argument("question", help="Question to answer")
    qa_parser.add_argument(
        "--provider", default=None,
        choices=["anthropic", "openai", "placeholder"],
        help="LLM provider (auto-detected from env vars if omitted)",
    )

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    if args.command == "encode":
        _cmd_encode(args)
    elif args.command == "qa":
        _cmd_qa(args)


if __name__ == "__main__":
    main()
