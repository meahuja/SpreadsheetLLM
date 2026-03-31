"""Print a clean summary of the encoded JSON."""
import json, sys, os
sys.path.insert(0, os.path.dirname(__file__))

with open("messy_2000_encoded.json", "r", encoding="utf-8") as f:
    data = json.load(f)

print("="*70)
print("MESSY 2000-ROW ENCODING SUMMARY")
print("="*70)

# Metrics
m = data["compression_metrics"]["overall"]
print(f"\n--- COMPRESSION METRICS ---")
print(f"  Original tokens:        {m['original_tokens']:>10,}")
print(f"  After anchor extraction:{m['after_anchor_tokens']:>10,}  ({m['anchor_ratio']:.2f}x)")
print(f"  After inverted index:   {m['after_inverted_index_tokens']:>10,}  ({m['inverted_index_ratio']:.2f}x)")
print(f"  After format aggregation:{m['after_format_tokens']:>10,}  ({m['format_ratio']:.2f}x)")
print(f"  Final encoding:         {m['final_tokens']:>10,}  ({m['overall_ratio']:.2f}x)")

for sheet_name, sheet_data in data["sheets"].items():
    print(f"\n--- SHEET: {sheet_name} ---")

    anchors = sheet_data["structural_anchors"]
    print(f"  Anchor rows: {len(anchors['rows'])} (first 10: {anchors['rows'][:10]}...)")
    print(f"  Anchor cols: {anchors['columns']}")

    cells = sheet_data["cells"]
    print(f"\n  CELLS (inverted index): {len(cells)} unique values")

    # Show formulas
    formulas = {k: v for k, v in cells.items() if k.startswith("=")}
    print(f"  Formulas preserved: {len(formulas)}")
    print(f"  Sample formulas:")
    for i, (val, refs) in enumerate(formulas.items()):
        if i >= 10: break
        refs_str = str(refs[:3]) + ("..." if len(refs) > 3 else "")
        print(f"    {val:40s} -> {refs_str}")

    # Show non-formula values
    non_formulas = {k: v for k, v in cells.items() if not k.startswith("=")}
    print(f"\n  Non-formula values: {len(non_formulas)}")
    print(f"  Sample values (with merged ranges):")
    range_examples = [(k, v) for k, v in non_formulas.items() if any(":" in r for r in v)]
    for i, (val, refs) in enumerate(range_examples[:8]):
        print(f"    {val:40s} -> {refs}")

    single_examples = [(k, v) for k, v in non_formulas.items() if all(":" not in r for r in v)]
    print(f"  Single-cell values: {len(single_examples)}")
    for i, (val, refs) in enumerate(single_examples[:5]):
        print(f"    {val:40s} -> {refs[:3]}{'...' if len(refs)>3 else ''}")

    # Show formats
    fmts = sheet_data["formats"]
    print(f"\n  FORMATS (semantic type aggregation): {len(fmts)} groups")
    for fmt_key, refs in fmts.items():
        parsed = json.loads(fmt_key)
        type_name = parsed.get("type", "?")
        nfs = parsed.get("nfs", "?")
        print(f"    {type_name:15s} nfs={nfs:25s} -> {len(refs)} region(s): {refs[:3]}{'...' if len(refs)>3 else ''}")

    # Numeric ranges
    nr = sheet_data["numeric_ranges"]
    if nr:
        print(f"\n  NUMERIC RANGES: {len(nr)} groups")
        for k, v in nr.items():
            parsed = json.loads(k)
            print(f"    {parsed.get('type','?'):15s} -> {v[:3]}{'...' if len(v)>3 else ''}")

print(f"\n--- JSON FILE SIZE ---")
file_size = os.path.getsize("messy_2000_encoded.json")
print(f"  File: messy_2000_encoded.json")
print(f"  Size: {file_size:,} bytes ({file_size/1024:.1f} KB)")
