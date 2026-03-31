---
paths:
  - "**/*.py"
---
# Python Performance Rules

## Algorithm Complexity
- Avoid O(n^2) when O(n) or O(n log n) is possible
- Use `set` membership for O(1) lookups instead of list scans
- Use `dict` for key-value lookups, never linear search through lists of tuples

## Caching & Precomputation
- Build lookup dicts once, pass them down — don't rebuild per-call
- Cache expensive computations (style keys, merged cell maps)
- Pre-compile regex at module level, not inside functions

## Loop Optimization
- Early-exit from loops on first mismatch when checking homogeneity
- Use `any()` / `all()` with generators for short-circuit evaluation
- Avoid nested loops when a single pass + dict can achieve the same result

## Memory
- Use generators for large data iteration (`yield` instead of building lists)
- Never materialize an entire large sheet into memory — stream rows
- Use openpyxl `read_only=True` mode for profiling passes on large sheets

## Profiling
- Use `len(json.dumps(...))` as token count proxy — fast and consistent
- Track metrics incrementally, don't re-serialize entire structures
