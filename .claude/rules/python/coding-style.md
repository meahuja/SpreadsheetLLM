---
paths:
  - "**/*.py"
  - "**/*.pyi"
---
# Python Coding Style

## Standards
- Follow **PEP 8** conventions
- Use **type annotations** on all function signatures
- Pre-compile all regex patterns at module level (`re.compile(...)`)
- Use `from __future__ import annotations` for forward references

## Immutability
- Prefer `tuple` over `list` for fixed-size data
- Use `frozenset` for hashable set constants
- Use `@dataclass(frozen=True)` for immutable data containers

## Naming
- Functions: `snake_case` with verb-noun pattern (`detect_semantic_type`, `merge_cell_ranges`)
- Private helpers: prefix with `_` (`_build_merged_cell_map`)
- Constants: `UPPER_SNAKE_CASE` (`EMAIL_REGEX`, `MAX_CANDIDATES`)
- Type aliases: `PascalCase` (`CellCoord`, `StyleFingerprint`)

## Formatting
- Max line length: 99 characters
- Use f-strings for string interpolation
- Imports: stdlib → third-party → local, separated by blank lines
