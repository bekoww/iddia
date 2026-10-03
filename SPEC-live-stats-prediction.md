# Spec: Live football statistics and explainable prediction

## Objective

Add a read-only live-match analysis flow that collects publicly available
statistics from one configured source (Nesine, SofaScore, or LiveScore),
normalizes them, and returns market probabilities with the factors that drove
each probability. The feature must never place a bet or present a result as a
guarantee.

## Scope

1. A source adapter fetches a match page and turns its statistics into a
   source-neutral match snapshot.
2. A deterministic scoring function converts snapshots into probabilities for
   `MS 1`, `MS X`, `MS 2`, `2.5 Üst`, `2.5 Alt`, `KG Var`, and `KG Yok`.
3. The UI/API consumer receives both probabilities and human-readable reasons.

The first slice will support one source only; the other sources use the same
adapter contract later. Source priority is intentionally unresolved.

## Commands

```powershell
python -m compileall -q .
python -m unittest discover -s tests -v
```

## Project structure

- `live_api.py` — existing live-data boundary.
- `live_signals.py` — existing live-signal calculations.
- `tests/` — deterministic unit tests.
- `live_match_analysis.py` — new source-neutral normalization and prediction
  logic.

## Code style

Use small, type-annotated Python functions with explicit input/output
dictionaries. Network calls stay at adapter boundaries; calculation functions
must receive ordinary Python data so they can be tested without the web.

```python
def predict(snapshot: dict) -> dict:
    return {"markets": [], "reasons": []}
```

## Testing strategy

- Fixture snapshots cover balanced and high-pressure matches.
- Unit tests assert probabilities stay in `[0, 1]`, sum to one for match
  result, and cite the expected reasons.
- Network collection is not part of unit tests.

## Boundaries

- Always: timestamp every source snapshot, retain the source URL, validate
  numeric fields, and show uncertainty.
- Ask first: adding dependencies, storing live data in a new database table,
  or changing the polling interval.
- Never: commit API keys, bypass source access controls, or automate betting.

## Success criteria

- A normalized snapshot can represent score, minute, shots on target,
  possession, corners, cards, and pre-match form when a source supplies them.
- The prediction returns each supported market with a probability and at least
  one explainable reason.
- Deterministic unit tests pass without web access.

## Open questions

1. Which source is first priority: Nesine, SofaScore, or LiveScore?
2. Is the initial release pre-match, live-match, or both?
3. What polling interval is acceptable for live analysis?
