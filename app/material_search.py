"""Multi-field fuzzy matching for material/substitute search.

Replaces a single contiguous-substring match with token-based, accent-folded,
multi-field ranking so an operator can type "mã + tên" (and HS) together and
narrow candidates the way they expect — e.g. "dientro chip" matches a row whose
code contains DIENTRO and whose name contains "chip", in any order.
"""
from __future__ import annotations

import unicodedata
from collections.abc import Callable, Sequence


def fold_text(value) -> str:
    """Casefold + strip Vietnamese diacritics so search ignores accents."""
    text = str(value or "")
    decomposed = unicodedata.normalize("NFD", text)
    stripped = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    stripped = stripped.replace("đ", "d").replace("Đ", "D")
    return stripped.casefold().strip()


# Subsequence (typo/abbreviation) matching only applies to short, space-free
# fields like codes/HS. Allowing it over long descriptions scatter-matches —
# e.g. "aptomat" or "5000" would hit an unrelated IC whose description happens
# to contain those characters in order.
_SUBSEQ_MAX_FIELD_LEN = 24


def _is_subsequence(token: str, field: str) -> bool:
    it = iter(field)
    return all(ch in it for ch in token)


def _subseq_eligible(field: str) -> bool:
    return " " not in field and len(field) <= _SUBSEQ_MAX_FIELD_LEN


def _token_allows_subseq(token: str) -> bool:
    # Only alphabetic tokens of length >= 3 may match by subsequence (typo'd
    # codes like 'bientn' -> BIENTAN). A numeric token like '5000' must match a
    # real digit run; otherwise it scatter-matches HS codes ('85044090' ->
    # 5,0,0,0) and any digit-rich field.
    return len(token) >= 3 and any(ch.isalpha() for ch in token)


def _token_score(token: str, folded_fields: Sequence[str]) -> float:
    best = 0.0
    field_count = len(folded_fields)
    allows_subseq = _token_allows_subseq(token)
    for idx, field in enumerate(folded_fields):
        if not field:
            continue
        if token == field:
            hit = 4.0
        elif field.startswith(token):
            hit = 3.0
        elif token in field:
            hit = 1.5
        elif allows_subseq and _subseq_eligible(field) and _is_subsequence(token, field):
            hit = 0.6
        else:
            hit = 0.0
        if hit:
            # Earlier fields (code first) carry slightly more weight so an
            # exact code hit outranks an incidental name hit.
            hit += max(0, field_count - idx) * 0.1
            best = max(best, hit)
    return best


def match_score(query: str, fields: Sequence) -> float | None:
    """Return a ranking score, or None when the row should be excluded.

    Every query token must match at least one field (AND across tokens, OR
    across fields). Matching is accent-insensitive and tiered:
    exact > prefix > substring > subsequence (typo/abbreviation tolerance).
    """
    tokens = fold_text(query).split()
    if not tokens:
        return None
    folded_fields = [fold_text(field) for field in fields]
    total = 0.0
    for token in tokens:
        token_best = _token_score(token, folded_fields)
        if token_best == 0.0:
            return None
        total += token_best
    return total


def rank_matches(
    query: str,
    rows: Sequence,
    fields_of: Callable[[object], Sequence],
    limit: int = 20,
) -> list:
    """Filter + rank rows by match_score, best first, capped at limit."""
    capped = max(1, min(limit, 100))
    if not (query or "").strip():
        return list(rows)[:capped]
    scored: list[tuple[float, int, object]] = []
    for position, row in enumerate(rows):
        score = match_score(query, fields_of(row))
        if score is not None:
            # position keeps the sort stable for equal scores.
            scored.append((score, position, row))
    scored.sort(key=lambda item: (-item[0], item[1]))
    return [row for _, _, row in scored[:capped]]
