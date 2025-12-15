import re
from typing import Dict, Iterable, List, Optional

from rapidfuzz import fuzz, process


def normalize_field_name(name: str) -> str:
    """
    Normalize field names for fuzzy matching:
    - strip
    - lowercase
    - replace underscores/hyphens with spaces
    - remove non-alphanumeric characters (except spaces)
    - collapse multiple spaces
    """
    normalized = name.strip().lower()
    normalized = re.sub(r"[_-]+", " ", normalized)
    normalized = re.sub(r"[^a-z0-9 ]+", "", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def suggest_mappings(
    asf_fields: Iterable[str],
    tape_fields: Iterable[str],
    threshold: int = 80,
) -> Dict[str, Dict[str, Optional[int]]]:
    """
    For each ASF field, find the best fuzzy match among tape_fields.
    Return a dict:
    {
        asf_field: {"source_field": <tape_field or None>, "score": <int>}
    }
    - Only accept a match if score >= threshold, otherwise source_field is None.
    - Use normalize_field_name and rapidfuzz.process.extractOne with fuzz.token_sort_ratio.
    """
    tape_fields_list: List[str] = list(tape_fields)
    normalized_tape_fields: List[str] = [
        normalize_field_name(field) for field in tape_fields_list
    ]

    results: Dict[str, Dict[str, Optional[int]]] = {}

    for asf_field in asf_fields:
        normalized_asf = normalize_field_name(asf_field)
        match = process.extractOne(
            normalized_asf,
            normalized_tape_fields,
            scorer=fuzz.token_sort_ratio,
        )

        source_field = None
        score = None

        if match:
            matched_value, matched_score, matched_index = match
            score = int(matched_score)
            if score >= threshold:
                source_field = tape_fields_list[matched_index]

        results[asf_field] = {"source_field": source_field, "score": score}

    return results
