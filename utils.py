import io
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Union

import yaml
from rapidfuzz import fuzz, process


def preprocess_value(asf_field, value):
    """
    Placeholder hook to adjust values before writing into ASF.
    For now, just return the value unchanged.
    Later we can add special handling for dates, percentages, etc.
    """

    return value


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
    overrides: Optional[Dict[str, Iterable[str]]] = None,
) -> Dict[str, Dict[str, Optional[int]]]:
    """
    For each ASF field, find the best fuzzy match among tape_fields.
    Return a dict:
    {
        asf_field: {"source_field": <tape_field or None>, "score": <int>}
    }
    - Only accept a match if score >= threshold, otherwise source_field is None.
    - Use normalize_field_name and rapidfuzz.process.extractOne with fuzz.token_sort_ratio.
    - If overrides are provided, they should map ASF field names to alternative
      labels to try during matching.
    """
    tape_fields_list: List[str] = list(tape_fields)
    normalized_tape_fields: List[str] = [
        normalize_field_name(field) for field in tape_fields_list
    ]

    results: Dict[str, Dict[str, Optional[int]]] = {}
    override_mapping: Dict[str, Iterable[str]] = overrides or {}
    normalized_overrides: Dict[str, Iterable[str]] = {
        normalize_field_name(key): value for key, value in override_mapping.items()
    }

    for asf_field in asf_fields:
        normalized_key = normalize_field_name(asf_field)
        candidate_labels = [asf_field] + list(
            normalized_overrides.get(normalized_key, [])
        )
        best_match = None

        for candidate in candidate_labels:
            normalized_candidate = normalize_field_name(candidate)
            match = process.extractOne(
                normalized_candidate,
                normalized_tape_fields,
                scorer=fuzz.token_sort_ratio,
            )

            if not match:
                continue

            matched_value, matched_score, matched_index = match
            score = int(matched_score)

            if best_match is None or score > best_match[1]:
                best_match = (matched_index, score)

        source_field = None
        score = None

        if best_match:
            matched_index, matched_score = best_match
            score = matched_score
            if score >= threshold:
                source_field = tape_fields_list[matched_index]

        results[asf_field] = {"source_field": source_field, "score": score}

    return results


def load_override_mapping(file: Union[str, Path, io.BytesIO, io.BufferedReader]) -> Dict[str, List[str]]:
    """
    Load a YAML mapping override file and normalize it to
    {ASF_field: [alias1, alias2, ...]}.

    Unsupported or malformed files raise ValueError.
    """

    def _load_from_bytes(raw_bytes: bytes, filename: str) -> Dict[str, List[str]]:
        if not filename.endswith((".yaml", ".yml")):
            raise ValueError("Override file must be .yaml or .yml")

        try:
            data = yaml.safe_load(raw_bytes)
        except yaml.YAMLError as exc:
            error_location = ""
            if hasattr(exc, "problem_mark") and exc.problem_mark is not None:
                mark = exc.problem_mark
                error_location = f" (line {mark.line + 1}, column {mark.column + 1})"
            raise ValueError(f"Invalid YAML in {filename}{error_location}: {exc}") from exc

        if not isinstance(data, dict):
            raise ValueError("Override file must contain a mapping of ASF fields")

        normalized: Dict[str, List[str]] = {}

        for key, value in data.items():
            if value is None:
                continue

            if isinstance(value, (list, tuple, set)):
                aliases = [str(item) for item in value]
            else:
                aliases = [str(value)]

            normalized[str(key)] = aliases

        return normalized

    if isinstance(file, (str, Path)):
        path = Path(file)
        if not path.exists():
            raise FileNotFoundError(path)
        return _load_from_bytes(path.read_bytes(), path.name.lower())

    if hasattr(file, "getvalue"):
        filename = str(getattr(file, "name", "override.yaml")).lower()
        return _load_from_bytes(file.getvalue(), filename)

    if hasattr(file, "read"):
        raw_bytes = file.read()
        if hasattr(file, "seek"):
            file.seek(0)
        filename = str(getattr(file, "name", "override.yaml")).lower()
        return _load_from_bytes(raw_bytes, filename)

    raise ValueError("Unsupported override file type")


def build_column_index_by_field(ws, header_row: int = 1) -> Dict[str, int]:
    """
    Return dict mapping ASF field name -> column index in the worksheet.
    Field name is str(cell.value).strip().
    """

    column_index: Dict[str, int] = {}

    for cell in ws[header_row]:
        if cell.value is None:
            continue

        field_name = str(cell.value).strip()
        if not field_name:
            continue

        column_index[field_name] = cell.column

    return column_index


def write_loan_data_to_asf(ws, start_row: int, asf_fields: Iterable[str], df, mapping):
    """
    ws: ASF worksheet
    start_row: first ASF data row (e.g., 2)
    asf_fields: ordered list of ASF header names
    df: pandas DataFrame of source tape
    mapping: dict {asf_field: {"source_field": <str or None>, "score": <int>}}
    Logic:
    - Build column index mapping using build_column_index_by_field.
    - For each row in df:
      - For each asf_field in asf_fields:
        - Get mapped source_field; skip if None or "(unmapped)".
        - Read value = df_row[source_field].
        - Write value to target cell at (excel_row, col_index).
    - Do NOT change number_format or styles of the target cells.
    """

    column_index = build_column_index_by_field(ws)

    for row_offset, (_, df_row) in enumerate(df.iterrows()):
        excel_row = start_row + row_offset

        for asf_field in asf_fields:
            mapping_info = mapping.get(asf_field, {})
            source_field = mapping_info.get("source_field")

            if source_field in (None, "(unmapped)"):
                continue

            if asf_field not in column_index:
                continue

            value = df_row.get(source_field)
            value = preprocess_value(asf_field, value)

            target_cell = ws.cell(row=excel_row, column=column_index[asf_field])
            target_cell.value = value


def build_asf_output_stream(wb):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
