import re
from typing import Dict, Iterable, List, Optional

from rapidfuzz import fuzz, process
import io


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
