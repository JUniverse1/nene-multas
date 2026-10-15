"""Busca direcciones parecidas en el CSV de fotodetección."""

import argparse
import csv
import json
import re
import sys
import unicodedata
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path


_CSV_DIRECTORY = Path(__file__).resolve().parent
_PROJECT_DIRECTORY = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else _CSV_DIRECTORY.parent
)
_CSV_SEARCH_DIRECTORIES = (
    _CSV_DIRECTORY,
    _PROJECT_DIRECTORY / "csv",
)
_CSV_CANDIDATES = (
    path
    for directory in _CSV_SEARCH_DIRECTORIES
    if directory.exists()
    for path in directory.glob("*.csv")
    if not path.name.endswith("-resultados.csv")
    and "fotodeteccion" in path.name.casefold()
)
DEFAULT_CSV = max(
    _CSV_CANDIDATES,
    key=lambda path: path.stat().st_mtime,
    default=_PROJECT_DIRECTORY / "csv" / "fotodeteccion.csv",
)
ADDRESS_COLUMN = "Dirección"

TOKEN_ALIASES = {
    "av": "avenida",
    "aven": "avenida",
    "cl": "calle",
    "c": "calle",
    "cra": "carrera",
    "cr": "carrera",
    "kr": "carrera",
    "k": "carrera",
    "dg": "diagonal",
    "diag": "diagonal",
    "tv": "transversal",
    "transv": "transversal",
}
IGNORED_TOKENS = {"con", "por", "en", "la", "el", "esquina", "a", "de"}
ROAD_TYPES = {"calle", "carrera", "avenida", "diagonal", "transversal"}
MIN_TOKEN_SIMILARITY = 0.75


def strip_accents(text):
    """Lowercase *text* and remove diacritics (``Cañón`` -> ``canon``)."""
    decomposed = unicodedata.normalize("NFKD", text.casefold())
    return "".join(
        character for character in decomposed
        if not unicodedata.combining(character)
    )


@lru_cache(maxsize=262_144)
def _normalize_address_cached(text):
    tokens = re.findall(r"[a-z]+|\d+[a-z]?", strip_accents(text))
    return tuple(TOKEN_ALIASES.get(token, token) for token in tokens)


def normalize_address(text):
    """Standardize accents, punctuation, and common Colombian road abbreviations."""
    return list(_normalize_address_cached(text))


def normalize_text(value):
    """Normalize free text for substring comparison (no road-type aliases)."""
    return " ".join(re.findall(r"[a-z0-9]+", strip_accents(value)))


def _token_similarity(query_token, address_token):
    """Similarity in [0, 1]; numbers and very short words must match exactly."""
    if query_token == address_token:
        return 1.0
    if query_token.isdigit() or address_token.isdigit():
        return 0.0
    if len(query_token) < 4 or len(address_token) < 4:
        return 0.0
    return SequenceMatcher(None, query_token, address_token).ratio()


def _street_pairs(tokens):
    """Find ``(type_index, id_index, type, identifier)`` for each street.

    A street is a road type followed by an identifier, e.g. ``calle 15``.
    """
    pairs = []
    for index, token in enumerate(tokens[:-1]):
        identifier = tokens[index + 1]
        if (
            token in ROAD_TYPES
            and identifier not in IGNORED_TOKENS
            and identifier not in ROAD_TYPES
        ):
            pairs.append((index, index + 1, token, identifier))
    return pairs


def score_address(query, address):
    """Return the percentage of meaningful query tokens matched in an address."""
    query_tokens = _normalize_address_cached(query)
    address_tokens = _normalize_address_cached(address)
    meaningful_query = [
        index for index, token in enumerate(query_tokens)
        if token not in IGNORED_TOKENS
    ]
    meaningful_address = [
        index for index, token in enumerate(address_tokens)
        if token not in IGNORED_TOKENS
    ]
    if not meaningful_query or not meaningful_address:
        return 0.0

    query_pairs = _street_pairs(query_tokens)
    address_pairs = _street_pairs(address_tokens)
    paired_query_indexes = {
        index for type_index, id_index, _, _ in query_pairs
        for index in (type_index, id_index)
    }
    used_address_indexes = set()
    matched_weight = 0.0
    total_weight = 0.0

    def token_weight(token):
        return 2.0 if any(character.isdigit() for character in token) else 1.0

    total_weight = sum(token_weight(query_tokens[index]) for index in meaningful_query)

    for _, _, road_type, identifier in query_pairs:
        best_pair = None
        best_similarity = 0.0
        for pair in address_pairs:
            address_type_index, address_id_index, address_type, address_id = pair
            if address_type != road_type or address_type_index in used_address_indexes:
                continue
            similarity = _token_similarity(identifier, address_id)
            if similarity > best_similarity:
                best_pair = pair
                best_similarity = similarity
        if best_pair is not None and best_similarity >= MIN_TOKEN_SIMILARITY:
            address_type_index, address_id_index, _, _ = best_pair
            used_address_indexes.update((address_type_index, address_id_index))
            matched_weight += (
                token_weight(road_type)
                + token_weight(identifier) * best_similarity
            )

    for query_index in meaningful_query:
        if query_index in paired_query_indexes:
            continue
        query_token = query_tokens[query_index]
        weight = token_weight(query_token)
        best_index = None
        best_similarity = 0.0
        for index in meaningful_address:
            if index in used_address_indexes:
                continue
            similarity = _token_similarity(query_token, address_tokens[index])
            if similarity > best_similarity:
                best_index = index
                best_similarity = similarity
        if best_index is not None and best_similarity >= MIN_TOKEN_SIMILARITY:
            used_address_indexes.add(best_index)
            matched_weight += weight * best_similarity

    return 100.0 * matched_weight / total_weight


def read_csv(csv_path):
    """Return ``(fieldnames, rows)`` and fail clearly if the file is unsuitable.

    Unlike :func:`load_records`, the column names survive even when the file
    has a header but no data rows.
    """
    with Path(csv_path).open(encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames or ADDRESS_COLUMN not in reader.fieldnames:
            raise ValueError(f"El CSV debe incluir la columna «{ADDRESS_COLUMN}».")
        return list(reader.fieldnames), list(reader)


def load_records(csv_path):
    """Read all original CSV rows (see :func:`read_csv`)."""
    return read_csv(csv_path)[1]


def find_matches(records, criteria):
    """Apply every ``(field, query)`` criterion and rank rows by address score.

    The address column uses fuzzy scoring; all other columns use accent- and
    case-insensitive substring matching. Returns ``(score, row_number, row)``
    tuples, best score first; score is 0.0 when no address criterion is used.
    """
    normalized_criteria = [
        (field, query, normalize_text(query))
        for field, query in criteria
    ]
    matches = []
    for row_number, record in enumerate(records):
        address_score = 0.0
        matched = True
        for field, query, normalized_query in normalized_criteria:
            if field == ADDRESS_COLUMN:
                address_score = score_address(query, record.get(field) or "")
                if address_score == 0:
                    matched = False
                    break
            elif normalized_query not in normalize_text(record.get(field) or ""):
                matched = False
                break
        if matched:
            matches.append((address_score, row_number, record))

    matches.sort(key=lambda match: (-match[0], match[1]))
    return matches


def search_addresses(query, records, limit=10):
    """Return matching original rows, ranked by address-token coverage."""
    if limit < 1:
        raise ValueError("El límite de resultados debe ser al menos 1.")

    matches = []
    for row_number, row in enumerate(records):
        address = row.get(ADDRESS_COLUMN) or ""
        if not address.strip():
            continue
        score = score_address(query, address)
        if score > 0:
            matches.append((score, row_number, row))

    matches.sort(key=lambda match: (-match[0], match[1]))
    return matches[:limit]


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Busca direcciones aproximadas y muestra todos los campos "
            "de cada fila coincidente."
        )
    )
    parser.add_argument("consulta", help='Ejemplo: "Calle 15 carrera 15"')
    parser.add_argument(
        "--csv",
        type=Path,
        default=DEFAULT_CSV,
        help=f"Archivo CSV (predeterminado: {DEFAULT_CSV.name})",
    )
    parser.add_argument(
        "--limite", type=int, default=5,
        help="Cantidad máxima de resultados (predeterminado: 5)",
    )
    parser.add_argument(
        "--json", action="store_true",
        help="Imprime los resultados completos en formato JSON",
    )
    arguments = parser.parse_args()
    if arguments.limite < 1:
        parser.error("--limite debe ser al menos 1")

    try:
        records = load_records(arguments.csv)
    except (OSError, UnicodeError, csv.Error, ValueError) as error:
        parser.error(str(error))

    matches = search_addresses(arguments.consulta, records, arguments.limite)
    if arguments.json:
        print(json.dumps(
            [
                {"coincidencia": round(score, 1), "campos": row}
                for score, _, row in matches
            ],
            ensure_ascii=False,
            indent=2,
        ))
        return

    if not matches:
        print("No se encontraron direcciones coincidentes.")
        return

    for rank, (score, _, row) in enumerate(matches, start=1):
        print(f"\n--- Resultado {rank} | coincidencia: {score:.1f}% ---")
        for field, value in row.items():
            print(f"{field}: {value or '-'}")


if __name__ == "__main__":
    main()