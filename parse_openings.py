#!/usr/bin/env python3

import csv
import json
import os
import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Tuple, Optional


# Input files from lichess chess-openings repo
TSV_FILES = [
    "/workspace/chess-openings/a.tsv",
    "/workspace/chess-openings/b.tsv",
    "/workspace/chess-openings/c.tsv",
    "/workspace/chess-openings/d.tsv",
    "/workspace/chess-openings/e.tsv",
]


MOVE_NUMBER_RE = re.compile(r"^\d+\.(\.{3})?$")
RESULT_TOKENS = {"1-0", "0-1", "1/2-1/2", "*"}


def count_san_moves(pgn: str) -> int:
    """Approximate count of SAN moves in a PGN fragment without using python-chess.

    We count tokens that are not move numbers (e.g., 1. or 1...) and not results.
    This is sufficient to compare line lengths within this curated dataset.
    """
    tokens = pgn.strip().split()
    count = 0
    for tok in tokens:
        if MOVE_NUMBER_RE.match(tok):
            continue
        if tok in RESULT_TOKENS:
            continue
        count += 1
    return count


def split_name(name: str) -> Tuple[str, List[str]]:
    """Split an opening name into family and variation components.

    - Family: part before the first ':' if present; otherwise the whole name.
    - Variations: parts after ':' split by ', '. Keep exact text (including 'with ...').
    """
    if ":" in name:
        family, rest = name.split(":", 1)
        family = family.strip()
        variations = [part.strip() for part in rest.split(",")]
    else:
        family = name.strip()
        variations = []
    return family, variations


@dataclass
class OpeningNode:
    label: str
    full_name: str
    level: int  # 0 family, 1 variation, ...
    eco_codes: set = field(default_factory=set)
    canonical_pgn: Optional[str] = None
    canonical_len: Optional[int] = None
    children: Dict[str, "OpeningNode"] = field(default_factory=dict)

    def to_dict(self) -> Dict:
        return {
            "label": self.label,
            "full_name": self.full_name,
            "level": self.level,
            "eco_codes": sorted(self.eco_codes),
            "canonical_pgn": self.canonical_pgn,
            "children": [child.to_dict() for _, child in sorted(self.children.items())],
        }


def ensure_child(parent: OpeningNode, label: str) -> OpeningNode:
    if label not in parent.children:
        full_name = label if parent.level == 0 and parent.full_name == parent.label else (
            f"{parent.full_name}: {label}" if parent.full_name else label
        )
        parent.children[label] = OpeningNode(
            label=label,
            full_name=full_name,
            level=parent.level + 1,
        )
    return parent.children[label]


def update_canonical(node: OpeningNode, eco: str, pgn: str) -> None:
    move_count = count_san_moves(pgn)
    node.eco_codes.add(eco)
    if node.canonical_len is None or move_count < node.canonical_len:
        node.canonical_len = move_count
        node.canonical_pgn = pgn


def build_hierarchy(rows: List[Tuple[str, str, str]]) -> Dict[str, OpeningNode]:
    # Root map of family name to node
    families: Dict[str, OpeningNode] = {}

    # Track canonical per full name across potential multiple identical names
    canonical_by_name: Dict[str, Tuple[int, str]] = {}

    # First pass: compute canonical PGN per full name
    for eco, name, pgn in rows:
        moves = count_san_moves(pgn)
        if name not in canonical_by_name or moves < canonical_by_name[name][0]:
            canonical_by_name[name] = (moves, pgn)

    # Second pass: build tree and attach canonical to each name level when present in data
    for eco, name, pgn in rows:
        family_label, parts = split_name(name)

        # Family node
        if family_label not in families:
            families[family_label] = OpeningNode(
                label=family_label,
                full_name=family_label,
                level=0,
            )
        family_node = families[family_label]
        update_canonical(family_node, eco, canonical_by_name.get(family_label, (10**9, None))[1] or pgn)

        # Walk or create variation nodes
        current = family_node
        built_name = family_label
        for part in parts:
            child = ensure_child(current, part)
            built_name = f"{built_name}: {part}" if built_name else part
            # update canonical for this level if we have an entry for that exact name
            canon = canonical_by_name.get(built_name)
            if canon is not None:
                update_canonical(child, eco, canon[1])
            else:
                update_canonical(child, eco, pgn)
            current = child

    return families


def read_tsv_rows(file_path: str) -> List[Tuple[str, str, str]]:
    rows: List[Tuple[str, str, str]] = []
    with open(file_path, newline="") as f:
        reader = csv.DictReader(f, delimiter='\t')
        for r in reader:
            eco = (r.get("eco") or "").strip()
            name = (r.get("name") or "").strip()
            pgn = (r.get("pgn") or "").strip()
            if not eco or not name or not pgn:
                continue
            rows.append((eco, name, pgn))
    return rows


def main() -> None:
    all_rows: List[Tuple[str, str, str]] = []
    for path in TSV_FILES:
        if not os.path.exists(path):
            raise FileNotFoundError(f"Missing input file: {path}")
        all_rows.extend(read_tsv_rows(path))

    families = build_hierarchy(all_rows)

    # Write JSON hierarchy
    json_path = "/workspace/openings_hierarchy.json"
    with open(json_path, "w", encoding="utf-8") as jf:
        json.dump({k: v.to_dict() for k, v in sorted(families.items())}, jf, ensure_ascii=False, indent=2)

    # Write flat CSV
    csv_path = "/workspace/openings_flat.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as cf:
        writer = csv.writer(cf)
        writer.writerow(["full_name", "level", "family", "label", "eco_codes", "canonical_pgn"]) 
        for fam_label, fam in sorted(families.items()):
            stack = [fam]
            while stack:
                node = stack.pop()
                family_root = fam_label
                writer.writerow([
                    node.full_name,
                    node.level,
                    family_root,
                    node.label,
                    "|".join(sorted(node.eco_codes)),
                    node.canonical_pgn or "",
                ])
                # push children in reverse alphabetical to get A..Z order when popping
                for _, child in sorted(node.children.items(), reverse=True):
                    stack.append(child)

    print(f"Wrote {json_path} and {csv_path}")


if __name__ == "__main__":
    main()

