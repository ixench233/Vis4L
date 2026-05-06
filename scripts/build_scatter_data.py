import json
import re
from pathlib import Path

import numpy as np
from sklearn.decomposition import TruncatedSVD
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.manifold import TSNE
from sklearn.preprocessing import normalize


ROOT = Path(__file__).resolve().parents[1]
CONTENT_PATH = ROOT / "data" / "content.json"
SCATTER_PATH = ROOT / "data" / "scatter.json"


TEXT_FIELDS = [
    "title",
    "authors",
    "keywords",
    "Keywords",
    "abstract",
    "anchor_summary",
    "dataset",
    "interaction_detail",
]

CATEGORY_FIELDS = [
    "salon",
    "task",
    "modality",
    "domain",
    "model",
    "vis_encoding",
    "specific_view",
    "interaction",
    "evaluation",
    "special_tags",
]


def as_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    value = str(value).strip()
    return [value] if value else []


def clean_text(value):
    value = str(value or "")
    value = re.sub(r"<[^>]+>", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def build_document(entry):
    parts = []

    for field in TEXT_FIELDS:
        text = clean_text(entry.get(field))
        if text:
            parts.append(text)

    # Repeat structured labels so the embedding respects survey taxonomy, not
    # only abstract wording.
    for field in CATEGORY_FIELDS:
        values = as_list(entry.get(field))
        if values:
            parts.extend(values * 3)

    return " ".join(parts) or clean_text(entry.get("title")) or entry.get("id", "")


def normalize_coordinates(coords):
    coords = np.asarray(coords, dtype=float)
    coords -= coords.mean(axis=0)

    cov = np.cov(coords, rowvar=False)
    eigvals, eigvecs = np.linalg.eigh(cov)
    eigvals[eigvals < 1e-9] = 1e-9
    coords = coords @ eigvecs @ np.diag(1 / np.sqrt(eigvals))
    coords -= coords.mean(axis=0)

    scale = np.percentile(np.abs(coords), 98, axis=0)
    scale[scale == 0] = 1
    coords = coords / scale * 5

    return np.clip(coords, -6, 6)


def main():
    entries = json.loads(CONTENT_PATH.read_text(encoding="utf-8"))
    documents = [build_document(entry) for entry in entries]

    vectorizer = TfidfVectorizer(
        lowercase=True,
        stop_words="english",
        token_pattern=r"(?u)\b[\w][\w\-]+\b",
        min_df=1,
        max_df=0.85,
        ngram_range=(1, 2),
        max_features=3000,
    )
    matrix = vectorizer.fit_transform(documents)

    svd_components = min(50, matrix.shape[0] - 1, matrix.shape[1] - 1)
    if svd_components >= 2:
        dense = TruncatedSVD(n_components=svd_components, random_state=42).fit_transform(matrix)
    else:
        dense = matrix.toarray()

    dense = normalize(dense)

    perplexity = min(30, max(5, (len(entries) - 1) // 3))
    coords = TSNE(
        n_components=2,
        perplexity=perplexity,
        learning_rate="auto",
        init="pca",
        metric="cosine",
        random_state=42,
        max_iter=1500,
    ).fit_transform(dense)

    coords = normalize_coordinates(coords)

    scatter_entries = []
    for entry, coord in zip(entries, coords):
        item = {
            "id": entry.get("id", ""),
            "title": entry.get("title", ""),
            "x": round(float(coord[0]), 6),
            "y": round(float(coord[1]), 6),
        }
        for field in CATEGORY_FIELDS:
            item[field] = entry.get(field, [])
        item["year"] = entry.get("year", "")
        item["venue"] = entry.get("venue", "")
        item["authors"] = entry.get("authors", "")
        item["model"] = entry.get("model", [])
        scatter_entries.append(item)

    SCATTER_PATH.write_text(
        json.dumps(scatter_entries, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Wrote {len(scatter_entries)} entries to {SCATTER_PATH}")


if __name__ == "__main__":
    main()
