import json
import re
from collections import Counter, defaultdict
from io import BytesIO
from pathlib import Path

import openpyxl
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
XLSX_DIR = ROOT / "vis_data"
DATA_DIR = ROOT / "data"
BIBTEX_DIR = ROOT / "bibtex"
THUMBS100_DIR = ROOT / "thumbs100"
THUMBS200_DIR = ROOT / "thumbs200"


COL = {
    "title": 2,
    "screenshot": 3,
    "year": 4,
    "venue": 5,
    "rank": 6,
    "citations": 7,
    "bibtex": 8,
    "abstract": 9,
    "keywords": 10,
    "modal_count": 11,
    "modality": 12,
    "dataset": 13,
    "domain": 14,
    "salon": 15,
    "task": 16,
    "model": 17,
    "anchor_summary": 18,
    "improved_model": 19,
    "fine_tuning": 20,
    "vis_encoding": 21,
    "specific_view": 22,
    "interaction": 23,
    "interaction_detail": 24,
    "evaluation": 25,
}


CATEGORY_GROUPS = [
    ("salon", "SALON Reference Model", "SALON Reference Model"),
    ("task", "Downstream Task", "Downstream Task"),
    ("modality", "Modality Combination", "Modality Combination"),
    ("model", "Model", "Model"),
    ("vis_encoding", "Visual Encoding", "Visual Encoding"),
    ("interaction", "Interaction Technique", "Interaction Technique"),
    ("evaluation", "Evaluation Method", "Evaluation Method"),
    ("domain", "Application Domain", "Application Domain"),
    ("special_tags", "Special Tags", "Special Tags"),
]


SPLIT_RE = re.compile(r"[,;；，、\n]+")
DOI_RE = re.compile(r"\b10\.\d{4,9}/[-._;()/:A-Z0-9]+", re.I)


def clean_text(value):
    if value is None:
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\s+\n", "\n", text)
    text = re.sub(r"\n\s+", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def split_values(value):
    text = clean_text(value)
    if not text or text == "无":
        return []
    values = []
    for item in SPLIT_RE.split(text):
        item = item.strip()
        if item and item != "无":
            values.append(item)
    return values


def slugify(value, prefix):
    text = clean_text(value).lower()
    text = re.sub(r"\s+", "-", text)
    text = re.sub(r"[^0-9a-z\u4e00-\u9fff-]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-")
    if not text:
        text = "unknown"
    return f"{prefix}-{text}"


def get_bib_field(bibtex, field):
    if not bibtex:
        return ""
    pattern = re.compile(
        rf"\b{re.escape(field)}\s*=\s*(\{{(?P<braced>.*?)\}}|\"(?P<quoted>.*?)\")\s*,",
        re.I | re.S,
    )
    match = pattern.search(bibtex)
    if not match:
        return ""
    value = match.group("braced") if match.group("braced") is not None else match.group("quoted")
    return clean_text(value)


def bib_key(bibtex):
    match = re.search(r"@\w+\s*\{\s*([^,\s]+)", bibtex or "")
    return clean_text(match.group(1)) if match else ""


def infer_url(bibtex):
    url = get_bib_field(bibtex, "url")
    if url:
        return url
    doi = get_bib_field(bibtex, "doi")
    if not doi:
        match = DOI_RE.search(bibtex or "")
        doi = match.group(0) if match else ""
    doi = doi.rstrip("}. ,")
    return f"https://doi.org/{doi}" if doi else ""


def infer_special_tags(title, is_colored_title):
    if not is_colored_title:
        return []
    text = title.lower()
    if "RAG" in title or "retrieval-augmented" in text or "retrieval augmented" in text:
        return ["RAG"]
    return ["Agent"]


def is_marked_title(cell):
    color = cell.font.color
    if color is None:
        return False
    if color.type == "rgb":
        return color.rgb not in (None, "FF000000", "00000000")
    return False


def image_map(ws):
    result = {}
    for image in getattr(ws, "_images", []):
        anchor = image.anchor
        try:
            row = anchor._from.row + 1
            col = anchor._from.col + 1
        except AttributeError:
            continue
        if col != COL["screenshot"] or row in result:
            continue
        result[row] = image
    return result


def save_thumbnail(image, entry_id):
    raw = image._data()
    with Image.open(BytesIO(raw)) as source:
        source = source.convert("RGB")
        for size, directory in ((100, THUMBS100_DIR), (200, THUMBS200_DIR)):
            thumb = Image.new("RGB", (size, size), "white")
            copy = source.copy()
            copy.thumbnail((size, size), Image.Resampling.LANCZOS)
            x = (size - copy.width) // 2
            y = (size - copy.height) // 2
            thumb.paste(copy, (x, y))
            thumb.save(directory / f"{entry_id}.png")


def save_placeholder(entry_id, text):
    for size, directory in ((100, THUMBS100_DIR), (200, THUMBS200_DIR)):
        image = Image.new("RGB", (size, size), "#f4f6f8")
        draw = ImageDraw.Draw(image)
        draw.rectangle((0, 0, size - 1, size - 1), outline="#c7cdd4")
        label = "Vis4VL"
        font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), label, font=font)
        draw.text(((size - (bbox[2] - bbox[0])) / 2, size / 2 - 8), label, fill="#69717a", font=font)
        image.save(directory / f"{entry_id}.png")


def safe_write_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def make_reference(entry):
    authors = entry.get("authors") or "Unknown authors"
    title = entry["title"]
    venue = entry.get("venue") or "Unknown venue"
    year = entry.get("year") or ""
    return f"{authors}. <i>{title}</i>. {venue}, {year}."


def clear_generated_outputs():
    DATA_DIR.mkdir(exist_ok=True)
    BIBTEX_DIR.mkdir(exist_ok=True)
    THUMBS100_DIR.mkdir(exist_ok=True)
    THUMBS200_DIR.mkdir(exist_ok=True)
    for directory in (BIBTEX_DIR, THUMBS100_DIR, THUMBS200_DIR):
        for path in directory.glob("*"):
            if path.is_file():
                path.unlink()


def build_categories(rows):
    values_by_group = defaultdict(Counter)
    labels = {}
    for row in rows:
        for key, title, _description in CATEGORY_GROUPS:
            for value in row.get(key, []):
                slug = slugify(value, key)
                values_by_group[key][slug] += 1
                labels[slug] = value

    categories = []
    for key, title, description in CATEGORY_GROUPS:
        entries = []
        for slug, _count in sorted(values_by_group[key].items(), key=lambda item: labels[item[0]].lower()):
            label = labels[slug]
            entries.append({
                "type": "category-entry",
                "title": slug,
                "description": label,
                "content": label,
            })
        if entries:
            categories.append({
                "type": "category",
                "title": key,
                "description": description,
                "childrenDescription": f"{description}: ",
                "entries": entries,
            })
    return categories


def main():
    xlsx_files = sorted(XLSX_DIR.glob("*.xlsx"))
    if not xlsx_files:
        raise FileNotFoundError("No .xlsx file found in vis_data")

    workbook = openpyxl.load_workbook(xlsx_files[0], read_only=False, data_only=True)
    worksheet = workbook.active
    images = image_map(worksheet)
    clear_generated_outputs()

    rows = []
    id_counts = Counter()
    missing_images = []
    marked_counts = Counter()

    for row_number in range(3, worksheet.max_row + 1):
        title = clean_text(worksheet.cell(row_number, COL["title"]).value)
        if not title:
            continue

        bibtex = clean_text(worksheet.cell(row_number, COL["bibtex"]).value)
        key = bib_key(bibtex)
        if not key:
            key = re.sub(r"[^A-Za-z0-9]+", "", title.title())[:32] or f"paper{row_number}"
        id_counts[key] += 1
        entry_id = key if id_counts[key] == 1 else f"{key}-{id_counts[key]}"

        marked = is_marked_title(worksheet.cell(row_number, COL["title"]))
        special_tags = infer_special_tags(title, marked)
        for tag in special_tags:
            marked_counts[tag] += 1

        row = {
            "id": entry_id,
            "title": title,
            "url": infer_url(bibtex),
            "venue": clean_text(worksheet.cell(row_number, COL["venue"]).value),
            "year": int(worksheet.cell(row_number, COL["year"]).value),
            "authors": get_bib_field(bibtex, "author"),
            "keywords": clean_text(worksheet.cell(row_number, COL["keywords"]).value),
            "Keywords": clean_text(worksheet.cell(row_number, COL["keywords"]).value),
            "abstract": clean_text(worksheet.cell(row_number, COL["abstract"]).value),
            "rank": clean_text(worksheet.cell(row_number, COL["rank"]).value),
            "citations": clean_text(worksheet.cell(row_number, COL["citations"]).value),
            "modal_count": clean_text(worksheet.cell(row_number, COL["modal_count"]).value),
            "modality": split_values(worksheet.cell(row_number, COL["modality"]).value),
            "dataset": clean_text(worksheet.cell(row_number, COL["dataset"]).value),
            "domain": split_values(worksheet.cell(row_number, COL["domain"]).value),
            "salon": split_values(worksheet.cell(row_number, COL["salon"]).value),
            "task": split_values(worksheet.cell(row_number, COL["task"]).value),
            "model": split_values(worksheet.cell(row_number, COL["model"]).value),
            "anchor_summary": clean_text(worksheet.cell(row_number, COL["anchor_summary"]).value),
            "improved_model": clean_text(worksheet.cell(row_number, COL["improved_model"]).value),
            "fine_tuning": clean_text(worksheet.cell(row_number, COL["fine_tuning"]).value),
            "vis_encoding": split_values(worksheet.cell(row_number, COL["vis_encoding"]).value),
            "specific_view": clean_text(worksheet.cell(row_number, COL["specific_view"]).value),
            "interaction": split_values(worksheet.cell(row_number, COL["interaction"]).value),
            "interaction_detail": clean_text(worksheet.cell(row_number, COL["interaction_detail"]).value),
            "evaluation": split_values(worksheet.cell(row_number, COL["evaluation"]).value),
            "special_tags": special_tags,
            "is_marked": marked,
        }

        category_values = []
        for key_name, _title, _description in CATEGORY_GROUPS:
            category_values.extend(slugify(value, key_name) for value in row.get(key_name, []))
        row["categories"] = category_values
        row["reference"] = make_reference(row)

        if bibtex:
            (BIBTEX_DIR / f"{entry_id}.bib").write_text(bibtex + "\n", encoding="utf-8")
        else:
            (BIBTEX_DIR / f"{entry_id}.bib").write_text("", encoding="utf-8")

        if row_number in images:
            save_thumbnail(images[row_number], entry_id)
        else:
            save_placeholder(entry_id, title)
            missing_images.append(row_number)

        rows.append(row)

    categories = build_categories(rows)
    safe_write_json(DATA_DIR / "content.json", rows)
    safe_write_json(DATA_DIR / "categories.json", categories)

    print(f"Generated {len(rows)} entries")
    print(f"Generated {sum(len(group['entries']) for group in categories)} category entries in {len(categories)} groups")
    print(f"Special tags: {dict(marked_counts)}")
    print(f"Missing images: {missing_images}")


if __name__ == "__main__":
    main()
