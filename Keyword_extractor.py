import yake
import re

def extract_keywords_yake(text, top_n=25):
    if not text or len(text.strip()) < 100:
        return []

    text = re.sub(r"[^A-Za-z\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    extractor = yake.KeywordExtractor(
        lan="en",
        n=3,
        dedupLim=0.9,
        top=top_n * 2
    )

    keywords = extractor.extract_keywords(text)

    results = []
    seen = set()

    for kw, _ in keywords:
        kw_norm = kw.lower().strip()
        if len(kw_norm) < 3:
            continue
        if any(char.isdigit() for char in kw_norm):
            continue
        if kw_norm in seen:
            continue

        seen.add(kw_norm)
        results.append(kw)

        if len(results) == top_n:
            break

    return results


import pandas as pd
from pathlib import Path
import docx


def collect_keywords_matrix_yake(root_dir, excel_path, top_n=25):
    root = Path(root_dir)

    keyword_map = {}  # document_id → [keywords]

    for docx_file in root.rglob("*.cleaned.docx"):

        # Read document text
        doc = docx.Document(docx_file)
        text = "\n".join(
            p.text for p in doc.paragraphs if p.text.strip()
        )

        if not text.strip():
            continue

        # Extract keywords
        keywords = extract_keywords_yake(text, top_n=top_n)

        # Unique, traceable document identifier
        doc_id = docx_file.relative_to(root).as_posix()
        doc_id = doc_id.replace(".cleaned.docx", "")

        keyword_map[doc_id] = keywords

    if not keyword_map:
        print("⚠️ No keywords extracted.")
        return

    # Normalize column lengths
    max_len = max(len(v) for v in keyword_map.values())

    data = {"keywords": [""] * max_len}

    for doc_id, kws in keyword_map.items():
        padded = kws + [""] * (max_len - len(kws))
        data[doc_id] = padded

    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False)

    print(f"✅ Keyword matrix saved to: {excel_path}")





