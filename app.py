import streamlit as st
from pathlib import Path
import tempfile
import shutil
import zipfile

from english_cleaner import process_docx_file as process_english_file
from english_cleaner import process_folder as process_english_folder

from urdu_cleaner import process_docx_file as process_urdu_file
from urdu_cleaner import process_folder as process_urdu_folder

from Keyword_extractor import collect_keywords_matrix_yake
from renamer import collect_and_rename_cleaned_files


# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="OCR Post-Processing Cleaner",
    page_icon="📄",
    layout="centered"
)

st.title("📄 OCR Post-Processing Cleaner")
st.subheader("English & Urdu OCR Cleaning • Keyword Extraction")
st.caption("Upload DOCX or ZIP files. Download clean, reproducible outputs.")
st.divider()


# -------------------------------------------------
# TABS (BEST UI)
# -------------------------------------------------
tab_en, tab_ur, tab_yake, tab_rename = st.tabs([
    "🧹 English OCR Cleaning",
    "🧹 Urdu OCR Cleaning",
    "🔑 Keyword Extraction (YAKE)",
    "🗂 Collect & Rename Cleaned Files"
])


# =================================================
# 🔹 ENGLISH OCR CLEANING
# =================================================
with tab_en:
    st.markdown("### English OCR Cleaning")
    st.info("Upload a DOCX or ZIP. Output will contain originals + `.cleaned.docx` files.")

    uploaded = st.file_uploader(
        "Upload DOCX or ZIP",
        type=["docx", "zip"],
        key="en_upload"
    )

    if uploaded and st.button("Run English OCR Cleaning"):
        workspace = Path(tempfile.mkdtemp())
        input_path = workspace / uploaded.name
        input_path.write_bytes(uploaded.read())

        if input_path.suffix == ".zip":
            extract_dir = workspace / "input"
            extract_dir.mkdir()

            with zipfile.ZipFile(input_path, "r") as z:
                z.extractall(extract_dir)

            entries = list(extract_dir.iterdir())
            if len(entries) == 1 and entries[0].is_dir():
                extract_dir = entries[0]

            # ARCHIVE MODE
            process_english_folder(extract_dir, extract_dir)
            zip_root = extract_dir

        else:
            extract_dir = workspace / "single"
            extract_dir.mkdir()

            shutil.copy2(input_path, extract_dir / input_path.name)
            output_file = extract_dir / f"{input_path.stem}.cleaned.docx"
            process_english_file(input_path, output_file)
            zip_root = extract_dir

        zip_out = workspace / "english_ocr_results.zip"
        with zipfile.ZipFile(zip_out, "w") as z:
            for p in zip_root.rglob("*"):
                if p.is_file():
                    z.write(p, arcname=p.relative_to(zip_root))

        st.download_button(
            "⬇ Download English OCR Results",
            data=zip_out.read_bytes(),
            file_name="english_ocr_results.zip",
            mime="application/zip"
        )


# =================================================
# 🔹 URDU OCR CLEANING
# =================================================
with tab_ur:
    st.markdown("### Urdu OCR Cleaning")
    st.info("Upload a DOCX or ZIP. Output will contain originals + `.cleaned.docx` files.")

    uploaded = st.file_uploader(
        "Upload DOCX or ZIP",
        type=["docx", "zip"],
        key="ur_upload"
    )

    if uploaded and st.button("Run Urdu OCR Cleaning"):
        workspace = Path(tempfile.mkdtemp())
        input_path = workspace / uploaded.name
        input_path.write_bytes(uploaded.read())

        if input_path.suffix == ".zip":
            extract_dir = workspace / "input"
            extract_dir.mkdir()

            with zipfile.ZipFile(input_path, "r") as z:
                z.extractall(extract_dir)

            entries = list(extract_dir.iterdir())
            if len(entries) == 1 and entries[0].is_dir():
                extract_dir = entries[0]

            process_urdu_folder(extract_dir, extract_dir)
            zip_root = extract_dir

        else:
            extract_dir = workspace / "single"
            extract_dir.mkdir()

            shutil.copy2(input_path, extract_dir / input_path.name)
            output_file = extract_dir / f"{input_path.stem}.cleaned.docx"
            process_urdu_file(input_path, output_file)
            zip_root = extract_dir

        zip_out = workspace / "urdu_ocr_results.zip"
        with zipfile.ZipFile(zip_out, "w") as z:
            for p in zip_root.rglob("*"):
                if p.is_file():
                    z.write(p, arcname=p.relative_to(zip_root))

        st.download_button(
            "⬇ Download Urdu OCR Results",
            data=zip_out.read_bytes(),
            file_name="urdu_ocr_results.zip",
            mime="application/zip"
        )


# =================================================
# 🔹 KEYWORD EXTRACTION (YAKE)
# =================================================
with tab_yake:
    st.markdown("### Keyword Extraction (YAKE)")
    st.warning("Upload a ZIP containing `.cleaned.docx` files only.")

    uploaded = st.file_uploader(
        "Upload ZIP",
        type=["zip"],
        key="yake_upload"
    )

    if uploaded and st.button("Extract Keywords"):
        workspace = Path(tempfile.mkdtemp())
        input_path = workspace / uploaded.name
        input_path.write_bytes(uploaded.read())

        extract_dir = workspace / "input"
        extract_dir.mkdir()

        with zipfile.ZipFile(input_path, "r") as z:
            z.extractall(extract_dir)

        entries = list(extract_dir.iterdir())
        if len(entries) == 1 and entries[0].is_dir():
            extract_dir = entries[0]

        output_dir = workspace / "output"
        output_dir.mkdir()

        collect_keywords_matrix_yake(
            root_dir=extract_dir,
            excel_path=output_dir / "keywords.xlsx"
        )

        excel_file = output_dir / "keywords.xlsx"
        if excel_file.exists():
            st.download_button(
                "⬇ Download keywords.xlsx",
                data=excel_file.read_bytes(),
                file_name="keywords.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Keyword extraction failed.")


# =================================================
# 🔹 COLLECT & RENAME CLEANED FILES
# =================================================
with tab_rename:
    st.markdown("### Collect & Rename Cleaned Files")
    st.info("Outputs a ZIP containing ONLY renamed `.cleaned.docx` files.")

    uploaded = st.file_uploader(
        "Upload ZIP",
        type=["zip"],
        key="rename_upload"
    )

    if uploaded and st.button("Collect & Rename"):
        workspace = Path(tempfile.mkdtemp())
        input_path = workspace / uploaded.name
        input_path.write_bytes(uploaded.read())

        extract_dir = workspace / "input"
        extract_dir.mkdir()

        with zipfile.ZipFile(input_path, "r") as z:
            z.extractall(extract_dir)

        entries = list(extract_dir.iterdir())
        if len(entries) == 1 and entries[0].is_dir():
            extract_dir = entries[0]

        output_dir = workspace / "output"
        output_dir.mkdir()

        collect_and_rename_cleaned_files(
            source_root=extract_dir,
            output_dir=output_dir
        )

        zip_out = workspace / "renamed_cleaned_files.zip"
        with zipfile.ZipFile(zip_out, "w") as z:
            for p in output_dir.glob("*.docx"):
                z.write(p, arcname=p.name)

        st.download_button(
            "⬇ Download renamed cleaned files",
            data=zip_out.read_bytes(),
            file_name="renamed_cleaned_files.zip",
            mime="application/zip"
        )
