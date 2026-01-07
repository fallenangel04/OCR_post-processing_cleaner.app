import streamlit as st
from pathlib import Path
import tempfile
import shutil
import zipfile

from english_cleaner import process_docx_file as process_english_file
from english_cleaner import process_folder as process_english_folder

from urdu_cleaner import process_docx_file as process_urdu_file
from urdu_cleaner import process_folder as process_urdu_folder

from Keyword_extractor import collect_keywords_matrix_yake as collect_keyword_matrix_yake
from renamer import collect_and_rename_cleaned_files



# -------------------- Page Config --------------------
st.set_page_config(
    page_title="OCR Cleaner",
    layout="wide"
)

st.title("📄 OCR Cleaner & Keyword Extractor")
st.caption("English OCR • Urdu OCR • YAKE Keyword Matrix")


# -------------------- Mode Selection --------------------
mode = st.selectbox(
    "Select task",
    [
        "English OCR Cleaning",
        "Urdu OCR Cleaning",
        "Keyword Extraction (YAKE)",
        "Collect & Rename Cleaned Files"
    ]
)


# -------------------- Upload --------------------
uploaded = st.file_uploader(
    "Upload a DOCX file or a ZIP folder",
    type=["docx", "zip"]
)

if uploaded is None:
    st.info("Upload a file or ZIP folder to continue.")
    st.stop()


# -------------------- Temp Workspace --------------------
workspace = Path(tempfile.mkdtemp())
input_path = workspace / uploaded.name

with open(input_path, "wb") as f:
    f.write(uploaded.read())

output_dir = workspace / "output"
output_dir.mkdir(exist_ok=True)


# -------------------- Run --------------------
zip_root = None
if st.button("▶ Run"):

    with st.spinner("Processing..."):

        # ---------- ZIP (Folder) ----------
        if input_path.suffix == ".zip":
            extract_dir = workspace / "input_folder"
            extract_dir.mkdir()

            with zipfile.ZipFile(input_path, "r") as z:
                z.extractall(extract_dir)

            entries = list(extract_dir.iterdir())
            if len(entries) == 1 and entries[0].is_dir():
                extract_dir = entries[0]

            zip_root = extract_dir


            if mode == "English OCR Cleaning":
                process_english_folder(extract_dir, extract_dir)

            elif mode == "Urdu OCR Cleaning":
                process_urdu_folder(extract_dir, extract_dir)

            elif mode == "Keyword Extraction (YAKE)":
                collect_keyword_matrix_yake(
                    root_dir=extract_dir,
                    excel_path=output_dir / "keywords.xlsx"
                )

            elif mode == "Collect & Rename Cleaned Files":
                collect_and_rename_cleaned_files(
                    source_root=extract_dir,
                    output_dir=output_dir
                )

        # ---------- Single DOCX ----------
        else:
            extract_dir = workspace / "single_file"
            extract_dir.mkdir()

            # copy original
            shutil.copy2(input_path, extract_dir / input_path.name)
            zip_root = extract_dir

            output_file = extract_dir / f"{input_path.stem}.cleaned.docx"

            if mode == "English OCR Cleaning":
                process_english_file(input_path, output_file)

            elif mode == "Urdu OCR Cleaning":
                process_urdu_file(input_path, output_file)

            else:
                st.error("Keyword extraction requires a folder (ZIP).")
                st.stop()

            
# -------------------- Download --------------------

if mode == "Keyword Extraction (YAKE)":
    excel_file = output_dir / "keywords.xlsx"

    if not excel_file.exists():
        st.error("Keyword extraction failed: Excel file not found.")
        st.stop()

    with open(excel_file, "rb") as f:
        st.download_button(
            "⬇ Download keywords.xlsx",
            data=f,
            file_name="keywords.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif mode == "Collect & Rename Cleaned Files":
    if not any(output_dir.glob("*.docx")):
        st.error("No cleaned files found to rename.")
        st.stop()

    zip_out = workspace / "renamed_cleaned_files.zip"
    with zipfile.ZipFile(zip_out, "w") as z:
        for p in output_dir.glob("*.docx"):
            z.write(p, arcname=p.name)

    with open(zip_out, "rb") as f:
        st.download_button(
            "⬇ Download renamed cleaned files",
            data=f,
            file_name="renamed_cleaned_files.zip",
            mime="application/zip"
        )

else:
    # English / Urdu OCR cleaning
    zip_out = workspace / "ocr_results.zip"
    with zipfile.ZipFile(zip_out, "w") as z:
        for p in zip_root.rglob("*"):
            if p.is_file():
                z.write(p, arcname=p.relative_to(zip_root))

    with open(zip_out, "rb") as f:
        st.download_button(
            "⬇ Download results",
            data=f,
            file_name="ocr_results.zip",
            mime="application/zip"
        )
