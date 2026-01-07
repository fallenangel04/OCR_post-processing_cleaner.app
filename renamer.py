from pathlib import Path
import shutil

def collect_and_rename_cleaned_files(
    source_root: str,
    output_dir: str
):
    source_root = Path(source_root)
    output_dir = Path(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    for file_path in source_root.rglob("*"):
        if file_path.is_file() and ".cleaned" in file_path.stem:
            # Remove ".cleaned" from filename
            new_name = file_path.name.replace(".cleaned", "")
            target_path = output_dir / new_name

            # Handle name collisions
            counter = 1
            while target_path.exists():
                target_path = output_dir / f"{file_path.stem.replace('.cleaned','')}_{counter}{file_path.suffix}"
                counter += 1

            shutil.copy2(file_path, target_path)

    print("✅ All .cleaned files collected and renamed successfully.")

