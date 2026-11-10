#!/usr/bin/env python3
import re
import sys
import traceback
from pathlib import Path


def clean_filepath(path_str):
    """Clean surrounding quotes or spaces added by drag-and-drop."""
    cleaned = path_str.strip().strip("'\"")
    return Path(cleaned).expanduser().resolve()


def get_base_and_part(path):
    """Extract base name and part number from filenames like 'quiz_pt2.txt'."""
    match = re.search(r"^(.*)_pt(\d+)$", path.stem, re.IGNORECASE)
    if match:
        return match.group(1), int(match.group(2))
    return None, None


def process_combining(input_path):
    """Find and combine split files back into a single text file."""
    if input_path.is_dir():
        search_dir = input_path
        # Create output folder outside 'Split Questions' if dropped directly, or inside the folder
        if input_path.name.lower() == "split questions":
            outdir = input_path.parent / "Combined Quizzes"
        else:
            outdir = input_path / "Combined Quizzes"

        outdir.mkdir(parents=True, exist_ok=True)
        txt_files = [
            f for f in search_dir.iterdir()
            if f.is_file() and f.suffix.lower() == ".txt"
        ]

    elif input_path.is_file():
        search_dir = input_path.parent
        base, _ = get_base_and_part(input_path)
        if not base:
            print(
                f"\nError: Selected file '{input_path.name}' does not match split pattern (e.g., name_pt1.txt)."
            )
            return False

        outdir = search_dir
        txt_files = [
            f for f in search_dir.iterdir()
            if f.is_file() and f.suffix.lower() == ".txt"
        ]
    else:
        print("\nError: Invalid path provided.")
        return False

    # Group files by base name
    groups = {}
    for f in txt_files:
        base, part = get_base_and_part(f)
        if base is not None:
            groups.setdefault(base, []).append((part, f))

    # If user dropped a specific file, only process that file's family
    if input_path.is_file():
        target_base, _ = get_base_and_part(input_path)
        groups = {k: v for k, v in groups.items() if k == target_base}

    if not groups:
        print(f"\nNo split files (_pt1, _pt2, etc.) were found in:\n  {search_dir}")
        return False

    print(f"\nFound {len(groups)} quiz family/families to combine:\n")

    for base, parts in groups.items():
        # Sort files numerically by part number
        parts.sort(key=lambda x: x[0])

        combined_blocks = []
        for num, file_path in parts:
            try:
                content = file_path.read_text(encoding="utf-8").strip()
                if content:
                    combined_blocks.append(content)
            except Exception as e:
                print(f"Error reading {file_path.name}: {e}")

        if combined_blocks:
            out_file = outdir / f"{base}_combined.txt"
            merged_text = "\n\n".join(combined_blocks) + "\n"
            out_file.write_text(merged_text, encoding="utf-8")

            print(
                f"  -> Merged {len(parts)} part(s) into: {out_file.name}"
            )

    return True


def main():
    print("=== Quiz File Combiner ===\n")

    path_input = input(
        "Drag and drop a split file (or folder containing split files) here, then press Enter: "
    )
    if not path_input.strip():
        print("Error: No file or folder provided.")
        return

    input_path = clean_filepath(path_input)
    if not input_path.exists():
        print(f"\nError: Path not found at:\n  {input_path}")
        return

    success = process_combining(input_path)
    if success:
        print("\nSuccessfully finished combining!")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n--- ERROR OCCURRED ---")
        traceback.print_exc()
    finally:
        input("\nPress Enter to exit...")