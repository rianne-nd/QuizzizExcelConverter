#!/usr/bin/env python3
import re
import subprocess
import sys
import traceback
from pathlib import Path


def parse_blocks(text):
    """Split the text into question blocks, preserving internal newlines."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    raw_blocks = re.split(r"\n\s*\n+", text.strip())
    return [b.strip() for b in raw_blocks if b.strip()]


def chunk(items, size):
    """Split a list into chunks of length `size`."""
    return [items[i : i + size] for i in range(0, len(items), size)]


def make_output_name(input_path, part, outdir):
    """quiz.txt + part 1 -> outdir/quiz_pt1.txt"""
    return outdir / f"{input_path.stem}_pt{part}{input_path.suffix}"


def clean_filepath(path_str):
    """Clean surrounding quotes or spaces added by drag-and-drop."""
    cleaned = path_str.strip().strip("'\"")
    return Path(cleaned).expanduser().resolve()


def run_converter_cmd():
    """Locate and run 'Convert My Quizzes.cmd' if present."""
    script_dir = Path(__file__).parent
    cmd_path = script_dir / "Convert My Quizzes.cmd"

    if not cmd_path.is_file():
        cmd_path = Path.cwd() / "Convert My Quizzes.cmd"

    if cmd_path.is_file():
        print(f"\n---> Launching: {cmd_path.name}\n")
        try:
            subprocess.run(f'"{cmd_path}"', shell=True, check=True)
        except Exception as e:
            print(f"\nError running CMD script: {e}")
    else:
        print(
            f"\nError: Could not find 'Convert My Quizzes.cmd' in:\n  {script_dir}"
        )


def process_file(input_path, size):
    """Process a single quiz .txt file and write its split parts."""
    outdir = input_path.parent

    try:
        text = input_path.read_text(encoding="utf-8")
    except Exception as e:
        print(f"\nError reading {input_path.name}: {e}")
        return False

    blocks = parse_blocks(text)

    if not blocks:
        print(f"\nSkipping {input_path.name}: No valid question blocks found.")
        return False

    chunks = chunk(blocks, size)
    print(
        f"\nFile: {input_path.name}"
        f"\nFound {len(blocks)} question(s). Creating {len(chunks)} file(s) (~{size} questions each):"
    )

    for i, piece in enumerate(chunks, start=1):
        out_path = make_output_name(input_path, i, outdir)
        out_text = "\n\n".join(piece) + "\n"
        out_path.write_text(out_text, encoding="utf-8")
        print(f"  -> Created: {out_path.name} ({len(piece)} questions)")

    return True


def main():
    print("=== Quiz File Splitter ===\n")

    # Prompt 1: Number of questions
    size_input = input("How many questions per split file? [Default 10]: ").strip()
    if not size_input:
        size = 10
    else:
        try:
            size = int(size_input)
            if size < 1:
                raise ValueError
        except ValueError:
            print("Error: Please enter a valid number greater than 0.")
            return

    # Prompt 2: Drag and drop file or folder
    path_input = input(
        "\nDrag and drop a .txt file OR a folder here, then press Enter: "
    )
    if not path_input.strip():
        print("Error: No file or folder provided.")
        return

    input_path = clean_filepath(path_input)
    if not input_path.exists():
        print(f"\nError: Path not found at:\n  {input_path}")
        return

    files_processed = 0

    if input_path.is_dir():
        # Find all .txt files in folder (excluding already split _pt files)
        txt_files = [
            f
            for f in sorted(input_path.glob("*.txt"))
            if not re.search(r"_pt\d+$", f.stem)
        ]

        if not txt_files:
            print(f"\nError: No valid .txt files found in folder:\n  {input_path}")
            return

        print(f"\nFound {len(txt_files)} file(s) to process...")
        for file in txt_files:
            if process_file(file, size):
                files_processed += 1

    elif input_path.is_file():
        if process_file(input_path, size):
            files_processed += 1
    else:
        print("\nError: Specified path is neither a file nor a folder.")
        return

    if files_processed > 0:
        print("\nSuccessfully finished splitting!")
    else:
        print("\nNo files were split.")

    # Prompt 3: Ask to run Convert My Quizzes.cmd
    run_cmd = (
        input(
            "\nDo you also want to run 'Convert My Quizzes.cmd'? (y/n) [n]: "
        )
        .strip()
        .lower()
    )
    if run_cmd in ["y", "yes"]:
        run_converter_cmd()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n--- ERROR OCCURRED ---")
        traceback.print_exc()
    finally:
        input("\nPress Enter to exit...")