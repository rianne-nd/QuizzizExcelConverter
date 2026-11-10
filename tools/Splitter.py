#!/usr/bin/env python3
"""Quiz File Splitter.

Splits a quiz .txt into parts for Quizizz / Wayground.

WHY THIS WAS REWRITTEN
The first version split the file on blank lines:

    re.split(r"\\n\\s*\\n+", text)

That treats every blank-line-separated chunk as a "question". But a question is
often more than one chunk. A source like this:

    Question: ...
    Option: ...
    Explanation: Topic 4.4.5
    <blank>
    Type: Extended ACLs are required for this scenario because ...
    <blank>
    Placement: According to network design guidelines, ...

counts as THREE blocks for ONE question. On Checkpoint Exam 3-5 that produced
84 blocks for 74 questions - 10 phantom "questions". So files came out as
10/10/9/7/10/9/8/8/3 questions instead of ten each, and a chunk boundary could
land in the middle of a question, putting the question in one file and its
explanation in the next.

This version finds questions by their "Question:" line, so a question is one
unit no matter how many paragraphs it spans.
"""

import re
import subprocess
import traceback
from pathlib import Path

# A question starts at a line beginning with this.
QUESTION_RE = re.compile(r"^Question:", re.IGNORECASE)


def split_into_questions(text):
    """Return (questions, preamble).

    A question runs from its "Question:" line up to the next "Question:" line,
    so blank lines inside a question are kept as part of it.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    questions = []
    preamble = []
    current = None

    for line in text.split("\n"):
        if QUESTION_RE.match(line.strip()):
            if current is not None:
                questions.append("\n".join(current).strip())
            current = [line]
        elif current is None:
            if line.strip():
                preamble.append(line)
        else:
            current.append(line)

    if current is not None:
        questions.append("\n".join(current).strip())

    return [q for q in questions if q], preamble


def parse_blocks(text):
    """Fallback for files with no "Question:" lines: split on blank lines."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    raw_blocks = re.split(r"\n\s*\n+", text.strip())
    return [b.strip() for b in raw_blocks if b.strip()]


def chunk_with_tail(items, size, tail_threshold):
    """Chunk into `size`, but if what is left would be `tail_threshold` or
    fewer, put all of it in the final file instead of splitting again.

    74 questions, size 10, threshold 15 -> 10,10,10,10,10,10,14   (7 files)
    without the rule it would be        -> 10 x7, 4              (8 files)
    """
    chunks = []
    i = 0
    n = len(items)
    while i < n:
        remaining = n - i
        if remaining <= tail_threshold:
            chunks.append(items[i:])
            break
        chunks.append(items[i : i + size])
        i += size
    return chunks


def make_output_name(input_path, part, outdir):
    """quiz.txt + part 1 -> outdir/quiz_pt1.txt"""
    return outdir / f"{input_path.stem}_pt{part}{input_path.suffix}"


def clean_filepath(path_str):
    """Clean surrounding quotes or spaces added by drag-and-drop."""
    cleaned = path_str.strip().strip("'\"")
    return Path(cleaned).expanduser().resolve()


def run_converter_cmd():
    """Locate and run 'Convert My Quizzes.cmd' if present.

    Splitter.py lives in tools\\, and the launcher sits in the folder one level
    up, so look there first, then beside this script, then the working dir.
    """
    here = Path(__file__).parent
    cmd_path = None
    for candidate in (
        here.parent / "Convert My Quizzes.cmd",
        here / "Convert My Quizzes.cmd",
        Path.cwd() / "Convert My Quizzes.cmd",
    ):
        if candidate.is_file():
            cmd_path = candidate
            break

    if cmd_path is not None:
        print(f"\n---> Launching: {cmd_path.name}\n")
        try:
            subprocess.run(f'"{cmd_path}"', shell=True, check=True)
        except Exception as e:
            print(f"\nError running CMD script: {e}")
    else:
        print(
            f"\nError: Could not find 'Convert My Quizzes.cmd' in:\n"
            f"  {here.parent}\n  {here}"
        )


def process_file(input_path, size, tail_threshold):
    """Process a single quiz .txt file and write its split parts."""
    outdir = input_path.parent

    try:
        text = input_path.read_text(encoding="utf-8")
    except Exception as e:
        print(f"\nError reading {input_path.name}: {e}")
        return False

    questions, preamble = split_into_questions(text)

    if not questions:
        # No "Question:" lines - fall back to the old blank-line behaviour.
        blocks = parse_blocks(text)
        if not blocks:
            print(f"\nSkipping {input_path.name}: No valid question blocks found.")
            return False
        print(
            f"\n  NOTE: {input_path.name} has no 'Question:' lines - "
            f"falling back to blank-line splitting ({len(blocks)} blocks)."
        )
        questions = blocks

    if preamble:
        print(
            f"\n  NOTE: {len(preamble)} line(s) before the first 'Question:' "
            f"line were ignored, e.g.: {preamble[0][:70]}"
        )

    chunks = chunk_with_tail(questions, size, tail_threshold)

    if len(chunks) == 1 and len(questions) <= size:
        print(
            f"\n  NOTE: {input_path.name} has only {len(questions)} question(s); "
            f"nothing to split."
        )
        return False

    total = len(questions)
    print(
        f"\nFile: {input_path.name}"
        f"\nFound {total} question(s). Creating {len(chunks)} file(s):"
    )

    for i, piece in enumerate(chunks, start=1):
        out_path = make_output_name(input_path, i, outdir)
        out_text = "\n\n".join(piece) + "\n"
        out_path.write_text(out_text, encoding="utf-8")
        start = sum(len(c) for c in chunks[: i - 1]) + 1
        end = start + len(piece) - 1
        print(
            f"  -> Created: {out_path.name} "
            f"({len(piece)} questions, #{start}-#{end})"
        )

    return True


def main():
    print("=== Quiz File Splitter ===\n")

    # Prompt 1: Number of questions per file
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

    # Prompt 2: leftover threshold. If the questions still to be placed are this
    # many or fewer, they all go in one final file instead of being split again.
    tail_input = input(
        "If the leftover is this many questions or fewer, keep them in one "
        "file. [Default 15]: "
    ).strip()
    if not tail_input:
        tail_threshold = 15
    else:
        try:
            tail_threshold = int(tail_input)
            if tail_threshold < 0:
                raise ValueError
        except ValueError:
            print("Error: Please enter a valid number (0 or more).")
            return

    # Prompt 3: Drag and drop a file or folder
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
            if process_file(file, size, tail_threshold):
                files_processed += 1

    elif input_path.is_file():
        if process_file(input_path, size, tail_threshold):
            files_processed += 1
    else:
        print("\nError: Specified path is neither a file nor a folder.")
        return

    if files_processed > 0:
        print("\nSuccessfully finished splitting!")
    else:
        print("\nNo files were split.")

    # Prompt 4: Ask to run Convert My Quizzes.cmd
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
