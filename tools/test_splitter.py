import importlib.util
import sys
from pathlib import Path

REPO = Path(r"C:\Users\Kath Puno\Downloads\DataComs_Reviewer")

spec = importlib.util.spec_from_file_location("Splitter", REPO / "tools" / "Splitter.py")
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

failures = []


def check(label, got, want):
    ok = got == want
    if not ok:
        failures.append(label)
    print(f"  {'OK  ' if ok else 'FAIL'} {label}: got {got}, want {want}")


print("=== 1. question detection on the real file ===")
src = REPO / "DataComs_PrelimExam" / "Checkpoint Exam 3-5.txt"
text = src.read_text(encoding="utf-8")
questions, preamble = m.split_into_questions(text)
check("questions found", len(questions), 74)
check("preamble lines", len(preamble), 0)

old_blocks = m.parse_blocks(text)
check("old parser saw (blocks)", len(old_blocks), 84)

print("\n=== 2. every split piece starts at a question ===")
bad = [i for i, q in enumerate(questions) if not q.lstrip().startswith("Question:")]
check("pieces not starting with Question:", bad, [])

print("\n=== 3. chunking with the 15-question tail rule ===")
for total, want in [(74, [10, 10, 10, 10, 10, 10, 14]),
                    (26, [10, 10, 6]),
                    (25, [10, 15]),
                    (20, [10, 10]),
                    (15, [15]),
                    (9, [9]),
                    (100, [10] * 10)]:
    fake = [f"Question: q{i}" for i in range(total)]
    got = [len(c) for c in m.chunk_with_tail(fake, 10, 15)]
    check(f"total {total}", got, want)

print("\n=== 4. no question is lost or duplicated ===")
chunks = m.chunk_with_tail(questions, 10, 15)
flat = [q for c in chunks for q in c]
check("total questions in chunks", len(flat), 74)
check("identical to source order", flat == questions, True)
check("no piece is empty", all(c for c in chunks), True)

print("\n=== 5. each question keeps its explanation ===")
missing = [i + 1 for i, q in enumerate(flat) if "Explanation:" not in q]
check("pieces with no Explanation:", missing, [])

print("\n=== 6. the multi-paragraph question stays whole ===")
for q in flat:
    if "Extended ACLs are required for this scenario" in q:
        has_q = q.lstrip().startswith("Question:")
        has_place = "Placement:" in q
        check("prose question is one piece, starts with Question", has_q, True)
        check("prose question keeps its Placement paragraph", has_place, True)
        break
else:
    failures.append("prose question not found")
    print("  FAIL prose question not found")

print()
if failures:
    print(f"RESULT: {len(failures)} FAILURE(S): {failures}")
    sys.exit(1)
print("RESULT: all checks passed.")
