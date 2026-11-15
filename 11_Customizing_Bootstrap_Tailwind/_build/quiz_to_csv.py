"""Convert WEBD1000_Quiz_Week_11.md into a Brightspace question-import CSV.

    python _build/quiz_to_csv.py ../temp/WEBD1000_Quiz_Week_11.md ../temp/WEBD1000_Quiz_Week_11.csv

The quiz lives in the git-ignored temp/ folder so students can't pull the answer key.

Import it in Brightspace: Quizzes -> Question Library -> Import -> Upload a File.

Matches the INET2005 quiz CSVs: 5 columns, every field quoted, a blank row between questions.
Quoting is done by the csv module (QUOTE_ALL), so a " inside code becomes "" automatically.
Code blocks are flattened to one line: Brightspace's CSV import is one field per line.
"""
import csv, io, re, sys
from pathlib import Path

src = Path(sys.argv[1]); out = Path(sys.argv[2])
md = src.read_text(encoding="utf-8")

def oneline(s):
    return " ".join(s.split())

def plain(s):
    # Strip Markdown emphasis from the prose ONLY. Code (fenced blocks and `inline`) is set
    # aside first and put back untouched, so /* comments */ and quotes survive exactly.
    saved = []
    def keep(text):
        saved.append(text)
        return f"\x00{len(saved) - 1}\x00"
    s = re.sub(r"```\w*\n(.*?)```", lambda m: " Code: " + keep(oneline(m.group(1))) + " ", s, flags=re.S)
    s = re.sub(r"`([^`]+)`", lambda m: keep(m.group(1)), s)
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"(?<![\w*])\*(\S.*?\S|\S)\*(?![\w*])", r"\1", s)   # *italic*
    return re.sub(r"\x00(\d+)\x00", lambda m: saved[int(m.group(1))], s)

TITLES = ["Hovering a recoloured button", "Why it's called JavaScript", "Opening a modal",
          "Deleting Bootstrap's script", "Tooltips are opt in", "A toast that never shows",
          "Tailwind's md: prefix", "Spotting Tailwind", "table-responsive", "When two rules tie"]
DIFF = [2, 1, 1, 2, 2, 2, 1, 1, 1, 2]

rows = []
blocks = re.split(r"^## Question (\d+)\s*$", md, flags=re.M)
for num, body in zip(blocks[1::2], blocks[2::2]):
    n = int(num)
    body = body.split("\n---")[0].split("\n## ")[0]
    stem_part, _, rest = body.partition("\n- a)")
    rest = "- a)" + rest
    opts_part, _, fb_part = rest.partition("**Feedback:**")
    # Options can wrap onto an indented second line
    options = [oneline(o) for o in re.split(r"^- [a-d]\) ", opts_part, flags=re.M) if o.strip()]
    stem = oneline(plain(stem_part))
    feedback = oneline(plain(fb_part))
    correct = [("(correct)" in o) for o in options]
    if len(options) != 4 or sum(correct) != 1:
        raise SystemExit(f"Question {n}: expected 4 options with exactly 1 correct, got {len(options)}/{sum(correct)}")
    rows += [["NewQuestion", "MC"], ["ID", f"WEBD1000-WK11-{n:02d}"], ["Title", TITLES[n - 1]],
             ["QuestionText", stem], ["Points", "1"], ["Difficulty", str(DIFF[n - 1])]]
    for o, ok in zip(options, correct):
        rows.append(["Option", "100" if ok else "0", plain(o.replace("**(correct)**", "")).strip(), "", ""])
    rows += [["Hint"], ["Feedback", feedback], [""]]

rows = [(r + [""] * 5)[:5] for r in rows]
buf = io.StringIO()
csv.writer(buf, quoting=csv.QUOTE_ALL, lineterminator="\n").writerows(rows)
out.write_text(buf.getvalue(), encoding="utf-8")
print(f"wrote {out} ({len(blocks) // 2} questions)")
