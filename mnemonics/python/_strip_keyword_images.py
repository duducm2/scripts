"""One-shot helper: drop keyword-image tests and practice picture rows."""

from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]

test_path = ROOT / "mnemonics" / "python" / "test_mnemonic_keywords.py"
lines = test_path.read_text(encoding="utf-8").splitlines(keepends=True)
kept = lines[:77] + ["\n"] + lines[325:508]
text = "".join(kept)
text = text.replace("import keyword_images\n", "")
text = text.replace("import palace_practice_render\n", "")
old = '    assert "Keywords" not in text\n' '    assert "No keywords yet" not in text\n'
new = (
    '    assert "Keywords" not in text\n'
    '    assert "No keywords yet" not in text\n'
    '    assert "keyword-images" not in text\n'
)
if old not in text:
    raise SystemExit("assertion anchor missing")
test_path.write_text(text.replace(old, new, 1), encoding="utf-8")
print("tests rewritten")

removed = 0
practice = ROOT / "mnemonics" / "output" / "practice"
for md in practice.glob("*.md"):
    src = md.read_text(encoding="utf-8")
    out_lines = []
    for line in src.splitlines(keepends=True):
        if "web/assets/keyword-images/" in line:
            removed += 1
            continue
        out_lines.append(line)
    body = "".join(out_lines)
    while "\n\n\n" in body:
        body = body.replace("\n\n\n", "\n\n")
    if body != src:
        md.write_text(body, encoding="utf-8")
        print("stripped", md.name)
print("image lines removed", removed)
