from pathlib import Path

p = Path("mnemonics/python/migrate_md_to_csv.py")
t = p.read_text(encoding="utf-8")
t = t.replace('"kind": "subtopic"', '"kind": "zoned"')
t = t.replace(
    "# Cap at 4 subtopics (technique max); extras get new beast later if needed — keep first 4",
    "# Cap at 4 zoned Knowledge Atoms (technique max)",
)
t = t.replace(
    'any(a.get("kind") == "subtopic" for a in atoms)',
    'any(a.get("kind") in ("zoned", "subtopic") for a in atoms)',
)
p.write_text(t, encoding="utf-8")
print(
    "ok",
    t.count('"kind": "zoned"'),
    "zoned;",
    t.count('"kind": "subtopic"'),
    "subtopic left",
)
