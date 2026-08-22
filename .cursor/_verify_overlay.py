from pathlib import Path

t = Path("mnemonics/output/dashboard.html").read_text(encoding="utf-8")
for c in (
    "field-concept",
    "beast-name",
    "gap: 1.75rem",
    "max-height: 52vh",
    "grid-template-columns: 1fr",
):
    print(c, c in t)
i = t.find('id="ovImage"')
a = t.find('id="ovAtoms"')
p = t.find('id="ovPrompt"')
print("order_ok", i < a < p, i, a, p)
