from pathlib import Path

p = Path(r"C:\Users\eduev\Meu Drive\17 - Projects\scripts\assets\data\prompts.ini")
raw = p.read_bytes()
print("bom", raw[:3].hex(), "size", len(raw))
t = raw.decode("utf-8-sig")
print("fffd", t.count("\ufffd"))
names = [ln for ln in t.splitlines() if ln.startswith("Name=")]
out = Path(r"C:\Users\eduev\Meu Drive\17 - Projects\scripts\.cursor\_names_now.txt")
lines = []
for n in names:
    prefix = n[5:20]
    cps = " ".join(f"U+{ord(c):04X}" for c in prefix[:8])
    lines.append(f"{n}\n  {cps}\n  bytes={n[5:20].encode('utf-8')[:24].hex()}")
out.write_text("\n".join(lines), encoding="utf-8")
print("names", len(names))
