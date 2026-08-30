from pathlib import Path

p = Path(r"mnemonics/technique/prompts/story-reduction-prompt.txt")
text = p.read_text(encoding="utf-8")
old = (
    "- Preserve every remaining atom\u2019s concept/quote/story/sensory/ipa (no knowledge drops).\n"
    "- SMASH: merged atoms become `kind=zoned` with `Z1`\u2013`Z4` (max 4); never encode knowledge on a smashed beast body alone.\n"
    "- REMOVE: drop atoms for requested pegs only; reattach survivors to re-pegged beasts."
)
new = (
    "- Preserve every remaining atom\u2019s knowledge payload (no knowledge drops). Re-apply `README \u00a7Knowledge Atom Structure` compression: keep or tighten `concept` to maximal brevity with full semantic fidelity; never expand concepts during re-peg.\n"
    "- **Quote:** keep the verbatim transcript excerpt intact; add or adjust the optional ` \u2014 Note: <stripped nuance>` suffix only when compression requires it (same field format as story generation). Do not invent new factual content in Notes.\n"
    "- SMASH: merged atoms become `kind=zoned` with `Z1`\u2013`Z4` (max 4); never encode knowledge on a smashed beast body alone. Each zoned atom keeps its own compressed `concept`; any Note stays on that atom\u2019s `quote`.\n"
    "- REMOVE: drop atoms for requested pegs only; reattach survivors to re-pegged beasts with the same compression + Note rules."
)
if old not in text:
    raise SystemExit("OLD NOT FOUND")
p.write_text(text.replace(old, new, 1), encoding="utf-8")
print("OK")
