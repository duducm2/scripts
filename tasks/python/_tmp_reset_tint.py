from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from task_store import TaskStore

store = TaskStore(Path(__file__).resolve().parent.parent / "data")
rows = store.load("projects")
changed = 0
for r in rows:
    if str(r.get("icon_tint") or "").strip() != "0":
        r["icon_tint"] = "0"
        changed += 1
store.save("projects", rows)
print("reset_tint", changed)
