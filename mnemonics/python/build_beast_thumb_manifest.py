"""Rebuild the full beast-thumb catalog.

Older copies of this script wrote only Quick Recall beasts and dropped icons
for every other palace. Import and the SPA need the complete catalog, so this
entry point always rebuilds every beast.
"""

from __future__ import annotations

from rebuild_full_beast_thumb_manifest import main


if __name__ == "__main__":
    main()
