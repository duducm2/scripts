"""One-shot: revise Data Analyst plan resources (reading sources + fix weak links)."""

from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path

data = Path(__file__).resolve().parent.parent / "data"
res_path = data / "plan_resources.csv"
plan_id = "PLAN_0007"

with res_path.open(newline="", encoding="utf-8") as f:
    rows = list(csv.DictReader(f))
fr = ["id", "plan_id", "section_path", "line", "sort_order"]

# Drop near-duplicate Phase 5 book summary (same video as 5.1)
rows = [r for r in rows if r["id"] != "PRES_0446"]

# Fix weak / mismatched rows (prefer durable docs when the video link was wrong)
fixes = {
    "PRES_0416": (
        "- [XLOOKUP function (Microsoft Support)]"
        "(https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929)"
    ),
    "PRES_0417": (
        "- [Create a PivotTable to analyze worksheet data (Microsoft Support)]"
        "(https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576)"
    ),
    "PRES_0428": (
        "- [Understand star schema and the importance for Power BI (Microsoft Learn)]"
        "(https://learn.microsoft.com/en-us/power-bi/guidance/star-schema)"
    ),
    "PRES_0434": (
        "- [Group by: split-apply-combine (pandas docs)]"
        "(https://pandas.pydata.org/docs/user_guide/groupby.html)"
    ),
    "PRES_0441": (
        "- [GitHub Docs: About READMEs]"
        "(https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/about-readmes)"
    ),
}
for r in rows:
    if r["id"] in fixes:
        r["line"] = fixes[r["id"]]

p1 = "Phase 1: Advanced Tabular Modeling and Automation (Main Corporate Atrium)"
p2 = "Phase 2: Relational Database Extraction and Querying (Subterranean Archive)"
p3 = "Phase 3: Business Intelligence and Dimensional Visualization (Grand Exhibition Hall)"
p4 = "Phase 4: Programmatic Manipulation and Advanced Transformation (Chemical Engineering Laboratory)"
p5 = (
    "Phase 5: Synthesis, Storytelling, and Portfolio Development (Narrative Auditorium)"
)

additions: list[tuple[str, str]] = [
    # Phase 1
    (
        f"{p1} > 2. Advanced Functions",
        "- [INDEX function (Microsoft Support)]"
        "(https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a3-be95-3a391f4e6b2f)",
    ),
    (
        f"{p1} > 2. Advanced Functions",
        "- [MATCH function (Microsoft Support)]"
        "(https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a)",
    ),
    (
        f"{p1} > 4. Automated ETL Pipelines",
        "- [Power Query documentation (Microsoft Learn)](https://learn.microsoft.com/en-us/power-query/)",
    ),
    (
        f"{p1} > 4. Automated ETL Pipelines",
        "- [Power Query M formula language reference](https://learn.microsoft.com/en-us/powerquery-m/)",
    ),
    # Phase 2
    (
        f"{p2} > 1. Database Architecture",
        "- [PostgreSQL Tutorial (official)](https://www.postgresql.org/docs/current/tutorial.html)",
    ),
    (
        f"{p2} > 1. Database Architecture",
        "- [Mode Analytics SQL Tutorial (written)](https://mode.com/sql-tutorial)",
    ),
    (
        f"{p2} > 3. Set Theory and Joins",
        "- [PostgreSQL: Joins](https://www.postgresql.org/docs/current/tutorial-join.html)",
    ),
    (
        f"{p2} > 4. Query Modularity",
        "- [PostgreSQL: WITH Queries (CTEs)](https://www.postgresql.org/docs/current/queries-with.html)",
    ),
    (
        f"{p2} > 5. Advanced Analytics",
        "- [PostgreSQL: Window Functions](https://www.postgresql.org/docs/current/tutorial-window.html)",
    ),
    # Phase 3
    (
        f"{p3} > 1. Interface and Data Ingestion",
        "- [Get started with Power BI Desktop (Microsoft Learn)]"
        "(https://learn.microsoft.com/en-us/power-bi/fundamentals/desktop-getting-started)",
    ),
    (
        f"{p3} > 3. Dimensional Modeling",
        "- [Kimball dimensional modeling techniques (primer)]"
        "(https://www.kimballgroup.com/data-warehouse-business-intelligence-resources/"
        "kimball-techniques/dimensional-modeling-techniques/)",
    ),
    (
        f"{p3} > 4. DAX and Evaluation Context",
        "- [Row context and filter context in DAX (SQLBI)]"
        "(https://www.sqlbi.com/articles/row-context-and-filter-context-in-dax/)",
    ),
    (
        f"{p3} > 4. DAX and Evaluation Context",
        "- [Introducing CALCULATE in DAX (SQLBI)]"
        "(https://www.sqlbi.com/articles/introducing-calculate-in-dax/)",
    ),
    (
        f"{p3} > 5. Dashboard UX and Portfolio",
        "- [Visualization types in Power BI (Microsoft Learn)]"
        "(https://learn.microsoft.com/en-us/power-bi/visuals/"
        "power-bi-visualization-types-for-reports-and-q-and-a)",
    ),
    # Phase 4
    (
        f"{p4} > 1. Environment Setup",
        "- [Installing pandas (official)](https://pandas.pydata.org/docs/getting_started/install.html)",
    ),
    (
        f"{p4} > 2. Data Structures",
        "- [10 minutes to pandas (official)](https://pandas.pydata.org/docs/user_guide/10min.html)",
    ),
    (
        f"{p4} > 3. Exploratory Data Analysis",
        "- [Essential basic functionality (pandas docs)]"
        "(https://pandas.pydata.org/docs/user_guide/basics.html)",
    ),
    (
        f"{p4} > 4. Vectorization and Aggregation",
        "- [Merge, join, concatenate and compare (pandas docs)]"
        "(https://pandas.pydata.org/docs/user_guide/merging.html)",
    ),
    (
        f"{p4} > 5. Statistical Visualization",
        "- [seaborn tutorial (official)](https://seaborn.pydata.org/tutorial.html)",
    ),
    # Phase 5
    (
        f"{p5} > 1. Audience Empathy and Context",
        "- [Storytelling with Data — Cole Nussbaumer Knaflic (book)]"
        "(https://www.storytellingwithdata.com/book)",
    ),
    (
        f"{p5} > 2. Cognitive Load Reduction",
        "- [storytellingwithdata.com blog / examples](https://www.storytellingwithdata.com/blog)",
    ),
    (
        f"{p5} > 5. Public Deployment",
        "- [Creating a GitHub Pages site (GitHub Docs)]"
        "(https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site)",
    ),
]

max_rid = max(int(r["id"].split("_")[1]) for r in rows if r["id"].startswith("PRES_"))
sec_ord: dict[str, int] = defaultdict(int)
for r in rows:
    if r["plan_id"] == plan_id:
        sec_ord[r["section_path"]] = max(
            sec_ord[r["section_path"]], int(r["sort_order"] or 0)
        )

rid = max_rid + 1
new_rows = []
for sec, line in additions:
    sec_ord[sec] += 1
    new_rows.append(
        {
            "id": f"PRES_{rid:04d}",
            "plan_id": plan_id,
            "section_path": sec,
            "line": line,
            "sort_order": str(sec_ord[sec]),
        }
    )
    rid += 1

rows.extend(new_rows)
with res_path.open("w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=fr, lineterminator="\n")
    w.writeheader()
    w.writerows(rows)

da = sum(1 for r in rows if r["plan_id"] == plan_id)
print(f"PLAN_0007 resources: {da}")
print(f"Added {len(new_rows)}; fixed {len(fixes)}; removed PRES_0446")
