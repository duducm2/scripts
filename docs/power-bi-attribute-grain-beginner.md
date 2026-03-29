# Attribute grain vs composite grain — beginner guide

This note explains **attribute grain table**, whether you must **duplicate** the `individual-composite-scores` query in Power Query, and the **order of steps** used in this project’s Pipeline B workflow.

---

## 1. What “attribute grain table” means

**Grain** = *what one row in a table represents*.

| Term | Meaning in this workflow |
|------|---------------------------|
| **Attribute grain** | One row = **one attribute** for one subject/website (and related keys). You have many rows (on the order of **~19,116** in this project). Columns like **`attribute_name`**, **`attribute_type`**, **`attribute_score`** are normal, flat columns you can drag into slicers and visuals. |
| **Composite grain** | One row = **one composite score** per subject/website group (after aggregation). You have far fewer rows (on the order of **~531**). The main measure here is **`composite_trust_score`**. |

The **attribute grain table** is simply the **table/query that stays at attribute grain** — the wide, detailed table — **not** the small aggregated table.

**Important:** After you **Group By** to build the composite, Power Query often creates a nested column like **`group_rows`**. Inside that nested table, `attribute_name` still exists, but it is **hidden inside a table column**, not as a top-level field in the model. For slicers and most DAX, you need **`attribute_name` as a normal column** on the **attribute-grain query**, not only inside `group_rows` on the composite query.

---

## 2. Required action: duplicate the query in Power Query?

**Yes.** For this workflow you should:

1. **Duplicate** the **`individual-composite-scores`** query in the Power Query Editor (right‑click the query → **Duplicate**).
2. **Leave one copy unchanged** at **attribute grain** (all original rows; flat columns including `attribute_name`, `attribute_type`, `attribute_score`). This is your **attribute grain table** for slicers and Section 4.3–style measures.
3. **Use only the other copy** for **Group By**, **`group_rows`**, **`composite_trust_score`**, and any steps that shrink the data to **~531 rows** (composite grain).
4. **Rename** the queries (or adjust **load names** in the model) so you can tell them apart — for example one might stay `individual-composite-scores` for the composite side and the other `individual-composite-scores-attr` (or similar). If two tables cannot share one display name, pick two clear names and use the **same names consistently** in DAX and slicers (as your main setup guide and `measures.md` describe).

You do **not** need two copies of the **Excel sheet** — only **two queries** in Power Query that start from the same source, then diverge (one flat, one aggregated).

---

## 3. Step-by-step (beginner)

1. **Import data**  
   Load the Excel sheet **`individual-composite-scores`** (with the exported TrustMate data) into Power BI as you already do.

2. **Open Power Query**  
   **Home → Transform data**.

3. **Duplicate the query**  
   In the left list, right‑click **`individual-composite-scores`** → **Duplicate**.  
   You now have two queries; rename them so the names reflect their role, e.g.  
   - `individual-composite-scores-attr` = attribute grain (keep flat)  
   - `individual-composite-scores` (or `-composite`) = composite grain (you will aggregate this one only)

4. **Do not touch the duplicate you reserved for attribute grain**  
   On that copy, **do not** run Group By to one row per composite group. It should stay at **~19k rows** with **`attribute_name`** visible as a normal column in the preview.

5. **Aggregate only the other copy**  
   On the **second** query only: **Group By** on the keys your guide specifies (e.g. `id_subject`, `id_website`, `subject_type`), add **`group_rows`** (All Rows), then add **`composite_trust_score`** (Custom Column / M as in your main guide).  
   Result: **~531 rows**.

6. **Load both**  
   **Close & apply**. In **Model** view you should see **two tables**: one large (attribute), one small (composite).

7. **Use the right table for each job**  
   - Slicers and fields like **`attribute_name`**, **`attribute_type`** → **attribute** table (`AttributeGrainTable` in docs = that attribute-level table’s name in *your* model).  
   - **`composite_trust_score`** and composite Section 4.2 measures → **composite** table.

8. **If `attribute_name` is missing in the Data pane** on a table where you only aggregated → you are looking at the **composite** table. Switch to the **attribute** table for attribute-level fields.

---

## 4. Quick mental model

- **One export, two queries:** same starting data; **duplicate** in Power Query; **one path flat**, **one path aggregated**.  
- **Attribute grain table** = the **flat, many-row** query.  
- **Composite grain** = the **small** query with **`composite_trust_score`**; **`group_rows`** is for calculation inside M, not for dragging **`attribute_name`** into reports.

For full M steps, naming alignment with `measures.md`, and DAX table names, follow your project’s **`POWER_BI_SETUP_GUIDE.md`** and **`measures.md`** — this file only clarifies **terms** and **duplicate vs rename** expectations.
