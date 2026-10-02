# Data Analyst

<details open>
<summary><strong>Memory Palace 18: Star Schema Best Practices</strong> · Character: Galileo Galilei · 1 beast · 1 atom</summary>

![Memory Palace 18](images/data-analyst/18.png)

<p><em>1 beast · 1 Knowledge Atom</em></p>

#### Knowledge Atoms

### [Bu] butterfly

<img src="../../web/assets/beast-thumbs/butterfly.png" alt="butterfly" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Lean Fact Table]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BLean%20Fact%20Table%5D&color=9fd4ff" />(razor) [I keep my fact table <img alt="[lean]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Blean%5D&color=ffd966" />(ruler)] [by storing only keys and <img alt="[measures]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bmeasures%5D&color=ffd966" />(safe)] [while placing descriptive fields in <img alt="[dimensions]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdimensions%5D&color=ffd966" />(drawer)]

**Quote**
“keep your fact table lean only include keys and measures if possible avoid putting descriptive fields like product name in your fact table that's what the dimension tables are for”

#### Notes

No, that quote describes **dimensional modeling** (specifically designing a **star schema** for data warehousing / analytical reporting), not traditional database normalization.

While both techniques separate data across multiple tables to avoid redundancy, their rules, goals, and results are fundamentally different.

---

### Key Distinctions

| Feature | Star Schema / Dimensional Design | Database Normalization (3NF / BCNF) |
| --- | --- | --- |
| **Primary Goal** | Fast, intuitive analytical queries (OLAP) and aggregations. | Eliminating update/insert/delete anomalies and write redundancy (OLTP). |
| **Fact Table Role** | Stores numeric metrics/measures and foreign keys to dimensions. | Not a concept in relational modeling (everything is an entity/relation). |
| **Dimension Structure** | **Intentionally denormalized.** A single `dim_product` table typically bundles category, subcategory, brand, and name into flat columns. | Decomposed into separate linked tables (e.g., `products`, `subcategories`, `categories`) to remove transitive dependencies. |
| **Join Complexity** | Low. Fact tables join directly to wide dimension tables in 1-hop joins. | High. Queries require multiple deep joins across normalized entities. |

---

### Why It Is Often Confused with Normalization

1. **Splitting attributes:** Moving descriptive text (e.g., `product_name`) out of a transaction table resembles decomposing a flat sheet into relational entities.
2. **Surrogate keys:** Both approaches use keys (`product_key`) to link transactional rows to descriptive attributes.

### Where Normalization Would Go Further (Snowflaking)

If you strictly normalized the dimension tables themselves (e.g., splitting `dim_product` into separate tables for `product`, `subcategory`, and `category` so no non-key attribute depends on another non-key attribute), that is known in data warehousing as a **snowflake schema**.

In modern data warehousing (e.g., Power BI, Snowflake, BigQuery), the standard best practice remains the **star schema**: keeping the fact table strictly lean (keys + numeric values) while keeping the dimension tables wide and deliberately denormalized.

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 17: Star Schema Core Architecture</strong> · Character: Isaac Newton · 5 beasts · 5 atoms</summary>

![Memory Palace 17](images/data-analyst/17.png)

<p><em>5 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [Bt] Bone toucan

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/toucan.png" alt="Bone toucan" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Surrogate Key]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BSurrogate%20Key%5D&color=9fd4ff" />(brass key) [I assign an artificial unique <img alt="[identifier]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bidentifier%5D&color=ffd966" />(barcode)] [inside a <img alt="[dimension]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdimension%5D&color=ffd966" />(cabinet) table] [to maintain <img alt="[consistent]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bconsistent%5D&color=ffd966" />(handcuffs) relationships] — Note: Has no business meaning and prevents broken links when natural names change.

**Quote**
“surrogate keys are made up they're unique IDs like customer key for example 101 102 103 they have no business meaning but they're great for performance and consistency”

### [Bs] Bone skull

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/skull.png" alt="Bone skull" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Denormalization]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BDenormalization%5D&color=9fd4ff" />(steamroller) [I <img alt="[combine]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcombine%5D&color=ffd966" />(glue) related data] [into fewer <img alt="[wider]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bwider%5D&color=ffd966" />(bench) tables] [to <img alt="[speed]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bspeed%5D&color=ffd966" />(speedometer) up analytical queries] — Note: Normalization splits data across many small tables to reduce redundancy for transactions.

**Quote**
“denormalization is about flattening the data you combine related information into fewer but wider tables this is ideal for analytics and it's exactly what the star schema does”

### [Br] brontosaurus

<img src="../../web/assets/beast-thumbs/brontosaurus.png" alt="brontosaurus" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Dimension Table]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BDimension%20Table%5D&color=9fd4ff" />(tag) [I store descriptive <img alt="[attributes]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Battributes%5D&color=ffd966" />(label)] [that give <img alt="[context]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcontext%5D&color=ffd966" />(lens) to numerical facts] [for <img alt="[filtering]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bfiltering%5D&color=ffd966" />(funnel) and grouping]

**Quote**
“radiating out from the center are your dimension tables these describe the facts dimension tables are things like customers products dates and regions”

### [Bq] Bone Quetzalcoatl

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/quetzalcoatl.png" alt="Bone Quetzalcoatl" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Fact Table]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BFact%20Table%5D&color=9fd4ff" />(abacus) [I store quantitative <img alt="[metrics]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bmetrics%5D&color=ffd966" />(chest)] [at the <img alt="[center]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcenter%5D&color=ffd966" />(bullseye) of the schema] [<img alt="[linked]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Blinked%5D&color=ffd966" />(chain) to surrounding dimensions] — Note: Kept lean and numeric by excluding descriptive text fields.

**Quote**
“picture a star at the center is your fact table this holds the numbers sales revenue quantities or whatever else you capture in a transaction”

### [Bp] Bone panther

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/panther.png" alt="Bone panther" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Star Schema]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BStar%20Schema%5D&color=9fd4ff" />(badge) [I <img alt="[organize]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Borganize%5D&color=ffd966" />(binder) data into a central fact table] [<img alt="[surrounded]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bsurrounded%5D&color=ffd966" />(fence) by dimension tables] [to <img alt="[optimize]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Boptimize%5D&color=ffd966" />(rocket) querying and reporting]

**Quote**
“the star schema is a widely used data modeling technique in PowerBI and other business intelligence tools it's designed to optimize querying and reporting by organizing data into a clear intuitive structure”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 16: API to Power BI</strong> · Character: Albert Einstein · 4 beasts · 4 atoms</summary>

![Memory Palace 16](images/data-analyst/16.jpg)

<p><em>4 beasts · 4 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [Bo] bower-bird

<img src="../../web/assets/beast-thumbs/bower_bird.png" alt="bower-bird" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Push Datasets]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BPush%20Datasets%5D&color=9fd4ff" />(button) [I <img alt="[stream data]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bstream%20data%5D&color=ffd966" />(hose)] [directly into <img alt="[Power BI]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BPower%20BI%5D&color=ffd966" />(battery)] [using its <img alt="[REST API]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BREST%20API%5D&color=ffd966" />(menu)] — Note: This method does not support relationships or joins, so all complex data models must be flattened before ingestion.

**Quote**
“you can only send data using power bi's push data sets method which doesn't support relationships or joins meaning complex data models must be flattened before ingestion”

### [Bn] Bone Neanderthal

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/neanderthal.png" alt="Bone Neanderthal" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Custom Scripting]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BCustom%20Scripting%5D&color=9fd4ff" />(scroll) [I <img alt="[write custom scripts]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bwrite%20custom%20scripts%5D&color=ffd966" />(quill)] [to <img alt="[fetch API data]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bfetch%20API%20data%5D&color=ffd966" />(fishing rod)] [and <img alt="[load it]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bload%20it%5D&color=ffd966" />(dump truck)] [into a <img alt="[data warehouse]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdata%20warehouse%5D&color=ffd966" />(forklift)] — Note: This provides unmatched flexibility and optimization for big data, but requires significant coding complexity and infrastructure maintenance.

**Quote**
“whether you send data directly to powerbi or via a data warehouse custom scripts provide unmatched flexibility compared to the other two options”

### [Bm] Bone marmoset

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/marmoset.png" alt="Bone marmoset" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Native Web Connection]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BNative%20Web%20Connection%5D&color=9fd4ff" />(spider web) [I <img alt="[connect a web endpoint]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bconnect%20a%20web%20endpoint%5D&color=ffd966" />(cable)] [to <img alt="[Power BI]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BPower%20BI%5D&color=ffd966" />(battery)] [by <img alt="[specifying parameters]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bspecifying%20parameters%5D&color=ffd966" />(dial)] <img alt="[directly in the URL]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdirectly%20in%20the%20URL%5D&color=ffd966" />(address bar) — Note: This method is free and simple for one-time imports but lacks scalability for rate limits or complex authentication.

**Quote**
“you can connect an endpoint to powerbi by specifying parameters directly in the URL”

### [Bl] bloodhound

<img src="../../web/assets/beast-thumbs/bloodhound.png" alt="bloodhound" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[No-Code Integration]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BNo-Code%20Integration%5D&color=9fd4ff" />(magic wand) [I use an <img alt="[intermediary platform]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bintermediary%20platform%5D&color=ffd966" />(bridge)] [to <img alt="[connect data sources]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bconnect%20data%20sources%5D&color=ffd966" />(plug)] [to <img alt="[Power BI]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BPower%20BI%5D&color=ffd966" />(battery)] [in a <img alt="[fully automated way]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bfully%20automated%20way%5D&color=ffd966" />(robot)] — Note: These platforms clean messy responses and schedule updates but may lack support for complex mechanics like pagination.

**Quote**
“coupler iio is a no code solution that allows you to connect almost any data source to powerbi and other Platforms in a fully automated way”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 15: SQL Window Functions</strong> · Character: Socrates · 4 beasts · 5 atoms</summary>

![Memory Palace 15](images/data-analyst/15.jpg)

<p><em>4 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [Bk] Bone kitten

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/kitten.png" alt="Bone kitten" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Inline Aggregate Shortcut]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BInline%20Aggregate%20Shortcut%5D&color=9fd4ff" />(syringe) [I <img alt="[inject aggregated metrics]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Binject%20aggregated%20metrics%5D&color=ffd966" />(needle)] [directly into an <img alt="[un-grouped query]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bun-grouped%20query%5D&color=ffd966" />(open book)] [using a <img alt="[single OVER clause]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bsingle%20OVER%20clause%5D&color=ffd966" />(blanket)]

**Quote**
“what the partition by is doing is basically taking this query right here and sticking it on one line in the select statement”

### [Bj] Bone jester

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/jester.png" alt="Bone jester" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Isolated Aggregation]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BIsolated%20Aggregation%5D&color=9fd4ff" />(test tube) [I <img alt="[isolate a single column]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bisolate%20a%20single%20column%5D&color=ffd966" />(tweezers)] [for an <img alt="[aggregate function]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Baggregate%20function%5D&color=ffd966" />(blender)] [without changing the <img alt="[query&#x27;s granularity]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bquery%27s%20granularity%5D&color=ffd966" />(sandglass)]

**Quote**
“because we're using the partition by we're able to isolate just one column that we want to perform our aggregate function on”

### [Bi] bison

<img src="../../web/assets/beast-thumbs/bison.png" alt="bison" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[PARTITION BY]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BPARTITION%20BY%5D&color=9fd4ff" />(glass divider) [I <img alt="[divide the result set]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdivide%20the%20result%20set%5D&color=ffd966" />(pizza slicer)] <img alt="[into partitions]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Binto%20partitions%5D&color=ffd966" />(cubicle) [and <img alt="[calculate the window function]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcalculate%20the%20window%20function%5D&color=ffd966" />(abacus)] [while <img alt="[preserving individual row details]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bpreserving%20individual%20row%20details%5D&color=ffd966" />(magnifying glass)] — Note: GROUP BY reduces the output by collapsing multiple rows into a single summary row.

**Quote**
“the group by statement is going to reduce the number of rows in our output by actually rolling them up and then calculating the sums or averages for each group whereas partition by actually divides the result set into partitions and changes how the window function is calculated”

### [Bh] Bone Hydra

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/hydra.png" alt="Bone Hydra" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

🟦 **Z1 · LAG Function**

**Concept**
💡 <img alt="[LAG Function]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BLAG%20Function%5D&color=9fd4ff" />(lagging foot) [I <img alt="[query data]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bquery%20data%5D&color=ffd966" />(magnifying glass)] [from <img alt="[preceding rows]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bpreceding%20rows%5D&color=ffd966" />(footprints)] [relative to the <img alt="[current evaluation row]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcurrent%20evaluation%20row%5D&color=ffd966" />(anchor)] — Note: It requires an explicit ordering clause and evaluates to null when a preceding row does not exist.

**Quote**
“a lag is going to be for previous days and so notice how on the first day in the results there's no previous day so you see null there”

---

🟦 **Z2 · LEAD Function**

**Concept**
💡 <img alt="[LEAD Function]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BLEAD%20Function%5D&color=9fd4ff" />(leash) [I <img alt="[query data]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bquery%20data%5D&color=ffd966" />(binoculars)] [from <img alt="[subsequent rows]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bsubsequent%20rows%5D&color=ffd966" />(stepping stone)] [relative to the <img alt="[current evaluation row]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcurrent%20evaluation%20row%5D&color=ffd966" />(compass)] — Note: It returns a null value at the final dataset record where no subsequent row exists.

**Quote**
“notice how in the last row there's no next day so you see null at the end and all I'm going to do is change lag to lead”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 14: SQL Joins</strong> · Character: Frédéric Chopin · 3 beasts · 3 atoms</summary>

![Memory Palace 14](images/data-analyst/14.jpg)

<p><em>3 beasts · 3 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [Bg] Bone goat

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/goat.png" alt="Bone goat" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[FULL OUTER JOIN]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BFULL%20OUTER%20JOIN%5D&color=9fd4ff" />(outer space) [I return all <img alt="[records]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Brecords%5D&color=ffd966" />(vinyl record)] [from <img alt="[both tables]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bboth%20tables%5D&color=ffd966" />(table)] [and fill any missing sides with <img alt="[NULLs]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BNULLs%5D&color=ffd966" />(ghost)] — Note: It essentially combines the results of both a left join and a right join.

**Quote**
“Commonly referred to as FULL OUTER JOIN, it returns all records when there is a match in either the left or the right table. It essentially combines the results of both a LEFT JOIN and a RIGHT JOIN. Any missing matches on either side are filled with NULL values.”

### [Bf] Bone frog

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/frog.png" alt="Bone frog" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[LEFT JOIN]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BLEFT%20JOIN%5D&color=9fd4ff" />(left hand) [I return all <img alt="[records]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Brecords%5D&color=ffd966" />(vinyl record)] [from the <img alt="[left table]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bleft%20table%5D&color=ffd966" />(table)] [and fill missing right matches with <img alt="[NULLs]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BNULLs%5D&color=ffd966" />(ghost)]

**Quote**
“Returns all records from the left table, along with the matched records from the right table. If a record in the left table has no match in the right table, the query still returns the left table's row, but populates the right table's columns with NULL values.”

### [Be] bee

<img src="../../web/assets/beast-thumbs/bee.png" alt="bee" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[INNER JOIN]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BINNER%20JOIN%5D&color=9fd4ff" />(bullseye) [I return only the <img alt="[records]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Brecords%5D&color=ffd966" />(vinyl record)] [that have <img alt="[matching values]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bmatching%20values%5D&color=ffd966" />(puzzle piece)] [in <img alt="[both tables]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bboth%20tables%5D&color=ffd966" />(twins)] — Note: Unmatched rows are completely excluded.

**Quote**
“Returns only the records that have matching values in both tables. If a row in the first table does not have a corresponding match in the second table based on the join condition, that row is completely excluded from the final result.”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 13: SQL Advanced Analytics</strong> · Character: Erik Satie · 4 beasts · 4 atoms</summary>

![Memory Palace 13](images/data-analyst/13.jpg)

<p><em>4 beasts · 4 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [Bd] Bone dragon

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/dragon.png" alt="Bone dragon" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[GROUP BY Scope]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BGROUP%20BY%20Scope%5D&color=9fd4ff" />(lasso) [I <img alt="[include]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Binclude%5D&color=ffd966" />(vacuum) any non-aggregated column] [from the <img alt="[SELECT statement]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BSELECT%20statement%5D&color=ffd966" />(menu)] [inside the <img alt="[GROUP BY clause]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BGROUP%20BY%20clause%5D&color=ffd966" />(folder)] — Note: This ensures identical data combinations correctly collapse into a single summary row.

**Quote**
“Any non-aggregated column present in the SELECT list must appear in the GROUP BY clause.”

### [Bc] Bone cat

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/cat.png" alt="Bone cat" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[HAVING Clause]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BHAVING%20Clause%5D&color=9fd4ff" />(funnel) [I <img alt="[filter]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bfilter%5D&color=ffd966" />(coffee filter) summary rows] [after they have been <img alt="[processed]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bprocessed%5D&color=ffd966" />(blender)] [by the GROUP BY <img alt="[aggregation]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Baggregation%5D&color=ffd966" />(snowball)] — Note: The WHERE clause filters individual rows before any data grouping occurs.

**Quote**
“The `HAVING` clause in SQL is used to filter records after they have been aggregated by a `GROUP BY` clause.”

### [Bb] Bone bird of paradise

<img src="../../web/assets/beast-thumbs/adj_bone.png" alt="bone" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" /><img src="../../web/assets/beast-thumbs/bird_of_paradise.png" alt="Bone bird of paradise" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[Temporary Tables]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BTemporary%20Tables%5D&color=9fd4ff" />(tent) [I <img alt="[store]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bstore%5D&color=ffd966" />(freezer) the output of heavy computation] [in a <img alt="[temporary table]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Btemporary%20table%5D&color=ffd966" />(clipboard)] [to prevent the database from <img alt="[re-executing]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bre-executing%5D&color=ffd966" />(hamster wheel) it] — Note: CTEs re-execute from scratch each time, which is inefficient for massive datasets.

**Quote**
“if you find yourself using the same CTE again and again especially if your data set is large and your queries are taking a really long time to run then consider creating a temp table”

### [Ba] bat

<img src="../../web/assets/beast-thumbs/bat.png" alt="bat" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 <img alt="[SQL Advanced Analytics]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BSQL%20Advanced%20Analytics%5D&color=9fd4ff" />(dashboard) [I <img alt="[extract]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bextract%5D&color=ffd966" />(tweezers) nested subqueries using CTEs] [and apply <img alt="[window functions]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bwindow%20functions%5D&color=ffd966" />(window)] [to <img alt="[evaluate]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bevaluate%5D&color=ffd966" />(scales) specific data subsets] — Note: This combines structural organization with advanced analytical evaluations in a single query.

**Quote**
“a window function always has two components... now this whole section is called a CTE and we know that because it has this with keyword”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 12: Location Probes & Rearrangement</strong> · Character: Maurice Ravel · 2 beasts · 2 atoms</summary>

![Memory Palace 12](images/data-analyst/12.jpg)

<p><em>2 beasts · 2 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [Az] Aztec

<img src="../../web/assets/beast-thumbs/aztec.png" alt="Aztec" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Rearrangement]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BRearrangement%5D&color=9fd4ff" />(shuffle): I change marks and <img alt="[axis]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Baxis%5D&color=ffd966" />(dial) values] [so the new layout can change what I <img alt="[understand]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bunderstand%5D&color=ffd966" />(lightbulb)]

**Quote**
“Permitindo ao usuário modificar a disposição de marcas e de valores de eixos no espaço, a nova visão formada desse rearranjo pode levar a diferentes compreensões dos fatos mostrados na estrutura visual”

### [Ay] aye-aye

<img src="../../web/assets/beast-thumbs/aye_aye.png" alt="aye-aye" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Viewpoint]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BViewpoint%5D&color=9fd4ff" />(twin) Controls: I show two windows together] [<img alt="[overview]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Boverview%5D&color=ffd966" />(map) plus enlarged <img alt="[detail]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdetail%5D&color=ffd966" />(lens) of one area]

**Quote**
“Consiste em mostrar duas janelas em conjunto: uma contendo uma visão geral da estrutura visual [e] outra apresentando em detalhes uma área específica dessa estrutura, com foco ampliado.”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 11: Hybrid Marks & View Transforms</strong> · Character: Claude Debussy · 5 beasts · 5 atoms</summary>

![Memory Palace 11](images/data-analyst/11.jpg)

<p><em>5 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [Ax] axolotl

<img src="../../web/assets/beast-thumbs/axolotl.png" alt="axolotl" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Distortions]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BDistortions%5D&color=9fd4ff" />(magnifier): show focus and <img alt="[context]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcontext%5D&color=ffd966" />(frame)] [in the <img alt="[same]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bsame%5D&color=ffd966" />(stretch) visual structure at once]

**Quote**
“Criam visões com foco e contexto simultaneamente em uma mesma estrutura visual.”

### [Aw] awassi sheep

<img src="../../web/assets/beast-thumbs/awassi_sheep.png" alt="awassi sheep" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Location]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BLocation%5D&color=9fd4ff" />(pin) Investigations: I use a data mark's location] [to <img alt="[reveal]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Breveal%5D&color=ffd966" />(flashlight) extra table <img alt="[information]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Binformation%5D&color=ffd966" />(card)]

**Quote**
“Usam o local em que um dado está em uma estrutura visual para revelar informações adicionais da tabela de dados.”

### [Av] avocet

<img src="../../web/assets/beast-thumbs/avocet.png" alt="avocet" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [View <img alt="[Transformation]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BTransformation%5D&color=9fd4ff" />(switch): creates new <img alt="[views]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bviews%5D&color=ffd966" />(window)] [of the visual structure] [for my <img alt="[needs]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bneeds%5D&color=ffd966" />(wrench)]

**Quote**
“Transformação de visão: cria novas visões da estrutura visual de acordo com a necessidade do usuário.”

### [Au] auroch

<img src="../../web/assets/beast-thumbs/auroch.png" alt="auroch" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Dense <img alt="[Pixel]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BPixel%5D&color=9fd4ff" />(mosaic) Displays: I <img alt="[map]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bmap%5D&color=ffd966" />(tile) each value to individual pixels] [and form a <img alt="[polygon]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bpolygon%5D&color=ffd966" />(shape) per data dimension]

**Quote**
“Mapeiam cada valor para pixels individuais e criam um polígono para representar cada dimensão dos dados.”

### [At] atlas

<img src="../../web/assets/beast-thumbs/atlas.png" alt="atlas" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Glyph]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BGlyph%5D&color=9fd4ff" />(puppet): a graphic entity] [whose <img alt="[attributes]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Battributes%5D&color=ffd966" />(dial) are <img alt="[driven]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdriven%5D&color=ffd966" />(remote) by data attributes]

**Quote**
“Glyph: representação visual de um pedaço de dados ou informação em que uma entidade gráfica e seus atributos são controlados por um ou mais atributos de dados”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 10: Multivariate Line & Table Views</strong> · Character: Wolfgang Amadeus Mozart · 5 beasts · 5 atoms</summary>

![Memory Palace 10](images/data-analyst/10.jpg)

<p><em>5 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [As] asp

<img src="../../web/assets/beast-thumbs/asp.png" alt="asp" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Parallel <img alt="[Sets]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BSets%5D&color=9fd4ff" />(ribbon): like Parallel Coordinates] [but focused on <img alt="[nominal]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bnominal%5D&color=ffd966" />(tag) variables]

**Quote**
“Similar a Coordenadas Paralelas, porém com uso focado em variáveis nominais”

### [Ar] armadillo

<img src="../../web/assets/beast-thumbs/armadillo.png" alt="armadillo" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Table <img alt="[Lens]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BLens%5D&color=9fd4ff" />(shuffle): I combine reordering, bar-sized <img alt="[marks]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bmarks%5D&color=ffd966" />(bar), and semantic <img alt="[zoom]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bzoom%5D&color=ffd966" />(lens)] [by row and column]

**Quote**
“Table Lens: combina características: Reordenação de linhas e de colunas Tamanho de marcas (estilo gráfico de barras, para dados quantitativos) Zoom semântico por linha e por coluna”

### [Aq] aquatic leech

<img src="../../web/assets/beast-thumbs/aquatic_leech.png" alt="aquatic leech" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Radial]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BRadial%5D&color=9fd4ff" />(clock) Axis Techniques: I use polar axes] [to study <img alt="[cyclical]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcyclical%5D&color=ffd966" />(loop) events and <img alt="[seasonality]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bseasonality%5D&color=ffd966" />(calendar)]

**Quote**
“Pode ser útil para estudar eventos de natureza cíclica Ex.: hipóteses sobre a sazonalidade de um evento”

### [Ap] ape

<img src="../../web/assets/beast-thumbs/ape.png" alt="ape" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Parallel]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BParallel%5D&color=9fd4ff" />(fence) Coordinates: I draw each variable as a parallel axis] [and turn each <img alt="[tuple]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Btuple%5D&color=ffd966" />(bead) into a <img alt="[polyline]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bpolyline%5D&color=ffd966" />(wire)]

**Quote**
“Representa cada variável por um eixo Eixos são paralelos entre si Tupla se transforma em linha poligonal (polyline).”

### [Ao] aoudad

<img src="../../web/assets/beast-thumbs/aoudad.png" alt="aoudad" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Multivariate]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BMultivariate%5D&color=9fd4ff" />(crayon) Line Charts: I tell dimensions apart] [by color, <img alt="[width]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bwidth%5D&color=ffd966" />(rope), or line <img alt="[style]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bstyle%5D&color=ffd966" />(stripe)]

**Quote**
“Diferenciação das dimensões por atributos gráficos como cor, largura ou estilo de linha”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 9: Encoding & Perception Basics</strong> · Character: Ludwig van Beethoven · 5 beasts · 5 atoms</summary>

![Memory Palace 9](images/data-analyst/9.jpg)

<p><em>5 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [An] angel

<img src="../../web/assets/beast-thumbs/angel.png" alt="angel" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[RadViz]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BRadViz%5D&color=9fd4ff" />(ring): I place N anchors on a circle] [and <img alt="[pull]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bpull%5D&color=ffd966" />(magnet) points by <img alt="[Hooke]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BHooke%5D&color=ffd966" />(spring) spring balance]

**Quote**
“Técnica baseada na lei de Hooke para equilíbrio. Tabela de dados N-dimensionais; M pontos. Define-se N âncoras em uma circunferência”

### [Am] amulet

<img src="../../web/assets/beast-thumbs/amulet.png" alt="amulet" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Effectiveness]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BEffectiveness%5D&color=9fd4ff" />(stopwatch): fast easy <img alt="[distinction]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdistinction%5D&color=ffd966" />(magnifier) of data] [with as few interpretation <img alt="[errors]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Berrors%5D&color=ffd966" />(eraser) as possible]

**Quote**
“Capacidade de permitir rápida interpretação dos dados e fácil distinção entre eles, levando à menor quantidade possível de erros de interpretação.”

### [Al] alligator

<img src="../../web/assets/beast-thumbs/alligator.png" alt="alligator" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Expressiveness]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BExpressiveness%5D&color=9fd4ff" />(truth): my visual mapping must express all table <img alt="[data]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdata%5D&color=ffd966" />(ledger)] [and <img alt="[only]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bonly%5D&color=ffd966" />(filter) that data]

**Quote**
“De acordo com esse conceito, o mapeamento visual deve fazer com que a estrutura visual expresse todos os dados da tabela de dados, e somente eles.”

### [Ak] Akita (dog breed)

<img src="../../web/assets/beast-thumbs/akita_dog_breed.png" alt="Akita (dog breed)" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Automatic]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BAutomatic%5D&color=9fd4ff" />(flashlight) Visual Processing: I aid search and pattern detection] [with automatically <img alt="[processed]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bprocessed%5D&color=ffd966" />(pop) properties] [like <img alt="[color]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcolor%5D&color=ffd966" />(paint) and size]

**Quote**
“Mapeamentos visuais que pretendem auxiliar buscas e detecção de padrões podem ser feitos usando propriedades processadas de maneira automática, como cores e tamanhos;”

### [Aj] Ajax

<img src="../../web/assets/beast-thumbs/ajax.png" alt="Ajax" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Visual <img alt="[Mapping]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BMapping%5D&color=9fd4ff" />(plug): I link each data-table <img alt="[variable]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bvariable%5D&color=ffd966" />(dial)] [to a graphical or spatial <img alt="[property]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bproperty%5D&color=ffd966" />(paint)]

**Quote**
“Objetivo: associar cada variável da tabela de dados a uma propriedade gráfica ou espacial.”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 8: Schema & Visual Structure Core</strong> · Character: Johann Sebastian Bach · 5 beasts · 5 atoms</summary>

![Memory Palace 8](images/data-analyst/8.jpg)

<p><em>5 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [Ai] Airedale terrier

<img src="../../web/assets/beast-thumbs/airedale_terrier.png" alt="Airedale terrier" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Small <img alt="[Multiples]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BMultiples%5D&color=9fd4ff" />(stamps): they force visual <img alt="[comparison]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcomparison%5D&color=ffd966" />(eyes)] [of changes, differences, and <img alt="[alternatives]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Balternatives%5D&color=ffd966" />(fork)]

**Quote**
“A técnica força a comparação visual de alterações, das diferenças entre objetos, do escopo de alternativas.”

### [Ah] Ah!—a sigh

<img src="../../web/assets/beast-thumbs/ah_a_sigh.png" alt="Ah!—a sigh" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [<img alt="[Marks]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BMarks%5D&color=9fd4ff" />(toy): objects present] [in the <img alt="[chart]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bchart%5D&color=ffd966" />(frame) <img alt="[space]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bspace%5D&color=ffd966" />(room)] — Note: Marks use graphical and spatial properties to show data values.

**Quote**
“Objetos presentes no espaço do gráfico”

### [Ag] Agaric fungi

<img src="../../web/assets/beast-thumbs/agaric_fungi.png" alt="Agaric fungi" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Spatial <img alt="[Substrate]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BSubstrate%5D&color=9fd4ff" />(stage): the area available] [to <img alt="[display]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdisplay%5D&color=ffd966" />(screen) the <img alt="[dataset]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdataset%5D&color=ffd966" />(box)]

**Quote**
“Área disponível para exibição do conjunto de dados.”

### [Af] Afghan hound

<img src="../../web/assets/beast-thumbs/afghan_hound.png" alt="Afghan hound" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Visual <img alt="[Structure]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BStructure%5D&color=9fd4ff" />(kit): the set of visual elements] [that <img alt="[represent]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Brepresent%5D&color=ffd966" />(mirror) a <img alt="[dataset]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdataset%5D&color=ffd966" />(box)]

**Quote**
“Conjunto de elementos visuais que representam um conjunto de dados.”

### [Ae] aerialist

<img src="../../web/assets/beast-thumbs/aerialist.png" alt="aerialist" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Database <img alt="[Schema]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BSchema%5D&color=9fd4ff" />(blueprint): I treat it as the structural blueprint] [of my <img alt="[database]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdatabase%5D&color=ffd966" />(building)] [including tables, fields, <img alt="[relationships]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Brelationships%5D&color=ffd966" />(chain), and constraints]

**Quote**
“A database schema is the structural blueprint of a database. It defines the logical organization of data, including tables, fields, relationships, and constraints.”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 7: M Code Pipeline & Database Architectures</strong> · Character: Marie Curie · 1 beast · 1 atom</summary>

![Memory Palace 7](images/data-analyst/7.jpg)

<p><em>1 beast · 1 Knowledge Atom</em></p>

#### Knowledge Atoms

### [Ad] adder

<img src="../../web/assets/beast-thumbs/adder.png" alt="adder" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [I store data without fixed <img alt="[tables]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Btables%5D&color=9fd4ff" />(cloud)] [using flexible <img alt="[formats]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bformats%5D&color=ffd966" />(origami)] [for horizontal <img alt="[scaling]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bscaling%5D&color=ffd966" />(accordion)] — Note: Uses documents, key-value pairs, wide columns, or graphs to adapt easily to changing schemas.

**Quote**
“A NoSQL database stores data without fixed tables, using flexible formats like documents, key-value pairs, wide columns, or graphs.”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 6: Power Query Transformations & Step Reuse</strong> · Character: Tim Berners-Lee · 5 beasts · 6 atoms</summary>

![Memory Palace 6](images/data-analyst/6.jpg)

<p><em>5 beasts · 6 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [Z] Zeus

<img src="../../web/assets/beast-thumbs/zeus.png" alt="Zeus" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [My unpivot step automatically <img alt="[deletes]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdeletes%5D&color=9fd4ff" />(trash can) all rows] [with <img alt="[null]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bnull%5D&color=ffd966" />(ghost) values] — Note: Power Query has no built-in setting or parameter to turn off this automatic removal.

**Quote**
“now the characteristic of the unpivot function in power query is that by the default it actually removes the null values”

### [Y] yak

<img src="../../web/assets/beast-thumbs/yak.png" alt="yak" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [I unpivot multiple <img alt="[columns]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcolumns%5D&color=9fd4ff" />(pillar) into rows] [to <img alt="[model]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bmodel%5D&color=ffd966" />(clay) my data more easily] — Note: Column headers become an attribute column paired with a single value column.

**Quote**
“so unpivoting means I have columns and I want to see those columns in the rows”

### [Ac] acorn

<img src="../../web/assets/beast-thumbs/acorn.png" alt="acorn" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [I <img alt="[store]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bstore%5D&color=9fd4ff" />(chest) data] [in rigid <img alt="[tables]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Btables%5D&color=ffd966" />(grid) of rows and columns] [linked by predefined <img alt="[relationships]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Brelationships%5D&color=ffd966" />(chain)] — Note: Enforces schemas and data integrity using SQL validation rules.

**Quote**
“A relational database stores data in fixed tables made of rows and columns, linked together by predefined relationships.”

### [Ab] Abyssinian cat

<img src="../../web/assets/beast-thumbs/abyssinian_cat.png" alt="Abyssinian cat" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

🟦 **Z1 · Reusing query steps**

**Concept**
💡 [I <img alt="[reuse]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Breuse%5D&color=9fd4ff" />(stamp) my query steps] [across different <img alt="[files]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bfiles%5D&color=ffd966" />(binder)] [sharing the exact same table <img alt="[structure]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bstructure%5D&color=ffd966" />(twin)] — Note: Identical column headers and data formats are required so the query steps run without error.

**Quote**
“since the format of both files are the same I want to apply the exact same steps to my second file”

---

🟦 **Z2 · Copying transformation steps**

**Concept**
💡 [I <img alt="[copy]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcopy%5D&color=9fd4ff" />(scissors) all transformation steps] [below the initial <img alt="[source]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bsource%5D&color=ffd966" />(anchor) line] [from the Advanced <img alt="[Editor]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5BEditor%5D&color=ffd966" />(scroll)] — Note: The first line contains the specific file source path that must not overwrite the new table's connection.

**Quote**
“the First Line Imports the CSV files so we don't want this step we want to grab all the steps below it Ctrl C to copy”

### [Aa] aardvark

<img src="../../web/assets/beast-thumbs/aardvark.png" alt="aardvark" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [I replace nulls with a temporary <img alt="[placeholder]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bplaceholder%5D&color=9fd4ff" />(scarecrow)] [before <img alt="[unpivoting]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bunpivoting%5D&color=ffd966" />(jack)] [to <img alt="[swap]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bswap%5D&color=ffd966" />(boomerang) them back afterward] — Note: This prevents Power Query from dropping rows during the unpivot step.

**Quote**
“you can select the columns where you have the null values and you need to replace those with a placeholder”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 5: M Engine & Syntax</strong> · Character: Stephen Hawking · 5 beasts · 5 atoms</summary>

![Memory Palace 5](images/data-analyst/5.jpg)

<p><em>5 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [X] Xena, warrior woman

<img src="../../web/assets/beast-thumbs/xena_warrior_woman.png" alt="Xena, warrior woman" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [The OData connector suffers from <img alt="[network]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bnetwork%5D&color=9fd4ff" />(spider web) latency] [because it makes redundant metadata <img alt="[calls]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcalls%5D&color=ffd966" />(megaphone)] [at <img alt="[runtime]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bruntime%5D&color=ffd966" />(running shoes).]

**Quote**
“the problem uh with odata is that at runtime it has to make metadata calls to basically get the metadata and that makes a second call”

### [W] wombat

<img src="../../web/assets/beast-thumbs/wombat.png" alt="wombat" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [The OData connector uses a <img alt="[discovery]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bdiscovery%5D&color=9fd4ff" />(binoculars) mechanism] [to automatically determine the <img alt="[schema]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bschema%5D&color=ffd966" />(skeleton)] [of the external <img alt="[table]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Btable%5D&color=ffd966" />(picnic table).]

**Quote**
“odata has a discovery mechanism you know where now power query is kind of looking at the table and figuring out what it is”

### [V] vulture

<img src="../../web/assets/beast-thumbs/vulture.png" alt="vulture" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [I use query <img alt="[folding]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bfolding%5D&color=9fd4ff" />(origami)] [to <img alt="[push]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bpush%5D&color=ffd966" />(bulldozer) transformation work] [back to the data <img alt="[source]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bsource%5D&color=ffd966" />(well)] [to maximize processing <img alt="[efficiency]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Befficiency%5D&color=ffd966" />(stopwatch).]

**Quote**
“the idea of query folding is that you want the power query mashup engine you know to be as efficient as possible so the mashup engine will push work back to the data source”

### [U] unicorn

<img src="../../web/assets/beast-thumbs/unicorn.png" alt="unicorn" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [I avoid using <img alt="[spaces]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bspaces%5D&color=9fd4ff" />(vacuum)] [in my step <img alt="[names]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bnames%5D&color=ffd966" />(name tag)] [to keep the underlying M code <img alt="[clean]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bclean%5D&color=ffd966" />(sponge).] - Note: Spaces force the variables to be wrapped in quotes and a hash sign.

**Quote**
“if you put spaces in your step names it makes the applied steps thing look better yeah but it makes the uh you know m code look a lot worse”

### [T] toucan

<img src="../../web/assets/beast-thumbs/toucan.png" alt="toucan" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [I use the power query <img alt="[mashup]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bmashup%5D&color=9fd4ff" />(blender) engine] [as the underlying <img alt="[technology]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Btechnology%5D&color=ffd966" />(engine block)] [to <img alt="[execute]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bexecute%5D&color=ffd966" />(lightning bolt) my data queries.]

**Quote**
“at the base you know of this technology is something called the power query mashup engine that's the thing that executes your query”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 4: Data Architectures 2</strong> · Character: Neo (The Matrix) · 5 beasts · 5 atoms</summary>

![Memory Palace 4](images/data-analyst/4.png)

<p><em>5 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [S] skull

<img src="../../web/assets/beast-thumbs/skull.png" alt="skull" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Data Mart: A <img alt="[specialized]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bspecialized%5D&color=9fd4ff" />(scalpel) subset of a data warehouse] [<img alt="[focused]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bfocused%5D&color=ffd966" />(spotlight) on a specific business line or department[cite: 1].] — Note: Examples include Finance or Marketing[cite: 1].

**Quote**
“A subset of a data warehouse focused on a specific business line or department (e.g., Finance, Marketing)."[cite: 1]”

### [R] rat

<img src="../../web/assets/beast-thumbs/rat.png" alt="rat" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Data Swamp: A poorly <img alt="[governed]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bgoverned%5D&color=9fd4ff" />(broken crown) data lake] [where data is <img alt="[uncataloged]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Buncataloged%5D&color=ffd966" />(shredder), undocumented,] [and difficult to <img alt="[retrieve]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bretrieve%5D&color=ffd966" />(fishing rod)[cite: 1].]

**Quote**
“A poorly governed data lake where data is uncataloged, undocumented, and difficult to retrieve."[cite: 1]”

### [Q] Quetzalcoatl

<img src="../../web/assets/beast-thumbs/quetzalcoatl.png" alt="Quetzalcoatl" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Data Lakehouse: A <img alt="[hybrid]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bhybrid%5D&color=9fd4ff" />(centaur) architecture] [combining the scale and <img alt="[flexibility]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bflexibility%5D&color=ffd966" />(rubber band) of a data lake] [with the <img alt="[reliability]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Breliability%5D&color=ffd966" />(vault) of a warehouse[cite: 1].] — Note: Includes ACID features like Delta Lake on top of cloud storage[cite: 1].

**Quote**
“Hybrid architecture combining the scale and flexibility of a data lake with the reliability and ACID features of a warehouse (e.g., Delta Lake on top of cloud storage)."[cite: 1]”

### [P] panther

<img src="../../web/assets/beast-thumbs/panther.png" alt="panther" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Data Warehouse: A highly <img alt="[structured]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bstructured%5D&color=9fd4ff" />(filing cabinet), schema-on-write repository] [optimized for SQL <img alt="[analytics]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Banalytics%5D&color=ffd966" />(magnifying glass)] [and business <img alt="[intelligence]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bintelligence%5D&color=ffd966" />(briefcase)[cite: 1].]

**Quote**
“Highly structured, schema-on-write repository optimized for SQL analytics and business intelligence."[cite: 1]”

### [O] owl

<img src="../../web/assets/beast-thumbs/owl.png" alt="owl" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 [Data Lake: I <img alt="[store]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bstore%5D&color=9fd4ff" />(bucket) structured, semi-structured, and unstructured raw data] [in a <img alt="[centralized]" src="https://img.shields.io/static/v1?style=flat-square&label=&message=%5Bcentralized%5D&color=ffd966" />(swimming pool), low-cost object storage system[cite: 1].]

**Quote**
“Centralized storage for structured, semi-structured, and unstructured raw data in low-cost object storage."[cite: 1]”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 3: Data Architectures 1</strong> · Character: GitHub Copilot · 5 beasts · 5 atoms</summary>

![Memory Palace 3](images/data-analyst/3.png)

<p><em>5 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [N] Neanderthal

<img src="../../web/assets/beast-thumbs/neanderthal.png" alt="Neanderthal" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 Data Store: Any repository that persists and manages a collection of data[cite: 1]. — Note: Ranges from traditional relational databases to cloud object storage[cite: 1].

**Quote**
“An umbrella term for any repository that persists and manages a collection of data, ranging from traditional relational databases and NoSQL key-value stores to file systems, caches, and cloud object storage."[cite: 1]”

### [M] marmoset

<img src="../../web/assets/beast-thumbs/marmoset.png" alt="marmoset" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 Load Phase: I write the fully processed data into the target repository[cite: 1]. — Note: This makes it immediately available for business intelligence tools and analytics[cite: 1].

**Quote**
“Writing the fully processed data into the target repository, making it immediately available for business intelligence tools and analytics."[cite: 1]”

### [L] lion

<img src="../../web/assets/beast-thumbs/lion.png" alt="lion" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 Transform Phase: I clean, structure, and enrich the data in a temporary staging area to meet business requirements[cite: 1]. — Note: Includes standardizing formats, filtering anomalies, joining tables, and mapping to a predefined schema[cite: 1].

**Quote**
“Cleaning, structuring, and enriching the data in a temporary staging area to meet business and analytical requirements. This includes standardizing date formats, filtering anomalies, joining tables, applying business logic, and mapping the data to a predefined schema."[cite: 1]”

### [K] kitten

<img src="../../web/assets/beast-thumbs/kitten.png" alt="kitten" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 Extract Phase: I retrieve raw data from various origin systems during the first pipeline phase[cite: 1]. — Note: Sources include operational databases, SaaS applications, APIs, or flat files[cite: 1].

**Quote**
“Retrieving raw data from various source systems, such as operational databases, SaaS applications, APIs, or flat files."[cite: 1]”

### [J] jester

<img src="../../web/assets/beast-thumbs/jester.png" alt="jester" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 ETL Pipeline: I move raw data from multiple disparate sources into a centralized target system using an automated integration pipeline[cite: 1]. — Note: The target is typically a Data Warehouse or Data Mart[cite: 1].

**Quote**
“ETL is an automated data integration process that moves data from multiple disparate sources into a centralized target system, typically a Data Warehouse or Data Mart."[cite: 1]”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 2: Data Analytics Fundamentals</strong> · Character: Alan Turing · 5 beasts · 5 atoms</summary>

![Memory Palace 2](images/data-analyst/2.jpg)

<p><em>5 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [I] imp

<img src="../../web/assets/beast-thumbs/imp.png" alt="imp" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 I dig deeper beyond number crunching to understand underlying causes. — Note: Balances quantitative technical skills like SQL with qualitative critical reasoning.

**Quote**
“It's not just about crunching the numbers and sharing your data. Sometimes you'll need to dig deeper to understand really what's going on.”

### [H] Hydra

<img src="../../web/assets/beast-thumbs/hydra.png" alt="Hydra" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 I turn refined data into valuable business insights. — Note: The culmination of a five-step pipeline spanning definition, collection, cleaning, analysis, and presentation.

**Quote**
“The final step in this process is where data is turned into valuable business insights.”

### [G] goat

<img src="../../web/assets/beast-thumbs/goat.png" alt="goat" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 I evaluate team data to shape future business strategies. — Note: Serves as the cross-functional bridge between raw metrics and executive planning.

**Quote**
“Work as part of a team to evaluate and analyze key data that will be used to shape future business strategies.”

### [F] frog

<img src="../../web/assets/beast-thumbs/frog.png" alt="frog" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 I apply analytics to speed decisions, cut costs, and develop products. — Note: Optimizes operational performance and predicts market behavior.

**Quote**
“Broadly speaking, data analytics is used to make faster and better business decisions, to reduce overall business costs, and to develop new and innovative products and services.”

### [E] eagle

<img src="../../web/assets/beast-thumbs/eagle.png" alt="eagle" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 I analyze raw data to extract useful business insights. — Note: Transforms unorganized datasets into actionable intelligence.

**Quote**
“Data analytics is the process of analyzing raw data so that we can pull out insights which are useful to companies.”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>

<details open>
<summary><strong>Memory Palace 1: Data Analyst vs Data Scientist Foundations</strong> · Character: Ada Lovelace · 4 beasts · 5 atoms</summary>

![Memory Palace 1](images/data-analyst/1.jpg)

<p><em>4 beasts · 5 Knowledge Atoms</em></p>

#### Knowledge Atoms

### [D] dragon

<img src="../../web/assets/beast-thumbs/dragon.png" alt="dragon" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 Core tooling and workflow: SQL/BI dashboards isolating friction vs Python/ML/pipelines training real-time automated churn classifiers.

**Quote**
“Queries an e-commerce database using SQL to identify that customer churn increased by 15% in Q3, isolates the drop to checkout friction, and presents a visual dashboard to product managers.”

### [C] cat

<img src="../../web/assets/beast-thumbs/cat.png" alt="cat" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 Data types handled: structured relational tables/spreadsheets vs structured and unstructured multi-modal streams (images, audio, clickstreams).

**Quote**
“Structured data from relational databases and warehouses”

### [B] bird of paradise

<img src="../../web/assets/beast-thumbs/bird_of_paradise.png" alt="bird of paradise" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

**Concept**
💡 Key objectives: dashboards and KPI tracking vs ML pipelines and algorithmic automation.

**Quote**
“Dashboards, KPI tracking, business reports, trend identification”

### [A] Arachne

<img src="../../web/assets/beast-thumbs/arachne.png" alt="Arachne" width="64" height="64" style="vertical-align:middle;height:64px;width:64px;" />

🟦 **Z1 · Head**

**Concept**
💡 Core objective difference: Data Analysts explain past/present to inform decisions; Data Scientists build predictive models to forecast and automate.

**Quote**
“The primary difference lies in their core focus: a Data Analyst examines historical data to answer specific business questions and identify existing trends, while a Data Scientist builds predictive models, algorithms, and experiments to forecast future outcomes and discover unknowns.”

---

🟦 **Z2 · Forelimbs**

**Concept**
💡 Primary focus: descriptive and diagnostic (What happened? Why?) vs predictive and prescriptive modeling (What will happen? How to optimize?).

**Quote**
“Descriptive and diagnostic analysis (answering "What happened?" and "Why?")”

#### Notes

_No notes._

#### Gallery

_No gallery images._

</details>
