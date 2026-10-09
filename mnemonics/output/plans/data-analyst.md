<!-- synced from plans.csv (PLAN_0007) / study data-analyst -->

# Data Analyst Transition Study Plan

## 📃 Backlog

- [✅] Data lake

- [✅] ETL (Extract, Transform, Load)

- [✅] schema

- [✅] Visual Information Class PDFs

- [ ] Fact tables and dimension tables

- [ ] snow flake

- [ ] Experiência operacional de vários anos em processos logísticos e SAP (ERP - módulos SD / PP / MM, APO)

- [ ] data bricks project

- [ ] ACID

- [ ] Delta Lake

- [ ] Spark

- [ ] Database view

## Phase 1: Advanced Tabular Modeling and Automation

### 1. Interface and Navigation

- [✅] Master Excel keyboard shortcuts for navigation without relying on the mouse
- [✅] Apply tabular best practices for large dataset management
- [✅] Set up a clean workbook structure suitable for analyst workflows
- [✅] Practice navigating large sheets with Ctrl/Shift selection patterns used by analysts

**🔗 Resources:**

- ▶ [Excel for Data Analysts | Module 1.1: Shortcuts and Best Practices](https://www.youtube.com/watch?v=Pwp1j-FN_DE)

### 2. Advanced Functions

- [✅] Master XLOOKUP with correct lookup-array vs return-array order
- [✅] Build INDEX/MATCH lookups for flexible retrieval
- [✅] Write nested IF / logical operators for conditional modeling

**🔗 Resources:**

- 🔗 [XLOOKUP function (Microsoft Support)](https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929)
- 🔗 [INDEX function (Microsoft Support)](https://support.microsoft.com/en-us/office/index-function-a5dcf0dd-996d-40a3-be95-3a391f4e6b2f)
- 🔗 [MATCH function (Microsoft Support)](https://support.microsoft.com/en-us/office/match-function-e8dffd45-c762-47d6-bf89-533f4a37673a)

### 3. Dimensional Aggregation

- [✅] Build Pivot Tables from transactional rows into high-level metrics
- [✅] Compute percentage distributions and relative shares in pivots
- [✅] Apply dynamic number formatting for stakeholder-ready summaries
- [✅] Place fields in Rows, Columns, and Values and check how each area changes the pivot

**🔗 Resources:**

- 🔗 [Create a PivotTable to analyze worksheet data (Microsoft Support)](https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576)
- ▶ [Excel Number Formatting for Data Analysts | Beginner Tutorial](https://www.youtube.com/watch?v=WqZ90oyB0hc)

### 4. Automated ETL Pipelines

- [✅] Use Power Query to extract, transform, and load dirty data
- [✅] Learn M language basics for repeatable transformations
- [✅] Unpivot data and handle null/missing values in the query editor
- [✅] Build a reusable Power Query that strips bad formatting and merges sources

**🔗 Resources:**

- ▶ [Excel for Data Analysts: Master ETL with Power Query](https://www.youtube.com/watch?v=RxYwM2x4VCs)
- 🔗 [Power Query documentation (Microsoft Learn)](https://learn.microsoft.com/en-us/power-query/)
- 🔗 [Power Query M formula language reference](https://learn.microsoft.com/en-us/powerquery-m/)

### 5. Capstone Synthesis

- [✅] Complete an end-to-end tabular analysis on a real-world dataset in Excel
- [✅] Produce a cleaned report ready for review
- [✅] Save the final cleaned Excel workbook as a stakeholder-facing summary

**🔗 Resources:**

- 🔗 [Luke Barousse: Excel for Data Analytics (Course Page)](https://www.lukebarousse.com/courses)

## Phase 2: Relational Database Extraction and Querying

### 1. Database Architecture

- [✅] Install and configure PostgreSQL (or practice MySQL) locally
- [✅] Understand schemas, tables, and primary/foreign key relationships
- [✅] Write SELECT, FROM, and WHERE queries to extract and filter rows

**🔗 Resources:**

- ▶ [SQL Tutorial for Beginners (Alex The Analyst)](https://www.youtube.com/watch?v=h0nxCDiD-zg)
- ▶ [SQL Basics Tutorial For Beginners | Select + From Statements](https://www.youtube.com/watch?v=PyYgERKq25I)
- 📄 [PostgreSQL Tutorial (official)](https://www.postgresql.org/docs/current/tutorial.html)
- 🔗 [Mode Analytics SQL Tutorial (written)](https://mode.com/sql-tutorial)

### 2. Aggregation Logic

- [✅] Aggregate with GROUP BY and filter groups with HAVING
- [✅] Apply SUM, AVG, COUNT, MIN, and MAX correctly

**🔗 Resources:**

- ▶ [Learn SQL Beginner to Advanced in Under 4 Hours](https://www.youtube.com/watch?v=OT1RErkfLNQ)

### 3. Set Theory and Joins

- [✅] Write INNER JOIN queries across related tables
- [✅] Write LEFT JOIN queries and interpret NULL placeholders for non-matches
- [✅] Use OUTER JOIN patterns and table aliases for readable multi-table SQL
- [✅] Contrast INNER JOIN and LEFT JOIN, including unmatched rows returned as NULL

**🔗 Resources:**

- ▶ [Intermediate SQL Tutorial | Aliasing](https://www.youtube.com/watch?v=Dk7he_yEs4U)
- 📄 [PostgreSQL: Joins](https://www.postgresql.org/docs/current/tutorial-join.html)

### 4. Query Modularity

- [✅] Refactor nested subqueries into readable Common Table Expressions (CTEs)
- [✅] Format long queries so each clause stays readable
- [✅] Replace a long nested query with named Common Table Expressions

**🔗 Resources:**

- ▶ [SQL for Data Analytics - Learn SQL in 4 Hours (Luke Barousse)](https://www.youtube.com/watch?v=7mz73uXD9DA)
- 📄 [PostgreSQL: WITH Queries (CTEs)](https://www.postgresql.org/docs/current/queries-with.html)

### 5. Advanced Analytics

- [✅] Apply Window Functions: ROW_NUMBER, RANK, SUM() OVER(PARTITION BY)
- [✅] Use LEAD and LAG for row-relative comparisons without collapsing grain
- [✅] Practice PARTITION BY windows that preserve row grain while adding running metrics

**🔗 Resources:**

- 🔗 [Intermediate SQL for Data Analytics (Luke Barousse)](https://www.lukebarousse.com/int-sql)
- 📄 [PostgreSQL: Window Functions](https://www.postgresql.org/docs/current/tutorial-window.html)

### 6. Portfolio Integration

- [✅] Run Exploratory Data Analysis in SQL on a real-world dataset (e.g. COVID / demographic data)
- [✅] Document findings in a clean GitHub markdown logbook
- [✅] Import an unstructured public dataset into PostgreSQL for portfolio EDA

**🔗 Resources:**

- ▶ [Data Analyst Portfolio Project | SQL Data Exploration | Project 1/4](https://www.youtube.com/watch?v=qfyynHBFOsM)

## Phase 3: Business Intelligence and Dimensional Visualization

### 1. Interface and Data Ingestion

- [✅] Install Power BI Desktop and import practice datasets
- [✅] Navigate the workspace: Report, Data, and Model views
- [✅] Leverage UX background: treat every dashboard as a user interface with journey mapping

**🔗 Resources:**

- ▶ [Power BI Tutorial for Beginners (Kevin Stratvert)](https://www.youtube.com/watch?v=NNSHu0rkew8)
- 🔗 [Get started with Power BI Desktop (Microsoft Learn)](https://learn.microsoft.com/en-us/power-bi/fundamentals/desktop-getting-started)

### 2. Advanced Power Query

- [✅] Transform datasets in Power BI Power Query (types, splits, merges)
- [✅] Handle API or external source integrations when needed
- [✅] Reuse Phase 1 Power Query skills in Power BI, which uses the same engine

**🔗 Resources:**

- ▶ [How to use Power Query in Power BI (Alex The Analyst)](https://www.youtube.com/watch?v=gP-AxNi6uxo)

### 3. Dimensional Modeling

- [✅] Construct a Star Schema with a central Fact table and Dimension tables
- [✅] Manage 1-to-Many cardinality and active vs inactive relationships
- [✅] Relate each dimension to the fact table so filters flow through one active relationship

**🔗 Resources:**

- 🔗 [Understand star schema and the importance for Power BI (Microsoft Learn)](https://learn.microsoft.com/en-us/power-bi/guidance/star-schema)
- 🔗 [Kimball dimensional modeling techniques (primer)](https://www.kimballgroup.com/data-warehouse-business-intelligence-resources/kimball-techniques/dimensional-modeling-techniques/)

### 4. DAX and Evaluation Context

- [ ] Distinguish row context from filter context in DAX
- [ ] Master CALCULATE, FILTER, and iterator functions like SUMX
- [ ] Build Time Intelligence measures with DATEADD and related functions
- [ ] Use CALCULATE to change the filter context of a measure

**🔗 Resources:**

- ▶ [Power BI Intro to Data Analysis Advanced Tutorial (Learnit)](https://www.youtube.com/watch?v=GpAXwNaUiLY)
- 🔗 [Row context and filter context in DAX (SQLBI)](https://www.sqlbi.com/articles/row-context-and-filter-context-in-dax/)
- 🔗 [Introducing CALCULATE in DAX (SQLBI)](https://www.sqlbi.com/articles/introducing-calculate-in-dax/)

### 5. Dashboard UX and Portfolio

- [ ] Apply UX principles: visual hierarchy, F/Z-pattern, Gestalt layout
- [ ] Choose chart types correctly (bar, line, scatter) and avoid chart junk
- [ ] Create drill-throughs and publish a portfolio dashboard to the service
- [ ] Lead stakeholder eye to the primary KPI first using hierarchy and contrast

**🔗 Resources:**

- ▶ [Learn Power BI in Under 3 Hours (Full Project)](https://www.youtube.com/watch?v=I0vQ_VLZTWg)
- ▶ [Power BI Full Course Tutorial (8+ Hours)](https://www.youtube.com/watch?v=e6QD8lP-m6E)
- 🔗 [Visualization types in Power BI (Microsoft Learn)](https://learn.microsoft.com/en-us/power-bi/visuals/power-bi-visualization-types-for-reports-and-q-and-a)

## Phase 4: Programmatic Manipulation and Advanced Transformation

### 1. Environment Setup

- [ ] Install Anaconda and open Jupyter Notebooks
- [ ] Import pandas and numpy in a clean analysis environment
- [ ] Use one notebook environment and declare imports at the top

**🔗 Resources:**

- 🔗 [Python for Data Analytics (Luke Barousse)](https://www.lukebarousse.com/python)
- 📄 [Installing pandas (official)](https://pandas.pydata.org/docs/getting_started/install.html)

### 2. Data Structures

- [ ] Understand a DataFrame as a collection of Series
- [ ] Index, select, filter, and sort DataFrame columns and rows
- [ ] Practice slicing and boolean filtering without mutating source frames carelessly

**🔗 Resources:**

- ▶ [Learn PANDAS in 5 minutes | Pandas Ultraquick Tutorial (Keith Galli)](https://www.youtube.com/watch?v=m1_33jhhiLE)
- ▶ [Keith Galli channel (Pandas tutorials)](https://www.youtube.com/c/KGMIT/videos)
- 📄 [10 minutes to pandas (official)](https://pandas.pydata.org/docs/user_guide/10min.html)

### 3. Exploratory Data Analysis

- [ ] Profile datasets with .info(), .describe(), and .isnull().sum()
- [ ] Interpret mean, std, and percentiles from statistical summaries
- [ ] Check dtypes, missing values, and row counts before transforming a new dataset

**🔗 Resources:**

- ▶ [Exploratory Data Analysis in Pandas (Alex The Analyst)](https://www.youtube.com/watch?v=Liv6eeb1VfE)
- 📄 [Essential basic functionality (pandas docs)](https://pandas.pydata.org/docs/user_guide/basics.html)

### 4. Vectorization and Aggregation

- [ ] Replace iterative for-loops with vectorized column operations
- [ ] Aggregate with groupby() and combine tables with merge()
- [ ] Refuse row-wise Python loops for column transforms; prefer vectorized arrays

**🔗 Resources:**

- 📄 [Group by: split-apply-combine (pandas docs)](https://pandas.pydata.org/docs/user_guide/groupby.html)
- 📄 [Merge, join, concatenate and compare (pandas docs)](https://pandas.pydata.org/docs/user_guide/merging.html)

### 5. Statistical Visualization

- [ ] Compute correlation matrices on cleaned numeric features
- [ ] Plot heatmaps with Seaborn for exploratory visualization
- [ ] Keep matplotlib/seaborn for EDA only — do not introduce ML libraries

**🔗 Resources:**

- ▶ [Data Analyst Portfolio Project | Correlation in Python](https://www.youtube.com/watch?v=iPYVYBtUTyE)
- 📄 [seaborn tutorial (official)](https://seaborn.pydata.org/tutorial.html)

## Phase 5: Synthesis, Storytelling, and Portfolio Development

### 1. Audience Empathy and Context

- [ ] Map stakeholder personas: Executive, Operations Manager, Marketing Lead
- [ ] Build a narrative arc that ends on actionable insight, not exploration process
- [ ] Tailor pitch length and detail to Executive, Operations, and Marketing audiences

**🔗 Resources:**

- ▶ [Master the Art of Data Storytelling (Official Book Ginger)](https://www.youtube.com/watch?v=53gZrM42ig8)
- 📖 [Storytelling with Data — Cole Nussbaumer Knaflic (book)](https://www.storytellingwithdata.com/book)

### 2. Cognitive Load Reduction

- [ ] Eliminate gridlines, excess labels, and 3D/exploding pie charts (Knaflic)
- [ ] Use pre-attentive attributes (color, size, contrast) to force focus on the key metric
- [ ] Apply Storytelling with Data: one stark highlight (red line) on otherwise muted slides

**🔗 Resources:**

- ▶ [Storytelling with Data | Cole Nussbaumer Knaflic | Talks at Google](https://www.youtube.com/watch?v=8EMW7io4rSI)
- 🔗 [storytellingwithdata.com blog / examples](https://www.storytellingwithdata.com/blog)

### 3. SQL and BI Project Execution

- [ ] Complete SQL portfolio project: import raw data, query, window functions, document in GitHub
- [ ] Feed SQL insights into a Power BI Star Schema with DAX measures and minimalist visuals
- [ ] Chain SQL exploration output into a dimensional Power BI model for executives

**🔗 Resources:**

- ▶ [Data Analyst Portfolio Project 1/4 SQL Exploration](https://www.youtube.com/watch?v=qfyynHBFOsM)
- ▶ [Learn Power BI in Under 3 Hours (Full Project)](https://www.youtube.com/watch?v=I0vQ_VLZTWg)

### 4. Python Project Execution

- [ ] Complete Jupyter portfolio: clean a messy dataset with Pandas vectorization
- [ ] Document correlations and transformations as a readable executive-facing notebook
- [ ] Prove programmatic capability beyond BI tools with messy-data cleaning notebook

**🔗 Resources:**

- ▶ [Data Analyst Portfolio Project | Correlation in Python](https://www.youtube.com/watch?v=iPYVYBtUTyE)

### 5. Public Deployment

- [ ] Create a free portfolio website showcasing the three projects
- [ ] Establish a public GitHub repository with clean README and documented queries
- [ ] Write portfolio READMEs that show commercial utility, not only academic theory

**🔗 Resources:**

- 🔗 [GitHub Docs: About READMEs](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/about-readmes)
- 🔗 [Creating a GitHub Pages site (GitHub Docs)](https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site)

## Phase 6: Azure Databricks (Principal Platform)

### 1. Must

- [ ] Orient the Databricks workspace, notebooks, and Unity Catalog catalogs and schemas
- [ ] Query lakehouse tables with Spark SQL
- [ ] Transform data with PySpark DataFrames
- [ ] Create Delta tables and use time travel to inspect earlier versions
- [ ] Build a medallion pipeline through bronze, silver, and gold
- [ ] Map how a Databricks workspace, storage, and Unity Catalog sit on Azure

**🔗 Resources:**

- 🔗 [Get Started with Databricks for Data Engineering](https://customer-academy.databricks.com/learn/course/external/view/classroom/1511/get-started-with-databricks-for-data-engineering)

### 2. Good to know

- [ ] Schedule pipelines with Lakeflow Jobs
- [ ] Ingest data with Lakeflow Connect
- [ ] Enforce schema and data-quality checks on Delta tables
- [ ] Version notebooks with Git on the workspace
- [ ] Connect Power BI to Databricks SQL

**🔗 Resources:**

- 🔗 [Data Engineering with Databricks](https://www.databricks.com/training/catalog/data-engineering-with-databricks-911)

### 3. Learn after

- [ ] Track experiments with MLflow
- [ ] Serve features from a feature store
- [ ] Ingest streaming data into Delta
- [ ] Apply Unity Catalog row filters and data sharing

**🔗 Resources:**

- 🔗 [Databricks training catalog](https://www.databricks.com/learn/training/home)

## Phase 7: Microsoft Fabric (Integration Layer)

### 1. Must

- [ ] Create a Fabric workspace and place data in OneLake
- [ ] Build a lakehouse on Delta tables
- [ ] Ingest and transform with Dataflows Gen2 using Power Query
- [ ] Query the SQL analytics endpoint and build a semantic model
- [ ] Connect a Power BI report with Direct Lake on that lakehouse

**🔗 Resources:**

- 🔗 [Implement a Lakehouse with Microsoft Fabric](https://learn.microsoft.com/en-us/training/paths/implement-lakehouse-microsoft-fabric/)

### 2. Good to know

- [ ] Orchestrate movement with Data Factory pipelines
- [ ] Transform data in Fabric Spark notebooks
- [ ] Reference existing data with OneLake shortcuts
- [ ] Maintain Delta tables with OPTIMIZE and VACUUM

**🔗 Resources:**

- 🔗 [Lakehouse end-to-end scenario](https://learn.microsoft.com/en-us/fabric/data-engineering/tutorial-lakehouse-introduction)

### 3. Learn after

- [ ] Discover streaming sources in Real-Time hub
- [ ] Combine Direct Lake tables with Import tables in a composite model
- [ ] Choose a warehouse versus a lakehouse for a given workload

**🔗 Resources:**

- 🔗 [Direct Lake overview](https://learn.microsoft.com/en-us/fabric/fundamentals/direct-lake-overview)

## Phase 8: SAP Data for Bosch Roles

### 1. Must

- [ ] Explain SAP Datasphere spaces, the Data Builder, and an analytic model
- [ ] Describe how SAP Analytics Cloud consumes a Datasphere model
- [ ] Explain SAP Business Data Cloud and that Datasphere sits inside it
- [ ] Map the ERP handoff among SD (sales), MM (procurement), and PP (production)

**🔗 Resources:**

- 🔗 [Exploring SAP Datasphere](https://learning.sap.com/courses/exploring-sap-datasphere)

### 2. Good to know

- [ ] Distinguish replication flows from transformation flows
- [ ] Describe how spaces and authorization separate data
- [ ] Explain Delta Share as zero-copy sharing out of Business Data Cloud
- [ ] Map APO planning on those modules: DP/GATP, SNP, and PPDS

**🔗 Resources:**

- 🔗 [SAP Business Data Cloud design principles](https://learning.sap.com/courses/introducing-sap-business-data-cloud/describing-sap-business-data-cloud-and-its-design-principles_d92361e3-c958-4010-9d0a-2cbcc7263d4e)

### 3. Learn after

- [ ] Explore the Business Builder and the data marketplace
- [ ] Build an SAP Analytics Cloud story on a Datasphere model
- [ ] Study module configuration and Transportation Management only when a role requires it

**🔗 Resources:**

- 🔗 [SAP Business Data Cloud learning](https://learning.sap.com/products/business-data-cloud)

## Phase 9: Snowflake (Secondary)

### 1. Must

- [ ] Explain Snowflake storage separated from compute
- [ ] Size a virtual warehouse and relate it to credit consumption
- [ ] Contrast Snowflake SQL analytics with Databricks engineering and AI

**🔗 Resources:**

- 🔗 [Snowflake Level Up: First Concepts](https://learn.snowflake.com/en/pages/level-up-track)

### 2. Good to know

- [ ] Share data securely without copying it
- [ ] Set a resource monitor and read credit cost
- [ ] Recover data with time travel and zero-copy clone

**🔗 Resources:**

- 🔗 [Snowflake virtual warehouses](https://docs.snowflake.com/en/user-guide/warehouses)

### 3. Learn after

- [ ] Defer Snowpark, dynamic tables, and Cortex unless a Bosch project asks for Snowflake
- [ ] Use Snowpark DataFrames
- [ ] Build a pipeline with dynamic tables
- [ ] Call Cortex functions on warehouse data

**🔗 Resources:**

- 🔗 [Zero to Snowflake](https://www.snowflake.com/en/developers/guides/zero-to-snowflake/)
