<!-- synced from plans.csv (PLAN_0007) / study data-analyst -->

# Data Analyst Transition Study Plan

## 📃 Backlog

- [ ] Archive source curriculum PDF (Data Analyst Transition Study Plan) under research/ for reference

- [ ] Design Memory Palace blueprint: campus or corporate HQ with five wings (Atrium, Archive, Exhibition Hall, Lab, Auditorium) — palaces left blank until ready

- [ ] Hard constraint: stay in Data Analysis only — no ML, predictive modeling, Scikit-Learn, or TensorFlow

## Phase 1: Advanced Tabular Modeling and Automation (Main Corporate Atrium)

### 1. Interface and Navigation

- [ ] Master Excel keyboard shortcuts for navigation without relying on the mouse
- [ ] Apply tabular best practices for large dataset management
- [ ] Set up a clean workbook structure suitable for analyst workflows
- [ ] Practice navigating large sheets with Ctrl/Shift selection patterns used by analysts

**🔗 Resources:**

- ▶ [Excel for Data Analysts | Module 1.1: Shortcuts and Best Practices](https://www.youtube.com/watch?v=Pwp1j-FN_DE)

### 2. Advanced Functions

- [ ] Master XLOOKUP with correct lookup-array vs return-array order
- [ ] Build INDEX/MATCH lookups for flexible retrieval
- [ ] Write nested IF / logical operators for conditional modeling
- [ ] Encode XLOOKUP mentally as receptionist: search key -> lookup array -> return array

**🔗 Resources:**

- 🔍 [Kevin Stratvert: Advanced Excel (search Full Course on channel)](https://www.youtube.com/results?search_query=Kevin+Stratvert+Advanced+Excel+Full+Course)

### 3. Dimensional Aggregation

- [ ] Build Pivot Tables from transactional rows into high-level metrics
- [ ] Compute percentage distributions and relative shares in pivots
- [ ] Apply dynamic number formatting for stakeholder-ready summaries
- [ ] Practice dragging fields into Rows, Columns, and Values as physical pivot rotation

**🔗 Resources:**

- ▶ [Excel Number Formatting for Data Analysts | Beginner Tutorial](https://www.youtube.com/watch?v=WqZ90oyB0hc)

### 4. Automated ETL Pipelines

- [ ] Use Power Query to extract, transform, and load dirty data
- [ ] Learn M language basics for repeatable transformations
- [ ] Unpivot data and handle null/missing values in the query editor
- [ ] Build a reusable Power Query that strips bad formatting and merges sources

**🔗 Resources:**

- ▶ [Excel for Data Analysts: Master ETL with Power Query](https://www.youtube.com/watch?v=RxYwM2x4VCs)

### 5. Capstone Synthesis

- [ ] Complete an end-to-end tabular analysis on a real-world dataset in Excel
- [ ] Produce a cleaned report ready for review (Atrium Display Board equivalent)
- [ ] Pin final cleaned Excel output as a stakeholder-facing summary board

**🔗 Resources:**

- 🔗 [Luke Barousse: Excel for Data Analytics (Course Page)](https://www.lukebarousse.com/courses)

## Phase 2: Relational Database Extraction and Querying (Subterranean Archive)

### 1. Database Architecture

- [ ] Install and configure PostgreSQL (or practice MySQL) locally
- [ ] Understand schemas, tables, and primary/foreign key relationships
- [ ] Write SELECT, FROM, and WHERE queries to extract and filter rows
- [ ] Map SELECT/FROM/WHERE to flashlight, aisle, and iron gate loci before coding

**🔗 Resources:**

- ▶ [SQL Tutorial for Beginners (Alex The Analyst)](https://www.youtube.com/watch?v=h0nxCDiD-zg)
- ▶ [SQL Basics Tutorial For Beginners | Select + From Statements](https://www.youtube.com/watch?v=PyYgERKq25I)

### 2. Aggregation Logic

- [ ] Aggregate with GROUP BY and filter groups with HAVING
- [ ] Apply SUM, AVG, COUNT, MIN, and MAX correctly
- [ ] Visualize GROUP BY as a binding machine that stamps totals on bundled books

**🔗 Resources:**

- ▶ [Learn SQL Beginner to Advanced in Under 4 Hours](https://www.youtube.com/watch?v=OT1RErkfLNQ)

### 3. Set Theory and Joins

- [ ] Write INNER JOIN queries across related tables
- [ ] Write LEFT JOIN queries and interpret NULL placeholders for non-matches
- [ ] Use OUTER JOIN patterns and table aliases for readable multi-table SQL
- [ ] Contrast INNER (narrow bridge) vs LEFT JOIN (ghost NULL placeholders) on practice tables

**🔗 Resources:**

- ▶ [Intermediate SQL Tutorial | Aliasing](https://www.youtube.com/watch?v=Dk7he_yEs4U)

### 4. Query Modularity

- [ ] Refactor nested subqueries into readable Common Table Expressions (CTEs)
- [ ] Format long queries to avoid monolithic bad-smell structures
- [ ] Treat hundred-line nested queries as architectural bad smells; extract CTE desks

**🔗 Resources:**

- ▶ [SQL for Data Analytics - Learn SQL in 4 Hours (Luke Barousse)](https://www.youtube.com/watch?v=7mz73uXD9DA)

### 5. Advanced Analytics

- [ ] Apply Window Functions: ROW_NUMBER, RANK, SUM() OVER(PARTITION BY)
- [ ] Use LEAD and LAG for row-relative comparisons without collapsing grain
- [ ] Practice PARTITION BY windows that preserve row grain while adding running metrics

**🔗 Resources:**

- 🔗 [Intermediate SQL for Data Analytics (Luke Barousse)](https://www.lukebarousse.com/int-sql)

### 6. Portfolio Integration

- [ ] Run Exploratory Data Analysis in SQL on a real-world dataset (e.g. COVID / demographic data)
- [ ] Document findings in a clean GitHub markdown logbook
- [ ] Import an unstructured public dataset into PostgreSQL for portfolio EDA

**🔗 Resources:**

- ▶ [Data Analyst Portfolio Project | SQL Data Exploration | Project 1/4](https://www.youtube.com/watch?v=qfyynHBFOsM)

## Phase 3: Business Intelligence and Dimensional Visualization (Grand Exhibition Hall)

### 1. Interface and Data Ingestion

- [ ] Install Power BI Desktop and import practice datasets
- [ ] Navigate the workspace: Report, Data, and Model views
- [ ] Leverage UX background: treat every dashboard as a user interface with journey mapping

**🔗 Resources:**

- ▶ [Power BI Tutorial for Beginners (Kevin Stratvert)](https://www.youtube.com/watch?v=NNSHu0rkew8)

### 2. Advanced Power Query

- [ ] Transform datasets in Power BI Power Query (types, splits, merges)
- [ ] Handle API or external source integrations when needed
- [ ] Reuse Phase 1 Power Query skills at the gallery loading dock (same ETL engine)

**🔗 Resources:**

- ▶ [How to use Power Query in Power BI (Alex The Analyst)](https://www.youtube.com/watch?v=gP-AxNi6uxo)

### 3. Dimensional Modeling

- [ ] Construct a Star Schema with a central Fact table and Dimension tables
- [ ] Manage 1-to-Many cardinality and active vs inactive relationships
- [ ] Model Fact as central sculpture and Dimensions as filter spotlights on the walls

**🔗 Resources:**

- ▶ [How to Install Power BI | Building First Visualization / Relationships](https://www.youtube.com/watch?v=g0m5sEHPU-s)

### 4. DAX and Evaluation Context

- [ ] Distinguish Row Context (magnifying glass) vs Filter Context (tarp)
- [ ] Master CALCULATE, FILTER, and iterator functions like SUMX
- [ ] Build Time Intelligence measures with DATEADD and related functions
- [ ] Use CALCULATE as a remote control that reshapes the Filter Context tarp

**🔗 Resources:**

- ▶ [Power BI Intro to Data Analysis Advanced Tutorial (Learnit)](https://www.youtube.com/watch?v=GpAXwNaUiLY)

### 5. Dashboard UX and Portfolio

- [ ] Apply UX principles: visual hierarchy, F/Z-pattern, Gestalt layout
- [ ] Choose chart types correctly (bar, line, scatter) and avoid chart junk
- [ ] Create drill-throughs and publish a portfolio dashboard to the service
- [ ] Lead stakeholder eye to the primary KPI first using hierarchy and contrast

**🔗 Resources:**

- ▶ [Learn Power BI in Under 3 Hours (Full Project)](https://www.youtube.com/watch?v=I0vQ_VLZTWg)
- ▶ [Power BI Full Course Tutorial (8+ Hours)](https://www.youtube.com/watch?v=e6QD8lP-m6E)

## Phase 4: Programmatic Manipulation and Advanced Transformation (Chemical Engineering Laboratory)

### 1. Environment Setup

- [ ] Install Anaconda and open Jupyter Notebooks
- [ ] Import pandas and numpy in a clean analysis environment
- [ ] Calibrate a sterile notebook workflow: one environment, reproducible imports

**🔗 Resources:**

- 🔗 [Python for Data Analytics (Luke Barousse)](https://www.lukebarousse.com/python)

### 2. Data Structures

- [ ] Understand DataFrames as collections of Series (vats vs test tubes)
- [ ] Index, select, filter, and sort DataFrame columns and rows
- [ ] Practice slicing and boolean filtering without mutating source frames carelessly

**🔗 Resources:**

- ▶ [Learn PANDAS in 5 minutes | Pandas Ultraquick Tutorial (Keith Galli)](https://www.youtube.com/watch?v=m1_33jhhiLE)
- ▶ [Keith Galli channel (Pandas tutorials)](https://www.youtube.com/c/KGMIT/videos)

### 3. Exploratory Data Analysis

- [ ] Profile datasets with .info(), .describe(), and .isnull().sum()
- [ ] Interpret mean, std, and percentiles from statistical summaries
- [ ] Run litmus diagnostics immediately on every new dataset before deep wrangling

**🔗 Resources:**

- ▶ [Exploratory Data Analysis in Pandas (Alex The Analyst)](https://www.youtube.com/watch?v=Liv6eeb1VfE)

### 4. Vectorization and Aggregation

- [ ] Replace iterative for-loops with vectorized column operations
- [ ] Aggregate with groupby() (centrifuge) and combine tables with merge()
- [ ] Refuse row-wise Python loops for column transforms; prefer vectorized arrays

**🔗 Resources:**

- ▶ [Complete Python Pandas Data Science Tutorial (Keith Galli)](https://www.youtube.com/watch?v=m1_33jhhiLE)

### 5. Statistical Visualization

- [ ] Compute correlation matrices on cleaned numeric features
- [ ] Plot heatmaps with Seaborn for exploratory visualization
- [ ] Keep matplotlib/seaborn for EDA only — do not introduce ML libraries

**🔗 Resources:**

- ▶ [Data Analyst Portfolio Project | Correlation in Python](https://www.youtube.com/watch?v=iPYVYBtUTyE)

## Phase 5: Synthesis, Storytelling, and Portfolio Development (Narrative Auditorium)

### 1. Audience Empathy and Context

- [ ] Map stakeholder personas: Executive, Operations Manager, Marketing Lead
- [ ] Build a narrative arc that ends on actionable insight, not exploration process
- [ ] Tailor pitch length and grain to Executive vs Operations vs Marketing seats

**🔗 Resources:**

- ▶ [Master the Art of Data Storytelling (Official Book Ginger)](https://www.youtube.com/watch?v=53gZrM42ig8)

### 2. Cognitive Load Reduction

- [ ] Eliminate gridlines, excess labels, and 3D/exploding pie charts (Knaflic)
- [ ] Use pre-attentive attributes (color, size, contrast) to force focus on the key metric
- [ ] Apply Storytelling with Data: one stark highlight (red line) on otherwise muted slides

**🔗 Resources:**

- ▶ [Storytelling with Data | Cole Nussbaumer Knaflic | Talks at Google](https://www.youtube.com/watch?v=8EMW7io4rSI)
- ▶ [Storytelling with Data (book summary)](https://www.youtube.com/watch?v=53gZrM42ig8)

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

- 🔗 [How to Create a Portfolio Website for FREE (Alex The Analyst)](https://www.classcentral.com/course/youtube-data-analyst-portfolio-projects-92262)
