<!-- synced from plans.csv (PLAN_0001) / study AI -->

# AI Study Plan

## 📃 Backlog

- [ ] dumb zone

- [ ] Meta prompt

- [ ] ETL (Extract, Transform, Load)

- [ ] Data lake

- [ ] data governance

## Phase 1: Mathematical Foundation and Data Manipulation (Months 1-3)

### Month 1: Python Programming and Applied Statistics

#### Python Fundamentals and IDE Mastery

- [ ] Master variable types, loops, and functional programming in Python
- [ ] Configure Cursor IDE for efficient code completion and debugging
- [ ] Establish virtual environment discipline (venv/conda)
- [ ] Execute terminal operations seamlessly within the editor

**🔗 Resources:**

- ▶ [Python for Beginners - freeCodeCamp](https://www.youtube.com/watch?v=rfscVS0vtbw)

#### Applied Statistics and Linear Algebra

- [ ] Master descriptive statistics (mean, variance, standard deviation)
- [ ] Map linear algebra concepts (matrices, vectors, dot products)
- [ ] Calculate probability distributions (Normal, Binomial, Poisson)
- [ ] Transition theoretical math into programmatic logic

**🔗 Resources:**

- ▶ [Statistics Fundamentals - StatQuest](https://www.youtube.com/watch?v=qBigTkBDMWs)
- ▶ [Linear Algebra Essence - 3Blue1Brown](https://www.youtube.com/watch?v=fNk_zzaMoSs)

### Month 2: Data Wrangling and SQL Integration

#### Data Manipulation with Pandas

- [ ] Apply Month 1 statistical logic using Pandas DataFrames
- [ ] Master filtering, grouping, and aggregations (groupby, merge)
- [ ] Handle missing values and execute data imputation techniques
- [ ] Optimize script execution for large in-memory datasets

**🔗 Resources:**

- ▶ [Python Pandas Tutorial - Corey Schafer](https://www.youtube.com/watch?v=vmEHCJwhgG0)

#### Relational Databases (SQL)

- [ ] Query structured data to feed Month 2 Pandas workflows
- [ ] Execute complex joins and window functions
- [ ] Optimize query performance for analytical extraction
- [ ] Integrate SQL queries directly into Python scripts

**🔗 Resources:**

- ▶ [SQL Full Course - freeCodeCamp](https://www.youtube.com/watch?v=HXV3cyQiF80)

### Month 3: Exploratory Data Analysis (EDA) and Visualization

#### Visual Storytelling

- [ ] Visualize the cleaned datasets prepared in Month 2 using Matplotlib/Seaborn
- [ ] Identify underlying distributions and detect outliers
- [ ] Eliminate chart junk to highlight core data narratives
- [ ] Build interactive dashboards for multidimensional data

**🔗 Resources:**

- ▶ [Data Visualization for Data Science - Ken Jee](https://www.youtube.com/watch?v=0UlihRp2Ads)

#### Feature Engineering

- [ ] Transform raw variables into predictive signals based on EDA findings
- [ ] Encode categorical variables (One-Hot, Target Encoding)
- [ ] Scale numerical features (Standardization, Normalization)
- [ ] Establish a reproducible feature pipeline

**🔗 Resources:**

- ▶ [Feature Engineering - StatQuest](https://www.youtube.com/watch?v=6PeL4tN9tVE)

## Phase 2: Predictive Modeling and Advanced Algorithms (Months 4-6)

### Month 4: Supervised Machine Learning

#### Regression and Classification

- [ ] Train predictive models on the engineered features from Month 3
- [ ] Master Linear and Logistic Regression architectures
- [ ] Execute tree-based models (Random Forest, Gradient Boosting)
- [ ] Isolate target variables and map decision boundaries

**🔗 Resources:**

- ▶ [Machine Learning Fundamentals - StatQuest](https://www.youtube.com/watch?v=GwIoYPEg3ok)

#### Model Evaluation Metrics

- [ ] Quantify model accuracy using precision, recall, and F1-score
- [ ] Plot ROC-AUC curves to measure classification thresholds
- [ ] Evaluate regression performance (RMSE, MAE, R-squared)
- [ ] Prevent data leakage during cross-validation

**🔗 Resources:**

- ▶ [ROC and AUC - StatQuest](https://www.youtube.com/watch?v=4jRBrMpHDlc)

### Month 5: Unsupervised Machine Learning and Web Scraping

#### Clustering and Dimensionality Reduction

- [ ] Apply K-Means and DBSCAN to segment unlabelled data from Month 4 pipelines
- [ ] Execute PCA (Principal Component Analysis) to reduce feature space
- [ ] Identify latent patterns in continuous data streams
- [ ] Validate cluster cohesion and separation (Silhouette Score)

**🔗 Resources:**

- ▶ [Principal Component Analysis - StatQuest](https://www.youtube.com/watch?v=FgakZw6K1QQ)

#### Web Scraping for Custom Datasets

- [ ] Scrape e-commerce metrics (SSL presence, customer opinions) using BeautifulSoup/Selenium
- [ ] Extract unstructured web data to fuel Month 5 clustering algorithms
- [ ] Automate pagination and handle dynamic DOM rendering
- [ ] Structure scraped outputs into relational schemas

**🔗 Resources:**

- ▶ [Python Web Scraping - Corey Schafer](https://www.youtube.com/watch?v=XGkEhqBRx-Y)

### Month 6: Experimentation and Causal Inference

#### A/B Testing and Hypothesis Validation

- [ ] Design randomized experiments to test variables like perceived trust
- [ ] Calculate sample size and statistical power requirements
- [ ] Evaluate experiment results using t-tests and p-values
- [ ] Distinguish correlation from causal impact

**🔗 Resources:**

- ▶ [A/B Testing - StatQuest](https://www.youtube.com/watch?v=UsYh8EqgAJE)

#### Advanced Feature Selection

- [ ] Prune suboptimal features from Month 4 and Month 5 models
- [ ] Implement Recursive Feature Elimination (RFE)
- [ ] Calculate feature importance metrics and SHAP values
- [ ] Optimize the trade-off between model complexity and interpretability

**🔗 Resources:**

- ▶ [Ridge, Lasso, ElasticNet - StatQuest](https://www.youtube.com/watch?v=Q81rr3jDCNE)

## Phase 3: Deep Learning and Unstructured Data (Months 7-9)

### Month 7: Neural Networks Foundation

#### Deep Learning Architectures

- [ ] Transition from Month 4 traditional ML to multi-layer perceptrons
- [ ] Master backpropagation and gradient descent mechanics
- [ ] Implement activation functions (ReLU, Sigmoid, Softmax)
- [ ] Build and train networks using PyTorch or TensorFlow

**🔗 Resources:**

- ▶ [Neural Networks - 3Blue1Brown](https://www.youtube.com/watch?v=aircAruvnKk)

#### Regularization and Optimization

- [ ] Mitigate neural network overfitting using Dropout and L2 regularization
- [ ] Tune hyperparameters (learning rate, batch size)
- [ ] Implement early stopping mechanisms
- [ ] Monitor training loss vs. validation loss epochs

**🔗 Resources:**

- ▶ [Gradient Descent - StatQuest](https://www.youtube.com/watch?v=sDv2f5w14Ns)

### Month 8: Natural Language Processing (NLP)

#### Text Processing and Embeddings

- [ ] Process the unstructured customer opinion data scraped in Month 5
- [ ] Execute tokenization, lemmatization, and stop-word removal
- [ ] Map text to dense vectors using Word2Vec or GloVe
- [ ] Transition text representations into deep learning inputs

**🔗 Resources:**

- ▶ [Word Embeddings - StatQuest](https://www.youtube.com/watch?v=fBekMeEAupI)

#### Sequential Modeling

- [ ] Train RNNs and LSTMs on temporal/sequential data
- [ ] Analyze sentiment polarity within text corpus
- [ ] Implement attention mechanisms for long-context understanding
- [ ] Evaluate language models against baseline heuristics

**🔗 Resources:**

- ▶ [Recurrent Neural Networks - StatQuest](https://www.youtube.com/watch?v=AsNTP8Kwu80)

### Month 9: Transformer Models and Generative AI Integration

#### Advanced NLP architectures

- [ ] Upgrade from Month 8 sequential models to Transformer architectures
- [ ] Fine-tune pre-trained models (BERT, RoBERTa) on domain-specific data
- [ ] Execute zero-shot and few-shot inference techniques
- [ ] Integrate LLM APIs into data pipelines for unstructured parsing

**🔗 Resources:**

- ▶ [Transformers Explained - StatQuest](https://www.youtube.com/watch?v=zxQyTKU3KM)

## Phase 4: MLOps and Production Systems (Months 10-12)

### Month 10: Model Deployment and APIs

#### API Development

- [ ] Wrap the Month 9 Transformer models into RESTful APIs using FastAPI
- [ ] Define request schemas and response validations
- [ ] Handle concurrent model inference requests
- [ ] Document endpoints automatically via Swagger/OpenAPI

**🔗 Resources:**

- ▶ [FastAPI Course - freeCodeCamp](https://www.youtube.com/watch?v=0sOvCWFmrtA)

#### Containerization

- [ ] Containerize the FastAPI application using Docker
- [ ] Establish isolated, reproducible production environments
- [ ] Manage environment variables and secrets securely
- [ ] Optimize image sizes for faster deployment cycles

**🔗 Resources:**

- ▶ [Docker Full Course - freeCodeCamp](https://www.youtube.com/watch?v=fG4ZBpHkMjM)

### Month 11: MLOps and Continuous Integration

#### Model Tracking and CI/CD

- [ ] Implement MLflow to track experiments and model registries from Month 10
- [ ] Automate testing and deployment pipelines via GitHub Actions
- [ ] Establish model drift detection mechanisms
- [ ] Execute automated retraining triggers based on data degradation

**🔗 Resources:**

- ▶ [MLOps Explained - TechWorld with Nana](https://www.youtube.com/watch?v=06-AZXmD0pQ)

### Month 12: Capstone Project and Portfolio Consolidation

#### End-to-End E-commerce Trust System

- [ ] Integrate Month 5 web scraping, Month 6 experimentation, and Month 9 NLP into a unified system
- [ ] Deploy the containerized solution to a cloud provider (AWS/GCP)
- [ ] Establish a visualization dashboard mapping real-time predictions
- [ ] Compile the repository structure, documentation, and architecture diagrams
- [ ] Complete clinical playback analysis of the model's predictive performance

**🔗 Resources:**

- ▶ [Data Science Portfolio Tips - Ken Jee](https://www.youtube.com/watch?v=4u4w7JWqz-U)

## Practice Schedule (Standard Analytical Session Architecture)

### Session Structure

- [ ] Warm-up (5 min): Code review, syntax drill, or terminal navigation
- [ ] Theoretical Block (10 min): Math foundations, algorithm mechanics, or paper reading
- [ ] Application Block (25 min): Writing scripts, training models, or debugging pipelines
- [ ] Review and Reflection (5 min): Metric logging, code optimization, weekly reflection

**🔗 Resources:**

- ▶ [How to Study Data Science - Ken Jee](https://www.youtube.com/watch?v=0L3ZzBTJFJc)

## Progress Tracking

### Practice Consistency

- [ ] 5-6 sessions per week completed

**🔗 Resources:**

- ▶ [Building Consistent Study Habits - Ali Abdaal](https://www.youtube.com/watch?v=K8H0pQn72N0)

### Execution Progression

- [ ] Data Wrangling Pipeline - Target: < 500ms execution
- [ ] Supervised Training - Target: > 85% F1-Score on baseline
- [ ] API Response Time - Target: < 200ms latency

**🔗 Resources:**

- ▶ [How to Track Your Learning - Data Science Jay](https://www.youtube.com/watch?v=0L3ZzBTJFJc)

### Accuracy Trend

- [ ] Monitor Data Leakage Incidents
- [ ] Monitor Prediction Residuals
- [ ] Monitor Pipeline Crash Rates

**🔗 Resources:**

- ▶ [Confusion Matrix - StatQuest](https://www.youtube.com/watch?v=fSytzGwwBVw)

### Portfolio Readiness

- [ ] EDA Notebook on E-commerce metrics
- [ ] Classification Model API
- [ ] End-to-End MLOps Pipeline

**🔗 Resources:**

- ▶ [Data Science Portfolio - Ken Jee](https://www.youtube.com/watch?v=4u4w7JWqz-U)

### Skill Radar (Semi-Monthly Self-Assessment)

- [ ] Python/SQL Engineering
- [ ] Statistical Rigor
- [ ] Machine Learning Accuracy
- [ ] Deep Learning / NLP
- [ ] MLOps & Deployment
- [ ] Data Storytelling

**🔗 Resources:**

- ▶ [Self-Assessment for Data Scientists - Ken Jee](https://www.youtube.com/watch?v=0L3ZzBTJFJc)

## Goals (End of Year 1)

- [ ] Fluent script execution across local and cloud environments
- [ ] 3-5 production-ready projects (mixed domains)
- [ ] Solid intuition for mapping algorithms to business problems
- [ ] Comfortable interpreting complex feature interactions
- [ ] Deployed portfolio documentation (monthly snapshots)
- [ ] Consistent 5-6 session/week analytical discipline
- [ ] Zero unmanaged technical debt in primary repositories

**🔗 Resources:**

- ▶ [Data Science Roadmap - Ken Jee](https://www.youtube.com/watch?v=ua-CiDNNj30)
