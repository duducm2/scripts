<!-- synced from plans.csv (PLAN_0008) / study data-scientist -->

# Senior Data Scientist Transition Architecture

## 📃 Backlog

- [ ] Archive source curriculum PDF (Senior Data Science Study Plan) under research/ for reference

- [ ] Defer Memory Palace blueprint until ready — leave palaces blank for now

- [ ] This track includes ML, deep learning, and MLOps (unlike the Data Analyst no-ML constraint)

- [ ] Leverage existing Python, SQL, Power BI, Azure, and UX experimentation strengths while closing senior gaps

## Phase 0: Solidifying Statistical and Machine Learning Fundamentals

### Topic 0.1: Algorithmic Intuition

- [ ] Master the mathematics and logic behind k-fold cross-validation
- [ ] Prevent data leakage during feature engineering and nested validation
- [ ] Deconstruct the bias-variance tradeoff and its impact on generalization
- [ ] Use StatQuest to rebuild intuition for baseline tabular ML algorithms without treating APIs as black boxes
- [ ] Tune hyperparameters with nested CV so validation folds never leak into model selection

**🔗 Resources:**

- ▶ [Cross Validation — Machine Learning Fundamentals (StatQuest)](https://www.youtube.com/watch?v=fSytzGwwBVw)
- ▶ [A Gentle Introduction to Machine Learning (StatQuest)](https://www.youtube.com/watch?v=Gv9_4yMHFhI)
- 📄 [StatQuest video index](https://statquest.org/video_index.html)
- 📄 [scikit-learn: Cross-validation](https://scikit-learn.org/stable/modules/cross_validation.html)
- 🔗 [An Introduction to Statistical Learning (ISL) — free PDF](https://www.statlearning.com/)

### Topic 0.2: Advanced Tree Models

- [ ] Understand XGBoost objective functions and sequential gradient boosting mechanics
- [ ] Trace how XGBoost calculates similarity scores during tree growth
- [ ] Explain how XGBoost handles missing data natively
- [ ] Connect mechanical XGBoost understanding to feature engineering and corporate tabular debugging
- [ ] Debug a misbehaving XGBoost model by inspecting leaf weights and feature importance patterns

**🔗 Resources:**

- 📄 [XGBoost: A Scalable Tree Boosting System (original paper)](https://arxiv.org/abs/1603.02754)
- 🔗 [XGBoost documentation (official)](https://xgboost.readthedocs.io/en/stable/)
- ▶ [StatQuest XGBoost playlist search (channel)](https://www.youtube.com/c/joshstarmer/search?query=xgboost)

### Topic 0.3: Dimensionality Reduction

- [ ] Calculate PCA eigenvectors and eigenvalues manually on a small matrix
- [ ] Interpret principal components as linear transformations of high-dimensional feature space
- [ ] Compare PCA vs t-SNE for visualization of complex feature spaces
- [ ] Link dimensionality reduction to InfoVis / PhraseNets-style interpretation of high-dimensional data
- [ ] Produce a PCA scree plot and decide how many components to retain for downstream modeling

**🔗 Resources:**

- ▶ [PCA Main Ideas (StatQuest)](https://www.youtube.com/watch?v=FgakZw6K1QQ)
- 📄 [scikit-learn: PCA](https://scikit-learn.org/stable/modules/generated/sklearn.decomposition.PCA.html)

## Phase 1: Rigorous Algorithmic Foundations

### Topic 1.1: Empirical Risk

- [ ] Define supervised learning using Stanford CS229 formal mathematical notation
- [ ] Derive cost functions for linear and logistic regression with matrix calculus
- [ ] Explain empirical risk minimization and generalized linear models (GLMs)
- [ ] Diagnose why models converge or fail to find a global minimum under geometric constraints
- [ ] Read one CS229 lecture note section and restate the supervised learning setup in your own notation

**🔗 Resources:**

- ▶ [Stanford CS229 Machine Learning — Spring 2022 playlist](https://www.youtube.com/playlist?list=PLoROMvodv4rNyWOpJg_Yh4NSqI4Z4vOYy)
- ▶ [CS229 Lecture 1: Overview / supervised learning / ERM](https://www.youtube.com/watch?v=I-tmjGFaaBg)
- 🔗 [CS229 Lecture Notes PDF (Stanford)](https://cs229.stanford.edu/main_notes.pdf)

### Topic 1.2: Optimization

- [ ] Analyze gradient descent variations and their geometric convergence properties
- [ ] Relate learning rate and curvature to underfitting / overfitting failure modes
- [ ] Construct intuition for bespoke regularization terms on proprietary data
- [ ] Compare batch, mini-batch, and stochastic gradient descent on a toy convex loss

**🔗 Resources:**

- 🔗 [CS229 Lecture Notes (Stanford)](https://cs229.stanford.edu/main_notes.pdf)
- 🔗 [CS229 Autumn 2018 notes — Linear Algebra / Probability review](https://cs229.stanford.edu/section/cs229-linalg.pdf)

### Topic 1.3: Advanced Classifiers

- [ ] Implement an SVM objective and identify support vectors
- [ ] Apply the Kernel Trick to map non-linear data into higher-dimensional spaces
- [ ] State the bias-variance tradeoff mathematically, not only conceptually
- [ ] Sketch why a linear SVM fails on XOR and how a kernel resolves it

**🔗 Resources:**

- 🔗 [CS229 Course materials / notes index](https://cs229.stanford.edu/)
- ▶ [Stanford CS229 Machine Learning — Spring 2022 playlist](https://www.youtube.com/playlist?list=PLoROMvodv4rNyWOpJg_Yh4NSqI4Z4vOYy)

## Phase 2: Advanced Experimentation and Causal Inference

### Topic 2.1: Variance Reduction (CUPED)

- [ ] Define the mathematical formula for CUPED (Controlled-experiment Using Pre-Experiment Data)
- [ ] Calculate a CUPED-adjusted metric using pre-experiment covariates
- [ ] Estimate traffic / sample-size reductions and time-to-decision gains from variance reduction
- [ ] Contrast CUPED with traditional fixed-horizon A/B testing for micro-conversion detection
- [ ] Write a short memo: when CUPED helps vs when pre-experiment covariates are unavailable

**🔗 Resources:**

- ▶ [What Exactly is CUPED? (Eppo)](https://www.youtube.com/watch?v=KlPYmPwVpyU)
- 🔗 [How to Double A/B Testing Speed with CUPED (Towards Data Science)](https://towardsdatascience.com/how-to-double-a-b-testing-speed-with-cuped-f80460825a90/)
- 🔗 [Deep Dive Into Variance Reduction (Microsoft Research)](https://www.microsoft.com/en-us/research/articles/deep-dive-into-variance-reduction/)
- 📝 [Online Experiments Tricks — Variance Reduction (Medium)](https://medium.com/data-science/online-experiments-tricks-variance-reduction-291b6032dcd7)
- 🔗 [Improving the Sensitivity of Online Controlled Experiments (Deng et al., CUPED paper)](https://www.microsoft.com/en-us/research/publication/improving-the-sensitivity-of-online-controlled-experiments-by-utilizing-pre-experiment-data/)

### Topic 2.2: Multi-Armed Bandits

- [ ] Contrast static A/B testing with Multi-Armed Bandit exploration–exploitation dynamics
- [ ] Implement Thompson Sampling to dynamically route traffic to high-performing variants
- [ ] Quantify regret minimization for short-lived campaigns and dynamic pricing scenarios
- [ ] List enterprise use cases where bandits beat fixed A/B (promos, headlines, pricing)

**🔗 Resources:**

- ▶ [How Multi-Armed Bandits Optimize Experiments in Real-Time (Amplitude)](https://www.youtube.com/watch?v=moJ_poxgATo)
- 🔗 [Going from A/B Testing to AI: Optimization as Reinforcement Learning](https://www.conductrics.com/going-from-ab-testing-to-ai-optimization-as-reinforcement-learning/)
- 🔗 [Lil'Log: Multi-Armed Bandits](https://lilianweng.github.io/posts/2018-01-23-multi-armed-bandit/)

### Topic 2.3: Causal Inference

- [ ] Construct a Directed Acyclic Graph (DAG) mapping causal assumptions and confounders
- [ ] Calculate Average Treatment Effect (ATE) using the Potential Outcomes framework
- [ ] State the fundamental problem of causal inference and when RCTs are impossible
- [ ] Estimate treatment effects from observational / historical logs when A/B tests are unethical or infeasible
- [ ] Identify confounders in a historical policy-change scenario using a DAG

**🔗 Resources:**

- ▶ [Brady Neal: Causal Inference (Lecture 1, Part 1/4)](https://www.youtube.com/watch?v=U7qhODzS1Bs)
- 🔗 [Brady Neal — Introduction to Causal Inference (course site)](https://www.bradyneal.com/causal-inference-course)
- 🔗 [The Book of Why — Pearl & Mackenzie (conceptual causal literacy)](https://www.basicbooks.com/titles/judea-pearl/the-book-of-why/9780465097609/)

## Phase 3: Deep Learning and Modern Architectures

### Topic 3.1: Micrograd

- [ ] Build a scalar-valued autograd engine from scratch (Karpathy micrograd)
- [ ] Implement manual backpropagation through a DAG using the chain rule
- [ ] Explain how loss derivatives cascade backward through individual weights
- [ ] Unit-test your autograd engine against PyTorch gradients on a tiny expression graph

**🔗 Resources:**

- 🔗 [Neural Networks: Zero To Hero (Karpathy hub)](https://karpathy.ai/zero-to-hero.html)
- ▶ [The spelled-out intro to neural networks and backpropagation (micrograd)](https://www.youtube.com/watch?v=VMj-3S1tku0)
- 🔗 [karpathy/micrograd (GitHub)](https://github.com/karpathy/micrograd)

### Topic 3.2: Makemore and MLPs

- [ ] Construct a bigram character-level language model with torch.Tensor broadcasting
- [ ] Evaluate with negative log-likelihood and proper train/val/test splits
- [ ] Build a Multi-Layer Perceptron with embedding tables and hyperparameter tuning
- [ ] Visually diagnose forward-pass activation statistics
- [ ] Log train vs validation loss curves and diagnose overfitting in the MLP

**🔗 Resources:**

- ▶ [Building makemore — character-level language modeling](https://www.youtube.com/watch?v=PaCmpygFfXo)
- 🔗 [karpathy/nn-zero-to-hero (GitHub notebooks)](https://github.com/karpathy/nn-zero-to-hero)
- 📄 [PyTorch: Tensors tutorial](https://pytorch.org/tutorials/beginner/basics/tensorqs_tutorial.html)

### Topic 3.3: Network Stabilization

- [ ] Implement Kaiming initialization to prevent vanishing / exploding gradients
- [ ] Build Batch Normalization from scratch including Bessel's correction
- [ ] Become a backprop ninja: manually backprop through cross-entropy, BatchNorm, and linear layers without autograd
- [ ] Visualize activation histograms before vs after BatchNorm

**🔗 Resources:**

- ▶ [Building makemore Part 4: Becoming a Backprop Ninja](https://www.youtube.com/watch?v=q8SA3rM6ckI)
- 📄 [Delving Deep into Rectifiers (Kaiming He et al. — init paper)](https://arxiv.org/abs/1502.01852)
- 📄 [Batch Normalization paper (Ioffe & Szegedy)](https://arxiv.org/abs/1502.03167)

### Topic 3.4: Convolutional Sequences

- [ ] Recreate DeepMind WaveNet-style architecture for hierarchical sequence modeling
- [ ] Implement causal dilated convolutions to expand receptive field efficiently
- [ ] Treat modern frameworks as optimized wrappers for math you can recreate, not opaque APIs

**🔗 Resources:**

- 🔗 [Karpathy Zero to Hero series index](https://karpathy.ai/zero-to-hero.html)
- 📝 [WaveNet: A Generative Model for Raw Audio (DeepMind blog / paper)](https://www.deepmind.com/blog/wavenet-a-generative-model-for-raw-audio)

## Phase 4: Machine Learning Operations (MLOps)

### Topic 4.1: Experiment Tracking

- [ ] Configure a local MLflow server to log hyperparameters, metrics, and artifacts
- [ ] Register, stage, and version models in the MLflow Model Registry
- [ ] Replace ad-hoc notebook tracking with fully reproducible experiment runs
- [ ] Compare two MLflow runs side-by-side and promote the winner to Staging

**🔗 Resources:**

- 📝 [MLOps Zoomcamp (DataTalks.Club overview)](https://datatalks.club/blog/mlops-zoomcamp.html)
- 🔗 [DataTalksClub/mlops-zoomcamp (GitHub)](https://github.com/DataTalksClub/mlops-zoomcamp)
- 📄 [MLflow Tracking documentation](https://mlflow.org/docs/latest/tracking.html)
- 📄 [MLflow Model Registry](https://mlflow.org/docs/latest/model-registry.html)

### Topic 4.2: Pipeline Orchestration

- [ ] Convert procedural Jupyter experiments into modular object-oriented Python scripts
- [ ] Build an automated scheduled training pipeline with Prefect, Mage, or Airflow
- [ ] Orchestrate data ingestion, feature engineering, and training as a DAG of tasks
- [ ] Add retries and failure alerts to at least one orchestrated training task

**🔗 Resources:**

- ▶ [MLOps Zoomcamp YouTube playlist](https://www.youtube.com/playlist?list=PL3MmuxUbc_hIUISrluw_A7wDSmfOhErJK)
- 🔗 [Prefect docs: Getting started](https://docs.prefect.io/latest/get-started/index.html)

### Topic 4.3: Cloud Deployment

- [ ] Containerize a prediction service with Docker for environment parity
- [ ] Deploy a scheduled batch scoring job on cloud infrastructure (AWS or Azure)
- [ ] Deploy a real-time FastAPI/Flask inference endpoint for low-latency serving
- [ ] Sketch a streaming path using cloud event infrastructure (e.g. Kinesis/Lambda patterns)
- [ ] Document latency budget for the REST inference path from client to model

**🔗 Resources:**

- 🔗 [Docker docs: Get started](https://docs.docker.com/get-started/)
- 🔗 [FastAPI documentation](https://fastapi.tiangolo.com/)

### Topic 4.4: Model Monitoring

- [ ] Configure Evidently AI (or similar) to compute covariate shift and concept drift metrics
- [ ] Wire Prometheus / Grafana-style alerts for production performance degradation
- [ ] Connect monitoring alerts back to tensor-level debugging intuition from Phase 3
- [ ] Define alert thresholds for data drift that trigger investigation not blind retrain

**🔗 Resources:**

- 🔗 [Evidently AI documentation](https://docs.evidentlyai.com/)
- 📄 [Prometheus Getting started](https://prometheus.io/docs/prometheus/latest/getting_started/)

## Phase 5: Machine Learning System Design

### Topic 5.1: Requirements Engineering

- [ ] Translate ambiguous business objectives into quantifiable offline ML evaluation metrics
- [ ] Align model success criteria with macro ROI / product metrics (CSI, NPS, conversion)
- [ ] Write an offline metric card that executives can map to a business KPI

**🔗 Resources:**

- 🔗 [Designing Machine Learning Systems — Chip Huyen (book site / buy)](https://www.oreilly.com/library/view/designing-machine-learning/9781098107956/)
- 🔗 [Rules of Machine Learning (Google — Martin Zinkevich)](https://developers.google.com/machine-learning/guides/rules-of-ml)

### Topic 5.2: Data Supply Chains

- [ ] Design architectures balancing batch processing vs streaming feature engineering
- [ ] Map how data supply chains dictate production model performance (Chip Huyen)
- [ ] Connect CUPED's need for pre-experiment covariates to low-latency join / feature-store architecture
- [ ] Draw a one-page architecture: batch training path vs online inference path

**🔗 Resources:**

- 🔗 [chiphuyen/dmls-book (companion materials)](https://github.com/chiphuyen/dmls-book)
- 🔗 [MLOps guide — Chip Huyen](https://huyenchip.com/mlops/)

### Topic 5.3: Feature Stores

- [ ] Diagram API and infrastructure boundaries of a centralized organizational feature store
- [ ] Explain how feature stores reduce redundant computation across isolated ML use cases
- [ ] List three features that must be point-in-time correct to avoid training-serving skew

**🔗 Resources:**

- 🔗 [MLOps guide — Chip Huyen](https://huyenchip.com/mlops/)
- 📄 [Feast documentation (open-source feature store)](https://docs.feast.dev/)

### Topic 5.4: Adaptive Environments

- [ ] Establish continuous retraining loops based on delayed feedback attribution
- [ ] Document anti-patterns in deployed ML systems and continuous evaluation strategies
- [ ] Practice communicating system-design trade-offs to executive and engineering stakeholders
- [ ] Propose a delayed-feedback labeling strategy for a conversion or trust metric

**🔗 Resources:**

- 📖 [Designing Machine Learning Systems — Chip Huyen (book)](https://www.oreilly.com/library/view/designing-machine-learning/9781098107956/)
- 🔗 [Hidden Technical Debt in Machine Learning Systems (Sculley et al., Google)](https://papers.nips.cc/paper_files/paper/2015/hash/86df7dcfd896fcaf2674f757a2463eba-Abstract.html)
