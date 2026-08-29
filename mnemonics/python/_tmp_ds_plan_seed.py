"""One-shot: append Senior Data Scientist plan items + resources for PLAN_0008."""

from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path

data = Path(__file__).resolve().parent.parent / "data"
plan_id = "PLAN_0008"

p0 = "Phase 0: Solidifying Statistical and Machine Learning Fundamentals"
p1 = "Phase 1: Rigorous Algorithmic Foundations"
p2 = "Phase 2: Advanced Experimentation and Causal Inference"
p3 = "Phase 3: Deep Learning and Modern Architectures"
p4 = "Phase 4: Machine Learning Operations (MLOps)"
p5 = "Phase 5: Machine Learning System Design"

sections: list[tuple[str, list[str]]] = []


def add(section: str, *texts: str) -> None:
    sections.append((section, list(texts)))


add(
    "Backlog",
    "Archive source curriculum PDF (Senior Data Science Study Plan) under research/ for reference",
    "Defer Memory Palace blueprint until ready — leave palaces blank for now",
    "This track includes ML, deep learning, and MLOps (unlike the Data Analyst no-ML constraint)",
    "Leverage existing Python, SQL, Power BI, Azure, and UX experimentation strengths while closing senior gaps",
)

# Phase 0
add(
    f"{p0} > Topic 0.1: Algorithmic Intuition",
    "Master the mathematics and logic behind k-fold cross-validation",
    "Prevent data leakage during feature engineering and nested validation",
    "Deconstruct the bias-variance tradeoff and its impact on generalization",
    "Use StatQuest to rebuild intuition for baseline tabular ML algorithms without treating APIs as black boxes",
)
add(
    f"{p0} > Topic 0.2: Advanced Tree Models",
    "Understand XGBoost objective functions and sequential gradient boosting mechanics",
    "Trace how XGBoost calculates similarity scores during tree growth",
    "Explain how XGBoost handles missing data natively",
    "Connect mechanical XGBoost understanding to feature engineering and corporate tabular debugging",
)
add(
    f"{p0} > Topic 0.3: Dimensionality Reduction",
    "Calculate PCA eigenvectors and eigenvalues manually on a small matrix",
    "Interpret principal components as linear transformations of high-dimensional feature space",
    "Compare PCA vs t-SNE for visualization of complex feature spaces",
    "Link dimensionality reduction to InfoVis / PhraseNets-style interpretation of high-dimensional data",
)

# Phase 1
add(
    f"{p1} > Topic 1.1: Empirical Risk",
    "Define supervised learning using Stanford CS229 formal mathematical notation",
    "Derive cost functions for linear and logistic regression with matrix calculus",
    "Explain empirical risk minimization and generalized linear models (GLMs)",
    "Diagnose why models converge or fail to find a global minimum under geometric constraints",
)
add(
    f"{p1} > Topic 1.2: Optimization",
    "Analyze gradient descent variations and their geometric convergence properties",
    "Relate learning rate and curvature to underfitting / overfitting failure modes",
    "Construct intuition for bespoke regularization terms on proprietary data",
)
add(
    f"{p1} > Topic 1.3: Advanced Classifiers",
    "Implement an SVM objective and identify support vectors",
    "Apply the Kernel Trick to map non-linear data into higher-dimensional spaces",
    "State the bias-variance tradeoff mathematically, not only conceptually",
)

# Phase 2
add(
    f"{p2} > Topic 2.1: Variance Reduction (CUPED)",
    "Define the mathematical formula for CUPED (Controlled-experiment Using Pre-Experiment Data)",
    "Calculate a CUPED-adjusted metric using pre-experiment covariates",
    "Estimate traffic / sample-size reductions and time-to-decision gains from variance reduction",
    "Contrast CUPED with traditional fixed-horizon A/B testing for micro-conversion detection",
)
add(
    f"{p2} > Topic 2.2: Multi-Armed Bandits",
    "Contrast static A/B testing with Multi-Armed Bandit exploration–exploitation dynamics",
    "Implement Thompson Sampling to dynamically route traffic to high-performing variants",
    "Quantify regret minimization for short-lived campaigns and dynamic pricing scenarios",
)
add(
    f"{p2} > Topic 2.3: Causal Inference",
    "Construct a Directed Acyclic Graph (DAG) mapping causal assumptions and confounders",
    "Calculate Average Treatment Effect (ATE) using the Potential Outcomes framework",
    "State the fundamental problem of causal inference and when RCTs are impossible",
    "Estimate treatment effects from observational / historical logs when A/B tests are unethical or infeasible",
)

# Phase 3
add(
    f"{p3} > Topic 3.1: Micrograd",
    "Build a scalar-valued autograd engine from scratch (Karpathy micrograd)",
    "Implement manual backpropagation through a DAG using the chain rule",
    "Explain how loss derivatives cascade backward through individual weights",
)
add(
    f"{p3} > Topic 3.2: Makemore and MLPs",
    "Construct a bigram character-level language model with torch.Tensor broadcasting",
    "Evaluate with negative log-likelihood and proper train/val/test splits",
    "Build a Multi-Layer Perceptron with embedding tables and hyperparameter tuning",
    "Visually diagnose forward-pass activation statistics",
)
add(
    f"{p3} > Topic 3.3: Network Stabilization",
    "Implement Kaiming initialization to prevent vanishing / exploding gradients",
    "Build Batch Normalization from scratch including Bessel's correction",
    "Become a backprop ninja: manually backprop through cross-entropy, BatchNorm, and linear layers without autograd",
)
add(
    f"{p3} > Topic 3.4: Convolutional Sequences",
    "Recreate DeepMind WaveNet-style architecture for hierarchical sequence modeling",
    "Implement causal dilated convolutions to expand receptive field efficiently",
    "Treat modern frameworks as optimized wrappers for math you can recreate, not opaque APIs",
)

# Phase 4
add(
    f"{p4} > Topic 4.1: Experiment Tracking",
    "Configure a local MLflow server to log hyperparameters, metrics, and artifacts",
    "Register, stage, and version models in the MLflow Model Registry",
    "Replace ad-hoc notebook tracking with fully reproducible experiment runs",
)
add(
    f"{p4} > Topic 4.2: Pipeline Orchestration",
    "Convert procedural Jupyter experiments into modular object-oriented Python scripts",
    "Build an automated scheduled training pipeline with Prefect, Mage, or Airflow",
    "Orchestrate data ingestion, feature engineering, and training as a DAG of tasks",
)
add(
    f"{p4} > Topic 4.3: Cloud Deployment",
    "Containerize a prediction service with Docker for environment parity",
    "Deploy a scheduled batch scoring job on cloud infrastructure (AWS or Azure)",
    "Deploy a real-time FastAPI/Flask inference endpoint for low-latency serving",
    "Sketch a streaming path using cloud event infrastructure (e.g. Kinesis/Lambda patterns)",
)
add(
    f"{p4} > Topic 4.4: Model Monitoring",
    "Configure Evidently AI (or similar) to compute covariate shift and concept drift metrics",
    "Wire Prometheus / Grafana-style alerts for production performance degradation",
    "Connect monitoring alerts back to tensor-level debugging intuition from Phase 3",
)

# Phase 5
add(
    f"{p5} > Topic 5.1: Requirements Engineering",
    "Translate ambiguous business objectives into quantifiable offline ML evaluation metrics",
    "Align model success criteria with macro ROI / product metrics (CSI, NPS, conversion)",
)
add(
    f"{p5} > Topic 5.2: Data Supply Chains",
    "Design architectures balancing batch processing vs streaming feature engineering",
    "Map how data supply chains dictate production model performance (Chip Huyen)",
    "Connect CUPED's need for pre-experiment covariates to low-latency join / feature-store architecture",
)
add(
    f"{p5} > Topic 5.3: Feature Stores",
    "Diagram API and infrastructure boundaries of a centralized organizational feature store",
    "Explain how feature stores reduce redundant computation across isolated ML use cases",
)
add(
    f"{p5} > Topic 5.4: Adaptive Environments",
    "Establish continuous retraining loops based on delayed feedback attribution",
    "Document anti-patterns in deployed ML systems and continuous evaluation strategies",
    "Practice communicating system-design trade-offs to executive and engineering stakeholders",
)

resources: list[tuple[str, str]] = [
    # Phase 0
    (
        f"{p0} > Topic 0.1: Algorithmic Intuition",
        "- [Cross Validation — Machine Learning Fundamentals (StatQuest)](https://www.youtube.com/watch?v=fSytzGwwBVw)",
    ),
    (
        f"{p0} > Topic 0.1: Algorithmic Intuition",
        "- [A Gentle Introduction to Machine Learning (StatQuest)](https://www.youtube.com/watch?v=Gv9_4yMHFhI)",
    ),
    (
        f"{p0} > Topic 0.1: Algorithmic Intuition",
        "- [StatQuest video index](https://statquest.org/video_index.html)",
    ),
    (
        f"{p0} > Topic 0.2: Advanced Tree Models",
        "- [StatQuest with Josh Starmer (XGBoost / ML search)](https://www.youtube.com/c/joshstarmer/search?query=xgboost)",
    ),
    (
        f"{p0} > Topic 0.3: Dimensionality Reduction",
        "- [StatQuest PCA search](https://www.youtube.com/c/joshstarmer/search?query=pca)",
    ),
    # Phase 1
    (
        f"{p1} > Topic 1.1: Empirical Risk",
        "- [Stanford CS229 Machine Learning — Spring 2022 playlist](https://www.youtube.com/playlist?list=PLoROMvodv4rNyWOpJg_Yh4NSqI4Z4vOYy)",
    ),
    (
        f"{p1} > Topic 1.1: Empirical Risk",
        "- [CS229 Lecture 1: Overview / supervised learning / ERM](https://www.youtube.com/watch?v=I-tmjGFaaBg)",
    ),
    (
        f"{p1} > Topic 1.2: Optimization",
        "- [Stanford CS229 Machine Learning | Spring 2026 | Lecture 1](https://www.youtube.com/watch?v=DATnpGoGhM8)",
    ),
    (
        f"{p1} > Topic 1.3: Advanced Classifiers",
        "- [Stanford CS229 — Neural Networks / kernels path](https://www.youtube.com/watch?v=ZMxfDWPXmjc)",
    ),
    # Phase 2
    (
        f"{p2} > Topic 2.1: Variance Reduction (CUPED)",
        "- [What Exactly is CUPED? (Eppo)](https://www.youtube.com/watch?v=KlPYmPwVpyU)",
    ),
    (
        f"{p2} > Topic 2.1: Variance Reduction (CUPED)",
        "- [How to Double A/B Testing Speed with CUPED (Towards Data Science)](https://towardsdatascience.com/how-to-double-a-b-testing-speed-with-cuped-f80460825a90/)",
    ),
    (
        f"{p2} > Topic 2.1: Variance Reduction (CUPED)",
        "- [Deep Dive Into Variance Reduction (Microsoft Research)](https://www.microsoft.com/en-us/research/articles/deep-dive-into-variance-reduction/)",
    ),
    (
        f"{p2} > Topic 2.2: Multi-Armed Bandits",
        "- [How Multi-Armed Bandits Optimize Experiments in Real-Time (Amplitude)](https://www.youtube.com/watch?v=moJ_poxgATo)",
    ),
    (
        f"{p2} > Topic 2.2: Multi-Armed Bandits",
        "- [Going from A/B Testing to AI: Optimization as Reinforcement Learning](https://www.conductrics.com/going-from-ab-testing-to-ai-optimization-as-reinforcement-learning/)",
    ),
    (
        f"{p2} > Topic 2.3: Causal Inference",
        "- [Brady Neal: Causal Inference (Lecture 1, Part 1/4)](https://www.youtube.com/watch?v=U7qhODzS1Bs)",
    ),
    # Phase 3
    (
        f"{p3} > Topic 3.1: Micrograd",
        "- [Neural Networks: Zero To Hero (Karpathy hub)](https://karpathy.ai/zero-to-hero.html)",
    ),
    (
        f"{p3} > Topic 3.1: Micrograd",
        "- [The spelled-out intro to neural networks and backpropagation (micrograd)](https://www.youtube.com/watch?v=VMj-3S1tku0)",
    ),
    (
        f"{p3} > Topic 3.2: Makemore and MLPs",
        "- [Building makemore — character-level language modeling](https://www.youtube.com/watch?v=PaCmpygFfXo)",
    ),
    (
        f"{p3} > Topic 3.3: Network Stabilization",
        "- [Building makemore Part 4: Becoming a Backprop Ninja](https://www.youtube.com/watch?v=q8SA3rM6ckI)",
    ),
    (
        f"{p3} > Topic 3.4: Convolutional Sequences",
        "- [Karpathy Zero to Hero series index](https://karpathy.ai/zero-to-hero.html)",
    ),
    # Phase 4
    (
        f"{p4} > Topic 4.1: Experiment Tracking",
        "- [MLOps Zoomcamp (DataTalks.Club overview)](https://datatalks.club/blog/mlops-zoomcamp.html)",
    ),
    (
        f"{p4} > Topic 4.1: Experiment Tracking",
        "- [DataTalksClub/mlops-zoomcamp (GitHub)](https://github.com/DataTalksClub/mlops-zoomcamp)",
    ),
    (
        f"{p4} > Topic 4.2: Pipeline Orchestration",
        "- [MLOps Zoomcamp YouTube playlist](https://www.youtube.com/playlist?list=PL3MmuxUbc_hIUISrluw_A7wDSmfOhErJK)",
    ),
    (
        f"{p4} > Topic 4.3: Cloud Deployment",
        "- [MLOps Zoomcamp (DataTalks.Club) — deployment modules](https://datatalks.club/blog/mlops-zoomcamp.html)",
    ),
    (
        f"{p4} > Topic 4.4: Model Monitoring",
        "- [MLOps Zoomcamp GitHub — monitoring / Evidently modules](https://github.com/DataTalksClub/mlops-zoomcamp)",
    ),
    # Phase 5
    (
        f"{p5} > Topic 5.1: Requirements Engineering",
        "- [Designing Machine Learning Systems — Chip Huyen (O'Reilly)](https://www.oreilly.com/library/view/designing-machine-learning/9781098107956/)",
    ),
    (
        f"{p5} > Topic 5.2: Data Supply Chains",
        "- [chiphuyen/dmls-book (companion materials)](https://github.com/chiphuyen/dmls-book)",
    ),
    (
        f"{p5} > Topic 5.3: Feature Stores",
        "- [MLOps guide — Chip Huyen](https://huyenchip.com/mlops/)",
    ),
    (
        f"{p5} > Topic 5.4: Adaptive Environments",
        "- [Designing Machine Learning Systems — Chip Huyen (book)](https://www.oreilly.com/library/view/designing-machine-learning/9781098107956/)",
    ),
]

# --- write items ---
items_path = data / "plan_items.csv"
with items_path.open(newline="", encoding="utf-8") as f:
    existing_items = list(csv.DictReader(f))
fi = ["id", "plan_id", "section_path", "text", "checked", "sort_order"]

start_item = 829
sort_order = 1
new_items = []
for section, texts in sections:
    for text in texts:
        new_items.append(
            {
                "id": f"PITEM_{start_item:04d}",
                "plan_id": plan_id,
                "section_path": section,
                "text": text,
                "checked": "0",
                "sort_order": str(sort_order),
            }
        )
        start_item += 1
        sort_order += 1

with items_path.open("w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=fi, lineterminator="\n")
    w.writeheader()
    w.writerows(existing_items)
    w.writerows(new_items)

# --- write resources ---
res_path = data / "plan_resources.csv"
with res_path.open(newline="", encoding="utf-8") as f:
    existing_res = list(csv.DictReader(f))
fr = ["id", "plan_id", "section_path", "line", "sort_order"]

start_res = 468
sec_ord: dict[str, int] = defaultdict(int)
new_res = []
for section, line in resources:
    sec_ord[section] += 1
    new_res.append(
        {
            "id": f"PRES_{start_res:04d}",
            "plan_id": plan_id,
            "section_path": section,
            "line": line,
            "sort_order": str(sec_ord[section]),
        }
    )
    start_res += 1

with res_path.open("w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=fr, lineterminator="\n")
    w.writeheader()
    w.writerows(existing_res)
    w.writerows(new_res)

print(f"Added {len(new_items)} items (PITEM_0829..PITEM_{start_item - 1:04d})")
print(f"Added {len(new_res)} resources (PRES_0468..PRES_{start_res - 1:04d})")
