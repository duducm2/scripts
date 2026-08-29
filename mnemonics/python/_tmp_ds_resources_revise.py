"""Revise PLAN_0008 resources: fix weak links + add essential reading."""

from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path

data = Path(__file__).resolve().parent.parent / "data"
res_path = data / "plan_resources.csv"
plan_id = "PLAN_0008"

with res_path.open(newline="", encoding="utf-8") as f:
    rows = list(csv.DictReader(f))
fr = ["id", "plan_id", "section_path", "line", "sort_order"]

# Drop Class Central aggregator mirrors (prefer Karpathy hub / direct YT)
drop_ids = {"PRES_0498", "PRES_0499"}
rows = [r for r in rows if r["id"] not in drop_ids]

# Fix weak / mismatched rows
fixes = {
    # Channel search → concrete StatQuest XGBoost intro + official docs added below
    "PRES_0471": (
        "- [XGBoost: A Scalable Tree Boosting System (original paper)]"
        "(https://arxiv.org/abs/1603.02754)"
    ),
    "PRES_0472": (
        "- [PCA Main Ideas (StatQuest)](https://www.youtube.com/watch?v=FgakZw6K1QQ)"
    ),
    # Optimization was wrongly linked to Lecture 1 intro — point to CS229 notes
    "PRES_0475": (
        "- [CS229 Lecture Notes (Stanford)](https://cs229.stanford.edu/main_notes.pdf)"
    ),
    # Kernels/SVM was a NN lecture — use notes + keep playlist on 1.1
    "PRES_0476": (
        "- [CS229 Course materials / notes index](https://cs229.stanford.edu/)"
    ),
    # Duplicate Zoomcamp overview on deployment → Docker docs
    "PRES_0491": ("- [Docker docs: Get started](https://docs.docker.com/get-started/)"),
    # Duplicate GitHub on monitoring → Evidently docs
    "PRES_0492": ("- [Evidently AI documentation](https://docs.evidentlyai.com/)"),
    # O'Reilly paywall-ish; keep but also have free companions elsewhere
    "PRES_0493": (
        "- [Designing Machine Learning Systems — Chip Huyen (book site / buy)]"
        "(https://www.oreilly.com/library/view/designing-machine-learning/9781098107956/)"
    ),
}

for r in rows:
    if r["id"] in fixes:
        r["line"] = fixes[r["id"]]

p0 = "Phase 0: Solidifying Statistical and Machine Learning Fundamentals"
p1 = "Phase 1: Rigorous Algorithmic Foundations"
p2 = "Phase 2: Advanced Experimentation and Causal Inference"
p3 = "Phase 3: Deep Learning and Modern Architectures"
p4 = "Phase 4: Machine Learning Operations (MLOps)"
p5 = "Phase 5: Machine Learning System Design"

additions: list[tuple[str, str]] = [
    # Phase 0 — essential reading
    (
        f"{p0} > Topic 0.1: Algorithmic Intuition",
        "- [scikit-learn: Cross-validation](https://scikit-learn.org/stable/modules/cross_validation.html)",
    ),
    (
        f"{p0} > Topic 0.1: Algorithmic Intuition",
        "- [An Introduction to Statistical Learning (ISL) — free PDF](https://www.statlearning.com/)",
    ),
    (
        f"{p0} > Topic 0.2: Advanced Tree Models",
        "- [XGBoost documentation (official)](https://xgboost.readthedocs.io/en/stable/)",
    ),
    (
        f"{p0} > Topic 0.2: Advanced Tree Models",
        "- [StatQuest XGBoost playlist search (channel)](https://www.youtube.com/c/joshstarmer/search?query=xgboost)",
    ),
    (
        f"{p0} > Topic 0.3: Dimensionality Reduction",
        "- [scikit-learn: PCA](https://scikit-learn.org/stable/modules/generated/sklearn.decomposition.PCA.html)",
    ),
    # Phase 1
    (
        f"{p1} > Topic 1.1: Empirical Risk",
        "- [CS229 Lecture Notes PDF (Stanford)](https://cs229.stanford.edu/main_notes.pdf)",
    ),
    (
        f"{p1} > Topic 1.2: Optimization",
        "- [CS229 Autumn 2018 notes — Linear Algebra / Probability review](https://cs229.stanford.edu/section/cs229-linalg.pdf)",
    ),
    (
        f"{p1} > Topic 1.3: Advanced Classifiers",
        "- [Stanford CS229 Machine Learning — Spring 2022 playlist](https://www.youtube.com/playlist?list=PLoROMvodv4rNyWOpJg_Yh4NSqI4Z4vOYy)",
    ),
    # Phase 2
    (
        f"{p2} > Topic 2.1: Variance Reduction (CUPED)",
        "- [Improving the Sensitivity of Online Controlled Experiments (Deng et al., CUPED paper)]"
        "(https://www.microsoft.com/en-us/research/publication/improving-the-sensitivity-of-online-controlled-experiments-by-utilizing-pre-experiment-data/)",
    ),
    (
        f"{p2} > Topic 2.2: Multi-Armed Bandits",
        "- [Lil'Log: Multi-Armed Bandits](https://lilianweng.github.io/posts/2018-01-23-multi-armed-bandit/)",
    ),
    (
        f"{p2} > Topic 2.3: Causal Inference",
        "- [Brady Neal — Introduction to Causal Inference (course site)](https://www.bradyneal.com/causal-inference-course)",
    ),
    (
        f"{p2} > Topic 2.3: Causal Inference",
        "- [The Book of Why — Pearl & Mackenzie (conceptual causal literacy)]"
        "(https://www.basicbooks.com/titles/judea-pearl/the-book-of-why/9780465097609/)",
    ),
    # Phase 3
    (
        f"{p3} > Topic 3.1: Micrograd",
        "- [karpathy/micrograd (GitHub)](https://github.com/karpathy/micrograd)",
    ),
    (
        f"{p3} > Topic 3.2: Makemore and MLPs",
        "- [karpathy/nn-zero-to-hero (GitHub notebooks)](https://github.com/karpathy/nn-zero-to-hero)",
    ),
    (
        f"{p3} > Topic 3.2: Makemore and MLPs",
        "- [PyTorch: Tensors tutorial](https://pytorch.org/tutorials/beginner/basics/tensorqs_tutorial.html)",
    ),
    (
        f"{p3} > Topic 3.3: Network Stabilization",
        "- [Delving Deep into Rectifiers (Kaiming He et al. — init paper)](https://arxiv.org/abs/1502.01852)",
    ),
    (
        f"{p3} > Topic 3.3: Network Stabilization",
        "- [Batch Normalization paper (Ioffe & Szegedy)](https://arxiv.org/abs/1502.03167)",
    ),
    (
        f"{p3} > Topic 3.4: Convolutional Sequences",
        "- [WaveNet: A Generative Model for Raw Audio (DeepMind blog / paper)]"
        "(https://www.deepmind.com/blog/wavenet-a-generative-model-for-raw-audio)",
    ),
    # Phase 4
    (
        f"{p4} > Topic 4.1: Experiment Tracking",
        "- [MLflow Tracking documentation](https://mlflow.org/docs/latest/tracking.html)",
    ),
    (
        f"{p4} > Topic 4.1: Experiment Tracking",
        "- [MLflow Model Registry](https://mlflow.org/docs/latest/model-registry.html)",
    ),
    (
        f"{p4} > Topic 4.2: Pipeline Orchestration",
        "- [Prefect docs: Getting started](https://docs.prefect.io/latest/get-started/index.html)",
    ),
    (
        f"{p4} > Topic 4.3: Cloud Deployment",
        "- [FastAPI documentation](https://fastapi.tiangolo.com/)",
    ),
    (
        f"{p4} > Topic 4.4: Model Monitoring",
        "- [Prometheus Getting started](https://prometheus.io/docs/prometheus/latest/getting_started/)",
    ),
    # Phase 5
    (
        f"{p5} > Topic 5.1: Requirements Engineering",
        "- [Rules of Machine Learning (Google — Martin Zinkevich)]"
        "(https://developers.google.com/machine-learning/guides/rules-of-ml)",
    ),
    (
        f"{p5} > Topic 5.2: Data Supply Chains",
        "- [MLOps guide — Chip Huyen](https://huyenchip.com/mlops/)",
    ),
    (
        f"{p5} > Topic 5.3: Feature Stores",
        "- [Feast documentation (open-source feature store)](https://docs.feast.dev/)",
    ),
    (
        f"{p5} > Topic 5.4: Adaptive Environments",
        "- [Hidden Technical Debt in Machine Learning Systems (Sculley et al., Google)]"
        "(https://papers.nips.cc/paper_files/paper/2015/hash/86df7dcfd896fcaf2674f757a2463eba-Abstract.html)",
    ),
]

max_rid = max(int(r["id"].split("_")[1]) for r in rows if r["id"].startswith("PRES_"))
sec_ord: dict[str, int] = defaultdict(int)
for r in rows:
    if r["plan_id"] == plan_id:
        sec_ord[r["section_path"]] = max(
            sec_ord[r["section_path"]], int(r["sort_order"] or 0)
        )

# Avoid duplicate lines on same section
existing_lines = {
    (r["section_path"], r["line"]) for r in rows if r["plan_id"] == plan_id
}

rid = max_rid + 1
new_rows = []
for sec, line in additions:
    if (sec, line) in existing_lines:
        continue
    # Also skip if MLOps guide already on 5.3 and we're adding to 5.2 — fine, different sections
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
    existing_lines.add((sec, line))
    rid += 1

rows.extend(new_rows)
with res_path.open("w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=fr, lineterminator="\n")
    w.writeheader()
    w.writerows(rows)

print(f"PLAN_0008 resources: {sum(1 for r in rows if r['plan_id'] == plan_id)}")
print(f"Added {len(new_rows)}; fixed {len(fixes)}; dropped {len(drop_ids)}")
