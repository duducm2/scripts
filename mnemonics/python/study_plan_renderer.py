"""Build dashboard Plans panel payload from parsed study plans."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from study_plan_parser import default_studies_root, load_all_plans

PLANS_GITHUB_URL = (
    "https://github.com/duducm2/scripts/tree/main/mnemonics/output/plans"
)


def build_plans_payload(
    studies_root: Path,
    data_dir: Path,
) -> dict[str, Any]:
    plans = load_all_plans(studies_root, data_dir)
    return {
        "plans": plans,
        "github_url": PLANS_GITHUB_URL,
    }


def build_plans_panel_shell() -> str:
    return """
  <div id="plansPanel" class="hidden" aria-hidden="true">
    <div class="plans-layout">
      <nav class="plans-toc" id="plansToc" aria-label="Plan sections">
        <h2>Sections</h2>
        <ul id="plansTocList"></ul>
      </nav>
      <div class="plans-main">
        <div class="plans-header">
          <div>
            <h2 id="plansTitle">Study Plan</h2>
            <p class="plans-sub" id="plansSub"></p>
          </div>
          <div class="plans-header-actions">
            <a id="plansGithub" class="plans-github" href="#" target="_blank" rel="noopener">Plans on GitHub</a>
            <button type="button" id="btnPlansReset" class="btn-plans-reset" title="Clear saved checkbox progress">Reset to file</button>
          </div>
        </div>
        <div class="plans-progress-wrap">
          <div class="plans-progress-bar" id="plansProgressBar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0"></div>
          <span class="plans-progress-label" id="plansProgressLabel">0 / 0 complete</span>
        </div>
        <div id="plansContent" class="plans-content"></div>
        <p class="plans-empty hidden" id="plansEmpty">No study plan file for this topic.</p>
      </div>
    </div>
  </div>
"""


def default_plans_payload(data_dir: Path | None = None) -> dict[str, Any]:
    data = data_dir or (Path(__file__).resolve().parent.parent / "data")
    return build_plans_payload(default_studies_root(), data)
