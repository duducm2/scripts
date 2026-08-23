"""Build dashboard Plans panel payload from parsed study plans."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from study_plan_parser import default_studies_root, load_all_plans

PLANS_GITHUB_URL = "https://github.com/duducm2/scripts/tree/main/mnemonics/output/plans"
PLANS_SAVE_PORT = 8765
PLANS_SAVE_URL = f"http://127.0.0.1:{PLANS_SAVE_PORT}"


def build_plans_payload(
    studies_root: Path,
    data_dir: Path,
) -> dict[str, Any]:
    plans = load_all_plans(studies_root, data_dir)
    return {
        "plans": plans,
        "github_url": PLANS_GITHUB_URL,
        "save_url": PLANS_SAVE_URL,
    }


def build_plans_panel_shell() -> str:
    return """
  <div id="plansPanel" class="hidden" aria-hidden="true">
    <div class="plans-layout" id="plansLayout">
      <nav class="plans-toc" id="plansToc" aria-label="Plan sections">
        <div class="plans-toc-head">
          <h2>Sections</h2>
          <button type="button" id="btnPlansTocToggle" class="btn-plans-toc-toggle" title="Collapse sections" aria-expanded="true" aria-controls="plansTocList">◀</button>
        </div>
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
            <button type="button" id="btnPlansSave" class="btn-plans-save" title="Write checkbox progress to plan .md files">Save</button>
            <button type="button" id="btnPlansReset" class="btn-plans-reset" title="Clear saved checkbox progress">Reset to file</button>
          </div>
        </div>
        <p class="plans-save-status" id="plansSaveStatus" aria-live="polite"></p>
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
