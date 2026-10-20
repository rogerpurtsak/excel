"""Claude API client: workbook analysis and revision-question generation.

Uses the Anthropic SDK with structured output via tool_use. Tracks token usage,
estimates cost, enforces a per-workbook budget, and retries transient failures
with exponential backoff.
"""

from __future__ import annotations

import json
import random
import time
from typing import Any

try:
    import anthropic
except Exception:  # pragma: no cover
    anthropic = None  # type: ignore[assignment]

from .schemas import (
    REVISION_QUESTIONS_TOOL_SCHEMA,
    WORKBOOK_ANALYSIS_TOOL_SCHEMA,
    WorkbookAnalysis,
)

ANALYZE_TOOL_NAME = "record_workbook_analysis"
REVISION_TOOL_NAME = "record_revision_questions"

SYSTEM_PROMPT = (
    "You are a statistics teaching assistant that analyzes Excel workbooks "
    "containing statistics exercises (in any language, not necessarily Estonian) "
    "and produces a structured solution plan.\n\n"
    "CRITICAL — CELL REFERENCES:\n"
    "The workbook payload includes every non-empty cell with its exact column "
    "letter and 1-based row number (e.g. row 3 column C → 'C3'). You MUST read "
    "these coordinates from the payload and use them in your plan. Never assume "
    "data is in columns A and B. Never assume a fixed row count. Never guess "
    "where data lives — read it from the payload.\n\n"
    "For every task that involves two data series (regression, correlation, "
    "two-sample test), set inputs.x_range and inputs.y_range to the exact cell "
    "ranges you observed, e.g. 'C2:C20' and 'E2:E20'. For tasks with one data "
    "series (one-sample test, descriptive statistics), set inputs.data_range.\n\n"
    "SUPPORTED TASK TYPES:\n"
    "  correlation, regression, confidence_interval, prediction, "
    "prediction_interval, hypothesis_test, descriptive_stats, chart, other\n\n"
    "OPTIONAL INPUT FIELDS (include only when relevant):\n"
    "  alpha (default 0.05), x_new (required for prediction/CI at a point), "
    "mu0 (required for one-sample hypothesis test), "
    "method ('pearson'|'spearman'), equal_var (boolean, two-sample t-test)\n\n"
    "For each task also provide:\n"
    "  • excel_formulas: exact Excel formulas referencing real cell addresses "
    "    (CORREL, SLOPE, INTERCEPT, RSQ, FORECAST, CONFIDENCE.T, T.TEST, etc.)\n"
    "  • steps: numbered solution steps\n"
    "  • explanation: plain-language explanation of what the task asks\n\n"
    "The Python engine performs the actual computation; your job is the plan. "
    "Always call the provided tool to return your analysis."
)


class BudgetExceededError(RuntimeError):
    """Raised when a request would exceed the per-workbook budget."""


class AIClient:
    def __init__(self, config, api_key: str | None = None):
        self.config = config
        self._api_key = api_key or config.get_api_key()
        self.input_tokens = 0
        self.output_tokens = 0
        self._client = None
        if anthropic is not None and self._api_key:
            self._client = anthropic.Anthropic(api_key=self._api_key)

    # ------------------------------------------------------------------ cost
    @property
    def total_cost(self) -> float:
        price = self.config.price_for()
        return (
            self.input_tokens / 1_000_000 * price.get("input", 0.0)
            + self.output_tokens / 1_000_000 * price.get("output", 0.0)
        )

    def _record_usage(self, usage) -> None:
        if usage is None:
            return
        self.input_tokens += int(getattr(usage, "input_tokens", 0) or 0)
        self.output_tokens += int(getattr(usage, "output_tokens", 0) or 0)

    def _check_budget(self) -> None:
        if self.total_cost >= self.config.budget_per_workbook:
            raise BudgetExceededError(
                f"Eelarve ${self.config.budget_per_workbook:.2f} täis "
                f"(kulutatud ${self.total_cost:.4f}). Tõsta eelarvet Seadetes."
            )

    # ------------------------------------------------------------------ core call
    def _call_with_tool(
        self,
        prompt: str,
        tool_name: str,
        tool_schema: dict[str, Any],
        max_retries: int = 3,
    ) -> dict[str, Any]:
        if self._client is None:
            if anthropic is None:
                raise RuntimeError("Anthropic SDK pole paigaldatud.")
            raise RuntimeError(
                "API võti puudub. Ava Seaded ja lisa oma Anthropic API võti."
            )
        self._check_budget()

        tools = [
            {
                "name": tool_name,
                "description": "Record the structured result.",
                "input_schema": tool_schema,
            }
        ]

        last_exc: Exception | None = None
        for attempt in range(max_retries):
            try:
                resp = self._client.messages.create(
                    model=self.config.model,
                    max_tokens=8000,
                    system=SYSTEM_PROMPT,
                    tools=tools,
                    tool_choice={"type": "auto"},
                    messages=[{"role": "user", "content": prompt}],
                )
                self._record_usage(getattr(resp, "usage", None))
                for block in resp.content:
                    if getattr(block, "type", None) == "tool_use" and block.name == tool_name:
                        return dict(block.input)
                raise RuntimeError(
                    f"Mudel ei kasutanud tööriista '{tool_name}'. Proovitakse uuesti."
                )
            except Exception as exc:
                last_exc = exc
                # Never retry authentication or explicit bad-request errors.
                if anthropic is not None and isinstance(
                    exc,
                    (
                        getattr(anthropic, "AuthenticationError", ()),
                        getattr(anthropic, "BadRequestError", ()),
                    ),
                ):
                    raise
                if isinstance(exc, BudgetExceededError):
                    raise
                if attempt == max_retries - 1:
                    break
                delay = min(2 ** attempt + random.uniform(0, 1), 30.0)
                time.sleep(delay)
        raise last_exc if last_exc else RuntimeError("AI päring ebaõnnestus")

    # ------------------------------------------------------------------ public API
    def analyze_workbook(self, inspection_data: dict[str, Any]) -> WorkbookAnalysis:
        """One API call: analyze the workbook and return a WorkbookAnalysis."""
        prompt = (
            "Analyze this Excel workbook and plan how to solve every statistics "
            "task it contains. Use the exact cell coordinates from the payload "
            "to set x_range, y_range, or data_range in each task's inputs — "
            "do not assume column layout. Workbook data (JSON):\n\n"
            + json.dumps(inspection_data, ensure_ascii=False, indent=2)
            + "\n\nReturn the plan via the tool."
        )
        raw = self._call_with_tool(
            prompt, ANALYZE_TOOL_NAME, WORKBOOK_ANALYSIS_TOOL_SCHEMA
        )
        return WorkbookAnalysis.model_validate(raw)

    def generate_revision_questions(
        self, tasks: list[dict[str, Any]], results: list[dict[str, Any]]
    ) -> list[dict[str, Any]]:
        """One API call: produce revision questions with answers for the solved tasks."""
        prompt = (
            "Based on these solved statistics tasks and their results, write short "
            "Estonian-language revision questions (with answers) a student could use "
            "to check understanding.\n\nTasks:\n"
            + json.dumps(tasks, ensure_ascii=False, indent=2)
            + "\n\nResults:\n"
            + json.dumps(results, ensure_ascii=False, indent=2)
        )
        raw = self._call_with_tool(
            prompt, REVISION_TOOL_NAME, REVISION_QUESTIONS_TOOL_SCHEMA
        )
        return list(raw.get("questions", []))
