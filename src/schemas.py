"""Pydantic schemas for structured data exchanged with the AI and across the pipeline."""

from __future__ import annotations

from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class TaskType(str, Enum):
    CORRELATION = "correlation"
    REGRESSION = "regression"
    CONFIDENCE_INTERVAL = "confidence_interval"
    PREDICTION = "prediction"
    PREDICTION_INTERVAL = "prediction_interval"
    HYPOTHESIS_TEST = "hypothesis_test"
    DESCRIPTIVE_STATS = "descriptive_stats"
    CHART = "chart"
    OTHER = "other"


class TaskPlan(BaseModel):
    """A single unit of work the AI identified in the workbook."""

    task_id: str
    source_sheet: str
    cell_refs: list[str] = Field(default_factory=list)
    task_type: TaskType
    description: str
    inputs: dict[str, Any] = Field(default_factory=dict)
    steps: list[str] = Field(default_factory=list)
    # Each entry: {"cell": "B10", "formula": "=CORREL(A2:A9,B2:B9)"}
    excel_formulas: list[dict[str, str]] = Field(default_factory=list)
    explanation: str = ""
    # {"sheet": "Arvutused", "start_row": 1}
    output_location: dict[str, Any] = Field(default_factory=dict)
    dependencies: list[str] = Field(default_factory=list)
    verified: bool = False


class WorkbookAnalysis(BaseModel):
    """The AI's complete analysis of a workbook."""

    sheets: list[str] = Field(default_factory=list)
    data_ranges: dict[str, Any] = Field(default_factory=dict)
    tasks: list[TaskPlan] = Field(default_factory=list)
    revision_questions: list[dict[str, Any]] = Field(default_factory=list)


class TaskResult(BaseModel):
    """The computed result of a single task, produced by the Python math engine."""

    task_id: str
    task_type: TaskType
    source_sheet: str
    description: str = ""
    values: dict[str, Any] = Field(default_factory=dict)
    # Each entry: {"cell": "B10", "formula": "=...", "value": 0.885}
    written_cells: list[dict[str, Any]] = Field(default_factory=list)
    output_location: dict[str, Any] = Field(default_factory=dict)
    explanation: str = ""
    verified: bool = False
    error: str | None = None


# JSON schema used as the tool_use input_schema for the analyze_workbook call.
# Kept in sync with WorkbookAnalysis / TaskPlan above.
WORKBOOK_ANALYSIS_TOOL_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "sheets": {"type": "array", "items": {"type": "string"}},
        "data_ranges": {"type": "object", "additionalProperties": True},
        "tasks": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "task_id": {"type": "string"},
                    "source_sheet": {"type": "string"},
                    "cell_refs": {"type": "array", "items": {"type": "string"}},
                    "task_type": {
                        "type": "string",
                        "enum": [t.value for t in TaskType],
                    },
                    "description": {"type": "string"},
                    "inputs": {
                        "type": "object",
                        "additionalProperties": True,
                        "properties": {
                            "x_range": {
                                "type": "string",
                                "description": "Cell range for X data, e.g. 'A2:A20'. Use exact column letter and row numbers from the payload."
                            },
                            "y_range": {
                                "type": "string",
                                "description": "Cell range for Y data, e.g. 'C2:C20'. Must span the same rows as x_range."
                            },
                            "data_range": {
                                "type": "string",
                                "description": "Cell range for single-column data (one-sample tests, descriptive stats), e.g. 'B2:B15'."
                            },
                            "alpha": {
                                "type": "number",
                                "description": "Significance level. Default 0.05 unless the task specifies otherwise."
                            },
                            "x_new": {
                                "type": "number",
                                "description": "Predictor value at which to compute a prediction or mean-response CI."
                            },
                            "mu0": {
                                "type": "number",
                                "description": "Null hypothesis mean for a one-sample t-test."
                            },
                            "method": {
                                "type": "string",
                                "enum": ["pearson", "spearman"],
                                "description": "Correlation method. Omit for default Pearson."
                            },
                            "equal_var": {
                                "type": "boolean",
                                "description": "If true, use pooled-variance (Student) t-test; if false (default), use Welch."
                            }
                        }
                    },
                    "steps": {"type": "array", "items": {"type": "string"}},
                    "excel_formulas": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "cell": {"type": "string"},
                                "formula": {"type": "string"},
                            },
                            "required": ["cell", "formula"],
                        },
                    },
                    "explanation": {"type": "string"},
                    "output_location": {
                        "type": "object",
                        "additionalProperties": True,
                    },
                    "dependencies": {"type": "array", "items": {"type": "string"}},
                },
                "required": [
                    "task_id",
                    "source_sheet",
                    "task_type",
                    "description",
                ],
            },
        },
    },
    "required": ["sheets", "tasks"],
}


REVISION_QUESTIONS_TOOL_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "questions": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "task_id": {"type": "string"},
                    "question": {"type": "string"},
                    "options": {
                        "type": "array",
                        "items": {"type": "string"},
                        "minItems": 4,
                        "maxItems": 4,
                        "description": "Exactly 4 Estonian-language multiple-choice options (labelled A–D internally).",
                    },
                    "correct_option": {
                        "type": "string",
                        "enum": ["A", "B", "C", "D"],
                        "description": "Letter of the one correct option.",
                    },
                    "answer": {"type": "string"},
                    "explanation": {
                        "type": "string",
                        "description": "Concise Estonian explanation of why the answer is correct.",
                    },
                },
                "required": ["question", "options", "correct_option", "answer", "explanation"],
            },
        }
    },
    "required": ["questions"],
}
