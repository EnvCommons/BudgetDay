from __future__ import annotations

import asyncio
import base64
import io
import json
import re
from pathlib import Path
from typing import Any

import openpyxl
import pptx
from openreward import AsyncOpenReward, SandboxBucketConfig, SandboxSettings
from openreward.chat_backends import resolve_backend
from openreward.environments import JSONObject, TextBlock, ToolOutput, tool, Split
from openreward.toolsets import WordToolset, ExcelToolset, PowerPointToolset, PDFToolset
from pydantic import BaseModel

from cli_environment import CLIEnvironment
from constants import ENV_PATH


# Load rubrics at module level
with open(ENV_PATH / "rubrics.json") as f:
    RUBRICS = json.load(f)

# Define tasks
TASKS = [
    {
        "task_id": "budget_2020_psnb_forecast",
        "task_type": "chart",
        "year": 2020,
        "budget_name": "Budget 2020",
        "data_dir": "2020",
        "output_path": "/home/ubuntu/psnb_changes.xlsx",
        "output_path_png": "/home/ubuntu/psnb_changes.png",
        "description": "Create spreadsheet showing changes to OBR's forecast for PSNB since restated 2019 forecast",
    },
    {
        "task_id": "budget_2022_initial_response",
        "task_type": "report",
        "year": 2022,
        "budget_name": "Autumn Statement 2022",
        "data_dir": "2022",
        "output_path": "/home/ubuntu/final_report.md",
        "description": "Draft initial response to Autumn Statement 2022",
    },
    {
        "task_id": "budget_2023_initial_response",
        "task_type": "report",
        "year": 2023,
        "budget_name": "Spring Budget 2023",
        "data_dir": "2023",
        "output_path": "/home/ubuntu/final_report.md",
        "description": "Draft initial response to Spring Budget 2023",
    },
    {
        "task_id": "budget_2024_initial_response",
        "task_type": "report",
        "year": 2024,
        "budget_name": "Autumn Budget 2024",
        "data_dir": "2024",
        "output_path": "/home/ubuntu/final_report.md",
        "description": "Draft initial response to Autumn Budget 2024",
    },
    {
        "task_id": "budget_2025_initial_response",
        "task_type": "report",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/final_report.md",
        "description": "Draft initial response to 2025 Budget",
    },
    {
        "task_id": "budget_2025_borrowing_chart",
        "task_type": "chart",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/changes_borrowing.xlsx",
        "output_path_png": "/home/ubuntu/changes_borrowing.png",
        "description": "Create borrowing forecast comparison chart and spreadsheet",
    },
    {
        "task_id": "budget_2025_productivity_effect",
        "task_type": "qa",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/explanation.txt",
        "description": "Identify the effect of productivity downgrade on 2029-30 revenues",
        "question": "What was the effect of the productivity downgrade on 2029-30 revenues?",
        "expected_answer": "-£16bn",
    },
    {
        "task_id": "budget_2025_income_change_2028_29",
        "task_type": "numerical_qa",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/income_change.txt",
        "description": "Calculate annual income change in 2028-29 for specific household scenario",
        "question": "As a result of the 2025 Budget, what is the annual income change in 2028-29, adjusted for inflation, for a single parent with three children, working full-time at National Living Wage, living in Wales, and driving a petrol car?",
        "expected_answer": 4960,
        "margin_percent": 2.0,
    },
    {
        "task_id": "budget_2025_income_change_single_parent_four_children",
        "task_type": "numerical_qa",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/income_change_sp4c.txt",
        "description": "Calculate annual income change in 2028-29 for single parent with four children",
        "question": "As a result of the 2025 Budget, what is the annual income change for 2028-29 (adjusted for inflation) for a single parent with four children, not working but is subject to the two child benefit cap, does not drive and lives in London?",
        "expected_answer": 135,
        "margin_percent": 2.0,
    },
    {
        "task_id": "budget_2025_income_change_single_top_earner",
        "task_type": "numerical_qa",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/income_change_top_earner.txt",
        "description": "Calculate annual income change in 2028-29 for single top earner",
        "question": "As a result of the 2025 Budget, what is the annual income change for 2028-29 (adjusted for inflation) for a single adult who works a full-time, top earning job, lives in London and drives an electric vehicle?",
        "expected_answer": -295,
        "margin_percent": 2.0,
    },
    {
        "task_id": "budget_2025_income_change_pensioner_couple",
        "task_type": "numerical_qa",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/income_change_pensioners.txt",
        "description": "Calculate annual income change in 2028-29 for pensioner couple",
        "question": "As a result of the 2025 Budget, what is the annual income change for 2028-29 (adjusted for inflation) for a pensioner couple who live in Yorkshire, both receive an average private pension and drive a petrol car?",
        "expected_answer": 65,
        "margin_percent": 2.0,
    },
    {
        "task_id": "budget_2025_income_change_couple_two_children",
        "task_type": "numerical_qa",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/income_change_couple_2c.txt",
        "description": "Calculate annual income change in 2028-29 for couple with two children",
        "question": "As a result of the 2025 Budget, what is the annual income change for 2028-29 (adjusted for inflation) for a couple with two children who work low-to-average wage jobs in the South East and commute to work on the trains?",
        "expected_answer": 235,
        "margin_percent": 2.0,
    },
    {
        "task_id": "budget_2025_income_change_couple_no_children",
        "task_type": "numerical_qa",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/income_change_couple_0c.txt",
        "description": "Calculate annual income change in 2028-29 for couple without children",
        "question": "As a result of the 2025 Budget, what is the annual income change for 2028-29 (adjusted for inflation) for a couple without children, where both work full time in average-to-high wage jobs in the South West and drive a petrol car?",
        "expected_answer": -70,
        "margin_percent": 2.0,
    },
    {
        "task_id": "budget_2025_income_change_poorest_10pct",
        "task_type": "numerical_qa",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/income_change_poorest_10pct.txt",
        "description": "Calculate forecasted annual income change for poorest 10% in 2030-31",
        "question": "As a result of the 2025 Budget, what is the forecasted annual change in income for the poorest 10% as a result of selected measures in 2030-31, adjusted for inflation?",
        "expected_answer": 225,
        "margin_percent": 2.0,
    },
    {
        "task_id": "budget_2025_income_change_richest_10pct",
        "task_type": "numerical_qa",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/income_change_richest_10pct.txt",
        "description": "Calculate forecasted annual income change for richest 10% in 2030-31",
        "question": "As a result of the 2025 Budget, what is the forecasted annual change in income for the richest 10% as a result of selected measures in 2030-31, adjusted for inflation?",
        "expected_answer": -709,
        "margin_percent": 2.0,
    },
    {
        "task_id": "budget_2025_policy_decisions",
        "task_type": "chart",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/policy_decisions.xlsx",
        "output_path_png": "/home/ubuntu/policy_decisions.png",
        "description": "Create spreadsheet and chart showing effect of spending and tax decisions on borrowing",
    },
    {
        "task_id": "budget_deficit_comparison",
        "task_type": "chart",
        "year": 2025,
        "budget_name": "Budget Deficit Comparison",
        "data_dir": None,  # Needs both 2024 and 2025 data
        "output_path": "/home/ubuntu/current_budget_deficit.xlsx",
        "output_path_png": "/home/ubuntu/current_budget_deficit.png",
        "description": "Compare current budget deficit (% GDP) forecasts between October 2024 and November 2025 budgets",
    },
    {
        "task_id": "budget_2024_tax_measures",
        "task_type": "presentation",
        "year": 2024,
        "budget_name": "Autumn Budget 2024",
        "data_dir": "2024",
        "output_path": "/home/ubuntu/tax_measures.pptx",
        "description": "Create a PowerPoint presentation summarizing key tax measures from the Autumn Budget 2024",
    },
    {
        "task_id": "ai_measures_summary_2020_2025",
        "task_type": "report",
        "year": 2025,
        "budget_name": "AI Measures Summary 2020-2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/ai_measures_report.md",
        "description": "Summarize all AI-related measures and announcements from budgets and autumn statements between 2020 and 2025, including specific funding amounts, initiatives, and strategic programmes.",
    },
    {
        "task_id": "budget_2025_tax_proposals",
        "task_type": "tax_proposal",
        "year": 2025,
        "budget_name": "Budget 2025 Tax Policy",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/tax_proposals.md",
        "description": "Propose tax changes to reduce borrowing for 2026-27 by half using provided tax raising guidelines",
    },
    {
        "task_id": "budget_2025_opposition_response",
        "task_type": "report",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/opposition_response.md",
        "description": "Write Leader of the Opposition's response to Budget 2025",
    },
]


BORROWING_GROUND_TRUTH = {
    "years": ["2025-26", "2026-27", "2027-28", "2028-29", "2029-30"],
    "march_2025": [117.7, 97.2, 80.2, 77.4, 74.0],
    "october_2025": [138.3, 112.1, 98.5, 86.9, 67.9],
    "difference": [20.6, 14.9, 18.3, 9.5, -6.2],
}

POLICY_DECISIONS_GROUND_TRUTH = {
    "years": ["2025-26", "2026-27", "2027-28", "2028-29", "2029-30"],
    "spending_decisions": [4.9, 6.6, 16.0, 12.9, 11.3],
    "tax_decisions": [-1.3, -0.7, -6.1, -13.9, -26.1],
}

CURRENT_BUDGET_DEFICIT_GROUND_TRUTH = {
    "years": ["2025-26", "2026-27", "2027-28", "2028-29", "2029-30"],
    "budget_2024": [0.9, 0.2, -0.3, -0.3, -0.3],  # October 2024 (green line)
    "budget_2025": [1.7, 0.9, 0.1, -0.1, -0.6],   # November 2025 (yellow line)
}

PSNB_2020_GROUND_TRUTH = {
    "years": ["2019-20", "2020-21", "2021-22", "2022-23", "2023-24"],
    "restated_march_2019": [47.6, 40.2, 37.6, 35.4, 33.3],
    "budget_2020": [47.4, 54.8, 66.7, 61.5, 60.2],
    "difference": [-0.2, 14.6, 29.1, 26.0, 26.9],
}


class SubmitAnswerInput(BaseModel):
    """Input for submit_answer tool (no parameters needed - reads from fixed path)."""

    pass


class BudgetDay(CLIEnvironment):
    """Budget Day environment - stub implementation"""

    toolsets = [WordToolset, ExcelToolset, PowerPointToolset, PDFToolset]

    @classmethod
    def list_splits(cls) -> list[Split]:
        """Return available data splits"""
        return [Split(name="train", type="train")]

    @classmethod
    def list_tasks(cls, split:str) -> list[JSONObject]:
        """Return list of tasks for a given split"""
        if split != "train":
            return []

        return TASKS

    def __init__(self, task_spec: JSONObject, secrets: dict[str, str] = {}) -> None:
        super().__init__(task_spec, secrets=secrets)

        # Load task data
        self.task_id = str(task_spec["task_id"])
        self.task_data = next(
            (t for t in TASKS if t["task_id"] == self.task_id),
            None,
        )
        if not self.task_data:
            raise ValueError(f"Unknown task_id: {self.task_id}")

        # Grader traffic goes through the SDK's chat-backend layer, which pins
        # base_url + key explicitly so OPENAI_* env rewrites (training gateway)
        # can't silently redirect the judge.
        try:
            self.grader_client = resolve_backend(secrets=secrets)
        except ValueError as e:
            raise ValueError(
                "OpenAI API key required in secrets for rubric grading. Provide via "
                "secrets={'openai_api_key': 'sk-...'} or 'chat_api_key'"
            ) from e

        self.sandbox_settings = SandboxSettings(
            environment="GeneralReasoning/BudgetDay",
            image="generalreasoning/knowledge-worker:latest",
            machine_size="2:4",
            block_network=False,
            bucket_config=SandboxBucketConfig(
                mount_path="/orwd_data",
                read_only=True,
                only_dir=self.task_data["data_dir"] if self.task_data.get("data_dir") else None,  # Only mount budgetday/2025
            ),
        )

        or_client = AsyncOpenReward(api_key=secrets.get("api_key", ""))
        self.sandbox = or_client.sandbox(self.sandbox_settings)

        self.submitted = False

    async def setup(self) -> None:
        """Start sandbox before agent interaction begins"""
        await self.sandbox.start()

    async def get_prompt(self) -> list[TextBlock]:
        """Return task prompt with data location and submission instructions."""

        task_type = self.task_data.get("task_type", "report")  # default to report for backwards compatibility

        # Route based on task type
        if task_type == "qa":
            return [TextBlock(text=self._get_qa_prompt())]
        elif task_type == "numerical_qa":
            return [TextBlock(text=self._get_numerical_qa_prompt())]
        elif task_type == "chart":
            # Check which chart task
            if self.task_data["task_id"] == "budget_2020_psnb_forecast":
                return [TextBlock(text=self._get_psnb_2020_prompt())]
            elif self.task_data["task_id"] == "budget_2025_policy_decisions":
                return [TextBlock(text=self._get_policy_decisions_prompt())]
            elif self.task_data["task_id"] == "budget_deficit_comparison":
                return [TextBlock(text=self._get_current_budget_deficit_prompt())]
            else:  # budget_2025_borrowing_chart
                return [TextBlock(text=self._get_borrowing_chart_prompt())]
        elif task_type == "presentation":
            return [TextBlock(text=self._get_presentation_prompt())]
        elif task_type == "tax_proposal":
            return [TextBlock(text=self._get_tax_proposals_prompt())]
        elif task_type == "report" and self.task_data["task_id"] == "ai_measures_summary_2020_2025":
            return [TextBlock(text=self._get_ai_measures_prompt())]
        elif task_type == "report" and self.task_data["task_id"] == "budget_2025_opposition_response":
            return [TextBlock(text=self._get_opposition_response_prompt())]

        # Base prompt common to all report-writing tasks
        base_prompt = f"""# Task: Draft Initial Response to {self.task_data['budget_name']}

You are a policy analyst at a fiscal policy research organization. Your task is to draft an "initial response" to the {self.task_data['budget_name']} - a rapid policy analysis meant to be written soon after Budget Day.

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only).

## Your Task

Write a comprehensive initial response that provides rapid but serious analysis of the Budget.
"""

        submission_instructions = f"""

## Submission

When ready, write your final report to: **{self.task_data['output_path']}**

Then call `submit_answer` tool to submit for evaluation.
"""

        prompt_text = base_prompt + submission_instructions

        return [TextBlock(text=prompt_text)]

    def _get_borrowing_chart_prompt(self) -> str:
        """Return the prompt for the borrowing forecast comparison task."""
        return """# Task: Create Borrowing Forecast Comparison Chart and Spreadsheet

You are a policy analyst. Your task is to extract borrowing forecast data from the Budget 2025 documents and create two output files comparing the March 2025 and October 2025 (Budget 2025) OBR borrowing forecasts for Public Sector Net Borrowing (PSNB).

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only). The borrowing forecast data can be found within these documents for the fiscal years 2025-26 through 2029-30.

You need to find and compare:
- The March 2025 OBR forecast for borrowing (PSNB) for each year
- The October 2025 (Budget 2025) OBR forecast for borrowing (PSNB) for each year
- The difference between the two forecasts for each year

## Required Outputs

### 1. Bar Chart: `/home/ubuntu/changes_borrowing.png`
Create a bar chart showing the DIFFERENCE in borrowing forecasts (October 2025 minus March 2025) for each fiscal year from 2025-26 to 2029-30.
- The bars MUST be green
- The x-axis should show the tax years (2025-26 through 2029-30)
- The y-axis should show the change in borrowing (in GBP billions)
- Include appropriate title and axis labels
- Note: most years show an increase in borrowing (positive difference), but 2029-30 shows a decrease (negative difference)

### 2. Spreadsheet: `/home/ubuntu/changes_borrowing.xlsx`
Create an Excel spreadsheet containing:
- The tax years (2025-26 through 2029-30)
- The March 2025 borrowing forecast values for each year
- The October 2025 borrowing forecast values for each year
- The difference between the October and March forecasts for each year

## Submission

When both files are ready at the paths above, call `submit_answer` to submit for evaluation.

Your submission will be graded on:
- Accuracy of the spreadsheet data against the official OBR figures
- Quality and correctness of the bar chart (green bars, correct years, accurate differences)
"""

    def _get_psnb_2020_prompt(self) -> str:
        """Return the prompt for the Budget 2020 PSNB forecast comparison task."""
        return """# Task: Create PSNB Forecast Comparison Chart and Spreadsheet

You are a policy analyst. Your task is to extract PSNB (Public Sector Net Borrowing) forecast data from Budget 2020 documents and create two output files comparing the restated March 2019 OBR forecast with the Budget 2020 forecast.

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only). The PSNB forecast data can be found within these documents for fiscal years 2019-20 through 2023-24.

You need to find and compare:
- The restated March 2019 OBR forecast for PSNB for each year
- The Budget 2020 OBR forecast for PSNB for each year
- The difference between the two forecasts (Budget 2020 minus restated March 2019)

## Required Outputs

### 1. Spreadsheet: `/home/ubuntu/psnb_changes.xlsx`
Create an Excel spreadsheet containing:
- The fiscal years (2019-20 through 2023-24)
- The restated March 2019 PSNB forecast values for each year (£ billion)
- The Budget 2020 PSNB forecast values for each year (£ billion)
- The difference between Budget 2020 and restated March 2019 forecasts for each year (£ billion)

The spreadsheet may be organized as rows or columns - either orientation is acceptable.

### 2. Bar Chart (Optional): `/home/ubuntu/psnb_changes.png`
Optionally create a bar chart showing the DIFFERENCE in PSNB forecasts (Budget 2020 minus restated March 2019) for each fiscal year from 2019-20 to 2023-24.
- The x-axis should show the fiscal years (2019-20 through 2023-24)
- The y-axis should show the change in PSNB (in £ billions)
- Include appropriate title and axis labels
- Note: 2019-20 shows a small decrease (-0.2bn), all other years show increases

## Submission

When the spreadsheet is ready (and optionally the chart), call `submit_answer` to submit for evaluation.

Your submission will be graded on:
- Accuracy of the spreadsheet data against the official OBR figures (tolerance: ±0.5 £bn)
- Quality and correctness of the bar chart if provided
"""

    def _get_qa_prompt(self) -> str:
        """Return the prompt for Q&A tasks."""
        question = self.task_data["question"]
        output_path = self.task_data["output_path"]

        return f"""# Task: Budget 2025 Analysis Question

You are a policy analyst. Your task is to answer the following question about Budget 2025 by analyzing the budget documents.

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only).

## Question

{question}

## Your Task

Write your answer with a brief explanation (1-3 sentences) to: **{output_path}**

Then call `submit_answer` to submit for evaluation.
"""

    def _get_numerical_qa_prompt(self) -> str:
        """Return the prompt for numerical Q&A tasks."""
        question = self.task_data["question"]
        output_path = self.task_data["output_path"]

        return f"""# Task: Budget 2025 Numerical Analysis Question

You are a policy analyst. Your task is to answer the following question by analyzing Budget 2025 documents.

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only).

## Question

{question}

## Important Instructions

1. Find the relevant data in the Budget 2025 documents
2. Calculate the annual income change for this specific scenario
3. Ensure the figure is adjusted for inflation
4. Write ONLY the number (as an integer) to: **{output_path}**
   - Do NOT include currency symbols (£)
   - Do NOT include commas or thousand separators
   - Do NOT include any explanatory text
   - Example: If the answer is £4,960, write: 4960

## Your Task

Write your numerical answer to the file path specified above, then call `submit_answer` to submit for evaluation.
"""

    def _get_policy_decisions_prompt(self) -> str:
        """Return the prompt for the policy decisions chart task."""
        return """# Task: Create Policy Decisions Impact Chart and Spreadsheet

You are a policy analyst. Your task is to extract data from the Budget 2025 documents showing the effect of new spending decisions and new tax decisions on the change in borrowing since March 2025.

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only). The data can be found within these documents for the fiscal years 2025-26 through 2029-30.

You need to find:
- The effect of new spending decisions on borrowing for each year
- The effect of new tax decisions on borrowing for each year

## Required Outputs

### 1. Bar Chart (Grouped or Stacked): `/home/ubuntu/policy_decisions.png`
Create a bar chart (either grouped or stacked) showing TWO data series for each fiscal year from 2025-26 to 2029-30:
- **Blue bars**: Effect of spending decisions on borrowing (£bn)
- **Purple bars**: Effect of tax decisions on borrowing (£bn)
- The x-axis should show the fiscal years (2025-26 through 2029-30)
- The y-axis should show the change in borrowing (£bn)
- Include appropriate title and axis labels
- Note: Tax decisions will have negative values (reducing borrowing), spending decisions will have positive values (increasing borrowing)

### 2. Spreadsheet: `/home/ubuntu/policy_decisions.xlsx`
Create an Excel spreadsheet containing:
- The fiscal years (2025-26 through 2029-30)
- The effect of spending decisions on borrowing for each year
- The effect of tax decisions on borrowing for each year

## Submission

When both files are ready at the paths above, call `submit_answer` to submit for evaluation."""

    def _get_current_budget_deficit_prompt(self) -> str:
        """Return the prompt for the current budget deficit comparison task."""
        return """# Task: Compare Current Budget Deficit Forecasts

You are a policy analyst. Your task is to extract and compare current budget deficit forecasts (as % of GDP) from two different UK budgets.

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only):
- `/orwd_data/2024/` - Autumn Budget 2024 (October 2024)
- `/orwd_data/2025/` - Budget 2025 (November 2025)

## Data to Extract

For fiscal years 2025-26 through 2029-30, extract the **current budget deficit as % of GDP** from:
1. **October 2024 Budget** (Autumn Budget 2024) - found in `/orwd_data/2024/`
2. **November 2025 Budget** (Budget 2025) - found in `/orwd_data/2025/`

## Required Outputs

### 1. Line Graph: `/home/ubuntu/current_budget_deficit.png`
Create a line graph comparing the two budget forecasts:
- **X-axis**: Fiscal years (2025-26, 2026-27, 2027-28, 2028-29, 2029-30)
- **Y-axis**: Current budget deficit (% of GDP)
- **Green line**: October 2024 budget forecast
- **Yellow line**: November 2025 budget forecast
- Include appropriate title, axis labels, and legend
- Note: Values can be positive (deficit) or negative (surplus)

### 2. Spreadsheet: `/home/ubuntu/current_budget_deficit.xlsx`
Create an Excel spreadsheet containing:
- Column 1: Fiscal years (2025-26 through 2029-30)
- Column 2: Current budget deficit (% GDP) from October 2024 budget
- Column 3: Current budget deficit (% GDP) from November 2025 budget

## Submission

When both files are ready at the paths above, call `submit_answer` to submit for evaluation."""

    def _get_presentation_prompt(self) -> str:
        """Return the prompt for the presentation task."""
        return f"""# Task: Create "Tax Measures" Presentation for {self.task_data['budget_name']}

You are a policy analyst at a fiscal policy research organization. Your task is to create a PowerPoint presentation (.pptx) titled "Tax Measures" that provides a comprehensive, analytical summary of the key tax measures announced in the {self.task_data['budget_name']}.

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only). These contain detailed analysis of the {self.task_data['budget_name']} tax measures.

## Required Output

Save your presentation as: **{self.task_data['output_path']}**

Then call `submit_answer` to submit for evaluation."""

    def _get_ai_measures_prompt(self) -> str:
        """Return the prompt for the AI measures summary task."""
        return f"""# Task: Comprehensive AI Measures Summary 2020-2025

You are tasked with creating a comprehensive summary of all AI-related measures and announcements from UK budgets and autumn statements between 2020 and 2025.

## Your Task

Review budget documents from 2020-2025 and find all AI-related initiatives, creating a comprehensive report.

## Output

Write your comprehensive report to: **{self.task_data['output_path']}**

When complete, call the `submit_answer` tool to submit for evaluation."""

    def _get_tax_proposals_prompt(self) -> str:
        """Return the prompt for the tax proposals task."""
        return """# Task: Propose Tax Changes to Reduce Borrowing by Half

You are a civil servant at HM Treasury. Your task is to propose a package of tax changes that will reduce public sector net borrowing (PSNB) for 2026-27 by half.

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only).

## Output Format

Write your proposals to `/home/ubuntu/tax_proposals.md`

## Submission

When ready, call `submit_answer` to submit your proposals for evaluation."""

    def _get_opposition_response_prompt(self) -> str:
        """Return the prompt for the Leader of the Opposition response task."""
        return f"""# Task: Leader of the Opposition Response to Budget 2025

You are the Leader of the Opposition in the UK Parliament. Your task is to write a formal response to the 2025 Budget, delivered from the dispatch box in the House of Commons.

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only). You have access to:
- Budget 2025 full document
- Previous budget/statement documents for comparison
- OBR economic forecasts

## Submission

Write your opposition response to: **{self.task_data['output_path']}**

Then call `submit_answer` tool to submit for evaluation."""

    async def _grade_spreadsheet(self, xlsx_bytes: bytes) -> dict[str, Any]:
        """Grade the borrowing spreadsheet using gpt-5-mini by extracting data to text."""
        # Read Excel file and extract data
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))

        # Extract all data from all sheets into text format
        spreadsheet_text = ""
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            spreadsheet_text += f"Sheet: {sheet_name}\n"
            spreadsheet_text += "-" * 40 + "\n"

            for row in ws.iter_rows(values_only=True):
                # Convert row to strings, handling None values
                row_strs = [str(cell) if cell is not None else "" for cell in row]
                spreadsheet_text += " | ".join(row_strs) + "\n"

            spreadsheet_text += "\n"

        gt = BORROWING_GROUND_TRUTH
        grader_prompt = f"""You are evaluating a spreadsheet that should contain UK government borrowing forecast data.

The ground truth data is:

Tax Years: {', '.join(gt['years'])}
March 2025 forecast (GBP bn): {', '.join(str(v) for v in gt['march_2025'])}
October 2025 forecast (GBP bn): {', '.join(str(v) for v in gt['october_2025'])}
Difference (Oct minus March): {', '.join(str(v) for v in gt['difference'])}

Here is the spreadsheet content:

{spreadsheet_text}

The spreadsheet may have rows or columns in any order, and headers may vary. Please evaluate the following criteria. For each, answer PASS or FAIL:

1. YEARS_PRESENT: Does the spreadsheet contain all 5 fiscal years (2025-26, 2026-27, 2027-28, 2028-29, 2029-30) in some form?
2. MARCH_VALUES: Does the spreadsheet contain the March 2025 forecast values (117.7, 97.2, 80.2, 77.4, 74.0) or values very close to them (within 0.5)?
3. OCTOBER_VALUES: Does the spreadsheet contain the October 2025 forecast values (138.3, 112.1, 98.5, 86.9, 67.9) or values very close to them (within 0.5)?
4. DIFFERENCE_VALUES: Does the spreadsheet contain the difference values (20.6, 14.9, 18.3, 9.5, -6.2) or values very close to them (within 0.5)?

Format your response as:
YEARS_PRESENT: PASS/FAIL
MARCH_VALUES: PASS/FAIL
OCTOBER_VALUES: PASS/FAIL
DIFFERENCE_VALUES: PASS/FAIL

Then provide a brief explanation."""

        response = await self.grader_client.create(
            model="gpt-5-mini",
            messages=[
                {
                    "role": "user",
                    "content": grader_prompt,
                }
            ],
            # NO temperature parameter (per CLAUDE.md)
        )

        grading_text = response.text or ""

        checks = {}
        for criterion in ["YEARS_PRESENT", "MARCH_VALUES", "OCTOBER_VALUES", "DIFFERENCE_VALUES"]:
            pattern = rf"{criterion}\s*:\s*(PASS|FAIL)"
            match = re.search(pattern, grading_text.upper())
            checks[criterion] = match.group(1) == "PASS" if match else False

        passed_checks = sum(checks.values())
        score = passed_checks / len(checks)

        return {
            "score": score,
            "checks": checks,
            "details": grading_text,
        }

    async def _grade_chart_image(self, png_bytes: bytes) -> dict[str, Any]:
        """Grade the borrowing chart image using gpt-5-mini vision."""
        image_b64 = base64.b64encode(png_bytes).decode("utf-8")

        grader_prompt = """You are evaluating a bar chart showing changes in UK government borrowing forecasts.

The chart should show the DIFFERENCE in borrowing between October 2025 and March 2025 OBR forecasts for each fiscal year from 2025-26 to 2029-30.

Expected differences (in GBP billions):
- 2025-26: +20.6
- 2026-27: +14.9
- 2027-28: +18.3
- 2028-29: +9.5
- 2029-30: -6.2

Please evaluate the following criteria. For each, answer PASS or FAIL:

1. IS_BAR_CHART: Is this a bar chart (not a line chart, pie chart, scatter plot, etc.)?
2. GREEN_BARS: Are the bars green (any shade of green is acceptable)?
3. CORRECT_YEARS: Does the chart show all 5 fiscal years (2025-26 through 2029-30)?
4. NEGATIVE_2029_30: Does the last year (2029-30) show a negative value (bar going below zero / below the x-axis)?
5. APPROXIMATE_VALUES: Do the bar heights approximately match the expected differences listed above? Every year should be positive except 2029-30.

Format your response as:
IS_BAR_CHART: PASS/FAIL
GREEN_BARS: PASS/FAIL
CORRECT_YEARS: PASS/FAIL
NEGATIVE_2029_30: PASS/FAIL
APPROXIMATE_VALUES: PASS/FAIL

Then provide a brief explanation."""

        response = await self.grader_client.create(
            model="gpt-5-mini",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{image_b64}"},
                        },
                        {
                            "type": "text",
                            "text": grader_prompt,
                        },
                    ],
                }
            ],
            # NO temperature parameter (per CLAUDE.md)
        )

        grading_text = response.text or ""

        checks = {}
        for criterion in ["IS_BAR_CHART", "GREEN_BARS", "CORRECT_YEARS",
                          "NEGATIVE_2029_30", "APPROXIMATE_VALUES"]:
            pattern = rf"{criterion}\s*:\s*(PASS|FAIL)"
            match = re.search(pattern, grading_text.upper())
            checks[criterion] = match.group(1) == "PASS" if match else False

        passed_checks = sum(checks.values())
        score = passed_checks / len(checks)

        return {
            "score": score,
            "checks": checks,
            "details": grading_text,
        }

    async def _grade_psnb_2020_spreadsheet(self, xlsx_bytes: bytes) -> dict[str, Any]:
        """Grade the PSNB 2020 spreadsheet using gpt-5-mini by extracting data to text."""
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))

        spreadsheet_text = ""
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            spreadsheet_text += f"Sheet: {sheet_name}\n"
            spreadsheet_text += "-" * 40 + "\n"
            for row in ws.iter_rows(values_only=True):
                row_strs = [str(cell) if cell is not None else "" for cell in row]
                spreadsheet_text += " | ".join(row_strs) + "\n"
            spreadsheet_text += "\n"

        gt = PSNB_2020_GROUND_TRUTH
        grader_prompt = f"""You are evaluating a spreadsheet that should contain UK Budget 2020 PSNB forecast data.

The ground truth data is:

Fiscal Years: {', '.join(gt['years'])}
Restated March 2019 forecast (£bn): {', '.join(str(v) for v in gt['restated_march_2019'])}
Budget 2020 forecast (£bn): {', '.join(str(v) for v in gt['budget_2020'])}
Difference (Budget 2020 minus March 2019): {', '.join(str(v) for v in gt['difference'])}

Here is the spreadsheet content:

{spreadsheet_text}

The spreadsheet may have rows or columns in any order, and headers may vary. Please evaluate the following criteria. For each, answer PASS or FAIL:

1. YEARS_PRESENT: Does the spreadsheet contain all 5 fiscal years (2019-20, 2020-21, 2021-22, 2022-23, 2023-24) in some form?
2. MARCH_2019_VALUES: Does the spreadsheet contain the restated March 2019 forecast values (47.6, 40.2, 37.6, 35.4, 33.3) or values very close to them (within 0.5)?
3. BUDGET_2020_VALUES: Does the spreadsheet contain the Budget 2020 forecast values (47.4, 54.8, 66.7, 61.5, 60.2) or values very close to them (within 0.5)?
4. DIFFERENCE_VALUES: Does the spreadsheet contain the difference values (-0.2, 14.6, 29.1, 26.0, 26.9) or values very close to them (within 0.5)?

Format your response as:
YEARS_PRESENT: PASS/FAIL
MARCH_2019_VALUES: PASS/FAIL
BUDGET_2020_VALUES: PASS/FAIL
DIFFERENCE_VALUES: PASS/FAIL

Then provide a brief explanation."""

        response = await self.grader_client.create(
            model="gpt-5-mini",
            messages=[{"role": "user", "content": grader_prompt}],
        )

        grading_text = response.text or ""

        checks = {}
        for criterion in ["YEARS_PRESENT", "MARCH_2019_VALUES", "BUDGET_2020_VALUES", "DIFFERENCE_VALUES"]:
            pattern = rf"{criterion}\s*:\s*(PASS|FAIL)"
            match = re.search(pattern, grading_text.upper())
            checks[criterion] = match.group(1) == "PASS" if match else False

        passed_checks = sum(checks.values())
        score = passed_checks / len(checks)

        return {"score": score, "checks": checks, "details": grading_text}

    async def _grade_psnb_2020_chart(self, png_bytes: bytes) -> dict[str, Any]:
        """Grade the PSNB 2020 chart image using gpt-5-mini vision."""
        image_b64 = base64.b64encode(png_bytes).decode("utf-8")

        grader_prompt = """You are evaluating a bar chart showing changes in UK PSNB forecasts between Budget 2020 and restated March 2019.

The chart should show the DIFFERENCE in PSNB forecasts (Budget 2020 minus restated March 2019) for fiscal years 2019-20 through 2023-24.

Expected differences (in £ billions):
- 2019-20: -0.2
- 2020-21: +14.6
- 2021-22: +29.1
- 2022-23: +26.0
- 2023-24: +26.9

Please evaluate the following criteria. For each, answer PASS or FAIL:

1. IS_BAR_CHART: Is this a bar chart (not a line chart, pie chart, scatter plot, etc.)?
2. CORRECT_YEARS: Does the chart show all 5 fiscal years (2019-20 through 2023-24)?
3. NEGATIVE_2019_20: Does the first year (2019-20) show a small negative value (bar slightly below zero)?
4. POSITIVE_REMAINING: Are all remaining years (2020-21 through 2023-24) showing positive values?
5. APPROXIMATE_VALUES: Do the bar heights approximately match the expected differences listed above?

Format your response as:
IS_BAR_CHART: PASS/FAIL
CORRECT_YEARS: PASS/FAIL
NEGATIVE_2019_20: PASS/FAIL
POSITIVE_REMAINING: PASS/FAIL
APPROXIMATE_VALUES: PASS/FAIL

Then provide a brief explanation."""

        response = await self.grader_client.create(
            model="gpt-5-mini",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}},
                        {"type": "text", "text": grader_prompt},
                    ],
                }
            ],
        )

        grading_text = response.text or ""

        checks = {}
        for criterion in ["IS_BAR_CHART", "CORRECT_YEARS", "NEGATIVE_2019_20", "POSITIVE_REMAINING", "APPROXIMATE_VALUES"]:
            pattern = rf"{criterion}\s*:\s*(PASS|FAIL)"
            match = re.search(pattern, grading_text.upper())
            checks[criterion] = match.group(1) == "PASS" if match else False

        passed_checks = sum(checks.values())
        score = passed_checks / len(checks)

        return {"score": score, "checks": checks, "details": grading_text}

    async def _grade_psnb_2020_task(self) -> dict[str, Any]:
        """Grade the PSNB 2020 task by validating both files."""
        spreadsheet_result: dict[str, Any] | None = None
        chart_result: dict[str, Any] | None = None
        errors: list[str] = []

        xlsx_bytes: bytes | None = None
        png_bytes: bytes | None = None

        try:
            xlsx_bytes = await self.sandbox.download("/home/ubuntu/psnb_changes.xlsx")
        except Exception as e:
            errors.append(f"Spreadsheet download failed: {str(e)}")

        try:
            png_bytes = await self.sandbox.download("/home/ubuntu/psnb_changes.png")
        except Exception:
            pass  # Chart is optional

        tasks = []
        if xlsx_bytes:
            tasks.append(self._grade_psnb_2020_spreadsheet(xlsx_bytes))
        if png_bytes:
            tasks.append(self._grade_psnb_2020_chart(png_bytes))

        if tasks:
            results = await asyncio.gather(*tasks)
            idx = 0
            if xlsx_bytes:
                spreadsheet_result = results[idx]
                idx += 1
            if png_bytes:
                chart_result = results[idx]

        spreadsheet_score = spreadsheet_result["score"] if spreadsheet_result else 0.0
        chart_score = chart_result["score"] if chart_result else 0.0

        # Spreadsheet is required (70%), chart is optional bonus (30%)
        if chart_result:
            combined_reward = 0.7 * spreadsheet_score + 0.3 * chart_score
        else:
            combined_reward = spreadsheet_score

        display_lines = [
            "PSNB 2020 Forecast Task Evaluation",
            "=" * 60,
            "",
            f"SPREADSHEET VALIDATION ({spreadsheet_score:.0%}):",
        ]

        if spreadsheet_result and spreadsheet_result.get("checks"):
            for check_name, passed in spreadsheet_result["checks"].items():
                status = "PASS" if passed else "FAIL"
                display_lines.append(f"  {status}: {check_name}")
        elif not xlsx_bytes:
            display_lines.append("  File not found: /home/ubuntu/psnb_changes.xlsx")

        if chart_result:
            display_lines.append("")
            display_lines.append(f"CHART IMAGE VALIDATION (OPTIONAL) ({chart_score:.0%}):")
            for check_name, passed in chart_result["checks"].items():
                status = "PASS" if passed else "FAIL"
                display_lines.append(f"  {status}: {check_name}")
        elif png_bytes is None:
            display_lines.append("")
            display_lines.append("CHART IMAGE: Not provided (optional)")

        display_lines.extend([
            "",
            "=" * 60,
            f"Combined Reward: {combined_reward:.2f}",
        ])

        if errors:
            display_lines.append(f"Errors: {'; '.join(errors)}")

        return {
            "display_text": "\n".join(display_lines),
            "metadata": {
                "task_id": "budget_2020_psnb_forecast",
                "spreadsheet_result": spreadsheet_result,
                "chart_result": chart_result,
                "combined_reward": combined_reward,
                "errors": errors,
            },
            "reward": combined_reward,
        }

    async def _grade_policy_decisions_spreadsheet(self, xlsx_bytes: bytes) -> dict[str, Any]:
        """Grade the policy decisions spreadsheet using gpt-5-mini by extracting data to text."""
        # Read Excel file and extract data
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))

        # Extract all data from all sheets into text format
        spreadsheet_text = ""
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            spreadsheet_text += f"Sheet: {sheet_name}\n"
            spreadsheet_text += "-" * 40 + "\n"

            for row in ws.iter_rows(values_only=True):
                # Convert row to strings, handling None values
                row_strs = [str(cell) if cell is not None else "" for cell in row]
                spreadsheet_text += " | ".join(row_strs) + "\n"

            spreadsheet_text += "\n"

        gt = POLICY_DECISIONS_GROUND_TRUTH
        grader_prompt = f"""You are evaluating a spreadsheet that should contain UK Budget 2025 policy decisions impact data.

The ground truth data is:

Fiscal Years: {', '.join(gt['years'])}
Spending decisions effect on borrowing (GBP bn): {', '.join(str(v) for v in gt['spending_decisions'])}
Tax decisions effect on borrowing (GBP bn): {', '.join(str(v) for v in gt['tax_decisions'])}

Here is the spreadsheet content:

{spreadsheet_text}

The spreadsheet may have rows or columns in any order, and headers may vary. Please evaluate the following criteria. For each, answer PASS or FAIL:

1. YEARS_PRESENT: Does the spreadsheet contain all 5 fiscal years (2025-26, 2026-27, 2027-28, 2028-29, 2029-30) in some form?
2. SPENDING_VALUES: Does the spreadsheet contain the spending decisions values (4.9, 6.6, 16.0, 12.9, 11.3) or values very close to them (within 0.5)?
3. TAX_VALUES: Does the spreadsheet contain the tax decisions values (-1.3, -0.7, -6.1, -13.9, -26.1) or values very close to them (within 0.5)?

Format your response as:
YEARS_PRESENT: PASS/FAIL
SPENDING_VALUES: PASS/FAIL
TAX_VALUES: PASS/FAIL

Then provide a brief explanation."""

        response = await self.grader_client.create(
            model="gpt-5-mini",
            messages=[{"role": "user", "content": grader_prompt}],
            # NO temperature parameter (per CLAUDE.md)
        )

        grading_text = response.text or ""

        checks = {}
        for criterion in ["YEARS_PRESENT", "SPENDING_VALUES", "TAX_VALUES"]:
            pattern = rf"{criterion}\s*:\s*(PASS|FAIL)"
            match = re.search(pattern, grading_text.upper())
            checks[criterion] = match.group(1) == "PASS" if match else False

        passed_checks = sum(checks.values())
        score = passed_checks / len(checks)

        return {
            "score": score,
            "checks": checks,
            "details": grading_text,
        }

    async def _grade_policy_decisions_chart(self, png_bytes: bytes) -> dict[str, Any]:
        """Grade the policy decisions chart image using gpt-5-mini vision."""
        image_b64 = base64.b64encode(png_bytes).decode("utf-8")

        grader_prompt = """You are evaluating a bar chart showing the effect of Budget 2025 policy decisions on borrowing.

The chart should show TWO data series for each fiscal year from 2025-26 to 2029-30:
- Blue bars: Effect of spending decisions on borrowing
- Purple bars: Effect of tax decisions on borrowing

The chart can be either grouped bars (side-by-side) OR stacked bars - both formats are acceptable.

Expected data (in GBP billions):

Fiscal Year | Spending (blue) | Tax (purple)
2025-26     | +4.9            | -1.3
2026-27     | +6.6            | -0.7
2027-28     | +16.0           | -6.1
2028-29     | +12.9           | -13.9
2029-30     | +11.3           | -26.1

Note: Spending values are POSITIVE (increasing borrowing), tax values are NEGATIVE (reducing borrowing).

Please evaluate the following criteria. For each, answer PASS or FAIL:

1. IS_BAR_CHART: Is this a bar chart (grouped or stacked) with two data series per year? (Not a line chart, pie chart, etc.)
2. CORRECT_COLORS: Are spending decisions shown in blue and tax decisions shown in purple (or very close shades)?
3. CORRECT_YEARS: Does the chart show all 5 fiscal years (2025-26 through 2029-30)?
4. TAX_NEGATIVE: Are all tax decision values negative (showing below zero or as negative contribution if stacked)?
5. APPROXIMATE_VALUES: Do the values approximately match the expected data listed above?

Format your response as:
IS_BAR_CHART: PASS/FAIL
CORRECT_COLORS: PASS/FAIL
CORRECT_YEARS: PASS/FAIL
TAX_NEGATIVE: PASS/FAIL
APPROXIMATE_VALUES: PASS/FAIL

Then provide a brief explanation."""

        response = await self.grader_client.create(
            model="gpt-5-mini",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{image_b64}"},
                        },
                        {
                            "type": "text",
                            "text": grader_prompt,
                        },
                    ],
                }
            ],
            # NO temperature parameter (per CLAUDE.md)
        )

        grading_text = response.text or ""

        checks = {}
        for criterion in ["IS_BAR_CHART", "CORRECT_COLORS", "CORRECT_YEARS",
                          "TAX_NEGATIVE", "APPROXIMATE_VALUES"]:
            pattern = rf"{criterion}\s*:\s*(PASS|FAIL)"
            match = re.search(pattern, grading_text.upper())
            checks[criterion] = match.group(1) == "PASS" if match else False

        passed_checks = sum(checks.values())
        score = passed_checks / len(checks)

        return {
            "score": score,
            "checks": checks,
            "details": grading_text,
        }

    async def _grade_policy_decisions_task(self) -> dict[str, Any]:
        """Grade the policy decisions task by validating both files."""
        spreadsheet_result: dict[str, Any] | None = None
        chart_result: dict[str, Any] | None = None
        errors: list[str] = []

        # Download both files from sandbox
        xlsx_bytes: bytes | None = None
        png_bytes: bytes | None = None

        try:
            xlsx_bytes = await self.sandbox.download("/home/ubuntu/policy_decisions.xlsx")
        except Exception as e:
            errors.append(f"Spreadsheet download failed: {str(e)}")

        try:
            png_bytes = await self.sandbox.download("/home/ubuntu/policy_decisions.png")
        except Exception as e:
            errors.append(f"Chart image download failed: {str(e)}")

        # Grade both files concurrently
        tasks = []
        if xlsx_bytes:
            tasks.append(self._grade_policy_decisions_spreadsheet(xlsx_bytes))
        if png_bytes:
            tasks.append(self._grade_policy_decisions_chart(png_bytes))

        if tasks:
            results = await asyncio.gather(*tasks)
            idx = 0
            if xlsx_bytes:
                spreadsheet_result = results[idx]
                idx += 1
            if png_bytes:
                chart_result = results[idx]

        # Default scores for missing files
        spreadsheet_score = spreadsheet_result["score"] if spreadsheet_result else 0.0
        chart_score = chart_result["score"] if chart_result else 0.0
        combined_reward = 0.5 * spreadsheet_score + 0.5 * chart_score

        # Format display text
        display_lines = [
            "Policy Decisions Task Evaluation",
            "=" * 60,
            "",
            f"SPREADSHEET VALIDATION ({spreadsheet_score:.0%}):",
        ]

        if spreadsheet_result and spreadsheet_result.get("checks"):
            for check_name, passed in spreadsheet_result["checks"].items():
                status = "PASS" if passed else "FAIL"
                display_lines.append(f"  {status}: {check_name}")
        elif not xlsx_bytes:
            display_lines.append("  File not found: /home/ubuntu/policy_decisions.xlsx")

        display_lines.append("")
        display_lines.append(f"CHART IMAGE VALIDATION ({chart_score:.0%}):")

        if chart_result and chart_result.get("checks"):
            for check_name, passed in chart_result["checks"].items():
                status = "PASS" if passed else "FAIL"
                display_lines.append(f"  {status}: {check_name}")
        elif not png_bytes:
            display_lines.append("  File not found: /home/ubuntu/policy_decisions.png")

        display_lines.append("")
        display_lines.append("=" * 60)
        display_lines.append(f"Combined Reward: {combined_reward:.2f}")

        if errors:
            display_lines.append(f"Errors: {'; '.join(errors)}")

        return {
            "display_text": "\n".join(display_lines),
            "metadata": {
                "task_id": "budget_2025_policy_decisions",
                "spreadsheet_result": spreadsheet_result,
                "chart_result": chart_result,
                "combined_reward": combined_reward,
                "errors": errors,
            },
            "reward": combined_reward,
        }

    async def _grade_current_budget_deficit_spreadsheet(self, xlsx_bytes: bytes) -> dict[str, Any]:
        """Grade the current budget deficit spreadsheet using gpt-5-mini by extracting data to text."""
        # Read Excel file and extract data
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))

        # Extract all data from all sheets into text format
        spreadsheet_text = ""
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            spreadsheet_text += f"Sheet: {sheet_name}\n"
            spreadsheet_text += "-" * 40 + "\n"

            for row in ws.iter_rows(values_only=True):
                # Convert row to strings, handling None values
                row_strs = [str(cell) if cell is not None else "" for cell in row]
                spreadsheet_text += " | ".join(row_strs) + "\n"

            spreadsheet_text += "\n"

        gt = CURRENT_BUDGET_DEFICIT_GROUND_TRUTH
        grader_prompt = f"""You are evaluating a spreadsheet that should contain UK current budget deficit forecasts (as % of GDP).

The ground truth data is:

Fiscal Years: {', '.join(gt['years'])}
October 2024 budget (% GDP): {', '.join(str(v) for v in gt['budget_2024'])}
November 2025 budget (% GDP): {', '.join(str(v) for v in gt['budget_2025'])}

Here is the spreadsheet content:

{spreadsheet_text}

The spreadsheet may have rows or columns in any order, and headers may vary. Please evaluate the following criteria. For each, answer PASS or FAIL:

1. YEARS_PRESENT: Does the spreadsheet contain all 5 fiscal years (2025-26, 2026-27, 2027-28, 2028-29, 2029-30) in some form?
2. BUDGET_2024_VALUES: Does the spreadsheet contain the October 2024 budget values (0.9, 0.2, -0.3, -0.3, -0.3) or values very close to them (within 0.1)?
3. BUDGET_2025_VALUES: Does the spreadsheet contain the November 2025 budget values (1.7, 0.9, 0.1, -0.1, -0.6) or values very close to them (within 0.1)?

Format your response as:
YEARS_PRESENT: PASS/FAIL
BUDGET_2024_VALUES: PASS/FAIL
BUDGET_2025_VALUES: PASS/FAIL

Then provide a brief explanation."""

        response = await self.grader_client.create(
            model="gpt-5-mini",
            messages=[{"role": "user", "content": grader_prompt}],
            # NO temperature parameter (per CLAUDE.md)
        )

        grading_text = response.text or ""

        checks = {}
        for criterion in ["YEARS_PRESENT", "BUDGET_2024_VALUES", "BUDGET_2025_VALUES"]:
            pattern = rf"{criterion}\s*:\s*(PASS|FAIL)"
            match = re.search(pattern, grading_text.upper())
            checks[criterion] = match.group(1) == "PASS" if match else False

        passed_checks = sum(checks.values())
        score = passed_checks / len(checks)

        return {
            "score": score,
            "checks": checks,
            "details": grading_text,
        }

    async def _grade_current_budget_deficit_chart(self, png_bytes: bytes) -> dict[str, Any]:
        """Grade the current budget deficit chart image using gpt-5-mini vision."""
        image_b64 = base64.b64encode(png_bytes).decode("utf-8")

        grader_prompt = """You are evaluating a line graph showing current budget deficit forecasts (as % of GDP).

The chart should show TWO lines for fiscal years 2025-26 through 2029-30:
- Green line: October 2024 budget forecast
- Yellow line: November 2025 budget forecast

Expected data (% of GDP):

Fiscal Year | Oct 2024 (green) | Nov 2025 (yellow)
2025-26     | 0.9              | 1.7
2026-27     | 0.2              | 0.9
2027-28     | -0.3             | 0.1
2028-29     | -0.3             | -0.1
2029-30     | -0.3             | -0.6

Note: Negative values indicate a surplus, positive values indicate a deficit.

Please evaluate the following criteria. For each, answer PASS or FAIL:

1. IS_LINE_GRAPH: Is this a line graph (not a bar chart, scatter plot, etc.) showing two distinct lines?
2. CORRECT_COLORS: Is one line green (for Oct 2024) and one line yellow (for Nov 2025)?
3. CORRECT_YEARS: Does the chart show all 5 fiscal years (2025-26 through 2029-30) on the x-axis?
4. APPROXIMATE_VALUES: Do the line values approximately match the expected data listed above? The green line should start around 0.9 and trend to -0.3, while the yellow line should start around 1.7 and trend to -0.6.
5. AXES_LABELED: Are the axes properly labeled (fiscal years on x-axis, % GDP or similar on y-axis)?

Format your response as:
IS_LINE_GRAPH: PASS/FAIL
CORRECT_COLORS: PASS/FAIL
CORRECT_YEARS: PASS/FAIL
APPROXIMATE_VALUES: PASS/FAIL
AXES_LABELED: PASS/FAIL

Then provide a brief explanation."""

        response = await self.grader_client.create(
            model="gpt-5-mini",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{image_b64}"},
                        },
                        {
                            "type": "text",
                            "text": grader_prompt,
                        },
                    ],
                }
            ],
            # NO temperature parameter (per CLAUDE.md)
        )

        grading_text = response.text or ""

        checks = {}
        for criterion in ["IS_LINE_GRAPH", "CORRECT_COLORS", "CORRECT_YEARS",
                          "APPROXIMATE_VALUES", "AXES_LABELED"]:
            pattern = rf"{criterion}\s*:\s*(PASS|FAIL)"
            match = re.search(pattern, grading_text.upper())
            checks[criterion] = match.group(1) == "PASS" if match else False

        passed_checks = sum(checks.values())
        score = passed_checks / len(checks)

        return {
            "score": score,
            "checks": checks,
            "details": grading_text,
        }

    async def _grade_current_budget_deficit_task(self) -> dict[str, Any]:
        """Grade the current budget deficit task by validating both files."""
        spreadsheet_result: dict[str, Any] | None = None
        chart_result: dict[str, Any] | None = None
        errors: list[str] = []

        # Download both files from sandbox
        xlsx_bytes: bytes | None = None
        png_bytes: bytes | None = None

        try:
            xlsx_bytes = await self.sandbox.download("/home/ubuntu/current_budget_deficit.xlsx")
        except Exception as e:
            errors.append(f"Spreadsheet download failed: {str(e)}")

        try:
            png_bytes = await self.sandbox.download("/home/ubuntu/current_budget_deficit.png")
        except Exception as e:
            errors.append(f"Chart image download failed: {str(e)}")

        # Grade both files concurrently
        tasks = []
        if xlsx_bytes:
            tasks.append(self._grade_current_budget_deficit_spreadsheet(xlsx_bytes))
        if png_bytes:
            tasks.append(self._grade_current_budget_deficit_chart(png_bytes))

        if tasks:
            results = await asyncio.gather(*tasks)
            idx = 0
            if xlsx_bytes:
                spreadsheet_result = results[idx]
                idx += 1
            if png_bytes:
                chart_result = results[idx]

        # Default scores for missing files
        spreadsheet_score = spreadsheet_result["score"] if spreadsheet_result else 0.0
        chart_score = chart_result["score"] if chart_result else 0.0
        combined_reward = 0.5 * spreadsheet_score + 0.5 * chart_score

        # Format display text
        display_lines = [
            "Current Budget Deficit Task Evaluation",
            "=" * 60,
            "",
            f"SPREADSHEET VALIDATION ({spreadsheet_score:.0%}):",
        ]

        if spreadsheet_result and spreadsheet_result.get("checks"):
            for check_name, passed in spreadsheet_result["checks"].items():
                status = "PASS" if passed else "FAIL"
                display_lines.append(f"  {status}: {check_name}")
        elif not xlsx_bytes:
            display_lines.append("  File not found: /home/ubuntu/current_budget_deficit.xlsx")

        display_lines.append("")
        display_lines.append(f"CHART IMAGE VALIDATION ({chart_score:.0%}):")

        if chart_result and chart_result.get("checks"):
            for check_name, passed in chart_result["checks"].items():
                status = "PASS" if passed else "FAIL"
                display_lines.append(f"  {status}: {check_name}")
        elif not png_bytes:
            display_lines.append("  File not found: /home/ubuntu/current_budget_deficit.png")

        display_lines.append("")
        display_lines.append("=" * 60)
        display_lines.append(f"Combined Reward: {combined_reward:.2f}")

        if errors:
            display_lines.append(f"Errors: {'; '.join(errors)}")

        return {
            "display_text": "\n".join(display_lines),
            "metadata": {
                "task_id": "budget_deficit_comparison",
                "spreadsheet_result": spreadsheet_result,
                "chart_result": chart_result,
                "combined_reward": combined_reward,
                "errors": errors,
            },
            "reward": combined_reward,
        }

    async def _grade_borrowing_task(self) -> dict[str, Any]:
        """Grade the borrowing chart task by sending both files to gpt-5-mini."""
        spreadsheet_result: dict[str, Any] | None = None
        chart_result: dict[str, Any] | None = None
        errors: list[str] = []

        # Download both files from sandbox
        xlsx_bytes: bytes | None = None
        png_bytes: bytes | None = None

        try:
            xlsx_bytes = await self.sandbox.download("/home/ubuntu/changes_borrowing.xlsx")
        except Exception as e:
            errors.append(f"Spreadsheet download failed: {str(e)}")

        try:
            png_bytes = await self.sandbox.download("/home/ubuntu/changes_borrowing.png")
        except Exception as e:
            errors.append(f"Chart image download failed: {str(e)}")

        # Grade both files concurrently
        tasks = []
        if xlsx_bytes:
            tasks.append(self._grade_spreadsheet(xlsx_bytes))
        if png_bytes:
            tasks.append(self._grade_chart_image(png_bytes))

        if tasks:
            results = await asyncio.gather(*tasks)
            idx = 0
            if xlsx_bytes:
                spreadsheet_result = results[idx]
                idx += 1
            if png_bytes:
                chart_result = results[idx]

        # Default scores for missing files
        spreadsheet_score = spreadsheet_result["score"] if spreadsheet_result else 0.0
        chart_score = chart_result["score"] if chart_result else 0.0
        combined_reward = 0.5 * spreadsheet_score + 0.5 * chart_score

        # Format display text
        display_lines = [
            "Borrowing Forecast Task Evaluation",
            "=" * 60,
            "",
            f"SPREADSHEET VALIDATION ({spreadsheet_score:.0%}):",
        ]

        if spreadsheet_result and spreadsheet_result.get("checks"):
            for check_name, passed in spreadsheet_result["checks"].items():
                status = "PASS" if passed else "FAIL"
                display_lines.append(f"  {status}: {check_name}")
        elif not xlsx_bytes:
            display_lines.append("  File not found: /home/ubuntu/changes_borrowing.xlsx")

        display_lines.append("")
        display_lines.append(f"CHART IMAGE VALIDATION ({chart_score:.0%}):")

        if chart_result and chart_result.get("checks"):
            for check_name, passed in chart_result["checks"].items():
                status = "PASS" if passed else "FAIL"
                display_lines.append(f"  {status}: {check_name}")
        elif not png_bytes:
            display_lines.append("  File not found: /home/ubuntu/changes_borrowing.png")

        display_lines.append("")
        display_lines.append("=" * 60)
        display_lines.append(f"Combined Reward: {combined_reward:.2f}")

        if errors:
            display_lines.append(f"Errors: {'; '.join(errors)}")

        return {
            "display_text": "\n".join(display_lines),
            "metadata": {
                "task_id": "budget_2025_borrowing_chart",
                "spreadsheet_result": spreadsheet_result,
                "chart_result": chart_result,
                "combined_reward": combined_reward,
                "errors": errors,
            },
            "reward": combined_reward,
        }

    def _extract_pptx_text(self, pptx_bytes: bytes) -> str:
        """Extract all text from all slides of a PowerPoint presentation."""
        presentation = pptx.Presentation(io.BytesIO(pptx_bytes))

        all_text_parts = []
        for slide_num, slide in enumerate(presentation.slides, 1):
            slide_text_parts = [f"=== Slide {slide_num} ==="]
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        text = paragraph.text.strip()
                        if text:
                            slide_text_parts.append(text)
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            text = cell.text.strip()
                            if text:
                                slide_text_parts.append(text)
            all_text_parts.append("\n".join(slide_text_parts))

        return "\n\n".join(all_text_parts)

    async def _grade_presentation_task(self) -> dict[str, Any]:
        """Grade the presentation task by extracting text from .pptx and evaluating against rubrics."""
        # Download .pptx from sandbox
        try:
            pptx_bytes = await self.sandbox.download(self.task_data["output_path"])
        except Exception as e:
            return {
                "display_text": f"Failed to download presentation at {self.task_data['output_path']}: {str(e)}\n"
                               f"Please ensure you've saved the .pptx file to this exact path.",
                "metadata": {
                    "task_id": self.task_data["task_id"],
                    "error": "file_not_found",
                    "details": str(e),
                },
                "reward": 0.0,
            }

        # Extract text from presentation
        try:
            presentation_text = self._extract_pptx_text(pptx_bytes)
        except Exception as e:
            return {
                "display_text": f"Failed to parse .pptx file: {str(e)}\n"
                               f"The file may be corrupted or not a valid PowerPoint file.",
                "metadata": {
                    "task_id": self.task_data["task_id"],
                    "error": "parse_error",
                    "details": str(e),
                },
                "reward": 0.0,
            }

        # Grade extracted text against rubrics using the existing rubric grading method
        grading_results = await self._grade_with_rubric(presentation_text)

        return grading_results

    async def _grade_tax_proposals_task(self) -> dict[str, Any]:
        """Grade the tax proposals task using o3-mini reasoning model."""
        # Download the tax proposals file
        try:
            proposals_bytes = await self.sandbox.download(self.task_data["output_path"])
            proposals_text = proposals_bytes.decode("utf-8")
        except Exception as e:
            return {
                "display_text": f"Failed to download tax proposals at {self.task_data['output_path']}: {str(e)}\n"
                               f"Please ensure you've written your proposals to this exact path.",
                "metadata": {
                    "task_id": self.task_data["task_id"],
                    "error": "file_not_found",
                    "details": str(e),
                },
                "reward": 0.0,
            }

        # Use o3-mini reasoning model to parse and calculate
        target_revenue = 56050  # £56,050m (half of £112.1bn)

        grader_prompt = f"""You are evaluating tax proposals designed to reduce UK public sector net borrowing for 2026-27 by half.

The target is to raise £56,050 million (£56.05 billion) in additional revenue.

Here are the tax proposals submitted:

{proposals_text}

## Tax Raising Guidelines for 2026-27 (all figures in £m)

### Income Tax Rates
- Change starting rate for savings income by 1p: £0m
- Change basic rate by 1p: £6,900m
- Change higher rate by 1p: £1,600m
- Increase additional rate by 1p (yield): £145m
- Decrease additional rate by 1p (cost): £175m

### Income Tax Allowances and Reliefs
- Change personal allowance by £100: £810m
- Change personal allowance by 1%: £1,000m
- Change personal allowance by 10%: £10,000m
- Change Savings allowance by £100 for BR and £50 for HR taxpayers: £0m
- Change dividend allowance by £100: £0m

### Income Tax Limits
- Change starting rate limit for savings income by £100: Neg (≈£0m)
- Change basic rate limit by 1%: £495m
- Increase basic rate limit by 10% (cost): £4,600m
- Decrease basic rate limit by 10% (yield): £5,400m

### Income Tax Allowances + Starting + Basic Rate Limits
- Change all main allowances, starting and basic rate limits by 1%: £1,450m
- Increase all main allowances, starting and basic rate limits by 10% (cost): £14,400m
- Decrease all main allowances, starting and basic rate limits by 10% (yield): £14,300m

### National Insurance Contributions Rates
- Change Class 1 employee main rate by 1 percentage point: £5,350m
- Change Class 1 employee additional rate by 1 percentage point: £2,000m
- Change Class 1 employer rate by 1 percentage point: £11,150m
- Change Class 4 main rate by 1 percentage point: £440m
- Change Class 4 additional rate by 1 percentage point: £295m

### National Insurance Contribution Limits
- Change employee entry threshold by £2 per week: £210m
- Change employer threshold by £2 per week: £420m
- Change lower profits limit by £104 per year: £15m
- Change upper profits limit by £520 per year: £10m
- Change upper earnings limit by £10 per week: £220m

### Child Benefit
- Increase first child rate by £1 per week (cost): £335m
- Decrease first child rate by £1 per week (yield): £335m
- Increase subsequent child rate by £1 per week (cost): £230m
- Decrease subsequent child rate by £1 per week (yield): £230m

### Corporation Tax
- Change main rate by 1 percentage point: £3,600m

### Capital Gains Tax
- Increase Business Asset Disposal Relief rate by 1 percentage point: £10m
- Increase Business Asset Disposal Relief rate by 5 percentage points: £40m
- Increase lower Capital Gains Tax rate by 1 percentage point: -£5m
- Increase lower Capital Gains Tax rate by 5 percentage points: -£40m
- Increase lower Capital Gains Tax rate by 10 percentage points: -£130m
- Increase higher Capital Gains Tax rate by 1 percentage point: -£15m
- Increase higher Capital Gains Tax rate by 5 percentage points: -£170m
- Increase higher Capital Gains Tax rate by 10 percentage points: -£540m
- Increase Annual Exempt Amount by £500 for individuals and £250 for trusts: £0m

### Inheritance Tax
- Increase standard rate for estates left on death by 1 percentage point: £105m
- Increase Nil Rate Band by £5,000 (cost): £60m
- Increase Residence Nil Rate Band by £5,000 (cost): £25m

### 1% Change in Various Duties
- Beer and cider and other fermented product duties: £40m
- Spirits duties: £40m
- Tobacco Duties: £5m
- Petrol: £100m
- Diesel: £140m
- Rebated oil: Neg (≈£0m)
- Climate change levy: £15m
- Carbon price support: £5m
- Aggregates levy: £5m
- Landfill tax: Neg (≈£0m)

### Vehicle Excise Duty
- Increase rates by £1 for motorbikes and £5 for all other vehicles: £205m

### Air Passenger Duty
- Increase reduced rate by £1: £115m

### VAT
- Change reduced rate by 1 percentage point: £490m
- Change standard rate by 1 percentage point: £8,800m

### Insurance Premium Tax
- Change standard rate by 1 percentage point: £630m
- Change higher rate by 1 percentage point: £20m

### Stamp Duty Land Tax - Residential
- Cut residential 2% marginal rate by 1 percentage point (Cost): £450m
- Raise residential 2% marginal rate by 1 percentage point (Yield): £420m
- Cut residential 5% marginal rate by 1 percentage point (Cost): £775m
- Raise residential 5% marginal rate by 1 percentage point (Yield): £785m
- Cut residential 10% marginal rate by 1 percentage point (Cost): £40m
- Raise residential 10% marginal rate by 1 percentage point (Yield): £35m
- Cut residential 12% marginal rate by 1 percentage point (Cost): -£10m
- Raise residential 12% marginal rate by 1 percentage point (Yield): -£25m
- Decrease Higher Rates of Duty on Additional Dwellings by 1 percentage point (Cost): £90m
- Increase Higher Rates of Duty on Additional Dwellings by 1 percentage point (Yield): -£45m
- Decrease NRSDLT by 1 percentage point (Cost): £10m
- Increase NRSDLT by 1 percentage point (Yield): -£10m

### Stamp Duty Land Tax - Non-Residential
- Decrease non-residential 5% marginal rate by 1 percentage point (Cost): £305m
- Increase non-residential 5% marginal rate by 1 percentage point (Yield): £150m

Your task:
1. Parse each tax proposal and identify the specific change
2. Calculate the revenue impact of each proposal using the guidelines above
3. For multi-unit changes (e.g., "increase by 2p"), multiply the base amount by the number of units
4. Sum up the total revenue raised
5. Return your analysis in the following JSON format:

{{
  "proposals": [
    {{
      "tax": "Tax name",
      "change": "Description of change",
      "revenue_impact_m": 1234
    }},
    ...
  ],
  "total_revenue_m": 12345,
  "target_revenue_m": 56050,
  "difference_m": -43705,
  "reasoning": "Brief explanation of your calculations"
}}

Important:
- Revenue impacts should be in millions (£m)
- Positive numbers = revenue raised, negative numbers = revenue lost
- Be precise in your calculations based on the guidelines above
- If a proposal is unclear or doesn't match the guidelines, note it in your reasoning
- Watch for "cost" vs "yield" - costs are negative revenue, yields are positive

Provide ONLY the JSON output, no other text."""

        try:
            response = await self.grader_client.create(
                model="gpt-5-mini",
                reasoning_effort="medium",
                messages=[{"role": "user", "content": grader_prompt}],
            )

            grading_text = response.text or ""

            # Parse JSON response
            import json
            # Extract JSON from response (handle potential markdown formatting)
            json_text = grading_text
            if "```json" in grading_text:
                json_text = grading_text.split("```json")[1].split("```")[0].strip()
            elif "```" in grading_text:
                json_text = grading_text.split("```")[1].split("```")[0].strip()

            result = json.loads(json_text)

            calculated_revenue = float(result.get("total_revenue_m", 0))
            target = float(target_revenue)

            # Calculate error and reward using squared difference
            error = abs(calculated_revenue - target)
            squared_error = error ** 2

            # Normalization factor: allow 50% error before reward approaches 0
            normalization = (target * 0.5) ** 2

            # Reward based on squared difference
            reward = max(0.0, 1.0 - (squared_error / normalization))

            # Format display
            display_lines = [
                "Tax Proposals Evaluation",
                "=" * 60,
                "",
                f"Target Revenue: £{target:,.0f}m (£{target/1000:.2f}bn)",
                f"Calculated Revenue: £{calculated_revenue:,.0f}m (£{calculated_revenue/1000:.2f}bn)",
                f"Difference: £{calculated_revenue - target:,.0f}m",
                f"Absolute Error: £{error:,.0f}m ({error/target*100:.1f}%)",
                "",
                "Proposals Analyzed:",
            ]

            for i, prop in enumerate(result.get("proposals", []), 1):
                display_lines.append(f"  {i}. {prop.get('tax', 'Unknown')}: {prop.get('change', 'N/A')}")
                display_lines.append(f"     Revenue impact: £{prop.get('revenue_impact_m', 0):,.0f}m")

            display_lines.extend([
                "",
                "Reasoning:",
                result.get("reasoning", "No reasoning provided"),
                "",
                "=" * 60,
                f"Reward: {reward:.3f}",
            ])

            return {
                "display_text": "\n".join(display_lines),
                "metadata": {
                    "task_id": self.task_data["task_id"],
                    "target_revenue_m": target,
                    "calculated_revenue_m": calculated_revenue,
                    "error_m": error,
                    "percentage_error": error / target * 100,
                    "proposals": result.get("proposals", []),
                    "reasoning": result.get("reasoning", ""),
                },
                "reward": reward,
            }

        except Exception as e:
            return {
                "display_text": f"Failed to evaluate tax proposals: {str(e)}\n"
                               f"The reasoning model may have encountered an error parsing your proposals.",
                "metadata": {
                    "task_id": self.task_data["task_id"],
                    "error": "evaluation_error",
                    "details": str(e),
                },
                "reward": 0.0,
            }

    async def _evaluate_criterion(
        self, report: str, criterion: str, criterion_id: str
    ) -> dict[str, Any]:
        """
        Use gpt-5-mini to evaluate a single criterion.
        Per CLAUDE.md: Use gpt-5-mini, no temperature parameter.
        """
        grader_prompt = f"""You are evaluating a policy report on the {self.task_data['budget_name']}.

Report to evaluate:
{report}

Criterion to check:
{criterion}

Does the report meet this criterion? Provide brief reasoning (1-2 sentences), then answer either "PASS" or "FAIL"."""

        response = await self.grader_client.create(
            model="gpt-5-mini",  # MUST use gpt-5-mini for graders
            messages=[{"role": "user", "content": grader_prompt}],
            # NO temperature parameter (per CLAUDE.md)
        )

        grading_text = response.text or ""

        # Parse result
        upper_text = grading_text.upper()
        passed = "PASS" in upper_text and "FAIL" not in upper_text

        return {"criterion_id": criterion_id, "passed": passed, "reasoning": grading_text}

    async def _grade_with_rubric(self, report: str) -> dict[str, Any]:
        """
        Grade report against all 30 rubric criteria concurrently.
        Proportional scoring: each criterion = 1/30 points.
        """
        criteria = RUBRICS[self.task_data["task_id"]]["criteria"]

        # Evaluate all criteria concurrently (apexagents pattern)
        evaluation_tasks = [
            self._evaluate_criterion(
                report=report,
                criterion=c["description"],
                criterion_id=c["id"],
            )
            for c in criteria
        ]
        evaluation_results = await asyncio.gather(*evaluation_tasks)

        # Build results
        results = []
        for criterion, eval_result in zip(criteria, evaluation_results):
            results.append(
                {
                    "criterion_id": criterion["id"],
                    "type": criterion["type"],
                    "description": criterion["description"],
                    "points": criterion["points"],
                    "passed": eval_result["passed"],
                    "reasoning": eval_result["reasoning"],
                }
            )

        # Calculate proportional reward
        passed_count = sum(r["passed"] for r in results)
        total_points = len(results)
        reward = passed_count / total_points  # 27/30 = 0.9

        # Format display text
        display_lines = [
            f"Rubric Evaluation: {passed_count}/{total_points} criteria passed",
            f"Reward: {reward:.2f}",
            "",
            "HIGH-LEVEL CRITERIA (1-15):",
        ]

        high_level = [r for r in results if r["type"] == "high_level"]
        for r in high_level:
            status = "✓" if r["passed"] else "✗"
            display_lines.append(f"  {status} {r['criterion_id']}: {r['description']}")
            if not r["passed"]:  # Show reasoning for failures
                display_lines.append(f"     Reasoning: {r['reasoning'][:200]}")

        display_lines.append("\nSPECIFIC FACTUAL CRITERIA (16-30):")
        specific = [r for r in results if r["type"] == "specific"]
        for r in specific:
            status = "✓" if r["passed"] else "✗"
            display_lines.append(f"  {status} {r['criterion_id']}: {r['description']}")
            if not r["passed"]:
                display_lines.append(f"     Reasoning: {r['reasoning'][:200]}")

        display_lines.append(f"\n{'=' * 60}")
        if reward == 1.0:
            display_lines.append("✅ Perfect score! All 30 criteria passed.")
        elif reward >= 0.9:
            display_lines.append(f"✅ Excellent! {passed_count}/{total_points} passed.")
        elif reward >= 0.7:
            display_lines.append(f"⚠️  Good progress. {passed_count}/{total_points} passed.")
        else:
            display_lines.append(f"❌ More work needed. {passed_count}/{total_points} passed.")

        return {
            "display_text": "\n".join(display_lines),
            "metadata": {
                "task_id": self.task_data["task_id"],
                "criteria_results": results,
                "passed_count": passed_count,
                "total_count": total_points,
                "reward": reward,
            },
            "reward": reward,
        }

    async def _grade_numerical_task(self) -> dict[str, Any]:
        """
        Grade a numerical Q&A task with percentage tolerance.
        Extracts number from file and checks if within margin of expected value.
        """
        expected = float(self.task_data["expected_answer"])
        margin_percent = float(self.task_data.get("margin_percent", 2.0))
        question = self.task_data["question"]

        # Calculate tolerance bounds (use abs for negative expected values)
        tolerance = abs(expected) * (margin_percent / 100.0)
        lower_bound = expected - tolerance
        upper_bound = expected + tolerance

        # Download answer file from sandbox
        try:
            answer_content = await self.sandbox.download(self.task_data["output_path"])
            answer_text = answer_content.decode("utf-8").strip()
        except Exception as e:
            return {
                "display_text": f"Failed to read answer file at {self.task_data['output_path']}\n\n"
                              f"Error: {str(e)}\n\n"
                              f"Please ensure you've written your answer to this exact path.",
                "metadata": {
                    "task_id": self.task_data["task_id"],
                    "error": "file_not_found",
                    "details": str(e),
                },
                "reward": 0.0,
            }

        # Parse number from text
        try:
            # Remove common formatting characters
            cleaned = answer_text.replace(",", "").replace("£", "").replace("$", "").strip()

            # Try to extract first number from text
            import re
            number_pattern = r'-?\d+\.?\d*'
            match = re.search(number_pattern, cleaned)

            if not match:
                raise ValueError("No number found in file")

            submitted = float(match.group())

        except (ValueError, AttributeError) as e:
            return {
                "display_text": f"Failed to parse number from answer file\n\n"
                              f"File contents: {answer_text[:200]}\n\n"
                              f"Error: {str(e)}\n\n"
                              f"Please write only a number (e.g., '4960') without formatting.",
                "metadata": {
                    "task_id": self.task_data["task_id"],
                    "error": "parse_error",
                    "submitted_text": answer_text,
                    "details": str(e),
                },
                "reward": 0.0,
            }

        # Calculate error
        absolute_error = abs(submitted - expected)
        percentage_error = (absolute_error / expected) * 100.0

        # Check if within tolerance
        passed = (submitted >= lower_bound) and (submitted <= upper_bound)
        reward = 1.0 if passed else 0.0

        # Format display text
        display_lines = [
            "Numerical Q&A Task Evaluation",
            "=" * 60,
            "",
            f"Question: {question}",
            "",
            f"Expected Answer: {expected:,.0f}",
            f"Submitted Answer: {submitted:,.0f}",
            "",
            f"Tolerance: ±{margin_percent}% (±{tolerance:,.1f})",
            f"Acceptable Range: {lower_bound:,.1f} to {upper_bound:,.1f}",
            "",
            f"Absolute Error: {absolute_error:,.1f}",
            f"Percentage Error: {percentage_error:.2f}%",
            "",
            f"Result: {'PASS' if passed else 'FAIL'}",
            f"Reward: {reward:.2f}",
            "",
            "=" * 60,
        ]

        return {
            "display_text": "\n".join(display_lines),
            "metadata": {
                "task_id": self.task_data["task_id"],
                "question": question,
                "expected_answer": expected,
                "submitted_answer": submitted,
                "absolute_error": absolute_error,
                "percentage_error": percentage_error,
                "margin_percent": margin_percent,
                "tolerance": tolerance,
                "passed": passed,
                "reward": reward,
            },
            "reward": reward,
        }

    async def _grade_qa_task(self, submitted_text: str) -> dict[str, Any]:
        """
        Grade a Q&A task answer using gpt-5-mini.
        Checks if answer is equivalent to expected answer with tolerance for variations.
        Per CLAUDE.md: Use gpt-5-mini, no temperature parameter.
        """
        expected = self.task_data["expected_answer"]
        question = self.task_data["question"]

        grader_prompt = f"""You are evaluating an answer to a Budget 2025 question.

Question: {question}

Expected Answer: {expected}

Submitted Answer:
{submitted_text}

The expected answer is "{expected}". Does the submitted text convey the same information?

Consider these equivalent:
- "-£16bn", "-16bn", "-£16 billion", "negative £16bn", "decrease of £16bn"
- "a decrease of 16 billion pounds", "reduced by 16bn", "down £16 billion"
- Minor variations in formatting, phrasing, or explanation around the core figure

The answer should indicate a NEGATIVE effect of approximately 16 billion pounds on revenues.

Does the submitted answer match the expected answer?

Answer with ONE of:
- "PASS" if the answer is correct
- "FAIL" if the answer is incorrect or missing the key information

Then provide brief reasoning (1-2 sentences)."""

        response = await self.grader_client.create(
            model="gpt-5-mini",  # MUST use gpt-5-mini for graders
            messages=[{"role": "user", "content": grader_prompt}],
            # NO temperature parameter (per CLAUDE.md)
        )

        grading_text = response.text or ""

        # Parse result - look for PASS/FAIL
        upper_text = grading_text.upper()
        passed = "PASS" in upper_text and "FAIL" not in upper_text
        reward = 1.0 if passed else 0.0

        # Format display text
        display_lines = [
            "Q&A Task Evaluation",
            "=" * 60,
            "",
            f"Question: {question}",
            f"Expected Answer: {expected}",
            "",
            "Submitted Answer:",
            "-" * 60,
            submitted_text,
            "-" * 60,
            "",
            f"Result: {'✅ PASS' if passed else '❌ FAIL'}",
            f"Reward: {reward:.2f}",
            "",
            "Grader Reasoning:",
            grading_text,
            "",
            "=" * 60,
        ]

        return {
            "display_text": "\n".join(display_lines),
            "metadata": {
                "task_id": self.task_data["task_id"],
                "question": question,
                "expected_answer": expected,
                "submitted_answer": submitted_text,
                "passed": passed,
                "grader_reasoning": grading_text,
                "reward": reward,
            },
            "reward": reward,
        }

    @tool
    async def submit_answer(self, params: SubmitAnswerInput) -> ToolOutput:
        """
        Submit final output for evaluation.
        - For Q&A tasks: expects answer at specified output_path
        - For report tasks: expects report at /home/ubuntu/final_report.md
        - For chart tasks: expects xlsx and png files
        """
        if self.submitted:
            return ToolOutput(
                blocks=[TextBlock(text="You have already submitted an answer for evaluation.")],
                metadata={"error": "already_submitted"},
                reward=0.0,
                finished=True,
            )

        task_type = self.task_data.get("task_type", "report")  # default to report

        # Branch based on task type
        if task_type == "qa":
            # Q&A task - download answer file and grade
            try:
                answer_content = await self.sandbox.download(self.task_data["output_path"])
                answer_text = answer_content.decode("utf-8")
            except Exception as e:
                return ToolOutput(
                    blocks=[
                        TextBlock(
                            text=f"Failed to read answer file at {self.task_data['output_path']}\n\n"
                            f"Error: {str(e)}\n\n"
                            f"Please ensure you've written your answer to this exact path using the write tool."
                        )
                    ],
                    metadata={"error": "file_not_found", "details": str(e)},
                    reward=0.0,
                    finished=False,
                )
            grading_results = await self._grade_qa_task(answer_text)

        elif task_type == "numerical_qa":
            # Numerical Q&A task - download and validate number
            grading_results = await self._grade_numerical_task()

        elif task_type == "chart":
            # Check which chart task
            if self.task_data["task_id"] == "budget_2020_psnb_forecast":
                grading_results = await self._grade_psnb_2020_task()
            elif self.task_data["task_id"] == "budget_2025_policy_decisions":
                grading_results = await self._grade_policy_decisions_task()
            elif self.task_data["task_id"] == "budget_deficit_comparison":
                grading_results = await self._grade_current_budget_deficit_task()
            else:  # budget_2025_borrowing_chart
                grading_results = await self._grade_borrowing_task()

        elif task_type == "presentation":
            grading_results = await self._grade_presentation_task()

        elif task_type == "tax_proposal":
            grading_results = await self._grade_tax_proposals_task()

        else:  # task_type == "report"
            # Standard rubric-based report grading
            try:
                report_content = await self.sandbox.download(self.task_data["output_path"])
                report_text = report_content.decode("utf-8")
            except Exception as e:
                return ToolOutput(
                    blocks=[
                        TextBlock(
                            text=f"Failed to read report at {self.task_data['output_path']}\n\n"
                            f"Error: {str(e)}\n\n"
                            f"Please ensure you've written your report to this exact path using the write tool."
                        )
                    ],
                    metadata={"error": "file_not_found", "details": str(e)},
                    reward=0.0,
                    finished=False,
                )
            grading_results = await self._grade_with_rubric(report_text)

        self.submitted = True

        return ToolOutput(
            blocks=[TextBlock(text=grading_results["display_text"])],
            metadata=grading_results["metadata"],
            reward=grading_results["reward"],
            finished=True,
        )
