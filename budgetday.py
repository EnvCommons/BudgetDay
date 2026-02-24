from __future__ import annotations

import asyncio
import json
from pathlib import Path
from typing import Any

import openai
from openreward import AsyncOpenReward, SandboxBucketConfig, SandboxSettings
from openreward.environments import JSONObject, TextBlock, ToolOutput, tool
from pydantic import BaseModel

from cli_environment import CLIEnvironment
from constants import ENV_PATH


# Load rubrics at module level
with open(ENV_PATH / "rubrics.json") as f:
    RUBRICS = json.load(f)

# Define tasks
TASKS = [
    {
        "task_id": "budget_2024_initial_response",
        "year": 2024,
        "budget_name": "Autumn Budget 2024",
        "data_dir": "2024",
        "output_path": "/home/ubuntu/final_report.md",
        "description": "Draft initial response to Autumn Budget 2024",
    },
    {
        "task_id": "budget_2025_initial_response",
        "year": 2025,
        "budget_name": "Budget 2025",
        "data_dir": "2025",
        "output_path": "/home/ubuntu/final_report.md",
        "description": "Draft initial response to 2025 Budget",
    },
]


class SubmitAnswerInput(BaseModel):
    """Input for submit_answer tool (no parameters needed - reads from fixed path)."""

    pass


class BudgetDay(CLIEnvironment):
    """Budget Day environment - stub implementation"""

    @classmethod
    def list_splits(cls) -> list[str]:
        """Return available data splits"""
        return ["test"]

    @classmethod
    def list_tasks(cls, split: str) -> list[JSONObject]:
        """Return list of tasks for a given split"""
        if split != "test":
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

        # Initialize OpenAI client for grading
        api_key = secrets.get("openai_api_key")
        if not api_key:
            raise ValueError("OpenAI API key required in secrets for rubric grading")
        self.grader_client = openai.AsyncClient(api_key=api_key)

        # Configure sandbox with selective data mounting
        self.sandbox_settings = SandboxSettings(
            environment="GeneralReasoning/BudgetDay",
            image="generalreasoning/python-ds:3.12-tools",
            machine_size="0.5:0.5",
            block_network=False,
            bucket_config=SandboxBucketConfig(
                mount_path="/orwd_data",
                read_only=True,
                only_dir=self.task_data["data_dir"],  # Only mount budgetday/2025
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

        # Base prompt common to all tasks
        base_prompt = f"""# Task: Draft Initial Response to {self.task_data['budget_name']}

You are a policy analyst at a fiscal policy research organization. Your task is to draft an "initial response" to the {self.task_data['budget_name']} - a rapid policy analysis meant to be written soon after Budget Day.

## Available Data

Budget documents are available at `/orwd_data/` (mounted read-only).

**File Format**: MHTML (MIME-encoded HTML) files are text-based. You can:
- Use `read <file_path>` to view contents
- Use `grep <pattern> <path>` to search for specific terms or numbers
- Use `bash` to run Python scripts for parsing if needed
- Use standard text processing tools

## Your Task

Write a comprehensive initial response that provides rapid but serious analysis of the Budget.
"""

        # Task-specific guidance
        if self.task_data["task_id"] == "budget_2024_initial_response":
            task_guidance = """
## Key Themes for Autumn Budget 2024

Your analysis should cover:

1. **"Broad Brush Strokes" Headlines**: Big tax rises, more cash for public services, more borrowing, more investment
2. **Central Organizing Device: "Two Gambles"**:
   - First gamble: Public services risk (can big cash injection improve services before pressures return?)
   - Second gamble: Borrowing efficiency risk (will increased borrowing "pay off" through growth?)
3. **Sequencing Theme**: Front-loaded support vs later restraint - tension between near-term spending/borrowing and later discipline
4. **Fiscal Credibility**: Not just rules compliance, but practical believability of slower future spending growth
5. **Growth Assessment**: Short-term vs long-term distinction (time-profiled and conditional, not simply pro/anti-growth)
6. **Tax Package**:
   - Incidence and economic effects (who ultimately bears tax rises)
   - Tax reform critique (revenue-raising vs improving tax design)
   - Policy inconsistency (especially climate goals vs fuel duty)
7. **Spending Profile**: Front-loading and later tightness, not just aggregate numbers
8. **Fiscal Inheritance**: Framing that inherited plans were unrealistic, tax rises near-inevitable
9. **Tone**: Analytically skeptical but balanced - commending some choices while warning about risks

## Important Figures to Find

Your analysis should include specific numbers such as:
- Tax-to-GDP ratio (highest ever?)
- Near-term borrowing increase (£28bn in 2025-26?)
- Public investment increase (£19bn?)
- Borrowing composition (day-to-day vs investment split)
- Stability rule timing (borrowing only to invest by when?)
- Pre-election vs current borrowing plans comparison
- Growth trajectory timeline (short-term, end of parliament, 2032)
- Employer NICs details (rate and threshold changes)
- Tax incidence estimates (e.g., three-quarters of NICs on employees)
- Spending growth rates over time (+4.8%, +3.1%, then 1.3%/year?)
- Fiscal headroom (£10bn?)
- Fiscal framework changes (rolling target, PSNFL rule)
"""
        else:  # budget_2025_initial_response
            task_guidance = """
## Key Themes for Budget 2025

Your analysis should cover:

1. **Economic Context**: OBR forecast changes, growth projections, fiscal position
2. **Key Policy Announcements**: Tax changes, spending plans, welfare reforms
3. **Fiscal Strategy**: "Spend now, pay later" sequencing, consolidation path credibility
4. **Detailed Analysis**:
   - Tax package design (distribution, incentives, reform quality)
   - Public spending pressures (including SEND funding specifically)
   - Welfare changes and poverty implications (especially two-child limit)
5. **Synthesis**: Big-picture assessment with zoomed-out perspective

## Requirements

Your response MUST:
- Frame this as an "initial response" (rapid policy analysis), not a retrospective
- Anchor clearly to Budget 2025 context
- Cover BOTH the economic forecast AND policy announcements
- Include specific figures and numbers from the documents
- Capture the headline theme: "big Budget, but not as expected"
- Explain the "spend now, pay later" sequencing
- Discuss both sides of fiscal picture (tax rises, headroom, borrowing pressures)
- Be analytical and evidence-based (mixed praise/critique), not political
- Use clear structure with headings/sections
- Include a synthesis paragraph stepping back to assess the whole package

## Important Figures to Include

Your analysis should reference specific numbers such as:
- Tax rise comparisons (this year vs previous: £26bn vs £32bn?)
- Headroom figures and changes (increased to £22bn?)
- OBR forecast revisions (revenues +£16bn, spending +£22bn, deterioration £6bn?)
- Fiscal consolidation figures (£12bn in 2029-30?)
- Time-profile of borrowing (higher 2025-29, lower in 2029-30?)
- Cumulative borrowing effect (£57bn over five years?)
- Threshold freeze details (3-year extension, yields)
- Taxpayer impacts (+5.2M taxpayers, +4.8M higher-rate)
- Tax burden trajectory (36.3% to 38.3% of GDP)
- Two-child limit abolition (£3bn/year, 560k families, £5,300 gain)
- SEND pressures (£6bn/year, 14% growth, 9% of schools budget)
"""

        submission_instructions = f"""
## Submission

When ready, write your final report to: **{self.task_data['output_path']}**

Then call `submit_answer` tool to submit for evaluation.

Your report will be graded against 30 criteria (15 broad/high-level + 15 specific factual). Each criterion contributes 1 point. Score 27/30 = 0.9 reward.

## Available Tools

- `bash` - Run shell commands, Python scripts
- `read` - Read file contents
- `write` - Write files (including your report)
- `grep` - Search for patterns in files
- `glob` - Find files matching patterns
- `submit_answer` - Submit your final report for grading
"""

        prompt_text = base_prompt + task_guidance + submission_instructions

        return [TextBlock(text=prompt_text)]

    async def _evaluate_criterion(
        self, report: str, criterion: str, criterion_id: str
    ) -> dict[str, Any]:
        """
        Use gpt-5-mini to evaluate a single criterion.
        Per CLAUDE.md: Use gpt-5-mini, no temperature parameter.
        """
        grader_prompt = f"""You are evaluating a policy report on the 2025 Budget.

Report to evaluate:
{report}

Criterion to check:
{criterion}

Does the report meet this criterion? Provide brief reasoning (1-2 sentences), then answer either "PASS" or "FAIL"."""

        response = await self.grader_client.chat.completions.create(
            model="gpt-5-mini",  # MUST use gpt-5-mini for graders
            messages=[{"role": "user", "content": grader_prompt}],
            # NO temperature parameter (per CLAUDE.md)
        )

        grading_text = response.choices[0].message.content or ""

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

    @tool
    async def submit_answer(self, params: SubmitAnswerInput) -> ToolOutput:
        """
        Submit final report for evaluation.
        Expects the report to be written to /home/ubuntu/final_report.md
        """
        if self.submitted:
            return ToolOutput(
                blocks=[TextBlock(text="You have already submitted an answer for evaluation.")],
                metadata={"error": "already_submitted"},
                reward=0.0,
                finished=True,
            )

        # Download report from sandbox
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

        # Grade against rubric
        grading_results = await self._grade_with_rubric(report_text)

        self.submitted = True

        return ToolOutput(
            blocks=[TextBlock(text=grading_results["display_text"])],
            metadata=grading_results["metadata"],
            reward=grading_results["reward"],
            finished=True,
        )
