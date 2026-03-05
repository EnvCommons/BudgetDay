# BudgetDay

[![OpenReward Environment](https://img.shields.io/badge/%E2%AD%90%20OpenReward-Environment-f7e6cc)](https://openreward.ai/GeneralReasoning/BudgetDay)

## Description

**BudgetDay** is an ORS environment for evaluating agents on UK fiscal policy analysis tasks. Agents are given a sandboxed environment with access to real UK Budget and Autumn Statement documents (2020-2025) and must produce policy reports, data visualizations, numerical analyses, presentations, and creative policy proposals. Tasks span drafting IFS-style initial responses, extracting and charting OBR borrowing forecasts, calculating household income impacts, creating PowerPoint presentations on tax measures, and writing parliamentary-style opposition responses. Grading uses a combination of LLM rubric evaluation (gpt-5-mini), numerical tolerance checks, and vision-based chart validation.

## Capabilities

- Drafting analytical policy reports from primary budget documents
- Extracting and comparing fiscal data across multiple budget years
- Creating charts (bar, line) and Excel spreadsheets from budget data
- Calculating household income impacts under specific demographic scenarios
- Creating PowerPoint presentations summarizing tax measures
- Proposing tax policy packages to meet fiscal targets
- Writing parliamentary-style speeches grounded in budget data

## Compute Requirements

Agents in BudgetDay are given a sandbox with 1 CPU and 2 GB RAM. The sandbox uses the `generalreasoning/knowledge-worker:latest` image, which includes tools for working with Word, Excel, PowerPoint, and PDF documents. Network access is enabled.

## Tasks

There is one split: train (21 tasks). Tasks span six types across UK budgets from 2020 to 2025:

- **Report** (6 tasks): Draft IFS-style initial responses to budgets (2022, 2023, 2024, 2025), compile a cross-year AI measures summary (2020-2025), and write a Leader of the Opposition parliamentary response to Budget 2025.
- **Chart** (4 tasks): Extract fiscal data and produce Excel spreadsheets and charts (borrowing forecasts, policy decisions, budget deficit comparisons, PSNB forecasts).
- **Numerical QA** (8 tasks): Calculate household income changes for specific demographic scenarios under Budget 2025.
- **QA** (1 task): Answer a qualitative question about the effect of productivity downgrade on revenues.
- **Presentation** (1 task): Create a PowerPoint presentation on Autumn Budget 2024 tax measures.
- **Tax proposal** (1 task): Propose tax changes to reduce borrowing by half using provided tax-raising guidelines.

## Reward Structure

BudgetDay uses a mixed reward structure that varies by task type:

- **Report tasks**: Continuous reward (0.0-1.0). Each report is graded against a 30-criterion rubric (15 high-level + 15 specific factual criteria) using an LLM grader (gpt-5-mini). Reward is the proportion of criteria passed.
- **Chart tasks**: Continuous reward (0.0-1.0). Spreadsheets are graded via LLM text extraction against ground-truth values. Chart images are graded via LLM vision against expected visual properties (chart type, colors, values). Scores are combined (typically 50/50 spreadsheet/chart weighting, or 70/30 when chart is optional).
- **Numerical QA tasks**: Binary reward. The agent's numerical answer must fall within +/-2% of the expected value. Reward is 1.0 (within tolerance) or 0.0 (outside tolerance).
- **QA tasks**: Binary reward. An LLM grader (gpt-5-mini) checks semantic equivalence to the expected answer.
- **Presentation tasks**: Continuous reward (0.0-1.0). PowerPoint text is extracted and graded against a 30-criterion rubric.
- **Tax proposal tasks**: Continuous reward (0.0-1.0). An LLM grader (gpt-5-mini) parses proposals, calculates total revenue using provided guidelines, and scores based on squared error from the target.
- **Opposition response tasks**: Continuous reward (0.0-1.0). Graded against a 35-criterion rubric covering content and stylistic requirements.

## Data

UK Budget and Autumn Statement documents are mounted read-only from `/orwd_data/` in the sandbox, organized by year (2020, 2022, 2023, 2024, 2025). Each task mounts only the relevant year's documents (or all years for cross-year tasks). Ground-truth data for chart tasks and rubrics for report tasks are embedded in the environment code.

## Tools

Agents are given CLI tools and document toolsets:

- `bash`: Run a bash command in the sandbox.
- `glob`: Find files matching a glob pattern.
- `grep`: Search for patterns in files.
- `ls`: List files and directories.
- `read`: Read file contents.
- `write`: Write content to a file.
- `edit`: Perform string replacement in a file.
- `multi_edit`: Perform multiple edits on a single file.
- `todo_write`: Manage a todo list for task planning.
- `submit_answer`: Submit output for evaluation. This tool can only be called once per task.

Additionally, four toolsets are included: **WordToolset** (creating and editing Word documents), **ExcelToolset** (creating and editing Excel spreadsheets), **PowerPointToolset** (creating and editing PowerPoint presentations), and **PDFToolset** (reading and extracting content from PDF files). Each toolset provides multiple tools for its respective document type.

## Time Horizon

BudgetDay is a multi-turn environment. The agent iterates using CLI tools to read budget documents, analyze data, write scripts, create output files (reports, spreadsheets, charts, presentations), and submit for evaluation.

[How many average tool calls?]

## Environment Difficulty

[Statistics on environment difficulty here]

## Other Environment Requirements

BudgetDay requires an OpenAI API key (`OPENAI_API_KEY` secret) for LLM-based rubric grading and chart validation.

## Safety

Agents in BudgetDay analyze publicly available UK government budget documents in a sandboxed environment. The environment does not present direct safety risks, as agents only interact with published fiscal policy documents and produce analytical outputs with no access to financial systems or real policy-making processes.

## Citations

```bibtex
@dataset{GRBudgetDay,
  author    = {General Reasoning Inc. Team},
  title     = {BudgetDay},
  year      = {2026},
  publisher = {OpenReward},
  url       = {https://openreward.ai/GeneralReasoning/BudgetDay}
}
```
