# AI Project Health Reporting Agent

> AI-powered project health analysis and executive reporting system Developed as a solution for the Zycus AI Engineer Intern take-home assessment.



---
## Key Design Principle

Project health is computed using deterministic business rules.

Large Language Models are used only for generating executive-friendly explanations and never for determining project health.

This ensures repeatable, explainable, and auditable results.
---

---

## Overview

Managing enterprise implementation projects often requires manually reviewing project plans, schedules, milestones, comments, and progress reports before stakeholders can understand project health.

This project automates that process by reading Microsoft Project Excel exports, analyzing project health using deterministic business rules, and generating executive-ready reports.

The system produces:

- Weekly Project Health Reports
- Portfolio Health Summary
- Executive PowerPoint Presentation
- Weekly Scheduled Execution

---

# Features

### Excel Project Parser

- Reads Microsoft Project Excel exports
- Automatically detects project sheets
- Parses task hierarchy
- Extracts milestones
- Parses project comments
- Handles inconsistent workbook layouts
- Handles missing data gracefully
- Supports multiple project workbooks

---

### Deterministic RAG Engine

Computes project health using explainable business rules instead of relying on LLM reasoning.

Signals evaluated:

- Schedule Health
- Milestone Health
- Blockers
- Stakeholder Sentiment
- Budget Availability

Output includes:

- Overall RAG Status
- Confidence Score
- Evidence
- Recommendations

---

### Weekly Report Generator

Automatically generates leadership-ready Markdown reports including:

- Executive Summary
- Project Status
- Signal Breakdown
- Key Evidence
- Recommendations
- Data Quality Notes

---

### Monthly Portfolio Synthesis

Combines multiple projects into a portfolio view.

Produces:

- Portfolio Health
- Health Distribution
- Cross Project Risk Themes
- Escalation Candidates
- Executive Recommendations

Outputs:

```
portfolio_summary.json
```

---

### Executive PowerPoint Generator

Automatically creates:

```
Executive_Project_Health_Report.pptx
```

Slides include:

- Title
- Portfolio Overview
- Health Distribution
- Projects Requiring Escalation
- Cross Project Risks
- Executive Recommendations
- Methodology

---

### Weekly Scheduler (Bonus)

Supports automated weekly execution.

Runs every Monday at **09:00** using the Python `schedule` library.

Can also be integrated with:

- Windows Task Scheduler
- Linux Cron

---

# Architecture

```
                     Excel Workbooks
                            │
                            ▼
                 parser/xlsx_adapter.py
                            │
                            ▼
                Normalized Project Model
                            │
                            ▼
                parser/rag_engine.py
                            │
        ┌───────────────────┼────────────────────┐
        ▼                   ▼                    ▼
Weekly Report      Portfolio Synthesis      JSON Export
 Generator           (Monthly)                 │
        │                   │                  ▼
        ▼                   ▼         ppt_generator.py
 Markdown Report     portfolio_summary.json      │
                                                  ▼
                                Executive PowerPoint
```

---

# Folder Structure

```
project/

├── parser/
│   ├── xlsx_adapter.py
│   ├── rag_engine.py
│
├── report_generator.py
├── monthly_synthesis.py
├── ppt_generator.py
├── scheduler.py
├── main.py
│
├── outputs/
│   ├── weekly/
│   └── monthly/
│
├── tests/
│
├── requirements.txt
└── README.md
```

---

# Technology Stack

- Python 3.12
- Pandas
- OpenPyXL
- python-pptx
- Schedule
- Pytest

Optional:

- Google Gemini API (Natural language explanations)

---

# Installation

Clone the repository

```bash
git clone https://github.com/HuwaisDon/zycus_test.git
cd zycus_test
```

Install dependencies

```bash
pip install -r requirements.txt
```

---

# Running the Project

## Analyze a Single Workbook

```bash
python main.py --input "S2P Project (2).xlsx"
```

---

## Analyze All Workbooks

```bash
python main.py --input data/
```

---

## Generate Weekly Reports

```bash
python report_generator.py data/
```

---

## Generate Portfolio Summary

```bash
python monthly_synthesis.py data/
```

---

## Generate Executive PowerPoint

```bash
python ppt_generator.py outputs/monthly/portfolio_summary.json
```

---

## Run Weekly Scheduler

```bash
python scheduler.py
```

---

# Output

The project automatically generates

```
outputs/

weekly/

    Project_Plan_weekly_report.md

    S2P_Project_weekly_report.md

monthly/

    portfolio_summary.json

    Executive_Project_Health_Report.pptx
```

---

# Design Decisions

## Why deterministic scoring?

Project health should be explainable and auditable.

The overall RAG status is therefore calculated using deterministic business rules rather than generated by an LLM.

This ensures consistent results for identical project data.

---

## Why use an LLM only for explanations?

LLMs excel at summarizing findings but should not determine project health.

The system computes all scores deterministically and optionally uses Gemini only to produce executive-friendly narratives.

---

## Handling Incomplete Data

The system never fails because of:

- Missing budget information
- Empty cells
- Malformed dates
- Missing comment sheets
- Different workbook layouts
- `#UNPARSEABLE` values

Instead, confidence is reduced and analysis continues.

---

# Testing

Unit tests cover:

- Excel parser
- RAG engine
- Report generator
- Portfolio synthesis
- PowerPoint generation

Integration testing verifies the complete pipeline from workbook parsing to report generation.

---

# Sample Workflow

```
Read Excel Files

↓

Normalize Project Data

↓

Compute RAG Health

↓

Generate Weekly Reports

↓

Build Portfolio Summary

↓

Generate Executive PowerPoint

↓

(Optional)
Run Weekly Scheduler
```

---

# Assumptions

- Workbooks follow Microsoft Project export format.
- Budget data may not always be available.
- Schedule health indicators may exist but are treated only as supporting evidence.
- Multiple workbooks may represent different projects within the same portfolio.

---

# Future Improvements

- Interactive Streamlit Dashboard
- Historical trend analysis
- Project risk forecasting using ML
- Email report distribution
- Teams / Slack integration
- Power BI dashboard integration
- PDF report generation
- Real-time project connectors (Jira, Azure DevOps, MS Project)

---

# Project Walkthrough

This repository contains complete documentation, architecture diagrams, generated reports, and reproducible execution instructions that demonstrate the end-to-end workflow from Excel ingestion to executive reporting.

---

# Author

**Huwais Al Qurni**

AI Engineer Intern Assessment Submission