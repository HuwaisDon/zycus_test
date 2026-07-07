import os
import unittest
from datetime import date, datetime
from unittest.mock import patch

from parser.xlsx_adapter import Project, ProjectMetadata, TaskComment, TaskInfo, MilestoneInfo
from parser.rag_engine import analyze_project_health


class ProjectHealthEngineTests(unittest.TestCase):
    def _make_project(self) -> Project:
        tasks = [
            TaskInfo(
                task_uid="1",
                task_name="Requirements gathering",
                status="In Progress",
                start_date=date(2026, 1, 1),
                finish_date=date(2026, 2, 10),
                percent_complete=55.0,
                variance=-5.0,
                total_float=-2.0,
                schedule_health="Red",
            ),
            TaskInfo(
                task_uid="2",
                task_name="Configuration",
                status="Completed",
                start_date=date(2026, 2, 1),
                finish_date=date(2026, 2, 20),
                percent_complete=100.0,
                variance=0.0,
                total_float=3.0,
                schedule_health="Green",
            ),
        ]
        milestones = [
            MilestoneInfo(milestone_name="Kickoff", due_date=date(2026, 1, 15)),
            MilestoneInfo(milestone_name="UAT", due_date=date(2026, 3, 1)),
            MilestoneInfo(milestone_name="Go Live", due_date=date(2026, 4, 1)),
        ]
        comments = [
            TaskComment(comment="Blocked by pending approval and dependency mapping", author="PM"),
            TaskComment(comment="Team is concerned about risk and issues", author="Stakeholder"),
            TaskComment(comment="Great progress and positive momentum", author="Sponsor"),
        ]
        return Project(
            metadata=ProjectMetadata(project_name="Sample Project", client_name="Contoso", project_manager="Alex"),
            tasks=tasks,
            milestones=milestones,
            comments=comments,
        )

    def test_returns_expected_structure_and_scores(self) -> None:
        project = self._make_project()

        with patch.dict(os.environ, {}, clear=False):
            result = analyze_project_health(project)

        self.assertIn("overall_rag", result)
        self.assertIn("confidence", result)
        self.assertIn("signal_scores", result)
        self.assertIn("evidence", result)
        self.assertIn("recommendations", result)

        self.assertEqual(result["overall_rag"], "Red")
        self.assertGreaterEqual(result["confidence"], 0.0)
        self.assertLessEqual(result["confidence"], 1.0)

        schedule = result["signal_scores"]["schedule_health"]
        self.assertEqual(schedule["rag"], "Red")
        self.assertIn("overdue_tasks", schedule)
        self.assertIn("variance", schedule)

        milestone_signal = result["signal_scores"]["milestone_health"]
        self.assertIn("completed", milestone_signal)
        self.assertIn("overdue", milestone_signal)
        self.assertIn("upcoming", milestone_signal)
        self.assertIn("missed", milestone_signal)

        blockers = result["signal_scores"]["blockers"]
        self.assertEqual(blockers["rag"], "Red")
        self.assertGreaterEqual(blockers["count"], 1)

        sentiment = result["signal_scores"]["stakeholder_sentiment"]
        self.assertIn("rag", sentiment)
        self.assertIn("score", sentiment)

        budget = result["signal_scores"]["budget"]
        self.assertEqual(budget["status"], "Not Available")
        self.assertLess(budget["confidence"], 0.9)

    def test_uses_keyword_sentiment_without_gemini(self) -> None:
        project = self._make_project()
        with patch.dict(os.environ, {"GEMINI_API_KEY": ""}, clear=False):
            result = analyze_project_health(project)
        sentiment = result["signal_scores"]["stakeholder_sentiment"]
        self.assertIn("rag", sentiment)
        self.assertIn("score", sentiment)


if __name__ == "__main__":
    unittest.main()
