from __future__ import annotations

import logging
from pathlib import Path

from parser.xlsx_adapter import parse_project_xlsx


def main() -> None:
    logging.basicConfig(level=logging.INFO)
    xlsx_path = Path(__file__).parent / "S2P Project (2).xlsx"
    project = parse_project_xlsx(xlsx_path)

    print(project.metadata.project_name or "")
    print(project.metadata.client_name or "")
    print(len([t for t in project.tasks if (t.task_name or t.task_uid)]))
    print(len([m for m in project.milestones if m.milestone_name]))
    print(len([c for c in project.comments if c.comment]))



if __name__ == "__main__":
    main()

