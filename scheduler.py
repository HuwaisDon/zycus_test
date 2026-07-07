from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

import schedule

from main import run_pipeline


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("project-health-scheduler")


def _run_once(input_path: Path | str = "d:/zycus") -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logger.info("Starting pipeline execution at %s", timestamp)
    try:
        result = run_pipeline(
            input_path=input_path,
            weekly_output_dir=Path("outputs/weekly"),
            monthly_output_dir=Path("outputs/monthly"),
            ppt_output_path=Path("outputs/monthly/Executive_Project_Health_Report.pptx"),
        )
        logger.info("Pipeline completed successfully: %s", result)
    except Exception as exc:  # pragma: no cover - defensive logging path
        logger.exception("Pipeline execution failed: %s", exc)


def run_continuously(interval_minutes: int = 60, input_path: Path | str = "d:/zycus") -> None:
    schedule.every().monday.at("09:00").do(_run_once, input_path=input_path)

    logger.info("Scheduler started. Next run is scheduled for Monday at 09:00.")
    while True:
        schedule.run_pending()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--once":
        _run_once()
    else:
        run_continuously()
