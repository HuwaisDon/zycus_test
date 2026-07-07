from pathlib import Path

from main import run_pipeline


def test_run_pipeline_writes_expected_outputs(tmp_path: Path) -> None:
    weekly_dir = tmp_path / "weekly"
    monthly_dir = tmp_path / "monthly"
    ppt_path = monthly_dir / "Executive_Project_Health_Report.pptx"

    result = run_pipeline(
        input_path=Path("d:/zycus"),
        weekly_output_dir=weekly_dir,
        monthly_output_dir=monthly_dir,
        ppt_output_path=ppt_path,
    )

    assert result["generated_reports"] >= 1
    assert (monthly_dir / "portfolio_summary.json").exists()
    assert ppt_path.exists()
