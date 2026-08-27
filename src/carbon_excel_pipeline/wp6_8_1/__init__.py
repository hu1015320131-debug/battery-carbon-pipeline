"""WP6-8.1 capability-driven end-to-end orchestration."""

from carbon_excel_pipeline.wp6_8_1.pipeline import run_end_to_end_pipeline
from carbon_excel_pipeline.wp6_8_1.router import decide_processing_routes

__all__ = ["decide_processing_routes", "run_end_to_end_pipeline"]
