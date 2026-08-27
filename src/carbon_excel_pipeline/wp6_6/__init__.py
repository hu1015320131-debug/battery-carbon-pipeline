"""WP6-6 cross-year historical factor scenario analysis."""

from .pipeline import (
    WP66AnalysisError,
    build_scenario_analysis,
    run_wp6_6_analysis,
)

__all__ = ["WP66AnalysisError", "build_scenario_analysis", "run_wp6_6_analysis"]
