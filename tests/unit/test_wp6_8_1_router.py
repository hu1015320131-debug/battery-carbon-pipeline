from carbon_excel_pipeline.wp6_8_1.router import (
    ACTIVITY_DIRECT,
    ACTIVITY_PCS,
    FACTOR_EMBEDDED,
    FACTOR_MISSING,
    decide_processing_routes,
)


OFFICIAL_2024 = "SYNTHETIC_INPUT_SHA_PLACEHOLDER"
OFFICIAL_2025 = "SYNTHETIC_INPUT_SHA_PLACEHOLDER"


def test_direct_mass_with_embedded_factor_and_official_boundary():
    decision = decide_processing_routes(
        capability={"direct_reported_mass_count": 2, "pcs_weight_derived_count": 0, "factor_ready_count": 2},
        input_sha256=OFFICIAL_2024,
    )
    assert decision["Activity_Route"] == ACTIVITY_DIRECT
    assert decision["Factor_Route"] == FACTOR_EMBEDDED
    assert decision["Boundary_Policy"] == "SYNTHETIC_DIRECT_MASS_BOUNDARY_V1"
    assert decision["status"] == "ROUTED"
    assert decision["Year_Used_As_Router"] is False


def test_pcs_without_embedded_factor_is_not_given_a_default():
    decision = decide_processing_routes(
        capability={"direct_reported_mass_count": 0, "pcs_weight_derived_count": 10, "factor_ready_count": 0},
        input_sha256=OFFICIAL_2025,
    )
    assert decision["Activity_Route"] == ACTIVITY_PCS
    assert decision["Factor_Route"] == FACTOR_MISSING
    assert decision["Factor_Ready"] is False
    assert decision["Boundary_Ready"] is False
    assert decision["status"] == "PARTIAL_RESULT"


def test_unknown_pcs_file_does_not_inherit_simulation_factor():
    decision = decide_processing_routes(
        capability={"direct_reported_mass_count": 0, "pcs_weight_derived_count": 10, "factor_ready_count": 0},
        input_sha256="A" * 64,
    )
    assert decision["Activity_Route"] == ACTIVITY_PCS
    assert decision["Activity_Ready"] is True
    assert decision["Factor_Route"] == FACTOR_MISSING
    assert decision["Factor_Ready"] is False
    assert decision["Boundary_Ready"] is False
    assert decision["status"] == "PARTIAL_RESULT"


def test_unknown_direct_mass_does_not_inherit_2024_boundary():
    decision = decide_processing_routes(
        capability={"direct_reported_mass_count": 5, "pcs_weight_derived_count": 0, "factor_ready_count": 5},
        input_sha256="B" * 64,
    )
    assert decision["Activity_Route"] == ACTIVITY_DIRECT
    assert decision["Factor_Route"] == FACTOR_EMBEDDED
    assert decision["Boundary_Policy"] == "BOUNDARY_POLICY_NOT_AVAILABLE"
    assert decision["Activity_Ready"] is True
    assert decision["Boundary_Ready"] is False
    assert decision["status"] == "PARTIAL_RESULT"


def test_both_activity_paths_are_ambiguous_without_explicit_policy():
    decision = decide_processing_routes(
        capability={"direct_reported_mass_count": 3, "pcs_weight_derived_count": 3, "factor_ready_count": 3},
        input_sha256="C" * 64,
    )
    assert decision["Activity_Route"] is None
    assert "MULTIPLE_VALID_ROUTES" in decision["Blocking_Reasons"]
    assert decision["status"] == "MULTIPLE_VALID_ROUTES"


def test_year_is_not_accepted_as_a_route_selector():
    decision = decide_processing_routes(
        capability={"direct_reported_mass_count": 2, "pcs_weight_derived_count": 0, "factor_ready_count": 2},
        input_sha256=OFFICIAL_2024,
    )
    assert "2024" not in str(decision["Activity_Route_Reason"])
    assert decision["Year_Used_As_Router"] is False


def test_no_activity_path_is_blocked():
    decision = decide_processing_routes(
        capability={"direct_reported_mass_count": 0, "pcs_weight_derived_count": 0, "factor_ready_count": 0},
        input_sha256="D" * 64,
    )
    assert decision["status"] == "BLOCKED"
    assert "NO_ACTIVITY_PATH" in decision["Blocking_Reasons"]
