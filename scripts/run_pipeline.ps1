param(
    [switch]$CheckOnly,
    [string]$InspectInput,
    [string]$RunRoot,
    [string]$ScopeCleanRun,
    [string]$ScopeConfig,
    [string]$StandardizeRun,
    [string]$PrivateIdBaseline,
    [string]$PrivateMappingWorkbook,
    [string]$UpstreamRun,
    [string]$StandardBaseline,
    [string]$ActivityBaseline,
    [string]$ThirdPartyBaseline,
    [string]$FactorMatchRun,
    [string]$FactorInput,
    [switch]$HistoricalSimulation,
    [string]$CalculationRun,
    [string]$Day8ExportRun,
    [string]$Day8RunAllInput,
    [string]$OutputDir,
    [string]$OutputRoot,
    [string]$ArtifactWorkDir,
    [string]$ArtifactWorkRoot,
    [string]$NodeExecutable,
    [string]$NodeModules,
    [string]$Wp5OpenItems
)

$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"

$env:PYTHONPATH = Join-Path $projectRoot "src"

if ($CheckOnly) {
    & $pythonExe -m carbon_excel_pipeline check
    exit $LASTEXITCODE
}

if ($InspectInput) {
    if (-not $RunRoot) {
        throw "-RunRoot is required with -InspectInput. It must point to the external isolated runs directory."
    }
    & $pythonExe -m carbon_excel_pipeline inspect --input $InspectInput --run-root $RunRoot
    exit $LASTEXITCODE
}

if ($ScopeCleanRun) {
    if ($ScopeConfig) {
        & $pythonExe -m carbon_excel_pipeline scope-clean --run-dir $ScopeCleanRun --scope-config $ScopeConfig
    } else {
        & $pythonExe -m carbon_excel_pipeline scope-clean --run-dir $ScopeCleanRun
    }
    exit $LASTEXITCODE
}

if ($StandardizeRun) {
    if (-not $PrivateIdBaseline -or -not $PrivateMappingWorkbook) {
        throw "Private Day 4 requires -PrivateIdBaseline and -PrivateMappingWorkbook."
    }
    & $pythonExe -m carbon_excel_pipeline standardize `
        --run-dir $StandardizeRun `
        --private-id-baseline $PrivateIdBaseline `
        --private-mapping-workbook $PrivateMappingWorkbook
    exit $LASTEXITCODE
}

if ($UpstreamRun) {
    if (-not $StandardBaseline -or -not $ActivityBaseline -or -not $ThirdPartyBaseline) {
        throw "Private Day 5 requires standard, activity and third-party baselines."
    }
    & $pythonExe -m carbon_excel_pipeline upstream `
        --run-dir $UpstreamRun `
        --standard-baseline $StandardBaseline `
        --activity-baseline $ActivityBaseline `
        --third-party-baseline $ThirdPartyBaseline
    exit $LASTEXITCODE
}

if ($FactorMatchRun) {
    if ($HistoricalSimulation) {
        & $pythonExe -m carbon_excel_pipeline factor-match `
            --run-dir $FactorMatchRun `
            --historical-simulation
    } elseif ($FactorInput) {
        & $pythonExe -m carbon_excel_pipeline factor-match `
            --run-dir $FactorMatchRun `
            --factor-input $FactorInput
    } else {
        throw "Choose -HistoricalSimulation or provide -FactorInput with -FactorMatchRun."
    }
    exit $LASTEXITCODE
}

if ($CalculationRun) {
    & $pythonExe -m carbon_excel_pipeline calculate-lineage --run-dir $CalculationRun
    exit $LASTEXITCODE
}

if ($Day8ExportRun) {
    if (-not $OutputDir -or -not $ArtifactWorkDir -or -not $NodeExecutable -or -not $NodeModules -or -not $Wp5OpenItems) {
        throw "Day 8 export requires -OutputDir, -ArtifactWorkDir, -NodeExecutable, -NodeModules and -Wp5OpenItems."
    }
    & $pythonExe -m carbon_excel_pipeline export-workbook `
        --run-dir $Day8ExportRun `
        --output-dir $OutputDir `
        --artifact-work-dir $ArtifactWorkDir `
        --node-executable $NodeExecutable `
        --node-modules $NodeModules `
        --wp5-open-items $Wp5OpenItems
    exit $LASTEXITCODE
}

if ($Day8RunAllInput) {
    $required = @(
        $RunRoot, $OutputRoot, $ArtifactWorkRoot, $NodeExecutable, $NodeModules,
        $PrivateIdBaseline, $PrivateMappingWorkbook, $StandardBaseline,
        $ActivityBaseline, $ThirdPartyBaseline, $Wp5OpenItems
    )
    if ($required -contains "" -or $required -contains $null) {
        throw "Day 8 run-all requires external run/output/work roots, Node runtime, three baselines, private ID/mapping evidence and WP5 Open Items."
    }
    & $pythonExe -m carbon_excel_pipeline run-all `
        --input $Day8RunAllInput `
        --run-root $RunRoot `
        --output-root $OutputRoot `
        --artifact-work-root $ArtifactWorkRoot `
        --node-executable $NodeExecutable `
        --node-modules $NodeModules `
        --private-id-baseline $PrivateIdBaseline `
        --private-mapping-workbook $PrivateMappingWorkbook `
        --standard-baseline $StandardBaseline `
        --activity-baseline $ActivityBaseline `
        --third-party-baseline $ThirdPartyBaseline `
        --wp5-open-items $Wp5OpenItems
    exit $LASTEXITCODE
}

throw "Choose one pipeline operation, including -Day8ExportRun or -Day8RunAllInput."
