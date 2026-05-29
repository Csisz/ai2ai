Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir
Push-Location $repoRoot
try {
    python -B ai_debate.py --regression-test
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}
