$ErrorActionPreference = "Stop"

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Command
    )

    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code $LASTEXITCODE."
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
Push-Location $repoRoot
try {
    Invoke-Checked { dotnet restore OutlookAligner.sln }
    Invoke-Checked { dotnet format OutlookAligner.sln --verify-no-changes --no-restore }
    Invoke-Checked { dotnet build OutlookAligner.sln -c Release --no-restore }
    Invoke-Checked { dotnet test OutlookAligner.sln -c Release --no-build --no-restore --coverage --coverage-output-format cobertura }

    Push-Location "web/calendar"
    try {
        Invoke-Checked { npm install --ignore-scripts }
        Invoke-Checked { npm run typecheck }
        Invoke-Checked { npm run lint }
        Invoke-Checked { npm run format:check }
        Invoke-Checked { npm run build }
    }
    finally {
        Pop-Location
    }
}
finally {
    Pop-Location
}
