[CmdletBinding()]
param(
    [string]$Repository = "AlexanderWilhelmsenBerg/Outlook-aligner",
    [string]$Branch = "phase-3/ui-assisted-testing",
    [string]$Workflow = "CI",
    [string]$ArtifactName = "OutlookAligner-Phase3-UI-TestHarness-win-x64",
    [string]$InstallDirectory = (Join-Path $env:LOCALAPPDATA "OutlookAligner\TestApp"),
    [switch]$Force,
    [switch]$NoLaunch,
    [switch]$UnblockFiles
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Step {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Invoke-GhJson {
    param([Parameter(Mandatory)][string[]]$Arguments)

    $output = & gh @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "GitHub CLI failed:`n$($output -join [Environment]::NewLine)"
    }

    return ($output -join [Environment]::NewLine) | ConvertFrom-Json
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw @"
GitHub CLI (gh) is required.
Install it once with:
  winget install --id GitHub.cli
Then authenticate once with:
  gh auth login
"@
}

Write-Step "Checking GitHub authentication"
& gh auth status *> $null
if ($LASTEXITCODE -ne 0) {
    throw "GitHub CLI is installed but not authenticated. Run 'gh auth login' once and retry."
}

Write-Step "Finding latest successful UI build on '$Branch'"
$runs = Invoke-GhJson @(
    "run", "list",
    "--repo", $Repository,
    "--workflow", $Workflow,
    "--branch", $Branch,
    "--event", "push",
    "--status", "success",
    "--limit", "1",
    "--json", "databaseId,headSha,createdAt,url"
)

$run = @($runs) | Select-Object -First 1
if ($null -eq $run) {
    throw "No successful push CI run was found for branch '$Branch'."
}

$runId = [long]$run.databaseId
$headSha = [string]$run.headSha
$createdAt = [DateTimeOffset]$run.createdAt
$runUrl = [string]$run.url

$metadataPath = Join-Path $InstallDirectory ".installed-run.json"
if (-not $Force -and (Test-Path $metadataPath)) {
    try {
        $installed = Get-Content $metadataPath -Raw | ConvertFrom-Json
        if ([long]$installed.runId -eq $runId -and [string]$installed.headSha -eq $headSha) {
            Write-Host "Already current: CI run $runId ($($headSha.Substring(0, [Math]::Min(12, $headSha.Length))))." -ForegroundColor Green
            if (-not $NoLaunch) {
                $existingExe = Join-Path $InstallDirectory "OutlookAligner.App.exe"
                if (Test-Path $existingExe) {
                    Write-Step "Launching Outlook Aligner"
                    Start-Process -FilePath $existingExe
                }
            }
            return
        }
    }
    catch {
        Write-Warning "Existing install metadata could not be read; the app will be refreshed. $($_.Exception.Message)"
    }
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("OutlookAligner-update-" + [Guid]::NewGuid().ToString("N"))
$downloadDirectory = Join-Path $tempRoot "artifact"
$installParent = Split-Path -Parent $InstallDirectory
$stagingDirectory = Join-Path $installParent (".TestApp.staging-" + $runId)
$previousDirectory = "$InstallDirectory.previous"

try {
    New-Item -ItemType Directory -Path $downloadDirectory -Force | Out-Null
    New-Item -ItemType Directory -Path $installParent -Force | Out-Null

    Write-Step "Downloading artifact '$ArtifactName' from CI run $runId"
    $downloadOutput = & gh run download $runId --repo $Repository --name $ArtifactName --dir $downloadDirectory 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Artifact download failed:`n$($downloadOutput -join [Environment]::NewLine)"
    }

    $appCandidate = Get-ChildItem -Path $downloadDirectory -Filter "OutlookAligner.App.exe" -File -Recurse |
        Select-Object -First 1
    if ($null -eq $appCandidate) {
        throw "Downloaded artifact does not contain OutlookAligner.App.exe."
    }

    $artifactRoot = $appCandidate.Directory.FullName
    $hostCandidate = Join-Path $artifactRoot "OutlookAligner.OutlookHost.exe"
    if (-not (Test-Path $hostCandidate)) {
        throw "Downloaded artifact does not contain OutlookAligner.OutlookHost.exe beside the UI executable."
    }

    if (Test-Path $stagingDirectory) {
        Remove-Item $stagingDirectory -Recurse -Force
    }
    New-Item -ItemType Directory -Path $stagingDirectory -Force | Out-Null

    Write-Step "Staging verified build"
    Copy-Item -Path (Join-Path $artifactRoot "*") -Destination $stagingDirectory -Recurse -Force

    $metadata = [ordered]@{
        repository = $Repository
        branch = $Branch
        workflow = $Workflow
        artifact = $ArtifactName
        runId = $runId
        headSha = $headSha
        runCreatedAt = $createdAt.ToString("O")
        runUrl = $runUrl
        installedAt = [DateTimeOffset]::Now.ToString("O")
    }
    $metadata | ConvertTo-Json | Set-Content -Path (Join-Path $stagingDirectory ".installed-run.json") -Encoding UTF8

    $stagedApp = Join-Path $stagingDirectory "OutlookAligner.App.exe"
    $stagedHost = Join-Path $stagingDirectory "OutlookAligner.OutlookHost.exe"
    if (-not (Test-Path $stagedApp) -or -not (Test-Path $stagedHost)) {
        throw "Staging verification failed; required executables are missing."
    }

    $installedApp = Join-Path $InstallDirectory "OutlookAligner.App.exe"
    $runningApp = Get-Process -Name "OutlookAligner.App" -ErrorAction SilentlyContinue |
        Where-Object {
            try {
                [string]::Equals($_.Path, $installedApp, [StringComparison]::OrdinalIgnoreCase)
            }
            catch {
                $false
            }
        }

    if ($runningApp) {
        Write-Step "Closing the currently installed test app"
        $runningApp | Stop-Process -Force
        Start-Sleep -Milliseconds 400
    }

    if (Test-Path $previousDirectory) {
        Remove-Item $previousDirectory -Recurse -Force
    }

    if (Test-Path $InstallDirectory) {
        Write-Step "Keeping current build as '$previousDirectory'"
        Move-Item -Path $InstallDirectory -Destination $previousDirectory
    }

    try {
        Write-Step "Activating CI run $runId"
        Move-Item -Path $stagingDirectory -Destination $InstallDirectory
    }
    catch {
        if (-not (Test-Path $InstallDirectory) -and (Test-Path $previousDirectory)) {
            Move-Item -Path $previousDirectory -Destination $InstallDirectory
        }
        throw
    }

    if ($UnblockFiles) {
        Write-Step "Removing downloaded-file zone markers"
        Get-ChildItem -Path $InstallDirectory -File -Recurse | Unblock-File -ErrorAction SilentlyContinue
    }

    $shortSha = $headSha.Substring(0, [Math]::Min(12, $headSha.Length))
    Write-Host "Updated Outlook Aligner to CI run $runId ($shortSha)." -ForegroundColor Green
    Write-Host "Install: $InstallDirectory"
    Write-Host "Previous build: $previousDirectory"
    Write-Host "Persistent app log: $(Join-Path $env:LOCALAPPDATA 'OutlookAligner\events.jsonl')"
    Write-Host "CI: $runUrl"

    if (-not $NoLaunch) {
        Write-Step "Launching Outlook Aligner"
        Start-Process -FilePath (Join-Path $InstallDirectory "OutlookAligner.App.exe")
    }
}
finally {
    if (Test-Path $tempRoot) {
        Remove-Item $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path $stagingDirectory) {
        Remove-Item $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
}
