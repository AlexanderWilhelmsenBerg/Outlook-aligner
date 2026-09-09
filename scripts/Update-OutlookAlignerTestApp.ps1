[CmdletBinding()]
param(
    [string]$Repository = "AlexanderWilhelmsenBerg/Outlook-aligner",
    [string]$Branch = "phase-3/ui-assisted-testing",
    [string]$Workflow = "CI",
    [string]$ArtifactName = "OutlookAligner-Phase3-UI-TestHarness-win-x64",
    [string]$InstallDirectory = (Join-Path $env:LOCALAPPDATA "OutlookAligner\TestApp"),
    [switch]$Force,
    [switch]$NoLaunch,
    [switch]$UnblockFiles,
    [switch]$PauseOnError
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:GhPath = $null

function Write-Step {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Get-GitHubCliArchitecture {
    $architecture = [string]$env:PROCESSOR_ARCHITECTURE
    if ($architecture -eq "AMD64") {
        return "amd64"
    }

    if ($architecture -eq "ARM64") {
        return "arm64"
    }

    if ($architecture -eq "x86") {
        return "386"
    }

    throw "Unsupported Windows processor architecture '$architecture'."
}

function Install-PortableGitHubCli {
    if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        throw "LOCALAPPDATA is not available in this Windows session."
    }

    $toolsRoot = Join-Path $env:LOCALAPPDATA "OutlookAligner\Tools\GitHubCLI"
    $portableGh = Join-Path $toolsRoot "gh.exe"
    if (Test-Path $portableGh) {
        return $portableGh
    }

    Write-Step "GitHub CLI is not installed system-wide; preparing a portable user copy"

    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("OutlookAligner-gh-" + [Guid]::NewGuid().ToString("N"))
    try {
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $toolsRoot -Force | Out-Null

        # Windows PowerShell 5.1 may otherwise negotiate an obsolete TLS version.
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        }
        catch {
            # PowerShell 7 does not need this and may ignore the legacy ServicePointManager path.
        }

        $headers = @{
            "User-Agent" = "OutlookAligner-TestUpdater"
            "Accept" = "application/vnd.github+json"
        }

        Write-Step "Finding the latest official GitHub CLI portable release"
        $release = Invoke-RestMethod `
            -Uri "https://api.github.com/repos/cli/cli/releases/latest" `
            -Headers $headers `
            -Method Get

        $architecture = Get-GitHubCliArchitecture
        $assetPattern = "^gh_[0-9.]+_windows_$architecture\.zip$"
        $asset = @($release.assets) |
            Where-Object { [string]$_.name -match $assetPattern } |
            Select-Object -First 1

        if ($null -eq $asset) {
            throw "The latest GitHub CLI release did not contain the expected Windows $architecture ZIP asset."
        }

        $zipPath = Join-Path $tempRoot ([string]$asset.name)
        $extractPath = Join-Path $tempRoot "extracted"

        Write-Step "Downloading portable GitHub CLI $($release.tag_name)"
        Invoke-WebRequest `
            -Uri ([string]$asset.browser_download_url) `
            -Headers @{ "User-Agent" = "OutlookAligner-TestUpdater" } `
            -OutFile $zipPath

        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        $downloadedGh = Get-ChildItem -Path $extractPath -Filter "gh.exe" -File -Recurse |
            Select-Object -First 1

        if ($null -eq $downloadedGh) {
            throw "Portable GitHub CLI archive did not contain gh.exe."
        }

        Copy-Item -Path $downloadedGh.FullName -Destination $portableGh -Force
        Unblock-File -Path $portableGh -ErrorAction SilentlyContinue

        if (-not (Test-Path $portableGh)) {
            throw "Portable GitHub CLI could not be installed in the current user profile."
        }

        Write-Host "Portable GitHub CLI: $portableGh" -ForegroundColor Green
        return $portableGh
    }
    finally {
        if (Test-Path $tempRoot) {
            Remove-Item $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Resolve-GitHubCli {
    $systemGh = Get-Command gh -ErrorAction SilentlyContinue
    if ($null -ne $systemGh) {
        return $systemGh.Source
    }

    return Install-PortableGitHubCli
}

function Invoke-GhJson {
    param([Parameter(Mandatory)][string[]]$Arguments)

    $output = & $script:GhPath @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "GitHub CLI failed:`n$($output -join [Environment]::NewLine)"
    }

    return ($output -join [Environment]::NewLine) | ConvertFrom-Json
}

function Ensure-GitHubAuthentication {
    Write-Step "Checking GitHub authentication"
    $authOutput = & $script:GhPath auth status 2>&1
    if ($LASTEXITCODE -eq 0) {
        return
    }

    Write-Host "GitHub authentication is required once for Actions artifact downloads." -ForegroundColor Yellow
    Write-Host "A browser window will open for GitHub login. No administrator rights are required." -ForegroundColor Yellow

    & $script:GhPath auth login --hostname github.com --git-protocol https --web
    if ($LASTEXITCODE -ne 0) {
        throw @"
GitHub authentication did not complete successfully.

Portable GitHub CLI is located at:
  $script:GhPath

You can retry manually with:
  & "$script:GhPath" auth login --hostname github.com --git-protocol https --web
"@
    }

    $authOutput = & $script:GhPath auth status 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "GitHub CLI login completed but authentication still could not be verified.`n$($authOutput -join [Environment]::NewLine)"
    }
}

function Invoke-OutlookAlignerUpdate {
    if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        throw "LOCALAPPDATA is not available in this Windows session."
    }

    $script:GhPath = Resolve-GitHubCli
    Write-Host "Using GitHub CLI: $script:GhPath"
    Ensure-GitHubAuthentication

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
    $existingApp = Join-Path $InstallDirectory "OutlookAligner.App.exe"
    $existingHost = Join-Path $InstallDirectory "OutlookAligner.OutlookHost.exe"
    if (-not $Force -and (Test-Path $metadataPath)) {
        try {
            $installed = Get-Content $metadataPath -Raw | ConvertFrom-Json
            $sameBuild = [long]$installed.runId -eq $runId -and [string]$installed.headSha -eq $headSha
            $installComplete = (Test-Path $existingApp) -and (Test-Path $existingHost)
            if ($sameBuild -and $installComplete) {
                Write-Host "Already current: CI run $runId ($($headSha.Substring(0, [Math]::Min(12, $headSha.Length))))." -ForegroundColor Green
                if (-not $NoLaunch) {
                    Write-Step "Launching Outlook Aligner"
                    Start-Process -FilePath $existingApp
                }
                return
            }

            if ($sameBuild -and -not $installComplete) {
                Write-Warning "The installed run metadata is current, but one or more required executables are missing. Reinstalling the build."
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
        $downloadOutput = & $script:GhPath run download $runId --repo $Repository --name $ArtifactName --dir $downloadDirectory 2>&1
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
}

try {
    Invoke-OutlookAlignerUpdate
    exit 0
}
catch {
    Write-Host ""
    Write-Host "OUTLOOK ALIGNER UPDATE FAILED" -ForegroundColor Red
    Write-Host "-----------------------------" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "If this is a managed work computer and company policy blocks downloaded executables, stop here rather than bypassing that policy." -ForegroundColor Yellow
    Write-Host "Otherwise, send me the text above and I can diagnose the updater directly." -ForegroundColor Yellow

    if ($PauseOnError) {
        Write-Host ""
        [void](Read-Host "Press Enter to close")
    }

    exit 1
}
