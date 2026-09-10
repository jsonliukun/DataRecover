[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Invoke-CheckedCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList
    )

    & $FilePath @ArgumentList
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code $LASTEXITCODE`: $FilePath $($ArgumentList -join ' ')"
    }
}

# $PSScriptRoot points to this file's directory, so callers do not need to
# change into the repository before running the script.
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$venvPath = Join-Path $projectRoot ".venv"
$venvPython = Join-Path $venvPath "Scripts\python.exe"

Push-Location -LiteralPath $projectRoot
try {
    if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
        $pyLauncher = Get-Command "py" -ErrorAction SilentlyContinue
        if ($null -ne $pyLauncher) {
            Invoke-CheckedCommand -FilePath $pyLauncher.Source -ArgumentList @("-3", "-m", "venv", $venvPath)
        }
        else {
            $pythonCommand = Get-Command "python" -ErrorAction Stop
            Invoke-CheckedCommand -FilePath $pythonCommand.Source -ArgumentList @("-m", "venv", $venvPath)
        }
    }

    if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
        throw "Virtual environment Python was not created at: $venvPython"
    }

    Invoke-CheckedCommand -FilePath $venvPython -ArgumentList @("-m", "pip", "install", "--upgrade", "pip")
    Invoke-CheckedCommand -FilePath $venvPython -ArgumentList @("-m", "pip", "install", ".")
    Invoke-CheckedCommand -FilePath $venvPython -ArgumentList @("-m", "pip", "check")
    Invoke-CheckedCommand -FilePath $venvPython -ArgumentList @(
        "-c",
        "import httpx, dotenv, pypdf, doc_renamer; print('Dependency imports OK')"
    )
    Invoke-CheckedCommand -FilePath $venvPython -ArgumentList @(
        "-m", "unittest", "discover", "-s", "tests", "-v"
    )
}
finally {
    Pop-Location
}
