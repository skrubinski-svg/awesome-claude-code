# Wheelworx FSC report launcher. First run creates .venv and installs requirements.
# Extra arguments are passed through, e.g.:  .\run_fsc_report.ps1 --csv prices.csv
$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot
$venv = Join-Path $PSScriptRoot '.venv'
$activate = Join-Path $venv 'Scripts\Activate.ps1'
$code = 0
try {
    if (-not (Test-Path $activate)) {
        Write-Host 'First run: creating .venv and installing requirements...'
        if (Get-Command py -ErrorAction SilentlyContinue) { py -3 -m venv $venv } else { python -m venv $venv }
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $activate)) { throw 'Could not create .venv. Is Python 3.11+ installed and on PATH?' }
        . $activate
        python -m pip install -r requirements.txt
        if ($LASTEXITCODE -ne 0) {
            deactivate
            Remove-Item -Recurse -Force $venv
            throw 'pip install -r requirements.txt failed (see above). .venv removed so the next run retries.'
        }
    } else {
        . $activate
    }
    python fsc_report.py --open @args
    $code = $LASTEXITCODE
} catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    $code = 1
}
if ($code -ne 0) { Write-Host "FSC report did not complete (exit code $code). See the message above." -ForegroundColor Red }
exit $code
