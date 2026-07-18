param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $InvoiceArgs
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

function Find-InvoicePython {
    $candidates = @()
    if ($env:INVOICE_SYSTEM_PYTHON) {
        $candidates += $env:INVOICE_SYSTEM_PYTHON
    }
    $candidates += Join-Path $repoRoot ".venv\Scripts\python.exe"
    if ($env:USERPROFILE) {
        $candidates += Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    }
    $candidates += "python"

    foreach ($candidate in $candidates) {
        try {
            & $candidate -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)" *> $null
            if ($LASTEXITCODE -eq 0) {
                return $candidate
            }
        }
        catch {
        }
    }

    throw "No usable Python 3.11+ runtime found. Set INVOICE_SYSTEM_PYTHON to python.exe."
}

$python = Find-InvoicePython
& $python -m invoice_system @InvoiceArgs
exit $LASTEXITCODE
