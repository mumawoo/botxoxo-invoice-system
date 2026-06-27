$ErrorActionPreference = "Stop"

$imageExtensions = @(".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp")
$trialImages = @(Get-ChildItem -Path "data\trial" -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
    $imageExtensions -contains $_.Extension.ToLowerInvariant()
})

if ($trialImages.Count -eq 0) {
    throw "No trial photos found in data\trial. Add real receipt photos and run .\scripts\run-trial.ps1 before comparing Ubuntu output."
}

$baselineWorkbooks = @(Get-ChildItem -Path "data\baseline" -Recurse -File -Filter "*.xlsx" -ErrorAction SilentlyContinue | Where-Object {
    $name = $_.Name.ToLowerInvariant()
    -not $name.StartsWith("~$") -and
    -not $name.Contains("comparison_report") -and
    -not $name.Contains("manually_checked")
})

if ($baselineWorkbooks.Count -eq 0) {
    throw "No Ubuntu baseline workbook found in data\baseline. Add the Ubuntu Excel output workbook before comparing."
}

if (-not (Test-Path "data\output\trial\Invoice_Output_Trial.xlsx")) {
    throw "No Windows trial workbook found at data\output\trial\Invoice_Output_Trial.xlsx. Run .\scripts\run-trial.ps1 first."
}

& "$PSScriptRoot\invoice.ps1" compare --baseline data/baseline --candidate data/output/trial --output data/output/ubuntu_comparison_report.xlsx
