param(
    [switch] $Multi
)

$ErrorActionPreference = "Stop"

if ($Multi) {
    & "$PSScriptRoot\invoice.ps1" sample --multi
    & "$PSScriptRoot\invoice.ps1" run --trial --input data\samples\synthetic_receipts_multi.jpg --output data\output\actual_ocr_smoke_multi
}
else {
    & "$PSScriptRoot\invoice.ps1" sample
    & "$PSScriptRoot\invoice.ps1" run --trial --input data\samples\synthetic_receipt.jpg --output data\output\actual_ocr_smoke
}
