$ErrorActionPreference = "Stop"
& "$PSScriptRoot\invoice.ps1" prepare
& "$PSScriptRoot\invoice.ps1" run --trial @args
