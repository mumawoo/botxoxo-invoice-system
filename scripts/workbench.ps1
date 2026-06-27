param()

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$EnvFile = Join-Path $ProjectRoot ".env"

function Get-DotEnvValue {
    param([string]$Name)
    if (-not (Test-Path -LiteralPath $EnvFile)) {
        return ""
    }
    foreach ($line in Get-Content -LiteralPath $EnvFile -Encoding UTF8) {
        if ($line -match "^\s*$([regex]::Escape($Name))\s*=\s*(.*)\s*$") {
            return $Matches[1].Trim().Trim('"').Trim("'")
        }
    }
    return ""
}

function Get-DefaultPython {
    if ($env:INVOICE_SYSTEM_PYTHON -and (Test-Path -LiteralPath $env:INVOICE_SYSTEM_PYTHON)) {
        return $env:INVOICE_SYSTEM_PYTHON
    }
    $codexPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    if (Test-Path -LiteralPath $codexPython) {
        return $codexPython
    }
    return "python"
}

function ConvertTo-CommandLiteral {
    param([string]$Value)
    return "'" + ($Value -replace "'", "''") + "'"
}

$Python = Get-DefaultPython
$ReimbursementWorkbookName = "$([char]0x62A5)$([char]0x9500)$([char]0x660E)$([char]0x7EC6)_2026_xlsx.xlsx"
$CheckedWorkbookName = "$([char]0x62A5)$([char]0x9500)_checked_2026.xlsx"
$DefaultUserIds = Get-DotEnvValue "TELEGRAM_ALLOWED_USER_IDS"
$DefaultUserId = ""
if ($DefaultUserIds -match "(\d{5,})") {
    $DefaultUserId = $Matches[1]
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Invoice System Workbench"
$Form.StartPosition = "CenterScreen"
$Form.Size = New-Object System.Drawing.Size(880, 790)
$Form.MinimumSize = New-Object System.Drawing.Size(760, 600)
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$Title = New-Object System.Windows.Forms.Label
$Title.Text = "Invoice System Workbench"
$Title.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$Title.AutoSize = $true
$Title.Location = New-Object System.Drawing.Point(18, 16)
$Form.Controls.Add($Title)

$ProjectLabel = New-Object System.Windows.Forms.Label
$ProjectLabel.Text = "Project: $ProjectRoot"
$ProjectLabel.AutoSize = $true
$ProjectLabel.Location = New-Object System.Drawing.Point(20, 52)
$Form.Controls.Add($ProjectLabel)

$UserLabel = New-Object System.Windows.Forms.Label
$UserLabel.Text = "Telegram user ID:"
$UserLabel.AutoSize = $true
$UserLabel.Location = New-Object System.Drawing.Point(20, 86)
$Form.Controls.Add($UserLabel)

$UserBox = New-Object System.Windows.Forms.TextBox
$UserBox.Text = $DefaultUserId
$UserBox.Location = New-Object System.Drawing.Point(150, 82)
$UserBox.Size = New-Object System.Drawing.Size(180, 28)
$Form.Controls.Add($UserBox)

$Status = New-Object System.Windows.Forms.Label
$Status.Text = "Ready. Buttons open commands in separate PowerShell windows."
$Status.AutoSize = $false
$Status.Location = New-Object System.Drawing.Point(20, 715)
$Status.Size = New-Object System.Drawing.Size(820, 40)
$Status.ForeColor = [System.Drawing.Color]::DarkSlateGray
$Form.Controls.Add($Status)

function Get-UserId {
    $value = $UserBox.Text.Trim()
    if (-not ($value -match "^\d+$")) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a numeric Telegram user ID.", "Missing user ID", "OK", "Warning") | Out-Null
        return $null
    }
    return $value
}

function Start-InvoiceCommand {
    param(
        [string]$Title,
        [string[]]$InvoiceArgs
    )
    $rootLiteral = ConvertTo-CommandLiteral $ProjectRoot
    $pythonLiteral = ConvertTo-CommandLiteral $Python
    $argsText = ($InvoiceArgs | ForEach-Object { ConvertTo-CommandLiteral $_ }) -join " "
    $command = "Set-Location -LiteralPath $rootLiteral; & $pythonLiteral -m invoice_system $argsText"
    Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-Command", $command) -WindowStyle Normal
    $Status.Text = "Started: $Title"
}

function Assert-TelegramDependency {
    try {
        & $Python -c "import telegram; raise SystemExit(0)" *> $null
        if ($LASTEXITCODE -eq 0) {
            return $true
        }
    } catch {
    }
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Telegram dependency is missing in this Python runtime.`n`nInstall python-telegram-bot now?",
        "Telegram dependency missing",
        "YesNo",
        "Question"
    )
    if ($answer -ne "Yes") {
        $Status.Text = "Telegram dependency missing. Install python-telegram-bot."
        return $false
    }
    $installCommand = "& " + (ConvertTo-CommandLiteral $Python) + " -m pip install `"python-telegram-bot>=22.0`"; if (`$LASTEXITCODE -eq 0) { Write-Host 'Telegram dependency installed. You can close this window.' } else { Write-Host 'Install failed. Keep this window open and check the error.' }"
    $process = Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-Command", $installCommand) -WindowStyle Normal -PassThru
    $process.WaitForExit()
    try {
        & $Python -c "import telegram; raise SystemExit(0)" *> $null
        if ($LASTEXITCODE -eq 0) {
            $Status.Text = "Telegram dependency installed."
            return $true
        }
    } catch {
    }
    [System.Windows.Forms.MessageBox]::Show("Telegram dependency install did not complete. Check the install window.", "Install failed", "OK", "Warning") | Out-Null
    $Status.Text = "Telegram dependency install failed."
    return $false
}

function Open-Path {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Invoke-Item -LiteralPath $Path
        $Status.Text = "Opened: $Path"
    } else {
        [System.Windows.Forms.MessageBox]::Show("Path does not exist yet:`n$Path", "Not found", "OK", "Information") | Out-Null
    }
}

function Add-Section {
    param([string]$Text, [int]$X, [int]$Y)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $Form.Controls.Add($label)
}

function Add-Button {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [scriptblock]$OnClick,
        [string]$ToolTip = ""
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size(255, 38)
    $button.Add_Click($OnClick)
    $Form.Controls.Add($button)
    if ($ToolTip) {
        $tip = New-Object System.Windows.Forms.ToolTip
        $tip.SetToolTip($button, $ToolTip)
    }
}

Add-Section "Daily Telegram" 20 130
Add-Button "Start Telegram Bot - Auto Scan" 20 160 {
    if (Assert-TelegramDependency) {
        Start-InvoiceCommand "Telegram bot auto scan" @("telegram", "--process")
    }
} "Polling bot. Saves photos and scans automatically."
Add-Button "Start Telegram Bot - Save Only" 300 160 {
    if (Assert-TelegramDependency) {
        Start-InvoiceCommand "Telegram bot save only" @("telegram", "--no-process")
    }
} "Polling bot. Saves photos but does not scan."
Add-Button "Telegram Config Check" 580 160 {
    Start-InvoiceCommand "Telegram config check" @("telegram", "--check")
}

Add-Button "Queue Status" 20 210 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Queue status" @("worker", "--user-id", $id) }
}
Add-Button "Process Pending Once" 300 210 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Process pending once" @("worker", "--user-id", $id, "--once") }
}
Add-Button "Retry Failed + Process" 580 210 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Retry failed queue" @("worker", "--user-id", $id, "--retry-failed") }
}

Add-Button "Reset Active Batch" 20 255 {
    $id = Get-UserId
    if ($id) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Archive current Telegram photos, queue, Excel, crops, and start fresh from 001?",
            "Confirm reset",
            "YesNo",
            "Warning"
        )
        if ($answer -eq "Yes") {
            Start-InvoiceCommand "Reset active batch" @("worker", "--user-id", $id, "--reset-active")
        }
    }
} "Moves active files into reset_archive and clears the live queue."

Add-Section "Finance / Reimbursement" 20 310
Add-Button "Report / Unsubmitted" 20 340 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Reimbursement report" @("reimburse", "report", "--user-id", $id) }
}
Add-Button "Submit Reimbursement" 300 340 {
    $id = Get-UserId
    if ($id) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Close the active Excel file first if it is open.`n`nSubmit all unsubmitted records, archive the active Excel, and start a new Excel batch?",
            "Confirm submit",
            "YesNo",
            "Question"
        )
        if ($answer -eq "Yes") {
            Start-InvoiceCommand "Submit reimbursement" @("reimburse", "submit", "--user-id", $id)
        }
    }
}
Add-Button "Submitted Batches" 580 340 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Submitted batches" @("reimburse", "submitted", "--user-id", $id) }
}

Add-Button "Open Manual Excel" 20 390 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id\$ReimbursementWorkbookName") }
}
Add-Button "Open Checked Excel" 300 390 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id\$CheckedWorkbookName") }
}
Add-Button "Open User Output Folder" 580 390 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id") }
}
Add-Button "Open Telegram Inbound" 580 435 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\inbound\telegram\$id") }
}

Add-Section "Maintenance / QA" 20 485
Add-Button "System Check" 20 515 {
    Start-InvoiceCommand "System check" @("check")
}
Add-Button "Audit Requirements" 300 515 {
    Start-InvoiceCommand "Audit" @("audit")
}
Add-Button "Compare Ubuntu Baseline" 580 515 {
    $id = Get-UserId
    if ($id) {
        Start-InvoiceCommand "Compare Ubuntu baseline" @(
            "compare",
            "--baseline", "data\baseline",
            "--candidate", "data\output\telegram\$id",
            "--output", "data\output\telegram\$id\comparison_report.xlsx"
        )
    }
}

Add-Button "Run Trial" 20 565 {
    Start-InvoiceCommand "Trial run" @("run", "--trial", "--resume")
}
Add-Button "A/B Test Latest Telegram" 300 565 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "A/B test latest Telegram" @("ab-test", "--user-id", $id) }
}
Add-Button "Open Project Folder" 580 565 {
    Open-Path $ProjectRoot
}

Add-Button "Open .env" 20 615 {
    Open-Path $EnvFile
}
Add-Button "Open Raw Crops" 300 615 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id\crops") }
}
Add-Button "Open Review Crops" 580 615 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id\review_crops") }
}
Add-Button "Open Final Crops" 300 660 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id\final_crops") }
}

[void]$Form.ShowDialog()
