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
$Form.Size = New-Object System.Drawing.Size(860, 620)
$Form.MinimumSize = New-Object System.Drawing.Size(760, 560)
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$Form.AutoScroll = $true
$Form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)

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
$Status.Location = New-Object System.Drawing.Point(20, 535)
$Status.Size = New-Object System.Drawing.Size(800, 42)
$Status.ForeColor = [System.Drawing.Color]::DarkSlateGray
$Status.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 247)
$Status.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 6)
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
        [string[]]$InvoiceArgs,
        [switch]$RestartExisting
    )
    if ($RestartExisting) {
        if (-not (Confirm-AndStopExistingInvoiceProcesses -ActionTitle $Title)) {
            return
        }
    }
    $rootLiteral = ConvertTo-CommandLiteral $ProjectRoot
    $pythonLiteral = ConvertTo-CommandLiteral $Python
    $argsText = ($InvoiceArgs | ForEach-Object { ConvertTo-CommandLiteral $_ }) -join " "
    $command = "Set-Location -LiteralPath $rootLiteral; & $pythonLiteral -m invoice_system $argsText"
    Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-Command", $command) -WindowStyle Normal
    $Status.Text = "Started: $Title"
}

function Get-MissingPythonModules {
    param([string[]]$Modules)
    $moduleCsv = $Modules -join ","
    $code = "import importlib.util; names='$moduleCsv'.split(','); print(','.join(n for n in names if importlib.util.find_spec(n) is None))"
    try {
        $output = (& $Python -c $code 2>$null | Out-String).Trim()
        if ($LASTEXITCODE -ne 0) {
            return @($Modules)
        }
    } catch {
        return @($Modules)
    }
    if (-not $output) {
        return @()
    }
    return @($output -split "," | Where-Object { $_ })
}

function Start-DependencyInstall {
    param([string[]]$MissingModules)
    $packageMap = @{
        "telegram" = "python-telegram-bot>=22.0"
        "cv2" = "opencv-python>=4.9.0"
        "PIL" = "Pillow>=10.0.0"
        "numpy" = "numpy>=1.24,<2.4"
        "openpyxl" = "openpyxl>=3.1.0"
    }
    $packages = @($MissingModules | ForEach-Object { $packageMap[$_] } | Where-Object { $_ } | Select-Object -Unique)
    $missingText = $MissingModules -join ", "
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Required modules are missing from the Python used by Workbench:`n$missingText`n`nInstall them now?",
        "Invoice runtime incomplete",
        "YesNo",
        "Question"
    )
    if ($answer -ne "Yes") {
        $Status.Text = "Cannot start. Missing: $missingText"
        return $false
    }
    $pythonLiteral = ConvertTo-CommandLiteral $Python
    $packageText = ($packages | ForEach-Object { ConvertTo-CommandLiteral $_ }) -join " "
    $installCommand = "& $pythonLiteral -m pip install $packageText; if (`$LASTEXITCODE -eq 0) { Write-Host 'Invoice runtime dependencies installed.' -ForegroundColor Green; Start-Sleep -Seconds 3; exit 0 } else { Write-Host 'Install failed. Keep this window open and check the error.' -ForegroundColor Red }"
    Start-Process -FilePath "powershell.exe" -ArgumentList @("-NoExit", "-ExecutionPolicy", "Bypass", "-Command", $installCommand) -WindowStyle Normal | Out-Null
    $Status.Text = "Installing missing modules: $missingText. Click Start again after the install window closes."
    return $false
}

function Assert-TelegramDependency {
    $missing = @(Get-MissingPythonModules -Modules @("telegram"))
    if ($missing.Count -eq 0) {
        return $true
    }
    return Start-DependencyInstall -MissingModules $missing
}

function Assert-ProcessingDependencies {
    $missing = @(Get-MissingPythonModules -Modules @("telegram", "cv2", "PIL", "numpy", "openpyxl"))
    if ($missing.Count -eq 0) {
        return $true
    }
    return Start-DependencyInstall -MissingModules $missing
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

function Limit-Text {
    param([string]$Text, [int]$MaxLength = 140)
    if ($Text.Length -le $MaxLength) {
        return $Text
    }
    return $Text.Substring(0, $MaxLength - 3) + "..."
}

function Get-InvoiceSystemProcesses {
    $matches = @()
    try {
        $processes = Get-CimInstance Win32_Process -Filter "Name = 'python.exe' OR Name = 'pythonw.exe' OR Name = 'py.exe' OR Name = 'powershell.exe' OR Name = 'pwsh.exe' OR Name = 'cmd.exe'"
    } catch {
        return @()
    }
    $rootLower = $ProjectRoot.ToLowerInvariant()
    foreach ($process in $processes) {
        $commandLine = [string]$process.CommandLine
        if ([string]::IsNullOrWhiteSpace($commandLine)) {
            continue
        }
        $lowerCommand = $commandLine.ToLowerInvariant()
        $isInvoiceCommand = $lowerCommand.Contains("-m invoice_system")
        $isProjectWrapper = $lowerCommand.Contains($rootLower) -and $lowerCommand.Contains("invoice_system")
        if ($isInvoiceCommand -or $isProjectWrapper) {
            $matches += $process
        }
    }
    return @($matches)
}

function Confirm-AndStopExistingInvoiceProcesses {
    param([string]$ActionTitle)
    $processes = @(Get-InvoiceSystemProcesses | Sort-Object ProcessId)
    if ($processes.Count -eq 0) {
        return $true
    }

    $processList = ($processes | ForEach-Object {
        "PID $($_.ProcessId): $(Limit-Text -Text ([string]$_.CommandLine))"
    }) -join "`n"
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Existing invoice_system bot/window process(es) are running.`n`n$processList`n`nStop them, then start: $ActionTitle?",
        "Restart invoice_system",
        "YesNo",
        "Warning"
    )
    if ($answer -ne "Yes") {
        $Status.Text = "Start canceled. Existing PID(s) were left running."
        return $false
    }

    $stopped = 0
    $failed = @()
    foreach ($process in $processes) {
        try {
            Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop
            $stopped += 1
        } catch {
            $failed += "PID $($process.ProcessId)"
        }
    }
    Start-Sleep -Milliseconds 800
    if ($failed.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Stopped $stopped process(es), but failed to stop: $($failed -join ', '). New start canceled.",
            "Could not stop all PIDs",
            "OK",
            "Warning"
        ) | Out-Null
        $Status.Text = "Restart canceled; some old PID(s) could not be stopped."
        return $false
    }

    $Status.Text = "Stopped $stopped old invoice_system bot/window process(es). Starting new process..."
    return $true
}

function Stop-InvoiceSystemProcessesAndClose {
    $processes = @(Get-InvoiceSystemProcesses | Sort-Object ProcessId)
    if ($processes.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No invoice_system bot/window process is running.`n`nThe workbench will close now.",
            "No running bot PID",
            "OK",
            "Information"
        ) | Out-Null
        $Form.Close()
        return
    }

    $processList = ($processes | ForEach-Object {
        "PID $($_.ProcessId): $(Limit-Text -Text ([string]$_.CommandLine))"
    }) -join "`n"
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Stop these invoice_system bot/window processes and close this workbench?`n`n$processList",
        "Confirm stop PID",
        "YesNo",
        "Warning"
    )
    if ($answer -ne "Yes") {
        $Status.Text = "Stop canceled."
        return
    }

    $stopped = 0
    $failed = @()
    foreach ($process in $processes) {
        try {
            Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop
            $stopped += 1
        } catch {
            $failed += "PID $($process.ProcessId)"
        }
    }
    if ($failed.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Stopped $stopped process(es), but failed to stop: $($failed -join ', ')",
            "Partial stop",
            "OK",
            "Warning"
        ) | Out-Null
        $Status.Text = "Stopped $stopped process(es); some PIDs failed."
        return
    }

    $Status.Text = "Stopped $stopped invoice_system bot/window process(es). Closing workbench."
    $Form.Close()
}

function Show-InvoiceSystemProcesses {
    $processes = @(Get-InvoiceSystemProcesses | Sort-Object ProcessId)
    if ($processes.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No invoice_system bot/window process is running.", "Running PIDs", "OK", "Information") | Out-Null
        $Status.Text = "No invoice_system bot/window process is running."
        return
    }
    $processList = ($processes | ForEach-Object {
        "PID $($_.ProcessId): $(Limit-Text -Text ([string]$_.CommandLine))"
    }) -join "`n"
    [System.Windows.Forms.MessageBox]::Show($processList, "Running invoice_system PIDs", "OK", "Information") | Out-Null
    $Status.Text = "Found $($processes.Count) invoice_system process(es)."
}

function New-WorkbenchTab {
    param([string]$Text)
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $Text
    $tab.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)
    $Tabs.TabPages.Add($tab) | Out-Null
    return $tab
}

function Add-Section {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [System.Windows.Forms.Control]$Parent = $Form
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
    $Parent.Controls.Add($label)
}

function Add-Button {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [scriptblock]$OnClick,
        [string]$ToolTip = "",
        [System.Windows.Forms.Control]$Parent = $Form,
        [int]$Width = 235,
        [string]$Kind = "Neutral"
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, 42)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 1
    $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(210, 216, 225)
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    if ($Kind -eq "Primary") {
        $button.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
        $button.ForeColor = [System.Drawing.Color]::White
        $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
    } elseif ($Kind -eq "Success") {
        $button.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
        $button.ForeColor = [System.Drawing.Color]::White
        $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
    } elseif ($Kind -eq "Danger") {
        $button.BackColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
        $button.ForeColor = [System.Drawing.Color]::White
        $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(185, 28, 28)
    } elseif ($Kind -eq "Warning") {
        $button.BackColor = [System.Drawing.Color]::FromArgb(245, 158, 11)
        $button.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
        $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(245, 158, 11)
    } else {
        $button.BackColor = [System.Drawing.Color]::White
        $button.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
    }
    $button.Add_Click($OnClick)
    $Parent.Controls.Add($button)
    if ($ToolTip) {
        $tip = New-Object System.Windows.Forms.ToolTip
        $tip.SetToolTip($button, $ToolTip)
    }
}

$Tabs = New-Object System.Windows.Forms.TabControl
$Tabs.Location = New-Object System.Drawing.Point(20, 130)
$Tabs.Size = New-Object System.Drawing.Size(800, 380)
$Tabs.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$Form.Controls.Add($Tabs)

$DailyTab = New-WorkbenchTab "Daily"
$FilesTab = New-WorkbenchTab "Files"
$FinanceTab = New-WorkbenchTab "Finance"
$AdvancedTab = New-WorkbenchTab "Advanced"

Add-Section "Telegram bot" 20 20 $DailyTab
Add-Button "Start / Restart Auto Scan" 20 55 {
    if (Assert-ProcessingDependencies) {
        Start-InvoiceCommand -Title "Telegram bot auto scan" -InvoiceArgs @("telegram", "--process") -RestartExisting
    }
} "Stops old invoice_system PIDs if any, then starts polling with auto scan." $DailyTab 235 "Primary"
Add-Button "Stop Bot PIDs + Close" 280 55 {
    Stop-InvoiceSystemProcessesAndClose
} "Stops running invoice_system Python processes, then closes this panel." $DailyTab 235 "Danger"
Add-Button "Queue Status" 540 55 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Queue status" @("worker", "--user-id", $id) }
} "" $DailyTab

Add-Button "Show Running PIDs" 20 110 {
    Show-InvoiceSystemProcesses
} "Shows running invoice_system Python PIDs." $DailyTab
Add-Button "Process Pending Once" 280 110 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Process pending once" @("worker", "--user-id", $id, "--once") }
} "" $DailyTab
Add-Button "Retry Failed + Process" 540 110 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Retry failed queue" @("worker", "--user-id", $id, "--retry-failed") }
} "" $DailyTab

Add-Button "Start / Restart Save Only" 20 165 {
    if (Assert-TelegramDependency) {
        Start-InvoiceCommand -Title "Telegram bot save only" -InvoiceArgs @("telegram", "--no-process") -RestartExisting
    }
} "Stops old invoice_system PIDs if any, then starts polling without scanning." $DailyTab
Add-Button "Telegram Config Check" 280 165 {
    Start-InvoiceCommand "Telegram config check" @("telegram", "--check")
} "" $DailyTab
Add-Button "Reset Active Batch" 540 165 {
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
} "Moves active files into reset_archive and clears the live queue." $DailyTab 235 "Warning"

Add-Section "Manual check files" 20 20 $FilesTab
Add-Button "Open Manual Excel" 20 55 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id\$ReimbursementWorkbookName") }
} "" $FilesTab 235 "Primary"
Add-Button "Open Crops" 280 55 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id\crops") }
} "Trace ID crop images used by the manual Excel." $FilesTab
Add-Button "Open Telegram Inbound" 540 55 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\inbound\telegram\$id") }
} "" $FilesTab
Add-Button "Open User Output Folder" 20 110 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id") }
} "" $FilesTab
Add-Button "Open .env" 280 110 {
    Open-Path $EnvFile
} "" $FilesTab
Add-Button "Open Project Folder" 540 110 {
    Open-Path $ProjectRoot
} "" $FilesTab

Add-Section "Finance export" 20 20 $FinanceTab
Add-Button "Report / Unsubmitted" 20 55 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Reimbursement report" @("reimburse", "report", "--user-id", $id) }
} "" $FinanceTab
Add-Button "Build Checked + Final Crops" 280 55 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Build checked Excel and final_crops" @("rebuild-final-crops", "--user-id", $id) }
} "Generate the finance checked Excel and final_crops on demand." $FinanceTab 235 "Primary"
Add-Button "Open Checked Excel" 540 55 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id\$CheckedWorkbookName") }
} "" $FinanceTab
Add-Button "Open Final Crops" 20 110 {
    $id = Get-UserId
    if ($id) { Open-Path (Join-Path $ProjectRoot "data\output\telegram\$id\final_crops") }
} "" $FinanceTab
Add-Button "Submit Reimbursement" 280 110 {
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
} "" $FinanceTab 235 "Success"
Add-Button "Submitted Batches" 540 110 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "Submitted batches" @("reimburse", "submitted", "--user-id", $id) }
} "" $FinanceTab

Add-Section "Maintenance / QA" 20 20 $AdvancedTab
Add-Button "System Check" 20 55 {
    Start-InvoiceCommand "System check" @("check")
} "" $AdvancedTab
Add-Button "Audit Requirements" 280 55 {
    Start-InvoiceCommand "Audit" @("audit")
} "" $AdvancedTab
Add-Button "Compare Ubuntu Baseline" 540 55 {
    $id = Get-UserId
    if ($id) {
        Start-InvoiceCommand "Compare Ubuntu baseline" @(
            "compare",
            "--baseline", "data\baseline",
            "--candidate", "data\output\telegram\$id",
            "--output", "data\output\telegram\$id\comparison_report.xlsx"
        )
    }
} "" $AdvancedTab
Add-Button "Run Trial" 20 110 {
    Start-InvoiceCommand "Trial run" @("run", "--trial", "--resume")
} "" $AdvancedTab
Add-Button "A/B Test Latest Telegram" 280 110 {
    $id = Get-UserId
    if ($id) { Start-InvoiceCommand "A/B test latest Telegram" @("ab-test", "--user-id", $id) }
} "" $AdvancedTab

[void]$Form.ShowDialog()
