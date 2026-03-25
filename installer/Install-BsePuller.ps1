Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'
[System.Windows.Forms.Application]::EnableVisualStyles()

function Show-InstallDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Install BSE Puller'
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ClientSize = New-Object System.Drawing.Size(540, 250)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.BackColor = [System.Drawing.Color]::White

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'BSE Puller setup'
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15, [System.Drawing.FontStyle]::Bold)
    $title.AutoSize = $true
    $title.Location = New-Object System.Drawing.Point(20, 18)

    $instructions = New-Object System.Windows.Forms.Label
    $instructions.Text = 'Enter the BILL Spend and Expense API token for this Windows user. The installer will save it and then install BSE Puller.'
    $instructions.AutoSize = $false
    $instructions.Size = New-Object System.Drawing.Size(500, 48)
    $instructions.Location = New-Object System.Drawing.Point(20, 56)

    $tokenLabel = New-Object System.Windows.Forms.Label
    $tokenLabel.Text = 'API token'
    $tokenLabel.AutoSize = $true
    $tokenLabel.Location = New-Object System.Drawing.Point(20, 116)

    $tokenBox = New-Object System.Windows.Forms.TextBox
    $tokenBox.Size = New-Object System.Drawing.Size(500, 26)
    $tokenBox.Location = New-Object System.Drawing.Point(20, 138)

    $installButton = New-Object System.Windows.Forms.Button
    $installButton.Text = 'Install'
    $installButton.Size = New-Object System.Drawing.Size(88, 32)
    $installButton.Location = New-Object System.Drawing.Point(340, 194)
    $installButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Cancel'
    $cancelButton.Size = New-Object System.Drawing.Size(88, 32)
    $cancelButton.Location = New-Object System.Drawing.Point(432, 194)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.AcceptButton = $installButton
    $form.CancelButton = $cancelButton
    $form.Controls.AddRange(@($title, $instructions, $tokenLabel, $tokenBox, $installButton, $cancelButton))

    while ($true) {
        $result = $form.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            return $null
        }

        $token = $tokenBox.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            return $token
        }

        [System.Windows.Forms.MessageBox]::Show(
            $form,
            'Enter a BILL API token before continuing.',
            'API token required',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
}

function Save-UserSettings([string]$apiToken) {
    $settingsFolder = Join-Path $env:LOCALAPPDATA 'BsePuller'
    $settingsPath = Join-Path $settingsFolder 'settings.json'
    New-Item -ItemType Directory -Path $settingsFolder -Force | Out-Null

    $payload = @{
        ApiToken = $apiToken
    } | ConvertTo-Json

    Set-Content -Path $settingsPath -Value $payload -Encoding UTF8
}

function Copy-AppFiles([string]$sourceDir, [string]$installDir) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null

    $appFiles = Get-ChildItem -Path $sourceDir -File | Where-Object {
        $_.Name -notin @('Install-BsePuller.cmd', 'Install-BsePuller.ps1', 'Build-Installer.ps1', 'BsePullerInstaller.sed')
    }

    foreach ($file in $appFiles) {
        Copy-Item $file.FullName (Join-Path $installDir $file.Name) -Force
    }
}

function New-Shortcut([string]$shortcutPath, [string]$targetPath, [string]$workingDirectory) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $targetPath
    $shortcut.WorkingDirectory = $workingDirectory
    $shortcut.IconLocation = $targetPath
    $shortcut.Save()
}

function Create-StartMenuShortcut([string]$installDir) {
    $programsFolder = [Environment]::GetFolderPath('Programs')
    $shortcutFolder = Join-Path $programsFolder 'BSE Puller'
    New-Item -ItemType Directory -Path $shortcutFolder -Force | Out-Null

    $exePath = Join-Path $installDir 'BsePuller.exe'
    $shortcutPath = Join-Path $shortcutFolder 'BSE Puller.lnk'
    New-Shortcut -shortcutPath $shortcutPath -targetPath $exePath -workingDirectory $installDir
}

function Show-Success([string]$installDir) {
    [System.Windows.Forms.MessageBox]::Show(
        "BSE Puller was installed successfully.`r`n`r`nLocation: $installDir`r`n`r`nYou can launch it from the Start menu.",
        'Install complete',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

$token = Show-InstallDialog
if ($null -eq $token) {
    exit 1
}

$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$installDir = Join-Path $env:LOCALAPPDATA 'Programs\BsePuller'

Save-UserSettings -apiToken $token
Copy-AppFiles -sourceDir $sourceDir -installDir $installDir
Create-StartMenuShortcut -installDir $installDir
Show-Success -installDir $installDir
