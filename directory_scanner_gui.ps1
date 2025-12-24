Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Modern Colors
$colorPrimary = [System.Drawing.Color]::FromArgb(79, 70, 229)
$colorPrimaryHover = [System.Drawing.Color]::FromArgb(67, 56, 202)
$colorSuccess = [System.Drawing.Color]::FromArgb(16, 185, 129)
$colorError = [System.Drawing.Color]::FromArgb(239, 68, 68)
$colorBg = [System.Drawing.Color]::FromArgb(249, 250, 251)
$colorCard = [System.Drawing.Color]::White
$colorText = [System.Drawing.Color]::FromArgb(17, 24, 39)
$colorTextSub = [System.Drawing.Color]::FromArgb(107, 114, 128)
$colorBorder = [System.Drawing.Color]::FromArgb(229, 231, 235)

$form = New-Object System.Windows.Forms.Form
$form.Text = "ディレクトリスキャナー"
$form.Size = New-Object System.Drawing.Size(480, 560)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = $colorBg
$form.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9)

# Header Panel
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(480, 80)
$headerPanel.BackColor = $colorPrimary
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(24, 18)
$titleLabel.Size = New-Object System.Drawing.Size(430, 30)
$titleLabel.Text = "ディレクトリスキャナー"
$titleLabel.Font = New-Object System.Drawing.Font("Yu Gothic UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::White
$headerPanel.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Location = New-Object System.Drawing.Point(24, 48)
$subtitleLabel.Size = New-Object System.Drawing.Size(430, 20)
$subtitleLabel.Text = "フォルダ構造をCSVファイルに出力"
$subtitleLabel.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9)
$subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 255)
$headerPanel.Controls.Add($subtitleLabel)

# Folder Section
$folderPanel = New-Object System.Windows.Forms.Panel
$folderPanel.Location = New-Object System.Drawing.Point(20, 100)
$folderPanel.Size = New-Object System.Drawing.Size(424, 100)
$folderPanel.BackColor = $colorCard
$form.Controls.Add($folderPanel)

$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Location = New-Object System.Drawing.Point(16, 14)
$pathLabel.Size = New-Object System.Drawing.Size(200, 18)
$pathLabel.Text = "スキャン対象フォルダ"
$pathLabel.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9, [System.Drawing.FontStyle]::Bold)
$pathLabel.ForeColor = $colorText
$folderPanel.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(16, 40)
$pathTextBox.Size = New-Object System.Drawing.Size(300, 28)
$pathTextBox.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10)
$pathTextBox.BorderStyle = "FixedSingle"
$folderPanel.Controls.Add($pathTextBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(324, 38)
$browseButton.Size = New-Object System.Drawing.Size(84, 32)
$browseButton.Text = "参照..."
$browseButton.FlatStyle = "Flat"
$browseButton.BackColor = $colorBorder
$browseButton.ForeColor = $colorText
$browseButton.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9)
$browseButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$browseButton.FlatAppearance.BorderSize = 0
$folderPanel.Controls.Add($browseButton)

# Settings Section
$settingsPanel = New-Object System.Windows.Forms.Panel
$settingsPanel.Location = New-Object System.Drawing.Point(20, 210)
$settingsPanel.Size = New-Object System.Drawing.Size(424, 100)
$settingsPanel.BackColor = $colorCard
$form.Controls.Add($settingsPanel)

$settingsTitle = New-Object System.Windows.Forms.Label
$settingsTitle.Location = New-Object System.Drawing.Point(16, 14)
$settingsTitle.Size = New-Object System.Drawing.Size(200, 18)
$settingsTitle.Text = "設定"
$settingsTitle.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9, [System.Drawing.FontStyle]::Bold)
$settingsTitle.ForeColor = $colorText
$settingsPanel.Controls.Add($settingsTitle)

$depthLabel = New-Object System.Windows.Forms.Label
$depthLabel.Location = New-Object System.Drawing.Point(16, 44)
$depthLabel.Size = New-Object System.Drawing.Size(80, 18)
$depthLabel.Text = "最大階層数"
$depthLabel.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9)
$depthLabel.ForeColor = $colorTextSub
$settingsPanel.Controls.Add($depthLabel)

$depthNumeric = New-Object System.Windows.Forms.NumericUpDown
$depthNumeric.Location = New-Object System.Drawing.Point(16, 64)
$depthNumeric.Size = New-Object System.Drawing.Size(70, 28)
$depthNumeric.Minimum = 1
$depthNumeric.Maximum = 8
$depthNumeric.Value = 8
$depthNumeric.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10)
$depthNumeric.BorderStyle = "FixedSingle"
$settingsPanel.Controls.Add($depthNumeric)

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(120, 44)
$outputLabel.Size = New-Object System.Drawing.Size(150, 18)
$outputLabel.Text = "出力ファイル名"
$outputLabel.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9)
$outputLabel.ForeColor = $colorTextSub
$settingsPanel.Controls.Add($outputLabel)

$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Location = New-Object System.Drawing.Point(120, 64)
$outputTextBox.Size = New-Object System.Drawing.Size(200, 28)
$outputTextBox.Text = "directory_structure.csv"
$outputTextBox.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10)
$outputTextBox.BorderStyle = "FixedSingle"
$settingsPanel.Controls.Add($outputTextBox)

# Result Section
$resultPanel = New-Object System.Windows.Forms.Panel
$resultPanel.Location = New-Object System.Drawing.Point(20, 320)
$resultPanel.Size = New-Object System.Drawing.Size(424, 110)
$resultPanel.BackColor = $colorCard
$resultPanel.Visible = $false
$form.Controls.Add($resultPanel)

$resultTitle = New-Object System.Windows.Forms.Label
$resultTitle.Location = New-Object System.Drawing.Point(16, 14)
$resultTitle.Size = New-Object System.Drawing.Size(200, 18)
$resultTitle.Text = "スキャン結果"
$resultTitle.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9, [System.Drawing.FontStyle]::Bold)
$resultTitle.ForeColor = $colorText
$resultPanel.Controls.Add($resultTitle)

$statusIcon = New-Object System.Windows.Forms.Label
$statusIcon.Location = New-Object System.Drawing.Point(16, 42)
$statusIcon.Size = New-Object System.Drawing.Size(50, 50)
$statusIcon.Font = New-Object System.Drawing.Font("Segoe UI", 28)
$statusIcon.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$resultPanel.Controls.Add($statusIcon)

$fileCountLabel = New-Object System.Windows.Forms.Label
$fileCountLabel.Location = New-Object System.Drawing.Point(70, 48)
$fileCountLabel.Size = New-Object System.Drawing.Size(200, 22)
$fileCountLabel.Font = New-Object System.Drawing.Font("Yu Gothic UI", 11)
$fileCountLabel.ForeColor = $colorText
$resultPanel.Controls.Add($fileCountLabel)

$timeLabel = New-Object System.Windows.Forms.Label
$timeLabel.Location = New-Object System.Drawing.Point(70, 72)
$timeLabel.Size = New-Object System.Drawing.Size(200, 20)
$timeLabel.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9)
$timeLabel.ForeColor = $colorTextSub
$resultPanel.Controls.Add($timeLabel)

$openButton = New-Object System.Windows.Forms.Button
$openButton.Location = New-Object System.Drawing.Point(290, 50)
$openButton.Size = New-Object System.Drawing.Size(118, 38)
$openButton.Text = "CSVを開く"
$openButton.FlatStyle = "Flat"
$openButton.BackColor = $colorSuccess
$openButton.ForeColor = [System.Drawing.Color]::White
$openButton.Font = New-Object System.Drawing.Font("Yu Gothic UI", 9, [System.Drawing.FontStyle]::Bold)
$openButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$openButton.FlatAppearance.BorderSize = 0
$openButton.Visible = $false
$resultPanel.Controls.Add($openButton)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 440)
$statusLabel.Size = New-Object System.Drawing.Size(424, 24)
$statusLabel.Text = ""
$statusLabel.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10)
$statusLabel.ForeColor = $colorTextSub
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($statusLabel)

# Scan Button
$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Location = New-Object System.Drawing.Point(130, 475)
$scanButton.Size = New-Object System.Drawing.Size(200, 50)
$scanButton.Text = "スキャン開始"
$scanButton.FlatStyle = "Flat"
$scanButton.BackColor = $colorPrimary
$scanButton.ForeColor = [System.Drawing.Color]::White
$scanButton.Font = New-Object System.Drawing.Font("Yu Gothic UI", 12, [System.Drawing.FontStyle]::Bold)
$scanButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$scanButton.FlatAppearance.BorderSize = 0
$form.Controls.Add($scanButton)

$script:lastOutputPath = ""

$browseButton.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = "スキャンするフォルダを選択"
    if ($fb.ShowDialog() -eq "OK") {
        $pathTextBox.Text = $fb.SelectedPath
    }
})

$openButton.Add_Click({
    if (Test-Path $script:lastOutputPath) {
        Start-Process $script:lastOutputPath
    }
})

$scanButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($pathTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("フォルダを選択してください", "エラー", "OK", "Warning")
        return
    }
    if (-not (Test-Path $pathTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("フォルダが見つかりません", "エラー", "OK", "Warning")
        return
    }

    $resultPanel.Visible = $false
    $openButton.Visible = $false
    $statusLabel.Text = "スキャン中..."
    $statusLabel.ForeColor = $colorPrimary
    $scanButton.Enabled = $false
    $scanButton.BackColor = [System.Drawing.Color]::FromArgb(156, 163, 175)
    $form.Refresh()

    $startTime = Get-Date

    $outFullPath = Join-Path $scriptDir $outputTextBox.Text
    $errFullPath = Join-Path $scriptDir "error.log"
    
    $cfg = @{
        rootPath = $pathTextBox.Text
        outputPath = $outFullPath
        errorLogPath = $errFullPath
        maxDepth = [int]$depthNumeric.Value
        excludeDirs = @(".git", "node_modules", "__pycache__", ".venv")
        excludeFiles = @(".DS_Store", "Thumbs.db")
        includeHidden = $false
        encoding = "UTF8"
        logLevel = "Info"
    }
    $cfgPath = Join-Path $scriptDir "config.json"
    $cfg | ConvertTo-Json -Depth 10 | Out-File -FilePath $cfgPath -Encoding UTF8 -Force

    $ps1 = Join-Path $scriptDir "directory_scanner.ps1"
    $out = $outFullPath
    $script:lastOutputPath = $out

    # Clear environment variables from old HTA
    $env:SCANNER_ROOT_PATH = $null
    $env:SCANNER_OUTPUT_FILE = $null
    $env:SCANNER_MAX_DEPTH = $null
    $env:SCANNER_EXCLUDE_DIRS = $null
    $env:SCANNER_APP_PATH = $null

    # Delete old CSV file
    if (Test-Path $out) {
        Remove-Item $out -Force -ErrorAction SilentlyContinue
    }

    & $ps1 -ConfigPath $cfgPath

    $endTime = Get-Date
    $elapsed = $endTime - $startTime
    $elapsedStr = "{0:N1} 秒" -f $elapsed.TotalSeconds

    $resultPanel.Visible = $true

    if (Test-Path $out) {
        $fileCount = (Import-Csv $out).Count
        $statusIcon.Text = [char]0x2713
        $statusIcon.ForeColor = $colorSuccess
        $fileCountLabel.Text = "ファイル数: " + $fileCount.ToString("N0")
        $timeLabel.Text = "処理時間: " + $elapsedStr
        $statusLabel.Text = "スキャンが完了しました"
        $statusLabel.ForeColor = $colorSuccess
        $openButton.Visible = $true
    } else {
        $statusIcon.Text = [char]0x2717
        $statusIcon.ForeColor = $colorError
        $fileCountLabel.Text = "ファイル数: -"
        $timeLabel.Text = "処理時間: " + $elapsedStr
        $statusLabel.Text = "エラー: CSVファイルを作成できませんでした"
        $statusLabel.ForeColor = $colorError
        $openButton.Visible = $false
    }

    $scanButton.Enabled = $true
    $scanButton.BackColor = $colorPrimary
})

[void]$form.ShowDialog()
