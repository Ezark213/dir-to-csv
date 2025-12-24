# ディレクトリスキャナー GUI版
# PowerShell Windows Forms を使用（HTAのセキュリティ問題を回避）

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# メインフォーム
$form = New-Object System.Windows.Forms.Form
$form.Text = "ディレクトリスキャナー"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# タイトル
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$titleLabel.Size = New-Object System.Drawing.Size(450, 30)
$titleLabel.Text = "フォルダ構造をCSVファイルに出力します"
$titleLabel.Font = New-Object System.Drawing.Font("Meiryo UI", 12)
$form.Controls.Add($titleLabel)

# フォルダパス入力
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Location = New-Object System.Drawing.Point(20, 60)
$pathLabel.Size = New-Object System.Drawing.Size(150, 20)
$pathLabel.Text = "スキャンするフォルダ:"
$form.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(20, 85)
$pathTextBox.Size = New-Object System.Drawing.Size(350, 25)
$form.Controls.Add($pathTextBox)

# 参照ボタン
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(380, 83)
$browseButton.Size = New-Object System.Drawing.Size(80, 28)
$browseButton.Text = "参照..."
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "スキャンするフォルダを選択してください"
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $pathTextBox.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($browseButton)

# 最大階層数
$depthLabel = New-Object System.Windows.Forms.Label
$depthLabel.Location = New-Object System.Drawing.Point(20, 125)
$depthLabel.Size = New-Object System.Drawing.Size(150, 20)
$depthLabel.Text = "最大階層数 (1-8):"
$form.Controls.Add($depthLabel)

$depthNumeric = New-Object System.Windows.Forms.NumericUpDown
$depthNumeric.Location = New-Object System.Drawing.Point(20, 148)
$depthNumeric.Size = New-Object System.Drawing.Size(80, 25)
$depthNumeric.Minimum = 1
$depthNumeric.Maximum = 8
$depthNumeric.Value = 8
$form.Controls.Add($depthNumeric)

# 除外フォルダ
$excludeLabel = New-Object System.Windows.Forms.Label
$excludeLabel.Location = New-Object System.Drawing.Point(20, 185)
$excludeLabel.Size = New-Object System.Drawing.Size(200, 20)
$excludeLabel.Text = "除外フォルダ（1行に1つ）:"
$form.Controls.Add($excludeLabel)

$excludeTextBox = New-Object System.Windows.Forms.TextBox
$excludeTextBox.Location = New-Object System.Drawing.Point(20, 208)
$excludeTextBox.Size = New-Object System.Drawing.Size(200, 80)
$excludeTextBox.Multiline = $true
$excludeTextBox.ScrollBars = "Vertical"
$excludeTextBox.Text = ".git`r`nnode_modules`r`n__pycache__`r`n.venv"
$form.Controls.Add($excludeTextBox)

# 出力ファイル名
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(250, 185)
$outputLabel.Size = New-Object System.Drawing.Size(150, 20)
$outputLabel.Text = "出力ファイル名:"
$form.Controls.Add($outputLabel)

$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Location = New-Object System.Drawing.Point(250, 208)
$outputTextBox.Size = New-Object System.Drawing.Size(210, 25)
$outputTextBox.Text = "directory_structure.csv"
$form.Controls.Add($outputTextBox)

# ステータスラベル
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 300)
$statusLabel.Size = New-Object System.Drawing.Size(440, 20)
$statusLabel.Text = ""
$form.Controls.Add($statusLabel)

# スキャン開始ボタン
$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Location = New-Object System.Drawing.Point(150, 330)
$scanButton.Size = New-Object System.Drawing.Size(180, 35)
$scanButton.Text = "スキャン開始"
$scanButton.Font = New-Object System.Drawing.Font("Meiryo UI", 10, [System.Drawing.FontStyle]::Bold)
$scanButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($pathTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("フォルダを選択してください", "エラー", "OK", "Warning")
        return
    }
    
    if (-not (Test-Path $pathTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("指定されたフォルダが存在しません", "エラー", "OK", "Warning")
        return
    }
    
    $statusLabel.Text = "スキャン中..."
    $form.Refresh()
    
    # 除外フォルダをカンマ区切りに
    $excludeDirs = ($excludeTextBox.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" }) -join ","
    
    # config.json を生成
    $config = @{
        rootPath = $pathTextBox.Text
        outputPath = $outputTextBox.Text
        errorLogPath = "error.log"
        maxDepth = [int]$depthNumeric.Value
        excludeDirs = @($excludeTextBox.Text -split "`r`n" | Where-Object { $_.Trim() -ne "" } | ForEach-Object { $_.Trim() })
        excludeFiles = @(".DS_Store", "Thumbs.db")
        includeHidden = $false
        encoding = "UTF8"
        logLevel = "Info"
    }
    
    $configPath = Join-Path $scriptDir "config.json"
    $config | ConvertTo-Json -Depth 10 | Out-File -FilePath $configPath -Encoding UTF8 -Force
    
    # スキャン実行
    $scannerPath = Join-Path $scriptDir "directory_scanner.ps1"
    try {
        & $scannerPath -ConfigPath $configPath
        
        $outputPath = Join-Path $scriptDir $outputTextBox.Text
        if (Test-Path $outputPath) {
            $statusLabel.Text = "完了しました！"
            $result = [System.Windows.Forms.MessageBox]::Show(
                "スキャンが完了しました。`n`n出力ファイル: $outputPath`n`nファイルを開きますか？",
                "完了",
                "YesNo",
                "Information"
            )
            if ($result -eq "Yes") {
                Start-Process $outputPath
            }
        } else {
            $statusLabel.Text = "エラーが発生しました"
            [System.Windows.Forms.MessageBox]::Show("CSVファイルが生成されませんでした。error.logを確認してください。", "エラー", "OK", "Error")
        }
    } catch {
        $statusLabel.Text = "エラーが発生しました"
        [System.Windows.Forms.MessageBox]::Show("エラー: $_", "エラー", "OK", "Error")
    }
})
$form.Controls.Add($scanButton)

# フォーム表示
[void]$form.ShowDialog()
