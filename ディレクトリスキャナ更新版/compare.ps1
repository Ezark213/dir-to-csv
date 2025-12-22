chcp 65001 | Out-Null

# CSVを読み込み
$dropbox = Import-Csv -Path "C:\Users\Y-IWAO\Desktop\ファイル移行\ドロップボックスリスト.csv"
$teams = Import-Csv -Path "C:\00_ExportDirectory\directory_structure.csv"

Write-Host "Dropbox: $($dropbox.Count) files"
Write-Host "Teams: $($teams.Count) files"

# Teamsにあるファイル名リストを取得
$teamsFiles = @{}
foreach ($t in $teams) {
    $teamsFiles[$t.filename] = $true
}

# Dropboxにあって、Teamsにないファイルを抽出
$missing = @()
foreach ($d in $dropbox) {
    if (-not $teamsFiles.ContainsKey($d.filename)) {
        $missing += $d
    }
}

Write-Host "Missing: $($missing.Count) files" -ForegroundColor Yellow

# 結果をCSVに出力
if ($missing.Count -gt 0) {
    $missing | Export-Csv -Path "C:\Users\Y-IWAO\Desktop\ファイル移行\missing_files.csv" -NoTypeInformation -Encoding UTF8
    Write-Host "Output: C:\Users\Y-IWAO\Desktop\ファイル移行\missing_files.csv"
} else {
    Write-Host "All files exist in Teams!"
}
