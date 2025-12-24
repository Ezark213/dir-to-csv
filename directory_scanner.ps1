<#
.SYNOPSIS
    Export directory structure to CSV

.DESCRIPTION
    Scan directory structure up to 8 levels and export to CSV format:
    level1,level2,level3,level4,level5,level6,level7,level8,filename,fullPath

.PARAMETER ConfigPath
    Path to config.json. Default is same directory as script.

.PARAMETER RootPath
    Root directory path to scan (passed from HTA).

.PARAMETER OutputFile
    Output CSV filename (passed from HTA).

.PARAMETER MaxDepth
    Maximum folder depth 1-8 (passed from HTA).

.PARAMETER ExcludeDirs
    Comma-separated list of directories to exclude (passed from HTA).

.PARAMETER AppPath
    Application directory path (passed from HTA).

.EXAMPLE
    .\directory_scanner.ps1
    .\directory_scanner.ps1 -ConfigPath "C:\config\myconfig.json"
    .\directory_scanner.ps1 -RootPath "C:\Users" -OutputFile "output.csv" -MaxDepth 5 -ExcludeDirs ".git,node_modules" -AppPath "C:\App"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ParamsFile = "",
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "",
    [Parameter(Mandatory = $false)]
    [string]$RootPath = "",
    [Parameter(Mandatory = $false)]
    [string]$OutputFile = "",
    [Parameter(Mandatory = $false)]
    [string]$MaxDepth = "",
    [Parameter(Mandatory = $false)]
    [string]$ExcludeDirs = "",
    [Parameter(Mandatory = $false)]
    [string]$AppPath = ""
)

$script:LogLevel = "Info"
$script:ErrorLogPath = ""
$script:ErrorLogInitialized = $false
$script:scannedCount = 0
$script:errorCount = 0
$script:ScanStartTime = ""
$script:CurrentAppPath = ""

function Read-Params {
    param([string]$ParamsFile)

    if (-not (Test-Path -Path $ParamsFile -PathType Leaf)) {
        throw "パラメータファイルが見つかりません: $ParamsFile"
    }

    try {
        # UTF-8 BOM対応: StreamReaderで読み込み
        $reader = New-Object System.IO.StreamReader($ParamsFile, [System.Text.Encoding]::UTF8, $true)
        $jsonContent = $reader.ReadToEnd()
        $reader.Close()
        $params = $jsonContent | ConvertFrom-Json

        if ([string]::IsNullOrEmpty($params.rootPath)) {
            throw "rootPathが指定されていません"
        }
        if ([string]::IsNullOrEmpty($params.appPath)) {
            throw "appPathが指定されていません"
        }

        $config = @{
            RootPath      = $params.rootPath
            OutputPath    = if ($params.outputFile) { $params.outputFile } else { "directory_structure.csv" }
            ErrorLogPath  = "error.log"
            MaxDepth      = if ($null -ne $params.maxDepth) { [int]$params.maxDepth } else { 8 }
            ExcludeDirs   = if ($params.excludeDirs) { @($params.excludeDirs) } else { @() }
            ExcludeFiles  = @(".DS_Store", "Thumbs.db")
            IncludeHidden = $false
            Encoding      = "UTF8"
            LogLevel      = "Info"
            AppPath       = $params.appPath
        }

        if ($config.MaxDepth -lt 1) { $config.MaxDepth = 1 }
        if ($config.MaxDepth -gt 8) { $config.MaxDepth = 8 }

        return $config
    }
    catch {
        throw "パラメータファイルの読み込みに失敗しました: $_"
    }
}

function Test-SymbolicLink {
    param([System.IO.FileSystemInfo]$Item)

    # ReparsePoint属性がない場合は確実にシンボリックリンクではない
    $reparsePoint = [System.IO.FileAttributes]::ReparsePoint
    if (($Item.Attributes -band $reparsePoint) -ne $reparsePoint) {
        return $false
    }

    # ReparsePoint属性がある場合、LinkTypeで判定
    # クラウドファイル（OneDrive, Dropbox等）はLinkTypeが空
    # シンボリックリンクは "SymbolicLink"、ジャンクションは "Junction"
    try {
        $linkType = $Item.LinkType
        if ($linkType -eq "SymbolicLink" -or $linkType -eq "Junction") {
            return $true
        }
    }
    catch {
        # LinkTypeプロパティがない古いPowerShellの場合
        # Modeプロパティで確認（'l'がシンボリックリンク）
        if ($Item.Mode -match 'l') {
            return $true
        }
    }

    return $false
}

function Write-Progress-File {
    param(
        [string]$AppPath,
        [int]$ScannedCount,
        [string]$CurrentPath,
        [int]$ErrorCount
    )

    $progressPath = Join-Path -Path $AppPath -ChildPath "progress.json"
    $progress = @{
        status = "running"
        scannedCount = $ScannedCount
        currentPath = $CurrentPath
        startTime = $script:ScanStartTime
        errorCount = $ErrorCount
        lastUpdate = (Get-Date).ToString("o")
    }
    $progress | ConvertTo-Json | Out-File -FilePath $progressPath -Encoding UTF8 -Force
}

function Write-Completion {
    param(
        [string]$AppPath,
        [string]$Status,
        [int]$FileCount = 0,
        [int]$ErrorCount = 0,
        [string]$OutputPath = "",
        [string]$ErrorCode = "",
        [string]$ErrorMessage = ""
    )

    $completionPath = Join-Path -Path $AppPath -ChildPath "completion.json"
    $completion = @{
        status = $Status
        fileCount = $FileCount
        errorCount = $ErrorCount
        outputPath = $OutputPath
        timestamp = (Get-Date).ToString("o")
    }

    if ($Status -eq "error") {
        $completion.errorCode = $ErrorCode
        $completion.errorMessage = $ErrorMessage
    }

    $completion | ConvertTo-Json | Out-File -FilePath $completionPath -Encoding UTF8 -Force
}

function Test-CancelRequested {
    param([string]$AppPath)
    $cancelFlagPath = Join-Path -Path $AppPath -ChildPath "cancel.flag"
    return (Test-Path -Path $cancelFlagPath -PathType Leaf)
}

function Initialize-ErrorLog {
    param([string]$Path)
    
    $script:ErrorLogPath = $Path
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $lines = @(
            "================================================================================"
            "Directory Scanner Error Log"
            "Start: $timestamp"
            "================================================================================"
            ""
        )
        $lines | Out-File -FilePath $Path -Encoding UTF8 -Force
        $script:ErrorLogInitialized = $true
    }
    catch {
        Write-Host "[WARNING] Failed to initialize error log: $_" -ForegroundColor Yellow
    }
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )

    $levelPriority = @{ "Debug" = 0; "Info" = 1; "Warning" = 2; "Error" = 3 }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    if ($levelPriority[$Level] -ge $levelPriority[$script:LogLevel]) {
        switch ($Level) {
            "Debug"   { Write-Host $logMessage -ForegroundColor Gray }
            "Info"    { Write-Host $logMessage -ForegroundColor White }
            "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
            "Error"   { Write-Host $logMessage -ForegroundColor Red }
        }
    }

    if ($script:ErrorLogInitialized -and ($Level -eq "Warning" -or $Level -eq "Error")) {
        try {
            $logMessage | Out-File -FilePath $script:ErrorLogPath -Encoding UTF8 -Append
        }
        catch { }
    }
}

function Write-ErrorLogSummary {
    param([int]$TotalFiles, [int]$ErrorCount)

    if (-not $script:ErrorLogInitialized) { return }

    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $lines = @(
            ""
            "================================================================================"
            "Completed"
            "End: $timestamp"
            "Total Files: $TotalFiles"
            "Error Count: $ErrorCount"
            "================================================================================"
        )
        $lines | Out-File -FilePath $script:ErrorLogPath -Encoding UTF8 -Append

        if ($ErrorCount -gt 0) {
            $fullPath = Resolve-Path -Path $script:ErrorLogPath
            Write-Log "Error log: $fullPath" -Level Info
        }
    }
    catch { }
}

function Read-Config {
    param([string]$Path)

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        throw "Config file not found: $Path"
    }

    try {
        $jsonContent = Get-Content -Path $Path -Raw -Encoding UTF8
        $jsonData = $jsonContent | ConvertFrom-Json

        $config = @{
            RootPath      = if ($jsonData.rootPath) { $jsonData.rootPath } else { "." }
            OutputPath    = if ($jsonData.outputPath) { $jsonData.outputPath } else { "directory_structure.csv" }
            ErrorLogPath  = if ($jsonData.errorLogPath) { $jsonData.errorLogPath } else { "error.log" }
            MaxDepth      = if ($null -ne $jsonData.maxDepth) { [int]$jsonData.maxDepth } else { 8 }
            ExcludeDirs   = if ($jsonData.excludeDirs) { @($jsonData.excludeDirs) } else { @() }
            ExcludeFiles  = if ($jsonData.excludeFiles) { @($jsonData.excludeFiles) } else { @() }
            IncludeHidden = if ($null -ne $jsonData.includeHidden) { [bool]$jsonData.includeHidden } else { $false }
            Encoding      = if ($jsonData.encoding) { $jsonData.encoding } else { "UTF8" }
            LogLevel      = if ($jsonData.logLevel) { $jsonData.logLevel } else { "Info" }
        }

        if ($config.MaxDepth -lt 1) { $config.MaxDepth = 1 }
        if ($config.MaxDepth -gt 8) { $config.MaxDepth = 8 }

        return $config
    }
    catch {
        throw "Failed to read config: $_"
    }
}

function Test-ShouldExclude {
    param(
        [string]$Name,
        [string[]]$ExcludePatterns,
        [bool]$IncludeHidden
    )

    if ($Name.StartsWith(".") -and -not $IncludeHidden) {
        return $true
    }

    foreach ($pattern in $ExcludePatterns) {
        if ($Name -like "*$pattern*") {
            return $true
        }
    }

    return $false
}

function Get-DirectoryStructure {
    param([hashtable]$Config)

    $script:scannedCount = 0
    $script:errorCount = 0
    $script:ScanStartTime = (Get-Date).ToString("o")
    $progressInterval = 500

    $rootPath = Resolve-Path -Path $Config.RootPath -ErrorAction Stop

    if (-not (Test-Path -Path $rootPath -PathType Container)) {
        throw "Path is not a directory: $($Config.RootPath)"
    }

    $rootDirName = Split-Path -Path $rootPath.Path -Leaf

    Write-Log "Scan start: $rootPath" -Level Info
    Write-Log "Root directory: $rootDirName" -Level Info
    Write-Log "Max depth: $($Config.MaxDepth)" -Level Info

    $results = New-Object System.Collections.ArrayList

    $stack = New-Object System.Collections.Stack
    $stack.Push(@{ Path = $rootPath.Path; Levels = @(); Depth = 1 })

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        $currentPath = $current.Path
        $levels = $current.Levels
        $depth = $current.Depth

        if ($depth -gt $Config.MaxDepth) { continue }

        # キャンセルチェック・進捗更新（500ファイル毎）
        if ($script:scannedCount % $progressInterval -eq 0 -and $script:scannedCount -gt 0) {
            if (Test-CancelRequested -AppPath $Config.AppPath) {
                Write-Log "キャンセルが要求されました" -Level Info
                return @{
                    Results    = $results
                    FileCount  = $script:scannedCount
                    ErrorCount = $script:errorCount
                    Cancelled  = $true
                }
            }
            Write-Progress-File -AppPath $Config.AppPath `
                                -ScannedCount $script:scannedCount `
                                -CurrentPath $currentPath `
                                -ErrorCount $script:errorCount
        }

        try {
            $items = Get-ChildItem -Path $currentPath -Force -ErrorAction Stop | Sort-Object { $_.PSIsContainer }, Name -Descending
        }
        catch [System.UnauthorizedAccessException] {
            Write-Log "Access denied: $currentPath" -Level Warning
            $script:errorCount++
            continue
        }
        catch {
            Write-Log "Read error: $currentPath - $_" -Level Warning
            $script:errorCount++
            continue
        }

        foreach ($item in $items) {
            $name = $item.Name

            if ($item.PSIsContainer) {
                # シンボリックリンク/ジャンクションチェック（新規）
                if (Test-SymbolicLink -Item $item) {
                    Write-Log "Skipped (symlink): $($item.FullName)" -Level Debug
                    continue
                }

                if (Test-ShouldExclude -Name $name -ExcludePatterns $Config.ExcludeDirs -IncludeHidden $Config.IncludeHidden) {
                    Write-Log "Excluded (dir): $($item.FullName)" -Level Debug
                    continue
                }

                $newLevels = @($levels) + @($name)
                $stack.Push(@{ Path = $item.FullName; Levels = $newLevels; Depth = ($depth + 1) })
            }
            else {
                if (Test-ShouldExclude -Name $name -ExcludePatterns $Config.ExcludeFiles -IncludeHidden $Config.IncludeHidden) {
                    Write-Log "Excluded (file): $($item.FullName)" -Level Debug
                    continue
                }

                $allLevels = @($rootDirName) + @($levels)

                $record = [PSCustomObject]@{
                    level1   = if ($allLevels.Count -ge 1) { $allLevels[0] } else { "" }
                    level2   = if ($allLevels.Count -ge 2) { $allLevels[1] } else { "" }
                    level3   = if ($allLevels.Count -ge 3) { $allLevels[2] } else { "" }
                    level4   = if ($allLevels.Count -ge 4) { $allLevels[3] } else { "" }
                    level5   = if ($allLevels.Count -ge 5) { $allLevels[4] } else { "" }
                    level6   = if ($allLevels.Count -ge 6) { $allLevels[5] } else { "" }
                    level7   = if ($allLevels.Count -ge 7) { $allLevels[6] } else { "" }
                    level8   = if ($allLevels.Count -ge 8) { $allLevels[7] } else { "" }
                    filename = $name
                    fullPath = $item.FullName
                }

                [void]$results.Add($record)
                $script:scannedCount++

                if ($script:scannedCount % 1000 -eq 0) {
                    Write-Log "Processing... $($script:scannedCount) files" -Level Info
                }
            }
        }
    }

    Write-Log "Scan completed: $($script:scannedCount) files" -Level Info
    if ($script:errorCount -gt 0) {
        Write-Log "Error count: $($script:errorCount)" -Level Warning
    }

    return @{
        Results    = $results
        FileCount  = $script:scannedCount
        ErrorCount = $script:errorCount
        Cancelled  = $false
    }
}

function Export-ToCsv {
    param(
        [System.Collections.ArrayList]$Data,
        [string]$OutputPath,
        [string]$Encoding
    )

    if ($Data.Count -eq 0) {
        Write-Log "No data to export" -Level Warning
        return
    }

    $outputDir = Split-Path -Path $OutputPath -Parent
    if ($outputDir -and -not (Test-Path -Path $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        Write-Log "Created output directory: $outputDir" -Level Info
    }

    $Data | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding $Encoding -Force

    $fullPath = Resolve-Path -Path $OutputPath
    Write-Log "CSV exported: $fullPath" -Level Info
    Write-Log "Total records: $($Data.Count)" -Level Info
}

# Main
$exitCode = 0
$fileCount = 0
$errCount = 0
$config = $null

try {
    # パラメータ読み込み（優先順位）
    # 1. params.json（新方式）
    # 2. 環境変数（後方互換）
    # 3. config.json（フォールバック）

    if (-not [string]::IsNullOrEmpty($ParamsFile) -and (Test-Path $ParamsFile)) {
        Write-Host "params.jsonから読み込み: $ParamsFile" -ForegroundColor Cyan
        $config = Read-Params -ParamsFile $ParamsFile
        $script:CurrentAppPath = $config.AppPath
        Set-Location -Path $config.AppPath
    }
    elseif (-not [string]::IsNullOrEmpty($env:SCANNER_ROOT_PATH) -and -not [string]::IsNullOrEmpty($env:SCANNER_APP_PATH)) {
        # 後方互換: 環境変数方式
        Write-Host "環境変数から読み込み（後方互換）" -ForegroundColor Yellow
        $RootPath = $env:SCANNER_ROOT_PATH
        $OutputFile = $env:SCANNER_OUTPUT_FILE
        $MaxDepth = $env:SCANNER_MAX_DEPTH
        $ExcludeDirs = $env:SCANNER_EXCLUDE_DIRS
        $AppPath = $env:SCANNER_APP_PATH

        Set-Location -Path $AppPath
        $script:CurrentAppPath = $AppPath

        $excludeDirList = @()
        if (-not [string]::IsNullOrEmpty($ExcludeDirs)) {
            $excludeDirList = $ExcludeDirs -split ',' | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
        }

        $maxDepthNum = 8
        if (-not [string]::IsNullOrEmpty($MaxDepth)) {
            $maxDepthNum = [int]$MaxDepth
            if ($maxDepthNum -lt 1) { $maxDepthNum = 1 }
            if ($maxDepthNum -gt 8) { $maxDepthNum = 8 }
        }

        $config = @{
            RootPath      = $RootPath
            OutputPath    = if (-not [string]::IsNullOrEmpty($OutputFile)) { $OutputFile } else { "directory_structure.csv" }
            ErrorLogPath  = "error.log"
            MaxDepth      = $maxDepthNum
            ExcludeDirs   = $excludeDirList
            ExcludeFiles  = @(".DS_Store", "Thumbs.db")
            IncludeHidden = $false
            Encoding      = "UTF8"
            LogLevel      = "Info"
            AppPath       = $AppPath
        }
    }
    else {
        # フォールバック: config.json
        Write-Host "config.jsonから読み込み（フォールバック）" -ForegroundColor Yellow
        if ([string]::IsNullOrEmpty($ConfigPath)) {
            $scriptDir = $PSScriptRoot
            if ([string]::IsNullOrEmpty($scriptDir)) {
                $scriptDir = Get-Location
            }
            $ConfigPath = Join-Path -Path $scriptDir -ChildPath "config.json"
        }
        $config = Read-Config -Path $ConfigPath
        $script:CurrentAppPath = Split-Path -Path $ConfigPath -Parent
        $config.AppPath = $script:CurrentAppPath
    }

    $script:LogLevel = $config.LogLevel

    Initialize-ErrorLog -Path (Join-Path -Path $config.AppPath -ChildPath $config.ErrorLogPath)

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Directory Structure Scanner v2.0" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Log "Config loaded" -Level Info

    $scanResult = Get-DirectoryStructure -Config $config
    $fileCount = $scanResult.FileCount
    $errCount = $scanResult.ErrorCount

    $outputFullPath = Join-Path -Path $config.AppPath -ChildPath $config.OutputPath

    if ($scanResult.Cancelled) {
        # キャンセル時
        Export-ToCsv -Data $scanResult.Results -OutputPath $config.OutputPath -Encoding $config.Encoding
        Write-Completion -AppPath $config.AppPath `
                         -Status "cancelled" `
                         -FileCount $fileCount `
                         -ErrorCount $errCount `
                         -OutputPath $outputFullPath
        Write-Host "キャンセルされました" -ForegroundColor Yellow
    }
    else {
        # 正常完了
        Export-ToCsv -Data $scanResult.Results -OutputPath $config.OutputPath -Encoding $config.Encoding
        Write-Completion -AppPath $config.AppPath `
                         -Status "success" `
                         -FileCount $fileCount `
                         -ErrorCount $errCount `
                         -OutputPath $outputFullPath
        Write-Host ""
        Write-Host "Completed successfully" -ForegroundColor Green
    }
}
catch {
    Write-Log "Error: $_" -Level Error

    # エラー時のcompletion.json出力
    if ($config -and $config.AppPath) {
        $errorCode = "ERR_UNKNOWN"
        if ($_.Exception.Message -like "*見つかりません*" -or $_.Exception.Message -like "*not found*") {
            $errorCode = "ERR_PATH_NOT_FOUND"
        }
        elseif ($_.Exception.Message -like "*アクセス*" -or $_.Exception.Message -like "*denied*") {
            $errorCode = "ERR_ACCESS_DENIED"
        }

        Write-Completion -AppPath $config.AppPath `
                         -Status "error" `
                         -ErrorCode $errorCode `
                         -ErrorMessage $_.Exception.Message
    }

    $exitCode = 1
}
finally {
    Write-ErrorLogSummary -TotalFiles $fileCount -ErrorCount $errCount
}

exit $exitCode
