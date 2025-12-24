<#
.SYNOPSIS
    Export directory structure to CSV
.PARAMETER ConfigPath
    Path to config.json
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = ""
)

function Read-Config {
    param([string]$Path)

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        throw "Config file not found: $Path"
    }

    $jsonContent = Get-Content -Path $Path -Raw -Encoding UTF8
    $jsonData = $jsonContent | ConvertFrom-Json

    return @{
        RootPath      = if ($jsonData.rootPath) { $jsonData.rootPath } else { "." }
        OutputPath    = if ($jsonData.outputPath) { $jsonData.outputPath } else { "directory_structure.csv" }
        MaxDepth      = if ($null -ne $jsonData.maxDepth) { [int]$jsonData.maxDepth } else { 8 }
        ExcludeDirs   = if ($jsonData.excludeDirs) { @($jsonData.excludeDirs) } else { @() }
        ExcludeFiles  = if ($jsonData.excludeFiles) { @($jsonData.excludeFiles) } else { @() }
    }
}

function Test-ShouldExclude {
    param(
        [string]$Name,
        [string[]]$ExcludePatterns
    )

    if ($Name.StartsWith(".")) { return $true }

    foreach ($pattern in $ExcludePatterns) {
        if ($Name -like "*$pattern*") { return $true }
    }
    return $false
}

function Get-DirectoryStructure {
    param([hashtable]$Config)

    $scannedCount = 0
    $rootPath = Resolve-Path -Path $Config.RootPath -ErrorAction Stop

    if (-not (Test-Path -Path $rootPath -PathType Container)) {
        throw "Path is not a directory: $($Config.RootPath)"
    }

    $rootDirName = Split-Path -Path $rootPath.Path -Leaf

    Write-Host "Scanning: $rootPath"
    Write-Host "Max depth: $($Config.MaxDepth)"

    $results = New-Object System.Collections.ArrayList

    $stack = New-Object System.Collections.Stack
    $stack.Push(@{ Path = $rootPath.Path; Levels = @(); Depth = 1 })

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        $currentPath = $current.Path
        $levels = $current.Levels
        $depth = $current.Depth

        if ($depth -gt $Config.MaxDepth) { continue }

        try {
            $items = Get-ChildItem -Path $currentPath -Force -ErrorAction Stop | Sort-Object { $_.PSIsContainer }, Name -Descending
        }
        catch {
            Write-Host "Access denied: $currentPath" -ForegroundColor Yellow
            continue
        }

        foreach ($item in $items) {
            $name = $item.Name

            if ($item.PSIsContainer) {
                if (Test-ShouldExclude -Name $name -ExcludePatterns $Config.ExcludeDirs) { continue }

                $newLevels = @($levels) + @($name)
                $stack.Push(@{ Path = $item.FullName; Levels = $newLevels; Depth = ($depth + 1) })
            }
            else {
                if (Test-ShouldExclude -Name $name -ExcludePatterns $Config.ExcludeFiles) { continue }

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
                $scannedCount++

                if ($scannedCount % 1000 -eq 0) {
                    Write-Host "Processing... $scannedCount files"
                }
            }
        }
    }

    Write-Host "Scan completed: $scannedCount files"

    return @{
        Results   = $results
        FileCount = $scannedCount
    }
}

# Main
try {
    if ([string]::IsNullOrEmpty($ConfigPath)) {
        $scriptDir = $PSScriptRoot
        if ([string]::IsNullOrEmpty($scriptDir)) {
            $scriptDir = Get-Location
        }
        $ConfigPath = Join-Path -Path $scriptDir -ChildPath "config.json"
    }

    Write-Host "========================================"
    Write-Host "  Directory Structure Scanner"
    Write-Host "========================================"
    Write-Host ""
    Write-Host "Config: $ConfigPath"

    $config = Read-Config -Path $ConfigPath

    $scanResult = Get-DirectoryStructure -Config $config

    if ($scanResult.Results.Count -gt 0) {
        $scanResult.Results | Export-Csv -Path $config.OutputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "CSV exported: $($config.OutputPath)"
        Write-Host "Total records: $($scanResult.Results.Count)"
    } else {
        Write-Host "No files found" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Completed successfully" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}
