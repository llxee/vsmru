#!/usr/bin/env pwsh
#requires -Version 7

param(
    [Parameter(ValueFromPipeline = $true, Position = 0)]
    [Alias('s')]
    [string]$Search,

    [ValidateRange(1, [int]::MaxValue)]
    [Alias('c')]
    [int]$n = 5,

    [Alias('d')]
    [switch]$Delete,

    [Alias('h')]
    [switch]$Help
)

begin {
    function Show-Help {
        Write-Host @"
vsmru - Visual Studio Most Recently Used Projects CLI

Usage:
    vsmru [<search-term>] [-c <count>] [-d]
    vsmru [-s <term>] [-c <count>] [-d]
    <term> | vsmru [-c <count>]

Parameters:
    -Search, -s <string>    Search term (matches name or path, wildcards supported)
    -Count, -c <int>        Number of entries to show (default: 5)
    -Delete, -d             Interactive delete mode
    -Help, -h, --help, -?   Show this help message

Examples:
    vsmru                    Show the 5 most recent VS projects
    vsmru -c 10              Show the 10 most recent
    vsmru -s Arctic          Search for projects matching "Arctic"
    "TicTacToe" | vsmru      Pipeline search
    vsmru -d                 Interactive delete mode

MRU data sources:
    VS 2022+ : %%LocalAppData%%\Microsoft\VisualStudio\<ver>_<id>\ApplicationPrivateSettings.xml
    VS 2017/9: HKCU:\Software\Microsoft\VisualStudio\<ver>\ProjectMRUList
"@
    }
    $script:showHelp = $Help

    if ($Search -match '^--(.+)$') {
        $flag = $Matches[1]
        if ($flag -in @('help', 'h', '?')) {
            $script:showHelp = $true
        } else {
            Write-Warning "Unknown option: --$flag"
            $script:showHelp = $true
        }
    }

    $i = 0
    while ($i -lt $args.Count) {
        $a = $args[$i]
        if ($a -match '^--?(.+)$') {
            $flag = $Matches[1]
            if ($flag -in @('h', 'help', '?')) {
                $script:showHelp = $true
            } else {
                Write-Warning "Unknown option: --$flag"
                $script:showHelp = $true
            }
        }
        $i++
    }

    if ($script:showHelp) {
        Show-Help
        exit 0
    }

    function Get-VSInstallations {
        $installs = [System.Collections.Generic.List[PSObject]]::new()

        $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
        if (Test-Path $vswhere) {
            try {
                $raw = & $vswhere -products * -format json -utf8 2>$null
                if ($raw) {
                    $parsed = $raw | ConvertFrom-Json
                    foreach ($vs in $parsed) {
                        $major = switch -Regex ($vs.catalog.productLineVersion) {
                            '2022' { 17 }; '2019' { 16 }; '2017' { 15 }
                            default { [int]($vs.installationVersion -replace '\..*$') }
                        }
                        $installs.Add([PSCustomObject]@{
                            VersionMajor = $major
                            DisplayName  = $vs.catalog.productLineVersion
                            InstanceId   = $vs.instanceId
                            InstallPath  = $vs.installationPath
                        })
                    }
                }
            } catch { Write-Debug "vswhere: $_" }
        }

        if ($installs.Count -eq 0) {
            $base = 'HKLM:\SOFTWARE\Microsoft\VisualStudio'
            if (Test-Path $base) {
                Get-ChildItem $base | ForEach-Object {
                    $verStr = $_.PSChildName
                    if ($verStr -match '^\d+\.0$' -and -not ($installs | Where-Object VersionMajor -eq ([int]($verStr -replace '\..*$')))) {
                        $major = [int]($verStr -replace '\..*$')
                        $installs.Add([PSCustomObject]@{
                            VersionMajor = $major
                            DisplayName  = "VS $major.0"
                            InstanceId   = $null
                            InstallPath  = (Get-ItemProperty "$($_.PSPath)\Setup\VS" -Name ProductDir -ErrorAction SilentlyContinue).ProductDir
                        })
                    }
                }
            }
        }

        return $installs | Sort-Object VersionMajor -Descending
    }

    function Get-AppPrivateSettingsPath {
        param([int]$VersionMajor, [string]$InstanceId)
        if ($VersionMajor -ge 17 -and $InstanceId) {
            $base = "$env:LocalAppData\Microsoft\VisualStudio"
            $pattern = "${VersionMajor}.0_${InstanceId}"
            $dir = Join-Path $base $pattern
            if (Test-Path $dir) {
                return Join-Path $dir "ApplicationPrivateSettings.xml"
            }
        }
        $base = "$env:LocalAppData\Microsoft\VisualStudio"
        $pattern = "${VersionMajor}.0_*"
        $dirs = Get-ChildItem $base -Directory -Filter $pattern -ErrorAction SilentlyContinue
        foreach ($d in $dirs) {
            $f = Join-Path $d.FullName "ApplicationPrivateSettings.xml"
            if (Test-Path $f) { return $f }
        }
        return $null
    }

    function Get-RegistryMRUPath {
        param([int]$VersionMajor, [string]$InstanceId)
        $baseKey = "HKCU:\Software\Microsoft\VisualStudio"
        if ($VersionMajor -ge 17) {
            $pattern = "$VersionMajor.0_*"
            $candidates = Get-ChildItem $baseKey -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -like $pattern }
            if ($InstanceId) {
                $exact = $candidates | Where-Object { $_.PSChildName -eq "$VersionMajor.0_$InstanceId" }
                if ($exact) {
                    $mru = Join-Path $exact.PSPath "ProjectMRUList"
                    if (Test-Path $mru) { return $mru }
                }
            }
            foreach ($c in $candidates) {
                $mru = Join-Path $c.PSPath "ProjectMRUList"
                if (Test-Path $mru) { return $mru }
            }
        }
        $path = "$baseKey\$VersionMajor.0\ProjectMRUList"
        if (Test-Path $path) { return $path }
        return $null
    }

    function Get-MRUEntriesFromXML {
        param([string]$XmlPath)
        if (-not (Test-Path $XmlPath)) { return @() }
        try {
            $xml = [System.Xml.XmlDocument]::new()
            $xml.Load($XmlPath)
            $collection = $xml.SelectSingleNode("//collection[@name='CodeContainers.Offline']")
            if (-not $collection) { return @() }
            $json = $collection.SelectSingleNode("value[@name='value']").'#text'
            if (-not $json) { return @() }
            $items = $json | ConvertFrom-Json
        } catch {
            Write-Debug "XML parse error: $_"
            return @()
        }
        $entries = [System.Collections.Generic.List[PSObject]]::new()
        $idx = 0
        foreach ($item in $items) {
            $idx++
            $lastAccessed = if ($item.Value.LastAccessed) {
                try { [DateTime]::Parse($item.Value.LastAccessed) } catch { $null }
            } else { $null }
            $path = $item.Key
            $entries.Add([PSCustomObject]@{
                DisplayIndex = $idx
                Name         = [System.IO.Path]::GetFileName($path)
                Path         = $path
                LastModified = $lastAccessed
                JsonObject   = $item
            })
        }
        return $entries | Sort-Object LastModified -Descending
    }

    function Get-MRUEntriesFromRegistry {
        param([string]$RegPath)
        if (-not $RegPath) { return @() }
        try {
            $key = Get-Item -Path $RegPath -ErrorAction Stop
        } catch { return @() }
        $mruBytes = $key.GetValue('MRUListEx', $null)
        if (-not $mruBytes -or $mruBytes.Length -lt 4) { return @() }
        $order = [System.Collections.Generic.List[int]]::new()
        for ($i = 0; $i -le $mruBytes.Length - 4; $i += 4) {
            $val = [System.BitConverter]::ToInt32($mruBytes, $i)
            if ($val -eq -1) { break }
            $order.Add($val)
        }
        if ($order.Count -eq 0) { return @() }
        $valueNames = $key.GetValueNames() | Where-Object { $_ -ne 'MRUListEx' }
        $pathLookup = @{}
        foreach ($name in $valueNames) {
            $regIdx = $null
            if ($name -match '^[a-z]$') {
                $regIdx = [int][char]$name - [int][char]'a'
            } elseif ($name -match '^File(\d+)$') {
                $regIdx = [int]$Matches[1]
            }
            if ($null -ne $regIdx) {
                $pathLookup[$regIdx] = $key.GetValue($name)
            }
        }
        $entries = [System.Collections.Generic.List[PSObject]]::new()
        $displayIdx = 0
        foreach ($regIdx in $order) {
            $path = $pathLookup[$regIdx]
            if (-not $path) { continue }
            $displayIdx++
            $lastMod = if (Test-Path $path -PathType Leaf) {
                (Get-Item $path).LastWriteTime
            } else { $null }
            $entries.Add([PSCustomObject]@{
                DisplayIndex = $displayIdx
                RegIndex     = $regIdx
                Name         = [System.IO.Path]::GetFileName($path)
                Path         = $path
                LastModified = $lastMod
            })
        }
        return $entries
    }

    function Show-Entries {
        param([array]$Entries, [int]$MaxCount)
        if ($Entries.Count -eq 0) { Write-Host "No MRU entries found."; return }
        $count = [Math]::Min($MaxCount, $Entries.Count)
        $Entries[0..($count - 1)] |
            Select-Object @{N='Idx';E={$_.DisplayIndex}},
                          @{N='Name';E={$_.Name}},
                          @{N='LastModified';E={if ($_.LastModified) { $_.LastModified.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }}},
                          @{N='Path';E={$_.Path}}
    }

    function Remove-MRUEntryXML {
        param([PSObject]$Entry, [string]$XmlPath)
        $xml = [System.Xml.XmlDocument]::new()
        $xml.Load($XmlPath)
        $collection = $xml.SelectSingleNode("//collection[@name='CodeContainers.Offline']")
        if (-not $collection) { Write-Error "Collection not found in XML."; return $false }
        $jsonNode = $collection.SelectSingleNode("value[@name='value']")
        if (-not $jsonNode) { Write-Error "Value not found in XML."; return $false }
        $items = $jsonNode.'#text' | ConvertFrom-Json
        $targetPath = $Entry.Path
        $newItems = $items | Where-Object { $_.Key -ne $targetPath }
        if ($newItems.Count -eq $items.Count) {
            Write-Host "Entry not found in XML data."
            return $false
        }
        $newJson = $newItems | ConvertTo-Json -Compress -Depth 10
        $jsonNode.'#text' = $newJson
        $xml.Save($XmlPath)
        Write-Host "Deleted: $($Entry.Name)"
        return $true
    }

    function Remove-MRUEntryRegistry {
        param([PSObject]$Entry, [string]$RegPath)
        $realPath = $RegPath -replace '^HKCU:\\', ''
        $realKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($realPath, $true)
        if (-not $realKey) { Write-Error "Cannot open registry key."; return $false }
        try {
            $valueName = [char]([int][char]'a' + $Entry.RegIndex)
            $existing = $realKey.GetValueNames() | Where-Object { $_ -eq $valueName }
            if (-not $existing) {
                $valueName = "File$($Entry.RegIndex)"
                $existing = $realKey.GetValueNames() | Where-Object { $_ -eq $valueName }
            }
            if (-not $existing) { Write-Error "Value not found."; return $false }
            $realKey.DeleteValue($valueName, $false)
            $mruBytes = $realKey.GetValue('MRUListEx')
            $indices = [System.Collections.Generic.List[int]]::new()
            for ($i = 0; $i -le $mruBytes.Length - 4; $i += 4) {
                $val = [System.BitConverter]::ToInt32($mruBytes, $i)
                if ($val -eq -1) { break }
                if ($val -ne $Entry.RegIndex) { $indices.Add($val) }
            }
            $indices.Add(-1)
            $newBytes = [System.Collections.Generic.List[byte]]::new()
            foreach ($val in $indices) { $newBytes.AddRange([System.BitConverter]::GetBytes($val)) }
            $realKey.SetValue('MRUListEx', $newBytes.ToArray(), [Microsoft.Win32.RegistryValueKind]::Binary)
            Write-Host "Deleted: $($Entry.Name)"
            return $true
        } catch { Write-Error "Failed: $_"; return $false }
        finally { if ($realKey) { $realKey.Close() } }
    }

    $script:searchTerm = $Search
}

process {
    if ($script:showHelp) { return }
    if ($_) { $script:searchTerm = $_ }
}

end {
    if ($script:showHelp) { return }

    $vsList = Get-VSInstallations
    if ($vsList.Count -eq 0) { Write-Error "No Visual Studio installation detected."; exit 1 }
    if ($vsList.Count -gt 1) {
        Write-Warning "Multiple VS installations detected: <$($vsList.DisplayName -join ', ')>"
    }
    $targetVS = $vsList[0]
    Write-Host "Using: $($targetVS.DisplayName) ($($targetVS.InstallPath))`n"

    $allEntries = @()
    $dataPath = $null
    $isRegistry = $false

    if ($targetVS.VersionMajor -ge 17) {
        $xmlPath = Get-AppPrivateSettingsPath -VersionMajor $targetVS.VersionMajor -InstanceId $targetVS.InstanceId
        if ($xmlPath) {
            $allEntries = Get-MRUEntriesFromXML -XmlPath $xmlPath
            $dataPath = $xmlPath
        }
    }

    if ($allEntries.Count -eq 0) {
        $regPath = Get-RegistryMRUPath -VersionMajor $targetVS.VersionMajor -InstanceId $targetVS.InstanceId
        if ($regPath) {
            $allEntries = Get-MRUEntriesFromRegistry -RegPath $regPath
            $dataPath = $regPath
            $isRegistry = $true
        }
    }

    if ($allEntries.Count -eq 0) {
        Write-Host "No MRU entries found."
        if ($targetVS.VersionMajor -ge 17) {
            Write-Host "Tip: Open a project in VS to populate the MRU list."
        }
        exit 0
    }

    if ($script:searchTerm) {
        $term = $script:searchTerm
        $allEntries = $allEntries | Where-Object {
            $_.Name -like "*$term*" -or $_.Path -like "*$term*"
        }
        if ($allEntries.Count -eq 0) {
            Write-Host "No entries matching '$term'."
            exit 0
        }
        $allEntries = $allEntries | Sort-Object LastModified -Descending
        $idx = 0
        $allEntries | ForEach-Object { $idx++; $_.DisplayIndex = $idx }
    }

    if ($Delete) {
        Show-Entries -Entries $allEntries -MaxCount $allEntries.Count
        Write-Host ""
        $userInput = Read-Host "Enter index to delete (or 'q' to cancel)"
        if ($userInput -eq 'q') { exit 0 }
        if (-not [int]::TryParse($userInput, [ref]$null)) {
            Write-Host "Invalid input. Cancelled."; exit 1
        }
        $selIdx = [int]$userInput
        $target = $allEntries | Where-Object { $_.DisplayIndex -eq $selIdx }
        if (-not $target) {
            Write-Host "Index $selIdx not found. Cancelled."; exit 1
        }
        $conf = Read-Host "Delete '$($target.Name)'? [y/N]"
        if ($conf -ne 'y') { Write-Host "Cancelled."; exit 0 }
        if ($isRegistry) {
            Remove-MRUEntryRegistry -Entry $target -RegPath $dataPath
        } else {
            Remove-MRUEntryXML -Entry $target -XmlPath $dataPath
        }
    } else {
        Show-Entries -Entries $allEntries -MaxCount $n
    }
}
