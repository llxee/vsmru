# vsmru - Visual Studio Most Recently Used Projects CLI

PowerShell 7 tool/module for viewing and managing Visual Studio's most recently used projects.

## Files

```
vsmru/
  vsmru.ps1    # standalone script (direct execution)
  vsmru.psm1   # module script (paired with psd1)
  vsmru.psd1   # module manifest
```

## Installation

### Option 1: Module (Recommended)

Copy the `vsmru` folder to any directory in `$env:PSModulePath`:

```powershell
# Copy to user module directory
Copy-Item .\vsmru $env:PSModulePath.Split(';')[0]\vsmru -Recurse

# Import
Import-Module vsmru

# Then use vsmru directly (PowerShell auto-loads on subsequent use)
vsmru
```

The module exports the `Get-VSMRU` function and the `vsmru` alias.

### Option 2: Standalone Script

```powershell
.\vsmru.ps1
```

Or add to PATH:
```powershell
$env:Path += ";D:\MiscProjects\vsmru"
```

Requires PowerShell 7+ (`pwsh`).

## Usage

```
vsmru [<search-term>] [-c <count>] [-d]
vsmru [-s <term>] [-c <count>] [-d]
<term> | vsmru [-c <count>]
```

### Show recent projects (default 5)

```powershell
vsmru
```

### Specify count

```powershell
vsmru -c 10
```

### Search

```powershell
vsmru -s Arctic
"TicTacToe" | vsmru
```

### Interactive delete

```powershell
vsmru -d
```

### Help

```powershell
vsmru -h
vsmru --help
vsmru -?
```

## Parameters

| Parameter | Alias | Description |
|-----------|-------|-------------|
| `-Search <string>` | `-s` | Search term (matches name or path, wildcards supported) |
| `-Count <int>` | `-c` | Number of entries to show (default: 5) |
| `-Delete` | `-d` | Interactive delete mode |
| `-Help` | `-h` | Show help |

## MRU Data Sources

| VS Version | Source |
|------------|--------|
| VS 2022+ | `%LocalAppData%\Microsoft\VisualStudio\<ver>_<id>\ApplicationPrivateSettings.xml` → `CodeContainers.Offline` |
| VS 2017/2019 | Registry `HKCU:\Software\Microsoft\VisualStudio\<ver>\ProjectMRUList` |
