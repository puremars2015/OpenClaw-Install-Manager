param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('status', 'install-prerequisites', 'install-openclaw', 'install-nanobot', 'uninstall-openclaw', 'stop-gateway', 'stop-nanobot')]
    [string]$Action
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

function Invoke-CommandAndCapture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [string[]]$Arguments = @()
    )

    $originalPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & $FilePath @Arguments 2>&1
        $exitCode = if ($null -ne $LASTEXITCODE) { $LASTEXITCODE } else { 0 }
        return [pscustomobject]@{
            Output = ($output | Out-String).Trim()
            ExitCode = $exitCode
        }
    }
    finally {
        $ErrorActionPreference = $originalPreference
    }
}

function Get-VersionInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandName,
        [string[]]$Arguments = @('--version')
    )

    $command = Get-Command $CommandName -ErrorAction SilentlyContinue
    if (-not $command) {
        return [ordered]@{
            installed = $false
            version = $null
            path = $null
        }
    }

    $result = Invoke-CommandAndCapture -FilePath $command.Source -Arguments $Arguments
    return [ordered]@{
        installed = $true
        version = $result.Output
        path = $command.Source
    }
}

function Get-NpmGlobalPrefix {
    $npm = Get-Command npm -ErrorAction SilentlyContinue
    if (-not $npm) {
        return $null
    }

    $result = Invoke-CommandAndCapture -FilePath $npm.Source -Arguments @('prefix', '-g')
    if ($result.ExitCode -ne 0) {
        return $null
    }

    return $result.Output.Trim()
}

function Resolve-OpenClawPath {
    $prefix = Get-NpmGlobalPrefix
    if ($prefix) {
        $cmdPath = Join-Path $prefix 'openclaw.cmd'
        if (Test-Path $cmdPath) {
            return $cmdPath
        }
    }

    $command = Get-Command openclaw -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    return $null
}

function Get-KnownPythonPathCandidates {
    $patterns = @()

    if ($env:LOCALAPPDATA) {
        $patterns += (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python*\python.exe')
    }
    if ($env:ProgramFiles) {
        $patterns += (Join-Path $env:ProgramFiles 'Python*\python.exe')
    }
    if (${env:ProgramFiles(x86)}) {
        $patterns += (Join-Path ${env:ProgramFiles(x86)} 'Python*\python.exe')
    }

    $matches = foreach ($pattern in $patterns) {
        Get-ChildItem -Path $pattern -File -ErrorAction SilentlyContinue
    }

    return @($matches | Sort-Object FullName -Descending | ForEach-Object { $_.FullName })
}

function Get-WorkingPythonCommand {
    $candidates = @()

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        $candidates += [pscustomobject]@{
            FilePath = $py.Source
            BaseArguments = @('-3')
            Display = "$($py.Source) -3"
        }
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        $candidates += [pscustomobject]@{
            FilePath = $python.Source
            BaseArguments = @()
            Display = $python.Source
        }
    }

    foreach ($pythonPath in Get-KnownPythonPathCandidates) {
        $candidates += [pscustomobject]@{
            FilePath = $pythonPath
            BaseArguments = @()
            Display = $pythonPath
        }
    }

    foreach ($candidate in $candidates) {
        if (-not (Test-Path $candidate.FilePath)) {
            continue
        }

        $versionResult = Invoke-CommandAndCapture -FilePath $candidate.FilePath -Arguments ($candidate.BaseArguments + @('--version'))
        if ($versionResult.ExitCode -eq 0 -and $versionResult.Output) {
            return [ordered]@{
                filePath = $candidate.FilePath
                baseArguments = @($candidate.BaseArguments)
                display = $candidate.Display
                version = $versionResult.Output
            }
        }
    }

    return $null
}

function Get-NanoBotRoot {
    $basePath = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { Join-Path $env:USERPROFILE 'AppData\Local' }
    return Join-Path $basePath 'OpenClawManager\NanoBot'
}

function Get-NanoBotVenvPath {
    return Join-Path (Get-NanoBotRoot) 'venv'
}

function Get-NanoBotPythonPath {
    return Join-Path (Get-NanoBotVenvPath) 'Scripts\python.exe'
}

function Get-NanoBotInfo {
    $venvPath = Get-NanoBotVenvPath
    $venvPython = Get-NanoBotPythonPath

    if (-not (Test-Path $venvPython)) {
        return [ordered]@{
            installed = $false
            version = $null
            path = $null
            pythonPath = $null
            environmentPath = $venvPath
        }
    }

    $showResult = Invoke-CommandAndCapture -FilePath $venvPython -Arguments @('-m', 'pip', 'show', 'nanobot')
    if ($showResult.ExitCode -ne 0 -or -not $showResult.Output) {
        return [ordered]@{
            installed = $false
            version = $null
            path = $null
            pythonPath = $venvPython
            environmentPath = $venvPath
        }
    }

    $versionLine = ($showResult.Output -split "`r?`n" | Where-Object { $_ -like 'Version:*' } | Select-Object -First 1)
    $version = if ($versionLine) { $versionLine -replace '^Version:\s*', '' } else { $null }

    return [ordered]@{
        installed = $true
        version = $version
        path = $venvPython
        pythonPath = $venvPython
        environmentPath = $venvPath
    }
}

function Convert-VersionSafely {
    param([string]$Value)

    if (-not $Value) {
        return $null
    }

    if ($Value -match '(\d+\.\d+\.\d+)') {
        return [version]$Matches[1]
    }

    return $null
}

function Test-VersionAtLeast {
    param(
        [string]$Actual,
        [string]$Minimum
    )

    $actualVersion = Convert-VersionSafely $Actual
    if (-not $actualVersion) {
        return $false
    }

    return $actualVersion -ge ([version]$Minimum)
}

function Get-GatewayProcesses {
    $processes = Get-CimInstance Win32_Process | Where-Object {
        $_.CommandLine -and (
            $_.CommandLine -match 'openclaw(\.cmd|\.mjs)?\s+gateway' -or
            $_.CommandLine -match 'node(.exe)? .*openclaw.*gateway'
        )
    }

    return $processes
}

function Get-NanoBotProcesses {
    $venvPython = Get-NanoBotPythonPath
    $escapedVenvPython = if ($venvPython) { [regex]::Escape($venvPython) } else { $null }

    $processes = Get-CimInstance Win32_Process | Where-Object {
        if (-not $_.CommandLine -or $_.ProcessId -eq $PID) {
            return $false
        }

        if ($_.CommandLine -match '(?i)(^|\s|["''])-m\s+nanobot(\s|$)') {
            return $true
        }

        if ($escapedVenvPython -and $_.CommandLine -match $escapedVenvPython -and $_.CommandLine -match '(?i)nanobot') {
            return $true
        }

        return $false
    }

    return $processes
}

function Test-WingetInstalled {
    return [bool](Get-Command winget -ErrorAction SilentlyContinue)
}

function Test-WingetNoApplicableUpgrade {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ExitCode,
        [string]$Output = ''
    )

    if ($ExitCode -eq -1978335189) {
        return $true
    }

    return $false
}

function Invoke-WingetPackage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageId,
        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    if (-not (Test-WingetInstalled)) {
        throw 'winget is required to install dependencies automatically.'
    }

    Write-Host "=== $DisplayName ($PackageId) ==="
    $listResult = Invoke-CommandAndCapture -FilePath 'winget' -Arguments @('list', '--id', $PackageId, '--exact', '--accept-source-agreements')
    $isInstalled = $listResult.ExitCode -eq 0 -and $listResult.Output -match [regex]::Escape($PackageId)

    if ($isInstalled) {
        $upgradeResult = Invoke-CommandAndCapture -FilePath 'winget' -Arguments @('upgrade', '--id', $PackageId, '--exact', '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity')
        if ($upgradeResult.Output) {
            Write-Host $upgradeResult.Output
        }
        if ($upgradeResult.ExitCode -ne 0 -and -not (Test-WingetNoApplicableUpgrade -ExitCode $upgradeResult.ExitCode -Output $upgradeResult.Output)) {
            throw "winget upgrade $PackageId failed."
        }
        return
    }

    & winget install --id $PackageId --exact --accept-package-agreements --accept-source-agreements --disable-interactivity
    if ($LASTEXITCODE -ne 0) {
        throw "winget install $PackageId failed."
    }
}

function Ensure-PythonInstalled {
    $pythonCommand = Get-WorkingPythonCommand
    if ($pythonCommand) {
        Write-Host "Python detected: $($pythonCommand.version)"
        return $pythonCommand
    }

    Write-Host 'Python was not found. Installing Python 3.12 via winget...'
    Invoke-WingetPackage -PackageId 'Python.Python.3.12' -DisplayName 'Python 3.12'

    $pythonCommand = Get-WorkingPythonCommand
    if ($pythonCommand) {
        Write-Host "Python installed: $($pythonCommand.version)"
        return $pythonCommand
    }

    throw 'Python installation completed but no usable python executable was found. Please reopen the app and try again.'
}

function Ensure-NanoBotEnvironment {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$PythonCommand
    )

    $nanobotRoot = Get-NanoBotRoot
    $venvPath = Get-NanoBotVenvPath
    $venvPython = Get-NanoBotPythonPath

    if (Test-Path $venvPython) {
        Write-Host "NanoBot virtual environment detected: $venvPath"
        return $venvPython
    }

    New-Item -ItemType Directory -Path $nanobotRoot -Force | Out-Null
    Write-Host "Creating NanoBot virtual environment: $venvPath"
    & $PythonCommand.filePath @($PythonCommand.baseArguments + @('-m', 'venv', $venvPath))
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $venvPython)) {
        throw 'Failed to create NanoBot virtual environment.'
    }

    return $venvPython
}

function Get-Status {
    $pwsh = Get-VersionInfo -CommandName 'pwsh'
    $node = Get-VersionInfo -CommandName 'node'
    $npm = Get-VersionInfo -CommandName 'npm'
    $git = Get-VersionInfo -CommandName 'git'
    $python = Get-VersionInfo -CommandName 'python'
    $py = Get-VersionInfo -CommandName 'py' -Arguments @('-3', '--version')
    $openClawPath = Resolve-OpenClawPath
    $nanobot = Get-NanoBotInfo

    $openclaw = if ($openClawPath) {
        $versionResult = Invoke-CommandAndCapture -FilePath $openClawPath -Arguments @('--version')
        [ordered]@{
            installed = $true
            version = $versionResult.Output
            path = $openClawPath
        }
    }
    else {
        [ordered]@{
            installed = $false
            version = $null
            path = $null
        }
    }

    $gatewayProcesses = @(Get-GatewayProcesses)
    $npmPrefix = Get-NpmGlobalPrefix

    [ordered]@{
        tools = [ordered]@{
            pwsh = $pwsh
            node = $node
            npm = $npm
            git = $git
            python = $python
            py = $py
            openclaw = $openclaw
            nanobot = $nanobot
            npmGlobalPrefix = $npmPrefix
            openclawResolvedPath = $openClawPath
            nanobotPythonPath = $nanobot.pythonPath
            nanobotEnvironmentPath = $nanobot.environmentPath
        }
        requirements = [ordered]@{
            wingetAvailable = (Test-WingetInstalled)
            pwsh7Installed = ($pwsh.installed -and (Test-VersionAtLeast -Actual $pwsh.version -Minimum '7.0.0'))
            nodeSatisfiesMinimum = ($node.installed -and (Test-VersionAtLeast -Actual $node.version -Minimum '22.16.0'))
            nodeRecommended = ($node.installed -and (Test-VersionAtLeast -Actual $node.version -Minimum '24.0.0'))
            pythonInstalled = ($python.installed -or $py.installed)
            nanobotEnvironmentReady = (Test-Path (Get-NanoBotPythonPath))
        }
        gateway = [ordered]@{
            running = ($gatewayProcesses.Count -gt 0)
            pids = @($gatewayProcesses | ForEach-Object { $_.ProcessId })
        }
    }
}

function Install-Prerequisites {
    Invoke-WingetPackage -PackageId 'Microsoft.PowerShell' -DisplayName 'PowerShell 7'
    Invoke-WingetPackage -PackageId 'OpenJS.NodeJS.LTS' -DisplayName 'Node.js LTS'
    Invoke-WingetPackage -PackageId 'Git.Git' -DisplayName 'Git'
    Invoke-WingetPackage -PackageId 'Python.Python.3.12' -DisplayName 'Python 3.12'
}

function Install-OpenClaw {
    $npm = Get-Command npm -ErrorAction SilentlyContinue
    if (-not $npm) {
        throw 'npm was not found. Install Node.js first.'
    }

    $status = Get-Status
    if (-not $status.requirements.nodeSatisfiesMinimum) {
        throw 'Node.js version is too old. Update to 22.16.0 or newer.'
    }

    & $npm.Source install -g openclaw@latest
    if ($LASTEXITCODE -ne 0) {
        throw 'npm install -g openclaw@latest failed.'
    }
}

function Install-NanoBot {
    $pythonCommand = Ensure-PythonInstalled
    $venvPython = Ensure-NanoBotEnvironment -PythonCommand $pythonCommand

    Write-Host 'Upgrading pip in NanoBot virtual environment...'
    & $venvPython -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to upgrade pip in the NanoBot virtual environment.'
    }

    Write-Host 'Installing NanoBot...'
    & $venvPython -m pip install --upgrade nanobot
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to install NanoBot.'
    }
}

function Uninstall-OpenClaw {
    $npm = Get-Command npm -ErrorAction SilentlyContinue
    if (-not $npm) {
        throw 'npm was not found. Cannot uninstall OpenClaw.'
    }

    & $npm.Source uninstall -g openclaw
    if ($LASTEXITCODE -ne 0) {
        throw 'npm uninstall -g openclaw failed.'
    }
}

function Stop-Gateway {
    $processes = @(Get-GatewayProcesses)
    if (-not $processes.Count) {
        Write-Host 'No running OpenClaw Gateway processes were detected.'
        return
    }

    $processes | ForEach-Object {
        Write-Host "Stopping PID $($_.ProcessId): $($_.CommandLine)"
        Stop-Process -Id $_.ProcessId -Force
    }
}

function Stop-NanoBot {
    $processes = @(Get-NanoBotProcesses)
    if (-not $processes.Count) {
        Write-Host 'No running NanoBot processes were detected.'
        return
    }

    $processes | Sort-Object ProcessId -Unique | ForEach-Object {
        Write-Host "Stopping PID $($_.ProcessId): $($_.CommandLine)"
        Stop-Process -Id $_.ProcessId -Force
    }
}

switch ($Action) {
    'status' {
        Get-Status | ConvertTo-Json -Depth 6
    }
    'install-prerequisites' {
        Install-Prerequisites
    }
    'install-openclaw' {
        Install-OpenClaw
    }
    'install-nanobot' {
        Install-NanoBot
    }
    'uninstall-openclaw' {
        Uninstall-OpenClaw
    }
    'stop-gateway' {
        Stop-Gateway
    }
    'stop-nanobot' {
        Stop-NanoBot
    }
}