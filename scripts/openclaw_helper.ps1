param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('status', 'install-prerequisites', 'install-openclaw', 'install-openclaw-latest', 'uninstall-openclaw')]
    [string]$Action,

    [string]$Version = '2026.4.1'
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

function Resolve-ToolPath {
    param(
        [string[]]$CommandNames = @(),
        [string[]]$LiteralPaths = @(),
        [string[]]$GlobPaths = @()
    )

    foreach ($commandName in $CommandNames) {
        $command = Get-Command $commandName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command -and $command.Source -and (Test-Path $command.Source)) {
            return $command.Source
        }
    }

    foreach ($path in $LiteralPaths) {
        if ($path -and (Test-Path $path)) {
            return $path
        }
    }

    foreach ($pattern in $GlobPaths) {
        $match = Get-ChildItem -Path $pattern -File -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1
        if ($match) {
            return $match.FullName
        }
    }

    return $null
}

function Get-VersionInfoFromPath {
    param(
        [string]$Path,
        [string[]]$Arguments = @('--version')
    )

    if (-not $Path -or -not (Test-Path $Path)) {
        return [ordered]@{
            installed = $false
            version = $null
            path = $null
        }
    }

    $result = Invoke-CommandAndCapture -FilePath $Path -Arguments $Arguments
    return [ordered]@{
        installed = ($result.ExitCode -eq 0)
        version = if ($result.ExitCode -eq 0) { $result.Output } else { $null }
        path = $Path
    }
}

function Resolve-PwshPath {
    return Resolve-ToolPath -CommandNames @('pwsh') -GlobPaths @(
        (Join-Path $env:ProgramFiles 'PowerShell\*\pwsh.exe')
    )
}

function Resolve-NodePath {
    return Resolve-ToolPath -CommandNames @('node') -LiteralPaths @(
        (Join-Path $env:ProgramFiles 'nodejs\node.exe')
    )
}

function Resolve-NpmPath {
    $literalPaths = @(
        (Join-Path $env:ProgramFiles 'nodejs\npm.cmd'),
        (Join-Path $env:ProgramFiles 'nodejs\npm')
    )

    if ($env:APPDATA) {
        $literalPaths += (Join-Path $env:APPDATA 'npm\npm.cmd')
    }

    return Resolve-ToolPath -CommandNames @('npm', 'npm.cmd') -LiteralPaths $literalPaths
}

function Resolve-GitPath {
    $literalPaths = @()

    if ($env:ProgramFiles) {
        $literalPaths += (Join-Path $env:ProgramFiles 'Git\cmd\git.exe')
        $literalPaths += (Join-Path $env:ProgramFiles 'Git\bin\git.exe')
    }
    if (${env:ProgramFiles(x86)}) {
        $literalPaths += (Join-Path ${env:ProgramFiles(x86)} 'Git\cmd\git.exe')
        $literalPaths += (Join-Path ${env:ProgramFiles(x86)} 'Git\bin\git.exe')
    }

    return Resolve-ToolPath -CommandNames @('git') -LiteralPaths $literalPaths
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

function Get-NpmGlobalPrefix {
    $npmPath = Resolve-NpmPath
    if (-not $npmPath) {
        return $null
    }

    $result = Invoke-CommandAndCapture -FilePath $npmPath -Arguments @('prefix', '-g')
    if ($result.ExitCode -ne 0) {
        return $null
    }

    return $result.Output.Trim()
}

function Resolve-OpenCodePath {
    $literalPaths = @()
    $prefix = Get-NpmGlobalPrefix
    if ($prefix) {
        $literalPaths += (Join-Path $prefix 'opencode.cmd')
        $literalPaths += (Join-Path $prefix 'opencode')
    }

    foreach ($path in $literalPaths) {
        if ($path -and (Test-Path $path)) {
            return $path
        }
    }

    return Resolve-ToolPath -CommandNames @('opencode', 'opencode.cmd')
}

function Resolve-OpenClawPath {
    $literalPaths = @()
    $prefix = Get-NpmGlobalPrefix
    if ($prefix) {
        $literalPaths += (Join-Path $prefix 'openclaw.cmd')
        $literalPaths += (Join-Path $prefix 'openclaw')
    }

    foreach ($path in $literalPaths) {
        if ($path -and (Test-Path $path)) {
            return $path
        }
    }

    return Resolve-ToolPath -CommandNames @('openclaw', 'openclaw.cmd')
}

function Test-WingetInstalled {
    return [bool](Get-Command winget -ErrorAction SilentlyContinue)
}

function Test-WingetPackageInstalled {
    param([Parameter(Mandatory = $true)][string]$PackageId)

    if (-not (Test-WingetInstalled)) {
        return $false
    }

    $result = Invoke-CommandAndCapture -FilePath 'winget' -Arguments @('list', '--id', $PackageId, '--exact', '--accept-source-agreements')
    return ($result.ExitCode -eq 0 -and $result.Output -match [regex]::Escape($PackageId))
}

function Install-WingetPackageIfMissing {
    param(
        [Parameter(Mandatory = $true)][string]$PackageId,
        [Parameter(Mandatory = $true)][string]$DisplayName
    )

    if (-not (Test-WingetInstalled)) {
        throw 'winget is required to install dependencies automatically.'
    }

    if (Test-WingetPackageInstalled -PackageId $PackageId) {
        Write-Host "$DisplayName 已安裝，略過。"
        return
    }

    Write-Host "安裝 $DisplayName ($PackageId)..."
    & winget install --id $PackageId --exact --scope machine --accept-package-agreements --accept-source-agreements --disable-interactivity
    if ($LASTEXITCODE -ne 0) {
        throw "winget install $PackageId failed."
    }
}

function Get-Status {
    $pwsh = Get-VersionInfoFromPath -Path (Resolve-PwshPath)
    $node = Get-VersionInfoFromPath -Path (Resolve-NodePath)
    $npm = Get-VersionInfoFromPath -Path (Resolve-NpmPath) -Arguments @('--version')
    $git = Get-VersionInfoFromPath -Path (Resolve-GitPath)

    $pythonCommand = Get-WorkingPythonCommand
    $python = if ($pythonCommand) {
        [ordered]@{
            installed = $true
            version = $pythonCommand.version
            path = $pythonCommand.display
        }
    }
    else {
        [ordered]@{
            installed = $false
            version = $null
            path = $null
        }
    }

    $opencode = Get-VersionInfoFromPath -Path (Resolve-OpenCodePath)
    $openclaw = Get-VersionInfoFromPath -Path (Resolve-OpenClawPath)

    [ordered]@{
        tools = [ordered]@{
            pwsh = $pwsh
            node = $node
            npm = $npm
            git = $git
            python = $python
            opencode = $opencode
            openclaw = $openclaw
        }
        requirements = [ordered]@{
            wingetAvailable = (Test-WingetInstalled)
            pwsh7Installed = $pwsh.installed
        }
    }
}

function Install-Dependency {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('powershell', 'nodejs', 'git', 'python', 'opencode')]
        [string]$Name
    )

    switch ($Name) {
        'powershell' {
            Install-WingetPackageIfMissing -PackageId 'Microsoft.PowerShell' -DisplayName 'PowerShell 7'
        }
        'nodejs' {
            Install-WingetPackageIfMissing -PackageId 'OpenJS.NodeJS.LTS' -DisplayName 'Node.js LTS'
        }
        'git' {
            Install-WingetPackageIfMissing -PackageId 'Git.Git' -DisplayName 'Git'
        }
        'python' {
            Install-WingetPackageIfMissing -PackageId 'Python.Python.3.12' -DisplayName 'Python 3.12'
        }
        'opencode' {
            $npmPath = Resolve-NpmPath
            if (-not $npmPath) {
                throw 'npm was not found. Install Node.js first.'
            }

            Write-Host 'Installing OpenCode (opencode-ai)...'
            & $npmPath install -g opencode-ai
            if ($LASTEXITCODE -ne 0) {
                throw 'npm install -g opencode-ai failed.'
            }
        }
    }
}

function Install-Prerequisites {
    $status = Get-Status

    if (-not $status.requirements.pwsh7Installed) {
        Install-Dependency -Name 'powershell'
    }
    if (-not $status.tools.node.installed -or -not $status.tools.npm.installed) {
        Install-Dependency -Name 'nodejs'
    }
    if (-not $status.tools.git.installed) {
        Install-Dependency -Name 'git'
    }
    if (-not $status.tools.python.installed) {
        Install-Dependency -Name 'python'
    }

    $status = Get-Status
    if (-not $status.tools.opencode.installed) {
        Install-Dependency -Name 'opencode'
    }
}

function Install-OpenClaw {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Version
    )

    $npmPath = Resolve-NpmPath
    if (-not $npmPath) {
        throw 'npm was not found. Install Node.js first.'
    }

    Write-Host "Installing OpenClaw version $Version..."
    & $npmPath install -g "openclaw@$Version"
    if ($LASTEXITCODE -ne 0) {
        throw "npm install -g openclaw@$Version failed."
    }
}

function Install-OpenClawLatest {
    $npmPath = Resolve-NpmPath
    if (-not $npmPath) {
        throw 'npm was not found. Install Node.js first.'
    }

    Write-Host 'Installing latest OpenClaw...'
    & $npmPath install -g openclaw
    if ($LASTEXITCODE -ne 0) {
        throw 'npm install -g openclaw failed.'
    }
}

function Uninstall-OpenClaw {
    $npmPath = Resolve-NpmPath
    if (-not $npmPath) {
        throw 'npm was not found. Install Node.js first.'
    }

    Write-Host 'Removing OpenClaw...'
    & $npmPath uninstall -g openclaw
    if ($LASTEXITCODE -ne 0) {
        throw 'npm uninstall -g openclaw failed.'
    }
}

switch ($Action) {
    'status' {
        Get-Status | ConvertTo-Json -Depth 4
    }
    'install-prerequisites' {
        Install-Prerequisites
    }
    'install-openclaw' {
        Install-OpenClaw -Version $Version
    }
    'install-openclaw-latest' {
        Install-OpenClawLatest
    }
    'uninstall-openclaw' {
        Uninstall-OpenClaw
    }
}