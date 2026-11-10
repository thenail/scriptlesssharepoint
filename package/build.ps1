# Builds the Microsoft Store package: ScriptlessSharePoint.exe (the app:
# WinUI 3 with the Windows App SDK inside it), scriptless-sharepoint-mcp.exe
# (the MCP server), their shared .NET runtime (self-contained, x64), the
# tiles and the MCP bundle manifest, as one MSIX.
#
#   ./package/build.ps1             build package/dist/package/ and the .msix
#   ./package/build.ps1 -Sign       also sign the .msix with a local test
#                                   certificate (to install it outside the Store)
#   ./package/build.ps1 -Register   register package/dist/package/ for this user
#                                   (needs Developer Mode; no signing), to try it
#                                   as installed
#
# The Store identity and built-in app ID come from identity.json, or
# SSP_STORE_NAME, SSP_STORE_PUBLISHER, SSP_STORE_PUBLISHER_NAME and
# SSP_CLIENT_ID. Windows only; needs the .NET 10 SDK and the Windows SDK.
param(
    [switch]$Sign,
    [switch]$Register
)
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$root = Split-Path $here
$dist = Join-Path $here 'dist'
$layout = Join-Path $dist 'package'
$publish = Join-Path $dist 'publish'

function Step([string]$s) { Write-Host "`n▸ $s" -ForegroundColor Cyan }
function Invoke([string]$exe, [string[]]$argv) {
    & $exe @argv
    if ($LASTEXITCODE -ne 0) { throw "$(Split-Path $exe -Leaf) failed ($LASTEXITCODE)" }
}
function Xml([string]$s) { [System.Security.SecurityElement]::Escape($s) }

# The newest Windows SDK's x64 copy of a tool.
function SdkTool([string]$name) {
    $bin = 'C:\Program Files (x86)\Windows Kits\10\bin'
    $found = Get-ChildItem $bin -Directory -ErrorAction SilentlyContinue |
        Where-Object Name -match '^10\.' |
        Sort-Object { [version]$_.Name } -Descending |
        ForEach-Object { Join-Path $_.FullName "x64\$name" } |
        Where-Object { Test-Path $_ } |
        Select-Object -First 1
    if (-not $found) { throw "$name not found: install the Windows SDK (winget install Microsoft.WindowsSDK.10.0.26100)." }
    $found
}

# 0. version and identity
$version = ([xml](Get-Content (Join-Path $root 'Directory.Build.props'))).Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
if ($version -notmatch '^\d+\.\d+\.\d+$') { throw "Version $version in Directory.Build.props isn't MAJOR.MINOR.PATCH." }
# MSIX versions have four parts, and the Store needs the last to be 0
$msixVersion = "$version.0"
$identity = Get-Content (Join-Path $here 'identity.json') -Raw | ConvertFrom-Json
if ($env:SSP_STORE_NAME) { $identity.name = $env:SSP_STORE_NAME }
if ($env:SSP_STORE_PUBLISHER) { $identity.publisher = $env:SSP_STORE_PUBLISHER }
if ($env:SSP_STORE_PUBLISHER_NAME) { $identity.publisherName = $env:SSP_STORE_PUBLISHER_NAME }
$clientId = if ($env:SSP_CLIENT_ID) { $env:SSP_CLIENT_ID } else { $identity.clientId }
if ($clientId -and -not [guid]::TryParse($clientId, [ref][guid]::Empty)) { throw "clientId '$clientId' isn't a GUID." }
Write-Host "Scriptless SharePoint $msixVersion, $($identity.name)"
Write-Host ($(if ($clientId) { "  Built-in app ID (override): $clientId" } else { '  Built-in app ID: the one in ScriptlessSharePoint.Core.csproj' }))

# 1. the executables, self-contained (the .NET runtime, and the Windows App
#    SDK for the app), so the package needs nothing else installed
Remove-Item $publish -Recurse -Force -ErrorAction SilentlyContinue
foreach ($p in @(
        @{ Project = 'src\ScriptlessSharePoint.Mcp'; Out = 'mcp' },
        @{ Project = 'src\ScriptlessSharePoint'; Out = 'app' })) {
    Step "dotnet publish $($p.Project)"
    $publishArgs = @('publish', (Join-Path $root $p.Project), '-c', 'Release', '-r', 'win-x64', '--self-contained', 'true',
        '-p:Platform=x64', '-p:DebugType=none', '-o', (Join-Path $publish $p.Out), '--nologo', '-v', 'quiet')
    # only an override: an empty -p would replace the project's built-in ID
    if ($clientId) { $publishArgs += "-p:SspClientId=$clientId" }
    Invoke 'dotnet' $publishArgs
}

# 2. the package folder: both publishes in one folder (they share the runtime)
Step 'package folder'
Remove-Item $layout -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $layout | Out-Null
Copy-Item (Join-Path $publish 'mcp\*') $layout -Recurse
$conflicts = @()
Get-ChildItem (Join-Path $publish 'app') -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring((Join-Path $publish 'app').Length + 1)
    $target = Join-Path $layout $rel
    if ((Test-Path $target) -and (Get-FileHash $target).Hash -ne (Get-FileHash $_.FullName).Hash) { $conflicts += $rel }
    New-Item -ItemType Directory -Force (Split-Path $target) | Out-Null
    Copy-Item $_.FullName $target -Force
}
if ($conflicts) { throw "The app and the MCP server ship different copies of: $($conflicts -join ', '). Give them the same package versions." }

& (Join-Path $here 'icons.ps1') -Assets (Join-Path $layout 'Assets') -McpIcon (Join-Path $layout 'mcp\icon.png') | Out-Null
$manifestJson = (& (Join-Path $layout 'scriptless-sharepoint-mcp.exe') --manifest) -join "`n"
if ($LASTEXITCODE -ne 0) { throw 'scriptless-sharepoint-mcp --manifest failed' }
[System.IO.File]::WriteAllText((Join-Path $layout 'mcp\manifest.json'), $manifestJson, [System.Text.UTF8Encoding]::new($false))

$appx = (Get-Content (Join-Path $here 'AppxManifest.xml') -Raw).
    Replace('{{NAME}}', (Xml $identity.name)).
    Replace('{{PUBLISHER}}', (Xml $identity.publisher)).
    Replace('{{PUBLISHER_NAME}}', (Xml $identity.publisherName)).
    Replace('{{VERSION}}', $msixVersion)
[System.IO.File]::WriteAllText((Join-Path $layout 'AppxManifest.xml'), $appx, [System.Text.UTF8Encoding]::new($false))

# 3. resources.pri: lets Windows pick the right scale of each tile, and
#    merges the app's own .pri files (its XAML, WinUI's controls) into it,
#    since a packaged app reads its resources from resources.pri
Step 'resources.pri (makepri)'
$makepri = SdkTool 'makepri.exe'
$priconfig = Join-Path $dist 'priconfig.xml'
Invoke $makepri @('createconfig', '/cf', $priconfig, '/dq', 'en-US', '/pv', '10.0.0', '/o') | Out-Null
# the tiles as files (Assets\), and the app's and WinUI's own .pri files
# merged in, one index each (they list their own files: indexing those
# again from the folder would make conflicting resources)
$config = [xml](Get-Content $priconfig)
# one package, not a bundle: every language and scale stays in resources.pri
[void]$config.resources.RemoveChild($config.resources.packaging)
$files = $config.resources.index
$template = $files.Clone()
$files.SetAttribute('startIndexAt', '\Assets')
foreach ($i in @($files.'indexer-config')) { if ($i.type -ne 'folder') { [void]$files.RemoveChild($i) } }
Get-ChildItem $layout -Filter *.pri | Where-Object Name -notlike 'resources*' | ForEach-Object {
    $pri = $template.Clone()
    $pri.SetAttribute('startIndexAt', "\$($_.Name)")
    foreach ($i in @($pri.'indexer-config')) { if ($i.type -ne 'PRI') { [void]$pri.RemoveChild($i) } }
    [void]$config.resources.AppendChild($pri)
}
$config.Save($priconfig)
Invoke $makepri @('new', '/pr', $layout, '/cf', $priconfig, '/mn', (Join-Path $layout 'AppxManifest.xml'), '/of', (Join-Path $layout 'resources.pri'), '/o') | Out-Null

if ($Register) {
    Step 'register for this user (Developer Mode)'
    Add-AppxPackage -Register (Join-Path $layout 'AppxManifest.xml') -ForceApplicationShutdown
    Write-Host "`nRegistered: start Scriptless SharePoint from the Start menu, or run ``scriptless-sharepoint-mcp --version``."
    return
}

# 4. the .msix
Step 'MSIX (makeappx)'
$msix = Join-Path $dist "ScriptlessSharePoint_${msixVersion}_x64.msix"
$out = & (SdkTool 'makeappx.exe') pack /d $layout /p $msix /o /h SHA256
if ($LASTEXITCODE -ne 0) { $out | Write-Host; throw 'makeappx failed' }
Write-Host ("{0}  {1:N0} MB" -f $msix, ((Get-Item $msix).Length / 1MB))

if ($Sign) {
    Step 'sign with a local test certificate'
    $pfx = Join-Path $dist 'test-signing.pfx'
    $cer = Join-Path $dist 'test-signing.cer'
    $password = 'ssp-local-test'
    if (-not (Test-Path $pfx)) {
        $c = New-SelfSignedCertificate -Type Custom -Subject $identity.publisher -KeyUsage DigitalSignature `
            -FriendlyName 'Scriptless SharePoint local test' -CertStoreLocation Cert:\CurrentUser\My `
            -TextExtension @('2.5.29.37={text}1.3.6.1.5.5.7.3.3', '2.5.29.19={text}')
        Export-PfxCertificate -Cert $c -FilePath $pfx -Password (ConvertTo-SecureString $password -Force -AsPlainText) | Out-Null
        Export-Certificate -Cert $c -FilePath $cer | Out-Null
    }
    $out = & (SdkTool 'signtool.exe') sign /fd SHA256 /f $pfx /p $password $msix
    if ($LASTEXITCODE -ne 0) { $out | Write-Host; throw 'signtool failed' }
    Write-Host "Signed. To install outside the Store, trust $cer once (as administrator:"
    Write-Host "  Import-Certificate -FilePath `"$cer`" -CertStoreLocation Cert:\LocalMachine\TrustedPeople)"
    Write-Host 'then open the .msix.'
}
