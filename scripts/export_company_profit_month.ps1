param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d{4}-(0[1-9]|1[0-2])$')]
    [string]$Month,

    [string]$SshAlias = 'cangfu_hk',
    [string]$RemoteAppRoot = '/www/wwwroot/woo-analysis',
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path (Split-Path $repoRoot -Parent) 'outputs\company-profit'
}
$outputRoot = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

$runtimeRoot = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\node'
$nodeExe = Join-Path $runtimeRoot 'bin\node.exe'
$nodeModules = Join-Path $runtimeRoot 'node_modules'
if (-not (Test-Path -LiteralPath $nodeExe)) {
    throw "未找到工作区 Node 运行时：$nodeExe"
}
if (-not (Test-Path -LiteralPath $nodeModules)) {
    throw "未找到工作区表格依赖：$nodeModules"
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("company-profit-" + [guid]::NewGuid().ToString('N'))
$resolvedTempParent = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
$resolvedTempRoot = [System.IO.Path]::GetFullPath($tempRoot)
if (-not $resolvedTempRoot.StartsWith($resolvedTempParent, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "临时目录不在系统临时目录内，拒绝继续：$resolvedTempRoot"
}
New-Item -ItemType Directory -Path $resolvedTempRoot | Out-Null

$token = [guid]::NewGuid().ToString('N')
$remoteJson = "/tmp/company-profit-$Month-$token.json"
$remotePycache = "/tmp/company-profit-pycache-$token"
$localJson = Join-Path $resolvedTempRoot "snapshot-$Month.json"
$localBuilder = Join-Path $resolvedTempRoot 'build_company_profit_workbook.mjs'
$localNodeModules = Join-Path $resolvedTempRoot 'node_modules'
$outputFile = Join-Path $outputRoot "公司经营月报-$Month.xlsx"

try {
    Copy-Item -LiteralPath (Join-Path $repoRoot 'tools\build_company_profit_workbook.mjs') -Destination $localBuilder
    New-Item -ItemType Junction -Path $localNodeModules -Target $nodeModules | Out-Null

    $remoteCommand = @"
set -eu
cd '$RemoteAppRoot'
if [ -x venv/bin/python ]; then PY=venv/bin/python
elif [ -x .venv/bin/python ]; then PY=.venv/bin/python
else PY=python3
fi
PYTHONPYCACHEPREFIX='$remotePycache' "`$PY" offline_company_profit_snapshot.py --month '$Month' --output '$remoteJson'
"@
    & ssh $SshAlias $remoteCommand
    if ($LASTEXITCODE -ne 0) {
        throw "远程只读快照生成失败，退出码：$LASTEXITCODE"
    }

    & scp "${SshAlias}:$remoteJson" $localJson
    if ($LASTEXITCODE -ne 0) {
        throw "下载离线快照失败，退出码：$LASTEXITCODE"
    }

    & $nodeExe $localBuilder --input $localJson --output $outputFile
    if ($LASTEXITCODE -ne 0) {
        throw "Excel 生成失败，退出码：$LASTEXITCODE"
    }
    if (-not (Test-Path -LiteralPath $outputFile)) {
        throw "Excel 生成命令完成，但未找到输出文件：$outputFile"
    }
    Write-Output $outputFile
}
finally {
    & ssh $SshAlias "set -eu; rm -f '$remoteJson'; case '$remotePycache' in /tmp/company-profit-pycache-*) rm -rf -- '$remotePycache' ;; *) exit 2 ;; esac" 2>$null
    if (Test-Path -LiteralPath $resolvedTempRoot) {
        $verifiedTemp = [System.IO.Path]::GetFullPath($resolvedTempRoot)
        if ($verifiedTemp.StartsWith($resolvedTempParent, [System.StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $verifiedTemp -Recurse -Force
        }
    }
}
