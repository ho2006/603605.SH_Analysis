<#
.SYNOPSIS
  通过 conda 环境顺序执行完整管线：数据落地与指标计算 -> 图表导出 -> Word/PDF 报告渲染。
  所有 Python 均经 `conda run -n <env>` 调用，不使用系统 python。

.EXAMPLE
  pwsh scripts/run_all.ps1                 # 默认环境 proya-fin
  pwsh scripts/run_all.ps1 -EnvName myenv
  pwsh scripts/run_all.ps1 -Step charts    # 只跑某一步：pipeline | charts | reports | all
#>
param(
    [string]$EnvName = 'proya-fin',
    [ValidateSet('all', 'pipeline', 'charts', 'reports')]
    [string]$Step = 'all'
)
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

if (-not (Get-Command conda -ErrorAction SilentlyContinue)) {
    throw "未找到 conda。请先安装 Miniconda/Miniforge 并执行 'conda init pwsh'，再运行: conda env create -f environment.yml"
}
$envs = conda env list | Out-String
if ($envs -notmatch "(?m)^\s*$([regex]::Escape($EnvName))\s") {
    throw "conda 环境 '$EnvName' 不存在。请先执行: conda env create -f environment.yml"
}

$env:PYTHONUTF8 = '1'
$steps = [ordered]@{
    pipeline = 'src/pipeline.py'
    charts   = 'src/charts.py'
    reports  = 'src/build_reports.py'
}
foreach ($name in $steps.Keys) {
    if ($Step -ne 'all' -and $Step -ne $name) { continue }
    Write-Host "==> [$name] conda run -n $EnvName python $($steps[$name])" -ForegroundColor Cyan
    conda run --no-capture-output -n $EnvName python $steps[$name]
    if ($LASTEXITCODE -ne 0) { throw "步骤 $name 失败 (exit $LASTEXITCODE)" }
}
Write-Host '==> 完成' -ForegroundColor Green
