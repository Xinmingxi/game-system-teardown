# autosave.ps1 - 每 2 小时提交一次，有远端就顺带推送。
$ErrorActionPreference = 'Continue'
# 没有凭据时立即失败，不要在隐藏窗口里静默挂起
$env:GIT_TERMINAL_PROMPT = '0'
$repo = 'D:\Desktop\重要文件\游戏系统拆解'
$log = Join-Path $repo '.autosave.log'

function Log($m) {
    $line = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '  ' + $m
    Add-Content -LiteralPath $log -Value $line -Encoding utf8
}

if (-not (Test-Path -LiteralPath (Join-Path $repo '.git'))) {
    Log 'skip: 不是 git 仓库'
    exit 1
}

Set-Location $repo

$changes = @(git status --porcelain)
if ($changes.Count -eq 0) {
    Log 'no changes'
    exit 0
}

git add -A 2>&1 | Out-Null
$msg = 'auto: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm')
git commit -q -m $msg 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Log 'commit failed'
    exit 1
}
Log ('committed ' + $changes.Count + ' path(s)')

$remotes = @(git remote)
if ($remotes -contains 'origin') {
    git push -q origin HEAD 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { Log 'pushed' } else { Log 'push failed（下次重试）' }
}
exit 0
