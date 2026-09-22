# 按版式计划从参考样板 deck 克隆页面，生成工作 deck 的空骨架。
# 用法：& clone_pages.ps1 -path_ref <reference-deck.pptx> -path_out <工作deck.pptx> -plan "1,2,5,6,6,7,8,9,10"
# plan 里的数字是参考 deck 的页号（版式编号，见 SKILL.md 版式表），可重复。
param(
    [Parameter(Mandatory = $true)][string]$path_ref,
    [Parameter(Mandatory = $true)][string]$path_out,
    [Parameter(Mandatory = $true)][string]$plan
)
$ErrorActionPreference = "Stop"

### check：结构性前提
if (-not (Test-Path $path_ref)) { Write-Output "❌ ERROR: 参考 deck 不存在 $path_ref"; exit 1 }
$lplan = @($plan -split "," | ForEach-Object { [int]$_.Trim() })
if ($lplan.Count -lt 1) { Write-Output "❌ ERROR: plan 为空"; exit 1 }
### to here

### prepare：以参考 deck 为底另存，保留母版与主题
Copy-Item $path_ref $path_out -Force
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open((Resolve-Path $path_out).Path, $false, $false, $false)
$n_ref = $pres.Slides.Count
Write-Output ("📁 参考版式页 " + $n_ref + " 页 ▶️ 计划生成 " + $lplan.Count + " 页")
### to here

### main：逐页 Duplicate 再 MoveTo 末尾；原始版式页最后统一删除
foreach ($n in $lplan) {
    if ($n -lt 1 -or $n -gt $n_ref) { Write-Output ("❌ 版式号越界：" + $n); $pres.Close(); $app.Quit(); exit 1 }
    $dup = $pres.Slides.Item($n).Duplicate()
    $dup.MoveTo($pres.Slides.Count)
    Write-Output ("  ✅ 版式 " + $n + " → 第 " + ($pres.Slides.Count - $n_ref) + " 页")
}
for ($i = $n_ref; $i -ge 1; $i--) { $pres.Slides.Item($i).Delete() }
### to here

$pres.Save()
$pres.Close()
$app.Quit()
Write-Output ("📊 完成：" + $path_out + " 共 " + $lplan.Count + " 页")
Write-Output "🎉 done"
