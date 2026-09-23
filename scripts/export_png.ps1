# PowerPoint COM 导出逐页 PNG，供 agent 目检版式（唯一可信的渲染真相）。
# 用法：& export_png.ps1 -path_deck <deck.pptx> -path_out <png目录>
param(
    [Parameter(Mandatory = $true)][string]$path_deck,
    [Parameter(Mandatory = $true)][string]$path_out
)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $path_deck)) { Write-Output "❌ ERROR: deck 不存在 $path_deck"; exit 1 }
New-Item -ItemType Directory -Force $path_out | Out-Null

# PowerPoint COM 是单实例：用户开着 PPT 时拿到的是同一个进程，Quit() 会连用户的文件一起关掉。
# 所以只关自己打开的这份，没有别的演示文稿时才 Quit。
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open((Resolve-Path $path_deck).Path, $true, $false, $false)
$pres.Export((Resolve-Path $path_out).Path, "PNG", 1600, 900)
$pres.Close()
if ($app.Presentations.Count -eq 0) { $app.Quit() }

$n = (Get-ChildItem $path_out -Filter *.PNG | Measure-Object).Count
Write-Output ("📊 导出 " + $n + " 张 → " + $path_out)
Write-Output "🎉 done"
