# PT 國考 115 年增量下載腳本
# 只下載 115 年新內容,不動 105-114 既有檔案
# 用法: PS> PowerShell -ExecutionPolicy Bypass -File .\download_pt_exams_115.ps1

param(
    [string]$BaseDir = "E:\OneDrive\物理治療國考"
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ProgressPreference = 'Continue'

$sortedDir = Join-Path $BaseDir "已分類"
$urlBase = "https://wwwq.moex.gov.tw/exam/wHandExamQandA_File.ashx"

$subjFolder = [ordered]@{
    'shen'  = '神經疾病物理治療學'
    'gu'    = '骨科疾病物理治療學'
    'xin'   = '心肺疾病與小兒疾病物理治療學'
    'ji'    = '物理治療基礎學'
    'gai'   = '物理治療學概論'
    'jishu' = '物理治療技術學'
}

$typeName = @{ 'Q' = '試題'; 'S' = '答案'; 'M' = '更正答案' }

# 115 年 2 場次(115020=第一次、115090=第二次),subject codes 沿用 114 格式 0701-0706
$exams = @(
    @{ year=115; sess=1; code='115020'; subjs=@{shen='0701';gu='0702';xin='0703';ji='0704';gai='0705';jishu='0706'}; m=@('shen','gu','xin','ji','gai','jishu') },
    @{ year=115; sess=2; code='115090'; subjs=@{shen='0701';gu='0702';xin='0703';ji='0704';gai='0705';jishu='0706'}; m=@('shen','gu','xin','ji','gai') }
)

$subjOrder = @('shen','gu','xin','ji','gai','jishu')

$total = 0
foreach ($exam in $exams) {
    foreach ($key in $subjOrder) {
        $types = @('Q','S')
        if ($exam.m -contains $key) { $types += 'M' }
        $total += $types.Count
    }
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  115 年增量下載" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "目標: $BaseDir\已分類\115\{科目}\"
Write-Host "預計檔案數: $total"
Write-Host ""

$downloaded = 0; $skipped = 0; $failed = 0
$failures = New-Object System.Collections.Generic.List[string]
$progress = 0
$startTime = Get-Date

foreach ($exam in $exams) {
    $sessionLabel = if ($exam.sess -eq 1) { '第一次' } else { '第二次' }
    $yearDir = Join-Path $sortedDir $exam.year.ToString()
    New-Item -ItemType Directory -Force -Path $yearDir | Out-Null

    foreach ($key in $subjOrder) {
        $folder = $subjFolder[$key]
        $subjDir = Join-Path $yearDir $folder
        New-Item -ItemType Directory -Force -Path $subjDir | Out-Null

        $s = $exam.subjs[$key]
        $types = @('Q','S')
        if ($exam.m -contains $key) { $types += 'M' }

        foreach ($t in $types) {
            $progress++
            $typeText = $typeName[$t]
            $finalName = "考選部_{0}_{1}_{2}{3}.pdf" -f $exam.year, $folder, $sessionLabel, $typeText
            $finalPath = Join-Path $subjDir $finalName
            $url = "{0}?t={1}&code={2}&c=311&s={3}&q=1" -f $urlBase, $t, $exam.code, $s

            Write-Progress -Activity "115 年下載" -Status "[$progress/$total] $folder $sessionLabel $typeText" -PercentComplete (($progress / $total) * 100)

            if (Test-Path $finalPath) { $skipped++; continue }

            $attempt = 0; $maxAttempts = 3; $success = $false
            while ($attempt -lt $maxAttempts -and -not $success) {
                $attempt++
                try {
                    Invoke-WebRequest -Uri $url -OutFile $finalPath -UseBasicParsing -TimeoutSec 60 `
                        -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

                    $bytes = [System.IO.File]::ReadAllBytes($finalPath)
                    if ($bytes.Length -lt 4 -or
                        $bytes[0] -ne 0x25 -or $bytes[1] -ne 0x50 -or $bytes[2] -ne 0x44 -or $bytes[3] -ne 0x46) {
                        throw "not a PDF ($($bytes.Length) bytes)"
                    }
                    $downloaded++; $success = $true
                } catch {
                    if ($attempt -lt $maxAttempts) { Start-Sleep -Seconds 2 }
                    else {
                        $failed++
                        $failures.Add("$folder $sessionLabel ${typeText}: $url -- $($_.Exception.Message)") | Out-Null
                        if (Test-Path $finalPath) { Remove-Item $finalPath -Force -ErrorAction SilentlyContinue }
                    }
                }
            }
            Start-Sleep -Milliseconds 200
        }
    }
}

Write-Progress -Activity "115 年下載" -Completed
$elapsed = (Get-Date) - $startTime

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  完成" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "下載成功 : $downloaded" -ForegroundColor Green
Write-Host "跳過(已存): $skipped"   -ForegroundColor Yellow
$failColor = if ($failed -eq 0) { 'Green' } else { 'Red' }
Write-Host "失敗     : $failed"     -ForegroundColor $failColor
Write-Host "耗時     : $([int]$elapsed.TotalMinutes) 分 $([int]$elapsed.Seconds) 秒"
Write-Host ""

if ($failures.Count -gt 0) {
    $logPath = Join-Path $BaseDir "download_failures_115.log"
    $failures | Out-File -FilePath $logPath -Encoding UTF8
    Write-Host "失敗清單: $logPath" -ForegroundColor Red
    Write-Host "(若失敗訊息說 not a PDF,可能考選部尚未上載該檔案,晚點再試)"
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Yellow
Write-Host "  按任意鍵關閉視窗..." -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
