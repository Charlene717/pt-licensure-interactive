# PT 國考歷屆試題下載與分類腳本
# 來源：考選部考畢試題查詢平臺 (wwwq.moex.gov.tw)
# 範圍：民國 105-114 年，20 場考試 × 6 科 ≈ 280 個 PDF
# 用法：右鍵此檔 → "用 PowerShell 執行" 或在 PowerShell 視窗中執行
#       PS> .\download_pt_exams.ps1

param(
    [string]$BaseDir = "E:\Nephron\物治國考"
)

# UTF-8 console output (避免中文亂碼)
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ProgressPreference = 'Continue'

$rawDir    = Join-Path $BaseDir "原始下載"
$sortedDir = Join-Path $BaseDir "已分類"
$urlBase   = "https://wwwq.moex.gov.tw/exam/wHandExamQandA_File.ashx"

New-Item -ItemType Directory -Force -Path $rawDir    | Out-Null
New-Item -ItemType Directory -Force -Path $sortedDir | Out-Null

# 科目 short-key → 中文資料夾名稱
$subjFolder = [ordered]@{
    'shen'  = '神經疾病物理治療學'
    'gu'    = '骨科疾病物理治療學'
    'xin'   = '心肺疾病與小兒疾病物理治療學'
    'ji'    = '物理治療基礎學'
    'gai'   = '物理治療學概論'
    'jishu' = '物理治療技術學'
}

# 檔案類型代碼 → 中文
$typeName = @{
    'Q' = '試題'
    'S' = '答案'
    'M' = '更正答案'
}

# 20 場考試清單
# subjs = 該年該場 6 科的 s 參數（科目 ID）
# m     = 該場有「更正答案」的科目
$exams = @(
    @{ year=105; sess=1; code='105020'; subjs=@{shen='44';gu='55';xin='66';ji='11';gai='22';jishu='33'}; m=@('gu','xin','ji','gai') },
    @{ year=105; sess=2; code='105100'; subjs=@{shen='44';gu='55';xin='66';ji='11';gai='22';jishu='33'}; m=@('shen','gu','xin','ji','gai') },
    @{ year=106; sess=1; code='106020'; subjs=@{shen='44';gu='55';xin='66';ji='11';gai='22';jishu='33'}; m=@('gu','xin','ji','gai') },
    @{ year=106; sess=2; code='106100'; subjs=@{shen='44';gu='55';xin='66';ji='11';gai='22';jishu='33'}; m=@('shen','gu','xin','gai') },
    @{ year=107; sess=1; code='107020'; subjs=@{shen='44';gu='55';xin='66';ji='11';gai='22';jishu='33'}; m=@('shen','ji') },
    @{ year=107; sess=2; code='107100'; subjs=@{shen='44';gu='55';xin='66';ji='11';gai='22';jishu='33'}; m=@('shen','gu','xin','ji','jishu') },
    @{ year=108; sess=1; code='108030'; subjs=@{shen='44';gu='55';xin='66';ji='11';gai='22';jishu='33'}; m=@('shen','jishu') },
    @{ year=108; sess=2; code='108100'; subjs=@{shen='44';gu='55';xin='66';ji='11';gai='22';jishu='33'}; m=@('gu','xin') },
    @{ year=109; sess=1; code='109020'; subjs=@{shen='44';gu='55';xin='66';ji='11';gai='22';jishu='33'}; m=@('gu','xin') },
    @{ year=109; sess=2; code='109100'; subjs=@{shen='11';gu='22';xin='33';ji='44';gai='55';jishu='66'}; m=@('shen','gu','ji','gai') },
    @{ year=110; sess=1; code='110020'; subjs=@{shen='11';gu='22';xin='33';ji='44';gai='55';jishu='66'}; m=@('gu','ji') },
    @{ year=110; sess=2; code='110101'; subjs=@{shen='11';gu='22';xin='33';ji='44';gai='55';jishu='66'}; m=@('gu','ji','gai','jishu') },
    @{ year=111; sess=1; code='111020'; subjs=@{shen='11';gu='22';xin='33';ji='44';gai='55';jishu='66'}; m=@('gu','ji') },
    @{ year=111; sess=2; code='111100'; subjs=@{shen='11';gu='22';xin='33';ji='44';gai='55';jishu='66'}; m=@('shen','ji','gai') },
    @{ year=112; sess=1; code='112020'; subjs=@{shen='11';gu='22';xin='33';ji='44';gai='55';jishu='66'}; m=@('gu','gai','jishu') },
    @{ year=112; sess=2; code='112100'; subjs=@{shen='11';gu='22';xin='33';ji='44';gai='55';jishu='66'}; m=@('shen','gu','ji') },
    @{ year=113; sess=1; code='113020'; subjs=@{shen='11';gu='22';xin='33';ji='44';gai='55';jishu='66'}; m=@('xin','ji') },
    @{ year=113; sess=2; code='113090'; subjs=@{shen='11';gu='22';xin='33';ji='44';gai='55';jishu='66'}; m=@('shen') },
    @{ year=114; sess=1; code='114020'; subjs=@{shen='0701';gu='0702';xin='0703';ji='0704';gai='0705';jishu='0706'}; m=@('gu','gai','jishu') },
    @{ year=114; sess=2; code='114090'; subjs=@{shen='0701';gu='0702';xin='0703';ji='0704';gai='0705';jishu='0706'}; m=@('shen','gu','xin','ji') }
)

$subjOrder = @('shen','gu','xin','ji','gai','jishu')

# 計算總數
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
Write-Host "  PT 國考歷屆試題下載 (105-114年)"        -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "目標目錄: $BaseDir"
Write-Host "預計檔案數: $total"
Write-Host ""

$downloaded = 0
$skipped    = 0
$failed     = 0
$failures   = New-Object System.Collections.Generic.List[string]
$progress   = 0
$startTime  = Get-Date

foreach ($exam in $exams) {
    $sessionLabel = if ($exam.sess -eq 1) { '第一次' } else { '第二次' }
    $yearStr = $exam.year.ToString()
    $yearDir = Join-Path $sortedDir $yearStr
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
            $rawName = "{0}_t{1}_c311_s{2}_q1.pdf" -f $exam.code, $t, $s
            $rawPath = Join-Path $rawDir $rawName
            $url = "{0}?t={1}&code={2}&c=311&s={3}&q=1" -f $urlBase, $t, $exam.code, $s

            $statusMsg = "[$progress/$total] $($exam.year) $sessionLabel - $folder - $typeText"
            Write-Progress -Activity "下載 PT 國考試題" -Status $statusMsg -PercentComplete (($progress / $total) * 100)

            if (Test-Path $finalPath) {
                $skipped++
                continue
            }

            $attempt = 0
            $maxAttempts = 3
            $success = $false
            while ($attempt -lt $maxAttempts -and -not $success) {
                $attempt++
                try {
                    Invoke-WebRequest -Uri $url -OutFile $rawPath -UseBasicParsing -TimeoutSec 60 `
                        -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

                    # 驗證是 PDF（前 4 byte 應是 %PDF）
                    $bytes = [System.IO.File]::ReadAllBytes($rawPath)
                    if ($bytes.Length -lt 4 -or
                        $bytes[0] -ne 0x25 -or $bytes[1] -ne 0x50 -or $bytes[2] -ne 0x44 -or $bytes[3] -ne 0x46) {
                        throw "Server did not return a PDF (got $($bytes.Length) bytes, header: $([System.Text.Encoding]::ASCII.GetString($bytes[0..[Math]::Min(15,$bytes.Length-1)])))"
                    }

                    Copy-Item -Path $rawPath -Destination $finalPath -Force
                    $downloaded++
                    $success = $true
                } catch {
                    if ($attempt -lt $maxAttempts) {
                        Start-Sleep -Seconds 2
                    } else {
                        $failed++
                        $msg = "[$($exam.year) $sessionLabel $folder $typeText] $url -- $($_.Exception.Message)"
                        $failures.Add($msg) | Out-Null
                        if (Test-Path $rawPath) { Remove-Item $rawPath -Force -ErrorAction SilentlyContinue }
                    }
                }
            }

            # 限速：每筆下載間隔 200ms 不要打爆伺服器
            Start-Sleep -Milliseconds 200
        }
    }
}

Write-Progress -Activity "下載 PT 國考試題" -Completed
$elapsed = (Get-Date) - $startTime

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  完成"                                     -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "下載成功 : $downloaded" -ForegroundColor Green
Write-Host "跳過(已存): $skipped"   -ForegroundColor Yellow
Write-Host "失敗     : $failed"    -ForegroundColor (if ($failed -eq 0) { 'Green' } else { 'Red' })
Write-Host "耗時     : $([int]$elapsed.TotalMinutes) 分 $([int]$elapsed.Seconds) 秒"
Write-Host ""

if ($failures.Count -gt 0) {
    $logPath = Join-Path $BaseDir "download_failures.log"
    $failures | Out-File -FilePath $logPath -Encoding UTF8
    Write-Host "失敗清單已寫入: $logPath" -ForegroundColor Red
}

Write-Host "原始檔: $rawDir"
Write-Host "已分類: $sortedDir"
