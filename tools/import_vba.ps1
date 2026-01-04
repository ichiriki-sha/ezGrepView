param(
    [string]$ExcelPath   = "..\excel\ezGrepView.xlsm",
    [string]$SrcDir      = "..\src"
)

Write-Host "=== VBA Import Tool ==="

#--------------------------------------------------
# Utility
#--------------------------------------------------
function Wait-ExcelIdle {
    param($excel)
    while ($true) {
        try {
            $null = $excel.Ready
            break
        } catch {
            Start-Sleep -Milliseconds 200
        }
    }
}

function Clean-VBAAttributes {
    param([string]$code)

    $lines = $code -split "`r?`n"
    $cleaned = @()
    $skip = $false

    foreach ($line in $lines) {
        if ($line -match '^BEGIN$') { $skip = $true; continue }
        if ($line -match '^END$')   { $skip = $false; continue }

        if (-not $skip) {
            if ($line -notmatch '^VERSION \d+\.\d+ CLASS$'      -and
                $line -notmatch '^Attribute VB_GlobalNameSpace' -and
                $line -notmatch '^Attribute VB_PredeclaredId'  -and
                $line -notmatch '^Attribute VB_Creatable'      -and
                $line -notmatch '^Attribute VB_Exposed'        -and
                $line -notmatch '^Attribute VB_Name') {
                $cleaned += $line
            }
        }
    }
    return ($cleaned -join "`r`n")
}

#--------------------------------------------------
# Type=100 Import (Sheet / ThisWorkbook)
#--------------------------------------------------
function Import-DocumentModule {
    param(
        $Workbook,
        [string]$FilePath,
        [string]$ModuleName
    )

    Write-Host "  [Type=100] Importing document module: $ModuleName"

    $vbcomps = $Workbook.VBProject.VBComponents

    # 仮 Import
    $imp = $vbcomps.Import($FilePath)

    try {
        # 既存 Document Module
        $orig = $vbcomps.Item($ModuleName)

        $srcMod = $imp.CodeModule
        $dstMod = $orig.CodeModule

        # 全削除 → コピー
        if ($dstMod.CountOfLines -gt 0) {
            $dstMod.DeleteLines(1, $dstMod.CountOfLines)
        }

        if ($srcMod.CountOfLines -gt 0) {
            $dstMod.AddFromString(
                $srcMod.Lines(1, $srcMod.CountOfLines)
            )
        }
    }
    finally {
        # 仮モジュール削除
        $vbcomps.Remove($imp)
    }
}

#--------------------------------------------------
# Excel 起動
#--------------------------------------------------
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$excel.AutomationSecurity = 1  # msoAutomationSecurityLow

$workbook = $excel.Workbooks.Open((Resolve-Path $ExcelPath))
Wait-ExcelIdle $excel

#--------------------------------------------------
# 標準 / クラスモジュール削除
#--------------------------------------------------
foreach ($comp in @($workbook.VBProject.VBComponents)) {
    if ($comp.Type -in 1, 2) {
        Write-Host "Removing module: $($comp.Name)"
        $workbook.VBProject.VBComponents.Remove($comp)
    }
}

#--------------------------------------------------
# Import 処理
#--------------------------------------------------
$SrcDir = (Resolve-Path $SrcDir).Path
$files = Get-ChildItem $SrcDir | Where-Object {
    $_.Extension -in ".bas", ".cls", ".dcm" 
}

foreach ($file in $files) {
    try {
        Write-Host "Importing: $($file.Name)"

        $raw = Get-Content $file.FullName -Raw -Encoding Default

        # モジュール名取得
        if ($raw -match 'Attribute VB_Name\s*=\s*"(.+?)"') {
            $modName = $Matches[1]
        } else {
            $modName = [IO.Path]::GetFileNameWithoutExtension($file.Name)
        }

        $code = Clean-VBAAttributes $raw

        # Type=100 判定
        $docComp = $null
        foreach ($c in $workbook.VBProject.VBComponents) {
            if ($c.Type -eq 100 -and $c.Name -eq $modName) {
                $docComp = $c
                break
            }
        }

        if ($docComp) {
            Import-DocumentModule `
                -Workbook   $workbook `
                -FilePath   $file.FullName `
                -ModuleName $modName
        }
        else {
            # 通常モジュール
            if ($file.Extension -eq ".bas") {
                $comp = $workbook.VBProject.VBComponents.Add(1)
            } else {
                $comp = $workbook.VBProject.VBComponents.Add(2)
            }

            $comp.Name = $modName
            $comp.CodeModule.AddFromString($code)
        }
    }
    catch {
        Write-Warning "Import failed: $($file.Name)"
        Write-Warning $_.Exception.Message
    }
}

#--------------------------------------------------
# 後処理
#--------------------------------------------------
$workbook.Save()
$workbook.Close($true)
$excel.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)    | Out-Null

[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "=== VBA Import Completed ==="
