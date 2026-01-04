param(
    [string]$DevExcel    = "..\dev\ezGrepView.xlsm",
    [string]$ClearnExcel = "..\excel\ezGrepView.xlsm",
    [string]$SrcDir      = "..\src"
)

Write-Host "=== VBA Export Tool ==="

#---------------------------------------
# Excel Busy 対策
#---------------------------------------
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

# ドライブレター絶対パス
$regexDrive = '^[A-Za-z]:[\\/].+'
# UNC パス
$regexUNC   = '^\\\\[^\\]+\\[^\\]+'

# 絶対パス判定
function IsAbsolutePath {
    param($path)
    return $path -match $regexDrive -or $path -match $regexUNC
}

function GetAbsolutePath {
    param($path)

	if (IsAbsolutePath($path)) { 
		return $path
	} else {
		$base = $PSScriptRoot
		$path = Join-Path $base $path
		return [System.IO.Path]::GetFullPath($path)
	}
}

#---------------------------------------
# Excel 起動
#---------------------------------------
$excel = New-Object -ComObject Excel.Application
$excel.Visible            = $false  # 非表示の方が安定
$excel.DisplayAlerts      = $false
$excel.EnableEvents       = $false
$excel.AutomationSecurity = 1       # マクロ有効

#---------------------------------------
# 絶対パスに変換
#---------------------------------------
$resolvedDevExcel    = GetAbsolutePath(Resolve-Path $DevExcel)
$resolvedClearnExcel = GetAbsolutePath($ClearnExcel)
$resolvedSrcDir      = GetAbsolutePath($SrcDir)

#---------------------------------------
# Workbook Open
#---------------------------------------
try {
    $workbook = $excel.Workbooks.Open($resolvedDevExcel)
    Write-Host "[OK] Workbook opened: $resolvedDevExcel"
}
catch {
    Write-Error "Workbook を開けません: $($_.Exception.Message)"
    $excel.Quit()
    exit 1
}

Wait-ExcelIdle $excel
Start-Sleep -Milliseconds 500

#---------------------------------------
# 出力フォルダ作成
#---------------------------------------
if (!(Test-Path $resolvedSrcDir)) {
    New-Item $resolvedSrcDir -ItemType Directory | Out-Null
}

#---------------------------------------
# VBA Export
#---------------------------------------
foreach ($comp in $workbook.VBProject.VBComponents) {

    Wait-ExcelIdle $excel

    $ext = switch ($comp.Type) {
        1   { ".bas" }   # vbext_ct_StdModule
        2   { ".cls" }   # vbext_ct_ClassModule
        100 { ".dcm" }   # vbext_ct_Document (ThisWorkbook / Sheet)
        default { $null }
    }

    if (-not $ext) {
        continue
    }

    $path = Join-Path $resolvedSrcDir ($comp.Name + $ext)

    try {
        $comp.Export($path)
        Write-Host "Exported: $($comp.Name)$ext"
        Start-Sleep -Milliseconds 300   # ★ 重要：COM 安定化
    }
    catch {
        Write-Warning "Export failed: $($comp.Name)"
        Write-Warning $_.Exception.Message
    }
}

#--------------------------------------------------
# 標準 / クラスモジュール削除
#--------------------------------------------------
foreach ($comp in @($workbook.VBProject.VBComponents)) {
    if ($comp.Type -in 1, 2) {
        Write-Host "Removing module: $($comp.Name)"
        $workbook.VBProject.VBComponents.Remove($comp)
    }
	elseif ($comp.Type -eq 100) {
        Write-Host "Removing module: $($comp.Name)"
        $vbcomps   = $Workbook.VBProject.VBComponents
        $docModule = $vbcomps.Item($comp.Name)
        $codModule = $docModule.CodeModule
        if ($codModule.CountOfLines -gt 0) {
            $codModule.DeleteLines(1, $codModule.CountOfLines)
        }
	}
}

#--------------------------------------------------
# すべてのワークシートを表示
#--------------------------------------------------
foreach ($sheet in $workbook.Sheets) {
    try {
        $sheet.Visible = -1   # xlSheetVisible
    }
    catch {
        Write-Warning "Failed to change visibility: $($sheet.Name)"
    }
}

#---------------------------------------
# 後処理
#---------------------------------------
$workbook.SaveAs($resolvedClearnExcel)
$workbook.Close($false)
$excel.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)    | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "=== VBA Export Completed ==="
