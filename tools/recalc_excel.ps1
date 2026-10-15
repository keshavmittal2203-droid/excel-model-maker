# Recalculate an XLSX with installed desktop Excel (COM) and save a new copy.
# Usage: pwsh -File tools/recalc_excel.ps1 -InputPath <in.xlsx> -OutputPath <out.xlsx>
# The input is opened read-only; output must differ from the original file.
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)
$ErrorActionPreference = 'Stop'
$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
$outputFull = Join-Path (Resolve-Path -LiteralPath (Split-Path -Parent $OutputPath)).Path (Split-Path -Leaf $OutputPath)
if ($inputFull -eq $outputFull) { throw 'Output must differ from the input file' }
$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
    $wb = $excel.Workbooks.Open($inputFull, 0, $true)
    $excel.CalculateFullRebuild()
    $wb.SaveAs($outputFull, 51)     # 51 = xlOpenXMLWorkbook
    $wb.Close($false)
    $wb = $null
    Write-Output ("RECALC_OK excel_version={0} output={1}" -f $excel.Version, $outputFull)
}
finally {
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    if ($excel -ne $null) { try { $excel.Quit() } catch {} }
}
