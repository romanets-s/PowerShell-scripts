Param(
    [string]$pathToInputFile,
    [system.collections.generic.dictionary[string, int32]]$indexOfHead,
    [system.collections.generic.dictionary[string, system.collections.generic.dictionary[string,string]]]$config
)
$excel = New-Object -ComObject excel.application
$excel.visible = $false
$excel.DisplayAlerts = $false
$errorMessage = ''

Try{
    $wb = $excel.workbooks.open($pathToInputFile, 0, 0, 5, $config['settings']['passwordWB'])
    $sh = $wb.Sheets(1)
    $sh.Range("A:Z").NumberFormat = "General"
    $sh.Range("B:E").NumberFormat = "MM/DD/YYYY"
    $sh.Range([char](65 + $indexOfHead["WeekStart"]) + ":" + [char](65 + $indexOfHead["WeekStart"])).NumberFormat = "MM/DD/YYYY"
    $sh.Range([char](65 + $indexOfHead["WeekEnd"]) + ":" + [char](65 + $indexOfHead["WeekEnd"])).NumberFormat = "MM/DD/YYYY"
    #$sh.Columns("W").NumberFormat = "YYYY/MM"
    $wb.Save()
    }
Catch{
    $errorMessage = $_.Exception.Message
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers | Out-Null
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable -Name excel
return $errorMessage