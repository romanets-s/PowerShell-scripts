Param(
[string] $fileName
)

$excel = New-Object -ComObject Excel.Application
$path = Get-Location
$res = [io.path]::combine($path, $fileName)

$WB = $excel.Workbooks.Open($res)
$sheets = $WB.Sheets
for($i = 1; $i -le $sheets.Count; $i++){
    $s = $sheets.Item($i)
    $s.UsedRange.Font.Name = "Verdana"
    $s.UsedRange.Font.Size = 9
}

$WB.Save()
$WB.Close()
$excel.Quit()
Stop-Process -processname EXCEL
