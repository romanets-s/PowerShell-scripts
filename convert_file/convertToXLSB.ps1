Param(
    [string]$pathToInputFile,
    [system.collections.generic.dictionary[string,system.collections.generic.dictionary[string,string]]]$config
)
$excel = New-Object -ComObject excel.application
$excel.visible = $false
$excel.DisplayAlerts = $false
$errorMessage = ''
$pathToFileSF1 = ''
$pathToFileSF2 = ''

Try{
    $wb = $excel.workbooks.open($pathToInputFile, 0, 0, 5, $config['settings']['passwordWB'])
    $timeStamp = [System.DateTime]::Now.ToString('yyyy-MM-dd-HH-mm-ss')
    $pathToFileSF1 = [IO.Path]::Combine($config['settings']['tmpFolder'], [string]::Format($config['settings']['fileNameSF1'], [IO.Path]::GetFileNameWithoutExtension($pathToInputFile), $timeStamp))
    $pathToFileSF2 = [IO.Path]::Combine($config['settings']['tmpFolder'], [string]::Format($config['settings']['fileNameSF2'], [IO.Path]::GetFileNameWithoutExtension($pathToInputFile), $timeStamp))
    $wb.SaveAs($pathToFileSF1, 50)
    $wb.SaveAs($pathToFileSF2, 50, $config['settings']['passwordWB'])
    }
Catch{
    $errorMessage = $_.Exception.Message
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers | Out-Null
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable -Name excel
return $pathToFileSF1, $pathToFileSF2, $errorMessage