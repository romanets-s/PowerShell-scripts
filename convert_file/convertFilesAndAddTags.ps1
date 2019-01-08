
Param(
    [string]$fullPathToInput,
    [string]$fullPathToOutput,
    [system.collections.generic.dictionary[string,string]]$keyValueCollectionName,
    [system.collections.generic.dictionary[string,system.collections.generic.dictionary[string,string]]]$config
)

#Load custom library
Add-Type -Path $config['settings']['pathToLibaryForPDF']

#init
$extentionAll = $config['settings']['fileExtentionForConvert']
$resultDic = New-Object 'system.collections.generic.dictionary[string,string]'
$renamed = 'Renamed'
$processed = 'processed'
$notValidFiles = New-Object 'system.collections.generic.dictionary[string,string]'

#create tmp folders
$failedFolder = [IO.Path]::Combine($fullPathToOutput, 'Failed')
if(!(Test-Path -Path $failedFolder)){
    New-Item -ItemType directory -path $failedFolder | Out-Null
}
$pathToProcessedFiles = [IO.Path]::Combine($fullPathToInput, $processed)
if(!(Test-Path -Path $pathToProcessedFiles)){
    New-Item -ItemType directory -path $pathToProcessedFiles | Out-Null
}

#Add keywords (tags) to file
Function add-keywords($appFile, $keyWords){
    $binding = "System.Reflection.BindingFlags" -as [type]
    $properties = $appFile.BuiltInDocumentProperties
    foreach($property in $properties){
        $pn = [System.__ComObject].invokemember("name", $binding::GetProperty, $null, $property, $null)
        if($pn -eq "Keywords"){
            [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $keyWords)
        }
    }
}

#Convert .xls or .xlsm to .xlsx
Function convert-xlsx($fileName, $fileNameNew, $keyWords){
    $excel = New-Object -ComObject excel.application
    $excel.visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.workbooks.open($fileName, 0, 0, 5, '')
    if($keyWords -ne $renamed){
        add-keywords $wb $keyWords
    }
    $wb.saveAs($fileNameNew, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault)
    $wb.close()
    $excel.Workbooks.Close()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers | Out-Null
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable -Name excel

    # Kill all excel processes
    $currentSessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    $processes = Get-Process -Name excel -ErrorAction Ignore
    foreach($process in $processes) {
        if ($process.SessionId -eq $currentSessionId) {
            Stop-Process -InputObject $process -Force -ErrorAction Ignore
            $timeouted = $null
            Wait-Process -InputObject $process -Timeout 5 -ea 0 -ev timeouted
            if($timeouted){
                $process.Kill() | Out-Null
            }
        }
    }
}

# Convert .doc to .docx
Function convert-docx($fileName, $fileNameNew, $keyWords){
    $word = new-object -comobject word.application
    $word.Visible = $False
    $doc = $word.documents.open($fileName, $null, $False, $null, '', '')
    if($keyWords -ne $renamed){
        add-keywords $doc $keyWords
    }
    $doc.saveas($fileNameNew, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)
    $doc.close();
    [gc]::collect()
    [gc]::WaitForPendingFinalizers() | Out-Null
    $word.quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    Remove-Variable word
    
    # Kill all word processes
    $currentSessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    $processes = Get-Process -Name winword -ErrorAction Ignore
    foreach($process in $processes) {
        if ($process.SessionId -eq $currentSessionId) {
            Stop-Process -InputObject $process -Force -ErrorAction Ignore
            $timeouted = $null
            Wait-Process -InputObject $process -Timeout 5 -ea 0 -ev timeouted
            if($timeouted){
                $process.Kill() | Out-Null
            }
        }
    }
}

# Convert .ppt to .pptx
Function convert-pptx($fileName, $fileNameNew, $keyWords){
    $powerPoint = New-Object -ComObject PowerPoint.application
    $presentation = $powerPoint.Presentations.open($fileName, $True, $True, $False)
    if($keyWords -ne $renamed){
        add-keywords $presentation $keyWords
    }
    $presentation.SaveAs($fileNameNew, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation)
    $presentation.Close()
    [gc]::collect()
    [gc]::WaitForPendingFinalizers() | Out-Null
    $powerPoint.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
    Remove-Variable -Name powerPoint

    #Kill all powerPoint processes
    $currentSessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    $processes = Get-Process -Name powerpnt -ErrorAction Ignore
    foreach($process in $processes) {
        if ($process.SessionId -eq $currentSessionId) {
            Stop-Process -InputObject $process -Force -ErrorAction Ignore
            $timeouted = $null
            Wait-Process -InputObject $process -Timeout 5 -ea 0 -ev timeouted
            if($timeouted){
                $process.Kill() | Out-Null
            }
        }
    }
}

#Convert .vsd to .vsdx
Function convert-vsdx($fileName, $fileNameNew, $keyWords){
    $visio = New-Object -ComObject Visio.Application
    $visio.visible = $false
    $doc = $visio.Documents.Open($fileName)
    if($keyWords -ne $renamed){
        add-keywords $doc $keyWords
    }
    $doc.ExportAsFixedFormat([Microsoft.Office.Interop.Visio.VisFixedFormatType]::xlTypePDF, $fileNameNew)
    $doc.Close()
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers() | Out-Null
    $visio.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($visio) | Out-Null
    Remove-Variable visio

    #Kill all visio processes
    $currentSessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    $processes = Get-Process -Name visio -ErrorAction Ignore
    foreach($process in $processes) {
        if ($process.SessionId -eq $currentSessionId) {
            Stop-Process -InputObject $process -Force -ErrorAction Ignore
            $timeouted = $null
            Wait-Process -InputObject $process -Timeout 5 -ea 0 -ev timeouted
            if($timeouted){
                $process.Kill() | Out-Null
            }
        }
    }
}

#Rename .pdf and add keywords
Function convert-pdf($fileName, $fileNameNew, $keyWords){
    $fileOld = New-Object -TypeName iTextSharp.text.pdf.PdfReader -ArgumentList ($fileName)
    $fileNew = New-Object -TypeName System.IO.FileStream -ArgumentList ($fileNameNew,[System.IO.FileMode]::OpenOrCreate)
    $reader = New-Object -TypeName iTextSharp.text.Document
    $pdfCopy = New-Object -TypeName iTextSharp.text.pdf.PdfCopy -ArgumentList ($reader, $fileNew)
    $reader.Open()
    
    #copy file and meta data
    $pdfCopy.AddDocument($fileOld)
    $reader.AddAuthor($fileOld.Info['Author']) | Out-Null
    $reader.AddCreator($fileOld.Info['Creator']) | Out-Null
    if($keyWords -ne $renamed){
        $reader.AddKeywords($keyWords) | Out-Null
    }
    $reader.AddSubject($fileOld.Info['Subject']) | Out-Null
    $reader.AddTitle($fileOld.Info['Title']) | Out-Null

    $reader.Close()

    $fileNew.Close()
    $fileOld.Dispose()
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers() | Out-Null
    Remove-Variable reader
}

#Convert image (.jpg, .jpeg, .png, .gif, .tif) to .pdf and add keywords
Function convert-img($fileName, $fileNameNew, $keyWords){
    $getImgSize = [System.Drawing.Bitmap]::FromFile($fileName)
    $setPageSize = [iTextSharp.text.Rectangle]::new(0, 0, $getImgSize.Width, $getImgSize.Height)
    $getImgSize.Dispose()
    $newFile = [iTextSharp.text.Document]::new($setPageSize, 0, 0, 0, 0)
    $ms = [System.IO.MemoryStream]::new()
    [iTextSharp.text.pdf.PdfWriter]::GetInstance($newFile, $ms).SetFullCompression()
    $newFile.Open()
    $image = [iTextSharp.text.Image]::GetInstance($fileName)
    $newFile.Add($image) | Out-Null
    $newFile.Close()
    $fileName = ([IO.Path]::Combine([IO.Path]::GetDirectoryName($fileName), ([IO.Path]::GetFileNameWithoutExtension($fileName) + ".pdf")))
    [System.IO.File]::WriteAllBytes($fileName, $ms.ToArray())
    convert-pdf $fileName $fileNameNew $keyWords
}

# Rename .msg, convert and add keywords
Function convert-msg($fileName, $fileNameNew, $keyWords, $pathToProcessedFiles){
    $currentSessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId

    #Kill all outlook processes
    $processes = Get-Process -Name outlook -ErrorAction Ignore
    foreach($process in $processes) {
        if ($process.SessionId -eq $currentSessionId) {
            Stop-Process -InputObject $process -Force -ErrorAction Ignore
            $timeouted = $null
            Wait-Process -InputObject $process -Timeout 5 -ea 0 -ev timeouted
            if($timeouted){
                $process.Kill() | Out-Null
            }
        }
    }        

    $pathToTempDoc = ([IO.Path]::Combine([IO.Path]::GetDirectoryName($fileName), ([IO.Path]::GetFileNameWithoutExtension($fileName) + ".doc")))     # Path to temp word document
    $pathToTempPdf = ([IO.Path]::Combine([IO.Path]::GetDirectoryName($fileName), ([IO.Path]::GetFileNameWithoutExtension($fileName) + ".pdf"))) # Path to temp pdf document
    $outlook = New-Object -ComObject Outlook.Application
    $msg = $outlook.CreateItemFromTemplate("$fileName")                                                  # Open mail message
    $msg.SaveAs($pathToTempDoc, 4);                                                                       # ConvertMailMessage to doc

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null                      # Cill outlook
    $processes = Get-Process -Name outlook -ErrorAction Ignore
    foreach($process in $processes) {
        if ($process.SessionId -eq $currentSessionId) {
            Stop-Process -InputObject $process -Force -ErrorAction Ignore
            $timeouted = $null
            Wait-Process -InputObject $process -Timeout 5 -ea 0 -ev timeouted
            if($timeouted){
                $process.Kill() | Out-Null
            }
        }
    }

    $word_app = New-Object -ComObject Word.Application                                                   # Open word app
    $word_app.Options.WarnBeforeSavingPrintingSendingMarkup = $false
    $doc = $word_app.Documents.Open($pathToTempDoc)                                                      # Open temp document
    $doc.ExportAsFixedFormat($pathToTempPdf,17,$false,0,3,1,1,0,$false, $false,0,$false, $true)          # Convert doc to pdf
    
    $doc.Close()                                                                                         # Close word
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word_app) | Out-Null
    $processes = Get-Process -Name winword -ErrorAction Ignore
    foreach($process in $processes) {
        if ($process.SessionId -eq $currentSessionId) {
            Stop-Process -InputObject $process -Force -ErrorAction Ignore
            $timeouted = $null
            Wait-Process -InputObject $process -Timeout 5 -ea 0 -ev timeouted
            if($timeouted){
                $process.Kill() | Out-Null
            }
        }
    }

    Remove-Item –path $pathToTempDoc | Out-Null                                                          # Remove temp word document

    convert-pdf $pathToTempPdf $fileNameNew $keyWords | Out-Null

    Remove-Item –path $pathToTempPdf | Out-Null                                                          # Remove temp pdf document
}

# Convert (.ics) to .pdf and add keywords
Function convert-ics($fileName, $fileNameNew, $keyWords){

    # Get file content
    $data = @{}
    $content = Get-Content $fileName -Encoding UTF8
    $previousElementKey = ""
    $content |
        foreach-Object {
            if ( -not $_.ToString().StartsWith("	")) {    # If row is not sub row
                if($_.Contains(':')){
                    $z=@{ $_.split( ':')[0] =( $_.split( ':')[1])}
                    $data. Add( $z. Keys, $z. Values)
                    $previousElementKey = $z. Keys
                }
            } else {
                if($previousElementKey -ne "") {             # If main row was found
                   $newValue = ($data.getEnumerator() | ?{ $_.Name -eq $previousElementKey}).Value += $_
                   $data.Remove($previousElementKey)
                   $data.Add($previousElementKey, $newValue) # Add sub row
                }
            }
        }

     #Check if file is valid
     $isFileWalid = $false
     foreach ($result in $data.Keys.Contains("DESCRIPTION")) {
        if($result) {
            $isFileWalid = $true
        }
     }
     if(-not $isFileWalid) {                                 # If description was not found
        throw New-Object System.Exception "Not valid file format"
     }

     # Get body
     $splitOption = [System.StringSplitOptions]::RemoveEmptyEntries
     $Body = ($data.getEnumerator() | ?{ $_.Name -eq "DESCRIPTION"}).Value
     if(-not ([string]::IsNullOrEmpty($Body))) {             # If body was found 
        $Body = $Body.replace("\,",",").replace("\\", "\")
     }

     # Get start
     $Start = ($data.getEnumerator() | ?{ $_.Name.Contains("DTSTART;")}).Value -replace "T"
     if(-not ([string]::IsNullOrEmpty($Start))) {            # If Start was found 
        Try{
            $Start = [datetime]::ParseExact($Start ,"yyyyMMddHHmmss" ,$null )
        } Catch{
            $Start = $Start.Substring(4, 2) + "/" + $Start.Substring(6, 2) + "/" + $Start.Substring(0, 4)
        }
     }

     #Get end
     $End = ($data.getEnumerator() | ?{ $_.Name.Contains("DTEND;")}).Value -replace "T"
     if(-not ([string]::IsNullOrEmpty($End))) {              # If End was found
        Try{
            $End = [datetime]::ParseExact($End ,"yyyyMMddHHmmss" ,$null )
        } Catch{
            $End = $End.Substring(4, 2) + "/" + $End.Substring(6, 2) + "/" + $End.Substring(0, 4)
        } 
     }

     $Subject = ($data.getEnumerator() | ?{ $_.Name.Contains("SUMMARY")}).Value
     $Location = ($data.getEnumerator() | ?{ $_.Name -eq "LOCATION"}).Value
     
     #Populate PDF file
     $newFile = [iTextSharp.text.Document]::new()
     $ms = [System.IO.MemoryStream]::new()
     [iTextSharp.text.pdf.PdfWriter]::GetInstance($newFile, $ms).SetFullCompression()
     $newFile.Open()
     $line = New-Object itextsharp.text.Paragraph("Subject: " + $Subject)
     $newFile.Add($line) | Out-Null
     $line = New-Object itextsharp.text.Paragraph("Location: " + $Location)
     $newFile.Add($line) | Out-Null
     $line = New-Object itextsharp.text.Paragraph("Start time: " + $Start)
     $newFile.Add($line) | Out-Null
     $line = New-Object itextsharp.text.Paragraph("End time: " + $End)
     $newFile.Add($line) | Out-Null
     
     # Add description
     $line = New-Object itextsharp.text.Paragraph("Description:")
     $newFile.Add($line) | Out-Null
     if(-not ([string]::IsNullOrEmpty($Body))) {             # If body was found 
        foreach ($line in $Body.split(@("\n"), [System.StringSplitOptions]::RemoveEmptyEntries)) {
            $line = New-Object itextsharp.text.Paragraph("             " + $line.Trim())
            $newFile.Add($line) | Out-Null
        }
     }

     #Save file
     $newFile.Close()
     $fileName = ([IO.Path]::Combine([IO.Path]::GetDirectoryName($fileName), ([IO.Path]::GetFileNameWithoutExtension($fileName) + ".pdf")))
     [System.IO.File]::WriteAllBytes($fileName, $ms.ToArray())

     #Add keywords
     convert-pdf $fileName $fileNameNew $keyWords
}

#get list of input files
$inputFilesOrigin = @(Get-ChildItem -Path $fullPathToInput -Recurse | where {!$_.PSIsContainer} | %{$_.FullName})
$inputFiles = @(Get-ChildItem -Path $fullPathToInput -Recurse | where {!$_.PSIsContainer} | %{$_.BaseName})
$inputFiles = @($inputFiles -replace ' ', '' | %{$_.ToUpper()})

foreach ($key in $keyValueCollectionName.Keys){
    if ($inputFiles.Contains($key)){
        $tmpFileName = [IO.Path]::GetFileNameWithoutExtension($inputFilesOrigin[[array]::indexof($inputFiles, $key)])
        foreach ($fileName in @(Get-ChildItem -Path $fullPathToInput -Recurse | where {!$_.PSIsContainer -and $_.BaseName -eq $tmpFileName} | %{$_.FullName})){
            $extention = ([IO.Path]::GetExtension($fileName)).ToLower()
            if($extentionAll.Contains($extention)){ # If file nead convert to +++x version (xlsx, docx, ...)
                $extention = $extention.Substring(0, 4) + 'x'
            }

            #get new file name and keywords from dictionary value
            $indexFirst = $keyValueCollectionName.$key.IndexOf('|')
            $indexLast = $keyValueCollectionName.$key.LastIndexOf('|')

            $fileNameTmp = $keyValueCollectionName.$key.Substring(0, $indexFirst)
            $keyWords = $keyValueCollectionName.$key.Substring($indexFirst + 1, $indexLast - $indexFirst - 1)
            $componet = $keyValueCollectionName.$key.Substring($indexLast + 1)
            
            if($componet.Contains('&')){
                $componet = $componet.Split('&', [System.StringSplitOptions]::RemoveEmptyEntries).Trim()
            }
            else{
                $componet = @($componet)
            }
            # if not componet value - add empty item
            if(!$componet){
                $componet = @(' ')
            }
            if($keyWords -eq $config['settings']['notValidValue']){
                $keyWords = ''
            }
            if([string]::IsNullOrEmpty($keyWords)){
                    $keyWords = $renamed
            }
                
            if(@('.msg', '.tif', '.gif', '.jpeg', '.png', '.ics', '.jpg').Contains($extention)){ # Files that must be converted to pdf format
                $tmpName = $fileNameTmp + ".pdf"
            }
            else{
                $tmpName = $fileNameTmp + $extention
            }
            Move-Item -Path $fileName -Destination ([IO.Path]::Combine($pathToProcessedFiles, [IO.Path]::GetFileName($fileName))) -Force
            $fileName = ([IO.Path]::Combine($pathToProcessedFiles, [IO.Path]::GetFileName($fileName)))
            foreach($mainFolder in $componet){
                $pathToOutput = [IO.Path]::Combine($fullPathToOutput, $mainFolder.Trim())
                $fileNameNew = [IO.Path]::Combine($pathToOutput, $tmpName)
                if(![System.IO.File]::Exists($fileNameNew)){
                    if(!(Test-Path -Path $pathToOutput)){
                        New-Item -ItemType directory -path $pathToOutput | Out-Null
                    }
                    Try{
                        if($extention.Contains('doc')){
                            convert-docx $fileName $fileNameNew $keyWords | Out-Null
                        }
                        elseIf($extention.Contains('ppt')){
                            convert-pptx $fileName $fileNameNew $keyWords | Out-Null
                        }
                        elseIf($extention.Contains('xls')){
                            convert-xlsx $fileName $fileNameNew $keyWords | Out-Null
                        }
                        elseIf($extention.Contains('pdf')){
                            convert-pdf $fileName $fileNameNew $keyWords | Out-Null
                        }
                        elseIf($extention.Contains('vsb')){
                            convert-vsdx $fileName $fileNameNew $keyWords | Out-Null
                        }
                        elseIf($extention.Contains('msg')){
                            convert-msg $fileName $fileNameNew $keyWords $pathToProcessedFiles | Out-Null
                        }
                        elseIf(@('.tif', '.gif', '.jpeg', '.png','.jpg').Contains($extention)){
                            convert-img $fileName $fileNameNew $keyWords | Out-Null
                        }
                        elseIf($extention.Contains('ics')){
                            convert-ics $fileName $fileNameNew $keyWords | Out-Null
                        }
                        else{
                            Copy-Item -Path $fileName -Destination $fileNameNew | Out-Null
                        }
                    }
                    Catch{
                        $proc = Get-Process -Name excel, winword, powerpnt, visio, outlook -ErrorAction Ignore
                        if($proc.count -ne 0){
                            $currentSessionId = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
                            foreach($process in $proc) {
                                if ($process.SessionId -eq $currentSessionId) {
                                    Stop-Process -InputObject $process -Force -ErrorAction Ignore
                                    $timeouted = $null
                                    Wait-Process -InputObject $process -Timeout 5 -ea 0 -ev timeouted
                                    if($timeouted){
                                        $process.Kill() | Out-Null
                                    }
                                }
                            }
                        }
                        [gc]::Collect()
                        [gc]::WaitForPendingFinalizers() | Out-Null
                        if([System.IO.File]::Exists($fileNameNew)){
                            Remove-Item -Path $fileNameNew -ErrorAction Ignore | Out-Null
                        }
                        Copy-Item -Path $fileName -Destination ([IO.Path]::Combine($failedFolder, [IO.Path]::GetFileName($fileName))) -Force | Out-Null
                        if (-not $notValidFiles.ContainsKey([IO.Path]::GetFileName($fileName))) {
                            $notValidFiles.Add([IO.Path]::GetFileName($fileName), $_.Exception.Message)
                            $resultDic.Add([IO.Path]::GetFileName($fileName), $config['unprocessed']['notPossibleRenameOrChangeExtension'])
                        }
                    }
                }
                else{
                    Copy-Item -Path $fileName -Destination ([IO.Path]::Combine($failedFolder, [IO.Path]::GetFileName($fileName))) -Force | Out-Null
                    if (-not $resultDic.ContainsKey([IO.Path]::GetFileName($fileName))) {
                        $resultDic.Add([IO.Path]::GetFileName($fileName), $config['unprocessed']['renamedAlreadyExists'])
                    }
                }
            }
            #Remove-Item -Path $fileName | Out-Null
        }
    }
    else{
        $resultDic.Add($key, $config['unprocessed']['fileNotAvailableInFolder'])
    }
}
foreach ($fileName in $inputFilesOrigin){
    if([System.IO.File]::Exists($fileName)){
        Move-Item -Path $fileName -Destination ([IO.Path]::Combine($failedFolder, [IO.Path]::GetFileName($fileName))) -Force | Out-Null
        if (-not $resultDic.ContainsKey([IO.Path]::GetFileName($fileName))) {
            $resultDic.Add([IO.Path]::GetFileName($fileName), $config['unprocessed']['notPresentInMapping'])
        }
    } 
}

return $resultDic, $notValidFiles


