Param(
    [string]$pathToZipFile,      # Full path to archive
    [string]$pathToOutFolder,    # Where will be unpacked (Full path)
    [string]$pathToDotNetZipDLL, # Full path to zip library
    [string]$password,           # Password for archive
    [string]$fileName            # Name of the file that must be extracted
)
[System.Reflection.Assembly]::UnsafeLoadFrom($pathToDotNetZipDLL) | Out-Null

$isPasswordValid = [Ionic.Zip.ZipFile]::CheckZipPassword($pathToZipFile, $password)
if ($isPasswordValid){
    try{
        Write-Host $pathToZipFile
        $zipFile = [Ionic.Zip.ZipFile]::Read($pathToZipFile)
        $zipFile.Password = $password
        foreach ($ZipEntry in $zipFile.Entries) {
            if($ZipEntry.FileName.ToLower().Contains($fileName.ToLower())) {
                $ZipEntry.Extract($pathToOutFolder, [Ionic.Zip.ExtractExistingFileAction]::OverwriteSilently)
                $zipFile.Dispose()
                return $true
            }
        }
    }catch{
                                 # do nothing
    }Finally{
        $zipFile.Dispose()
    }
}
return $false