Param(
    [string]$pathToZipFile,    # Full path to archive file
    [string]$pathToUnZipFiles, #Full path to output folder
    [system.collections.generic.dictionary[string,system.collections.generic.dictionary[string,string]]]$config
)
[System.Reflection.Assembly]::UnsafeLoadFrom([IO.Path]::Combine([IO.Directory]::GetCurrentDirectory(), $config["settings"]["pathToDotNetZipDLL"])) | Out-Null

$isPasswordValid = [Ionic.Zip.ZipFile]::CheckZipPassword($pathToZipFile, $config["runtime"]["zipPassword"])
if ($isPasswordValid){
    try{
        $zipFile = [Ionic.Zip.ZipFile]::Read($pathToZipFile)
        $zipFile.Password = $config["runtime"]["zipPassword"]
        $zipFile.ExtractAll($pathToUnZipFiles, [Ionic.Zip.ExtractExistingFileAction]::OverwriteSilently)
        $zipFile.Dispose()
        return $true
    }catch{
                               # do nothing
    }Finally{
        $zipFile.Dispose()
    }
}
return $false